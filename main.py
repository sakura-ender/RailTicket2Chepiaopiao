import imaplib
import email
from email.header import decode_header
import os
import time
from bs4 import BeautifulSoup
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import threading
import re
import subprocess
import platform
from email import policy
from email.parser import BytesParser
import ssl
# Outlook 所需的库
import sys
import requests
from azure.identity import InteractiveBrowserCredential

# ===== 以下是 Outlook 登录所需的配置信息 =====
CLIENT_ID = "1a7a3bf9-b338-4b17-b10e-10149e48df97"  # 必填，来自你的应用注册
TENANT_ID = "common"  # 可选，若想固定在本租户，也可改为具体租户 ID
SCOPES = ["https://graph.microsoft.com/Mail.Read"]  # 需要申请的权限范围

is_paused = False
matching_emails = []
is_processing = False

def log_message(message):
    """在界面和控制台同时输出日志信息。"""
    global log_text, email_count_label
    print(message)
    if log_text:
        log_text.insert(tk.END, message + '\n')
        log_text.see(tk.END)
    if email_count_label:
        email_count_label.config(text=f"当前匹配的邮件数: {len(matching_emails)}")

# ========== 以下是 IMAP 登录所需的部分 ==========
def login_email(username, password, imap_server):
    """使用 IMAP4_SSL 方式登录邮箱，返回 imap 对象。"""
    try:
        imaplib.Commands['ID'] = ('AUTH')

        if imap_server == 'imap.139.com':
            ctx = ssl.create_default_context()
            ctx.set_ciphers('DEFAULT')
            imap = imaplib.IMAP4_SSL(imap_server, ssl_context=ctx)
            imap.login(username, password)
            print(f"Logged in as {username}")
            return imap
        imap = imaplib.IMAP4_SSL(imap_server)
        imap.login(username, password)
        # 163邮箱需要发送 ID 命令来标识客户端信息
        if imap_server == 'imap.163.com':
            args = ("name", "myclient", "contact", username, "version", "1.0.0", "vendor", "myclient")
            typ, dat = imap._simple_command('ID', '("' + '" "'.join(args) + '")')
            imap._untagged_response(typ, dat, 'ID')

        return imap

    except imaplib.IMAP4.error as e:
        log_message(f"Login failed: {e}")
        messagebox.showerror("登录失败", f"登录失败: {e}")
        return None
    except Exception as e:
        log_message(f"Connection error: {e}")
        messagebox.showerror("连接错误", f"连接到IMAP服务器时发生错误: {e}")
        return None

def decode_mime_words(s):
    """解码邮件主题中的 MIME 词。"""
    decoded_fragments = decode_header(s)
    return ''.join([
        fragment.decode(encoding if encoding else 'utf-8') if isinstance(fragment, bytes) else fragment
        for fragment, encoding in decoded_fragments
    ])

def decode_payload(payload):
    """尝试用多种常见编码解码邮件正文。"""
    for charset in ['utf-8', 'gbk', 'gb2312', 'iso-8859-1']:
        try:
            return payload.decode(charset.strip()).strip()
        except UnicodeDecodeError:
            continue
    return None

def extract_body_from_msg(msg):
    """从邮件中提取纯文本正文内容。"""
    if msg.is_multipart():
        for part in msg.walk():
            content_type = part.get_content_type()
            content_disposition = str(part.get("Content-Disposition"))
            if "attachment" not in content_disposition and content_type.startswith("text/"):
                payload = part.get_payload(decode=True)
                if payload:
                    body = decode_payload(payload)
                    if body and '<' in body and '>' in body:
                        text = BeautifulSoup(body, 'html.parser').get_text()
                        text = re.sub(r'\s+', ' ', text).strip()
                        return text
    else:
        payload = msg.get_payload(decode=True)
        if payload:
            body = decode_payload(payload)
            if body and '<' in body and '>' in body:
                text = BeautifulSoup(body, 'html.parser').get_text()
                text = re.sub(r'\s+', ' ', text).strip()
                return text
    return None

def process_email(email_id, imap, keyword, retries=3, delay=5):
    """根据 email_id 获取邮件并解析是否包含关键词。"""
    global is_paused
    for attempt in range(1, retries + 1):
        if is_paused:
            return None
        try:
            status, msg_data = imap.fetch(email_id, '(RFC822)')
            if status == 'OK':
                break
            else:
                time.sleep(delay)
        except imaplib.IMAP4.abort:
            time.sleep(delay)
        except Exception:
            time.sleep(delay)
    else:
        return None

    if not msg_data:
        return None

    for response_part in msg_data:
        if isinstance(response_part, tuple):
            try:
                msg = email.message_from_bytes(response_part[1])
            except Exception:
                continue

            body = extract_body_from_msg(msg)
            if body and keyword in body:
                return msg
    return None

def fetch_emails_with_keyword_in_body(imap, keyword):
    """从 IMAP 收件箱搜索所有包含 keyword 的邮件，返回匹配列表。"""
    global matching_emails
    try:
        imap.select("inbox")
    except Exception as e:
        log_message(f"无法选择收件箱: {e}")
        return []

    try:
        status, messages = imap.search(None, 'ALL')
    except Exception as e:
        log_message(f"搜索邮件时出错: {e}")
        return []

    if status != "OK":
        log_message("获取邮件失败")
        return []

    email_ids = messages[0].split()
    log_message(f"收件箱中共有 {len(email_ids)} 封邮件。")
    matching_emails = []

    for email_id in email_ids:
        if is_paused:
            break
        result = process_email(email_id, imap, keyword)
        if result:
            matching_emails.append(result)
            log_message(f"找到匹配的邮件(ID: {email_id.decode()}), 包含关键词 '{keyword}'。")
        else:
            log_message(f"未找到匹配的邮件(ID: {email_id.decode()})。")
    log_message(f"总共找到 {len(matching_emails)} 封正文中包含 '{keyword}' 的邮件。")
    return matching_emails

# ========== 以下是保存邮件的逻辑 ==========
def save_all_emails_to_single_txt(emails, save_path):
    """将传入的邮件列表统一保存到一个 txt 文件中。"""
    try:
        with open(save_path, 'w', encoding='utf-8') as f:
            for i, msg in enumerate(emails, 1):
                subject = msg.get("Subject", "")
                decoded_subject = decode_mime_words(subject)
                f.write(f"\n###\n")
                body = extract_body_from_msg(msg)
                if body:
                    f.write(f"正文:\n{body}\n")
                else:
                    f.write("正文: 无法解码或为空。\n")
                f.write("\n###\n")
        log_message(f"所有邮件已导出到 {save_path}")
        messagebox.showinfo("完成", f"所有邮件已导出到 {save_path}")
        folder_path = os.path.dirname(save_path)
        if platform.system() == "Windows":
            os.startfile(folder_path)
        elif platform.system() == "Darwin":  # macOS
            subprocess.call(["open", folder_path])
        else:  # Linux
            subprocess.call(["xdg-open", folder_path])
    except Exception as e:
        log_message(f"Error saving emails: {e}")
        messagebox.showerror("保存错误", f"保存邮件到文件时出错: {e}")

# ========== 以下是 Outlook（Graph API） 获取邮件的逻辑 ==========
def fetch_outlook_emails_with_keyword(keyword="12306"):
    """
    1. 使用 InteractiveBrowserCredential 登录 Microsoft 帐号并获取 token。
    2. 调用 Graph API 获取邮件并筛选出包含 keyword 的邮件。
    3. 返回这些邮件（格式尽量与 IMAP 返回相似，方便后续统一保存）。
    """
    global matching_emails
    matching_emails = []  # 重置

    try:
        log_message("正在使用 Outlook 模式登录，请稍候...")
        credential = InteractiveBrowserCredential(
            client_id=CLIENT_ID,
            tenant_id=TENANT_ID,
            redirect_uri="http://localhost:8400"  # 这里与 Azure 注册的重定向 URI 保持一致
        )
        token = credential.get_token(*SCOPES).token
        log_message("获取 Outlook 令牌成功，正在获取邮件...")

        # 调用 Microsoft Graph API 获取邮件
        headers = {
            'Authorization': f'Bearer {token}',
            'Content-Type': 'application/json'
        }
        # 这里拉取前 50 封，扩展可加 $top 参数，或分页获取
        url = "https://graph.microsoft.com/v1.0/me/messages?$select=id,subject,body,from&$top=50"
        response = requests.get(url, headers=headers)
        data = response.json()

        messages = data.get("value", [])
        log_message(f"从 Outlook 获取到 {len(messages)} 封邮件，开始筛选包含 '{keyword}' 的邮件...")

        for msg_item in messages:
            subject = msg_item.get("subject", "")
            body_content = msg_item.get("body", {}).get("content", "")
            sender = msg_item.get("from", {}).get("emailAddress", {}).get("address", "")
            # 如果正文中包含关键词，则将该条“伪装”为 email.message.Message 对象
            # 以便和 IMAP 返回的 msg 结构一致
            if keyword in body_content:
                # 构造一个最简单的 email.message.Message 用来保存最终导出
                msg = email.message.EmailMessage()
                msg["Subject"] = subject
                msg["From"] = sender
                # 将 body_content 转成 HTML 或纯文本再解析，这里直接当作纯文本
                msg.set_content(body_content)
                matching_emails.append(msg)

        log_message(f"总共找到 {len(matching_emails)} 封正文中包含 '{keyword}' 的邮件(Outlook)。")
        return matching_emails

    except Exception as ex:
        log_message(f"获取 Outlook 邮件时出错: {ex}")
        return []

# ========== 以下是处理 EML 文件的逻辑 ==========
def process_eml_files(folder_path):
    """遍历文件夹中所有 .eml 文件，并将其解析为邮件对象后导出。"""
    eml_files = [os.path.join(folder_path, f) for f in os.listdir(folder_path) if f.endswith('.eml')]
    if not eml_files:
        messagebox.showwarning("无文件", "所选文件夹中没有EML文件！")
        return

    log_message(f"发现 {len(eml_files)} 个EML文件，开始解析...")
    emails = []
    for eml_file in eml_files:
        try:
            with open(eml_file, 'rb') as f:
                msg = BytesParser(policy=policy.default).parse(f)
                emails.append(msg)
                log_message(f"已解析: {os.path.basename(eml_file)}")
        except Exception as e:
            log_message(f"解析失败: {eml_file} - {str(e)}")

    if emails:
        save_path = filedialog.asksaveasfilename(
            defaultextension=".txt",
            filetypes=[("Text files", "*.txt")],
            title="保存解析结果"
        )
        if save_path:
            save_all_emails_to_single_txt(emails, save_path)
    else:
        messagebox.showinfo("结果", "未找到有效邮件内容")

# ========== 主窗口函数 ==========
def main():
    global is_processing, is_paused, matching_emails, email_count_label, log_text

    is_processing = False
    root = tk.Tk()
    root.title("邮箱登录与邮件处理")

    window_width = 550
    window_height = 450
    screen_width = root.winfo_screenwidth()
    screen_height = root.winfo_screenheight()
    position_top = int(screen_height / 2 - window_height / 2)
    position_right = int(screen_width / 2 - window_width / 2)
    root.geometry(f'{window_width}x{window_height}+{position_right}+{position_top}')

    notebook = ttk.Notebook(root)
    notebook.pack(expand=True, fill='both')

    # 两个主要的 Tab 页
    login_frame = ttk.Frame(notebook)
    folder_frame = ttk.Frame(notebook)

    notebook.add(login_frame, text="登录(自动获取)")
    notebook.add(folder_frame, text="选择文件夹(EML)")

    # ==================== 登录界面布局 ====================

    # 邮箱账号、授权码
    tk.Label(login_frame, text="邮箱账号:").grid(row=0, column=0, padx=10, pady=5, sticky='e')
    entry_username = tk.Entry(login_frame)
    entry_username.grid(row=0, column=1, padx=10, pady=5)

    tk.Label(login_frame, text="授权码:").grid(row=1, column=0, padx=10, pady=5, sticky='e')
    entry_password = tk.Entry(login_frame, show="*")
    entry_password.grid(row=1, column=1, padx=10, pady=5)

    # 选择 IMAP 邮箱类型（去掉 Outlook，因为让它成为单独按钮）
    email_provider = tk.StringVar(value="imap.qq.com")
    tk.Radiobutton(login_frame, text="QQ邮箱", variable=email_provider, value="imap.qq.com").grid(row=2, column=0, padx=10, pady=5)
    tk.Radiobutton(login_frame, text="163邮箱", variable=email_provider, value="imap.163.com").grid(row=2, column=1, padx=10, pady=5)
    tk.Radiobutton(login_frame, text="Gmail",  variable=email_provider, value="imap.gmail.com").grid(row=3, column=0, padx=10, pady=5)
    tk.Radiobutton(login_frame, text="自定义IMAP", variable=email_provider, value="custom").grid(row=3, column=1, padx=10, pady=5)

    tk.Label(login_frame, text="自定义IMAP服务器:").grid(row=4, column=0, padx=10, pady=5, sticky='e')
    entry_custom_imap = tk.Entry(login_frame)
    entry_custom_imap.grid(row=4, column=1, padx=10, pady=5)
    entry_custom_imap.config(state=tk.DISABLED)

    def toggle_custom_imap(*args):
        if email_provider.get() == "custom":
            entry_custom_imap.config(state=tk.NORMAL)
        else:
            entry_custom_imap.config(state=tk.DISABLED)

    email_provider.trace("w", toggle_custom_imap)

    # 单独的按钮：通过 Outlook 获取邮件
    def fetch_outlook_button_click():
        """点击按钮后直接触发 Outlook 流程。"""
        global matching_emails
        # 每次点击先清空一下 matching_emails
        matching_emails = []

        emails = fetch_outlook_emails_with_keyword(keyword="12306")
        if emails:
            # 如果获取到符合条件的邮件，则弹出保存对话框
            save_path = filedialog.asksaveasfilename(
                defaultextension=".txt",
                filetypes=[("Text files", "*.txt")]
            )
            if save_path:
                save_all_emails_to_single_txt(emails, save_path)
        else:
            log_message("没有找到包含 '12306' 的 Outlook 邮件。")

    outlook_button = tk.Button(login_frame, text="通过 Outlook 获取邮件", command=fetch_outlook_button_click)
    outlook_button.grid(row=5, column=0, columnspan=2, padx=10, pady=5)

    # 开始处理按钮（IMAP专用）
    def start_imap_processing():
        global is_processing, is_paused, matching_emails

        if is_processing:
            # 如果已经在处理中，则此时点击按钮为“暂停并导出”
            is_paused = True
            if matching_emails:
                save_path = filedialog.asksaveasfilename(defaultextension=".txt",
                                                         filetypes=[("Text files", "*.txt")])
                if save_path:
                    log_message("正在导出当前匹配到的邮件...")
                    save_all_emails_to_single_txt(matching_emails, save_path)
            messagebox.showinfo("暂停", "程序已暂停并导出当前匹配的邮件。")
            root.quit()
        else:
            # 如果尚未开始，则此时点击按钮为“开始IMAP处理”
            def process():
                global is_paused, matching_emails
                is_paused = False
                username = entry_username.get()
                password = entry_password.get()
                keyword = "12306"
                selected_provider = email_provider.get()

                # 若选择自定义IMAP，则获取服务器地址
                if selected_provider == "custom":
                    imap_server = entry_custom_imap.get()
                else:
                    imap_server = selected_provider

                if not imap_server:
                    messagebox.showwarning("输入错误", "请填写自定义 IMAP 服务器地址。")
                    return

                if not username or not password:
                    messagebox.showwarning("输入错误", "请填写邮箱账号和授权码。")
                    return

                log_message(f"正在登录(IMAP) -> {imap_server}")
                imap = login_email(username, password, imap_server)
                if not imap:
                    return

                log_message("正在获取邮件中...")
                emails = fetch_emails_with_keyword_in_body(imap, keyword)
                if len(emails) > 0 and not is_paused:
                    save_path = filedialog.asksaveasfilename(defaultextension=".txt",
                                                             filetypes=[("Text files", "*.txt")])
                    if save_path and not is_paused:
                        log_message("正在导出 IMAP 邮件...")
                        save_all_emails_to_single_txt(emails, save_path)
                else:
                    log_message(f"未找到包含 '{keyword}' 的邮件。")
                    messagebox.showinfo("结果", f"未找到包含 '{keyword}' 的邮件。")

                try:
                    imap.logout()
                    log_message("已登出 IMAP。")
                except Exception as e:
                    log_message(f"登出时出错: {e}")
                    messagebox.showerror("登出错误", f"登出时出错: {e}")

            threading.Thread(target=process).start()
            is_processing = True
            start_button.config(text="中途暂停并导出目前的邮件(不推荐)")

    start_button = tk.Button(login_frame, text="开始IMAP处理", command=start_imap_processing)
    start_button.grid(row=6, column=0, columnspan=2, pady=10)

    # 当前匹配邮件数
    email_count_label = tk.Label(login_frame, text="当前匹配的邮件数: 0")
    email_count_label.grid(row=7, column=0, columnspan=2, pady=5)

    # 日志输出框
    log_text = tk.Text(login_frame, height=10, width=65)
    log_text.grid(row=8, column=0, columnspan=2, padx=10, pady=5)

    # 右侧说明（可自行调整布局）
    desc_text = (
        "使用说明：\n\n"
        "1). 使用 IMAP 获取：\n"
        "   - 在上面输入邮箱账号、授权码\n"
        "   - 勾选对应服务器或自定义\n"
        "   - 点击 [开始IMAP处理]\n\n"
        "2). 使用 Outlook 获取：\n"
        "   - 点击 [通过 Outlook 获取邮件]，会弹出浏览器交互登录\n"
        "   - 成功登录后自动获取邮件\n\n"
        "3). 文件夹 EML 导入：\n"
        "   - 切换到 [选择文件夹(EML)] 标签\n"
        "   - 选择包含 EML 文件的文件夹进行导入"
    )
    usage_label = ttk.Label(login_frame, text=desc_text, justify="left")
    usage_label.grid(row=0, column=2, rowspan=9, padx=10, pady=5, sticky="n")

    # ==================== 文件夹选择界面布局 ====================
    def select_folder():
        folder_path = filedialog.askdirectory()
        if folder_path:
            log_message(f"Selected folder: {folder_path}")
            process_eml_files(folder_path)

    select_button = tk.Button(folder_frame, text="选择文件夹", command=select_folder)
    select_button.pack(pady=20)

    folder_desc = ttk.Label(folder_frame,
                            text="适用于不支持 IMAP 登录或其他特殊场景。\n在邮箱客户端导出 EML 文件后，放到同一文件夹再导入。")
    folder_desc.pack(pady=10)

    root.mainloop()

if __name__ == "__main__":
    main()