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
log_text_eml = None
def log_eml_message(message):
    global log_text_eml,email_count_label
    print(message)
    if log_text_eml:
        log_text_eml.insert(tk.END, message + '\n')
        log_text_eml.see(tk.END)
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
        if imap_server == 'imap.163.com' or imap_server == 'imap.126.com':
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
    2. 调用 Graph API 分页获取所有邮件并筛选出包含 keyword 的邮件。
    3. 返回这些邮件（格式尽量与 IMAP 返回相似，方便后续统一保存）。
    """
    global matching_emails
    matching_emails = []  # 重置

    try:
        log_message("正在使用 Outlook 模式登录，请稍候...")
        credential = InteractiveBrowserCredential(
            client_id=CLIENT_ID,
            tenant_id=TENANT_ID,
            redirect_uri="http://localhost:8400"
        )
        token = credential.get_token(*SCOPES).token
        log_message("获取 Outlook 令牌成功，正在获取邮件...")

        headers = {
            'Authorization': f'Bearer {token}',
            'Content-Type': 'application/json'
        }

        # 分页获取所有邮件
        messages = []
        url = "https://graph.microsoft.com/v1.0/me/messages?$select=id,subject,body,from&$top=1000"  # 单次最大数量
        while url:
            response = requests.get(url, headers=headers)
            if response.status_code != 200:
                log_message(f"获取邮件失败，状态码：{response.status_code}")
                break
            data = response.json()
            current_page_messages = data.get("value", [])
            messages.extend(current_page_messages)
            log_message(f"已获取 {len(current_page_messages)} 封邮件，累计总数：{len(messages)}")
            url = data.get('@odata.nextLink', None)  # 处理分页

        log_message(f"从 Outlook 共获取到 {len(messages)} 封邮件，开始筛选包含 '{keyword}' 的邮件...")

        for msg_item in messages:
            body_content = msg_item.get("body", {}).get("content", "")
            if keyword in body_content:
                subject = msg_item.get("subject", "")
                sender = msg_item.get("from", {}).get("emailAddress", {}).get("address", "")

                msg = email.message.EmailMessage()
                msg["Subject"] = subject
                msg["From"] = sender
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

    log_eml_message(f"发现 {len(eml_files)} 个EML文件，开始解析...")
    emails = []
    for eml_file in eml_files:
        try:
            with open(eml_file, 'rb') as f:
                msg = BytesParser(policy=policy.default).parse(f)
                emails.append(msg)
                log_eml_message(f"已解析: {os.path.basename(eml_file)}")
        except Exception as e:
            log_eml_message(f"解析失败: {eml_file} - {str(e)}")

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

outlook_button_pressed = False
def fetch_outlook_button_click():
    """点击按钮后直接触发 Outlook 流程。"""
    global matching_emails,outlook_button_pressed

    if outlook_button_pressed:
        log_message("请勿重复登录，程序即将退出....")
        root.update()
        time.sleep(1)
        os._exit(0)


    outlook_button_pressed = True
    token = None
    def process():
        emails = fetch_outlook_emails_with_keyword(keyword="12306")
        if emails:
            save_path = filedialog.asksaveasfilename(
                defaultextension=".txt",
                filetypes=[("Text files", "*.txt")]
            )
            if save_path:
                save_all_emails_to_single_txt(emails, save_path)
        else:
            log_message("没有找到包含 '12306' 的 Outlook 邮件。")

    threading.Thread(target=process).start()
def select_folder():
        folder_path = filedialog.askdirectory()
        if folder_path:
            log_eml_message(f"Selected folder: {folder_path}")
            process_eml_files(folder_path)


# ========== 主窗口函数 ==========
# ...（前面的导入和全局变量保持不变）

def main():
    global is_processing, is_paused, matching_emails, email_count_label, log_text, root,log_text_eml

    is_processing = False
    root = tk.Tk()
    root.title("邮箱处理工具 v1.0")

    window_width = 700  # 加宽窗口
    window_height = 500
    screen_width = root.winfo_screenwidth()
    screen_height = root.winfo_screenheight()
    position_top = int(screen_height / 2 - window_height / 2)
    position_right = int(screen_width / 2 - window_width / 2)
    root.geometry(f'{window_width}x{window_height}+{position_right}+{position_top}')

    notebook = ttk.Notebook(root)
    notebook.pack(expand=True, fill='both')

    # 三个Tab页
    login_frame = ttk.Frame(notebook)
    folder_frame = ttk.Frame(notebook)
    about_frame = ttk.Frame(notebook)

    notebook.add(login_frame, text="常规登录获取")
    notebook.add(folder_frame, text="EML导入处理")
    notebook.add(about_frame, text="关于")

    # ==================== 常规登录界面 ====================
    left_panel = ttk.Frame(login_frame)
    left_panel.pack(side=tk.LEFT, padx=20, pady=10, fill=tk.BOTH, expand=True)

    right_panel = ttk.Frame(login_frame)
    right_panel.pack(side=tk.RIGHT, padx=20, pady=10, fill=tk.Y)

    # 左侧登录组件
    tk.Label(left_panel, text="邮箱账号:").grid(row=0, column=0, padx=5, pady=5, sticky='e')
    entry_username = tk.Entry(left_panel, width=25)
    entry_username.grid(row=0, column=1, padx=5, pady=5)

    tk.Label(left_panel, text="授权码:").grid(row=1, column=0, padx=5, pady=5, sticky='e')
    entry_password = tk.Entry(left_panel, show="*", width=25)
    entry_password.grid(row=1, column=1, padx=5, pady=5)

    # 邮箱类型选择
    email_provider = tk.StringVar(value="imap.qq.com")
    providers = [
        ("QQ邮箱", "imap.qq.com"),
        ("163邮箱", "imap.163.com"),
        ("126邮箱", "imap.126.com"),
        ("139邮箱", "imap.139.com"),
        ("Gmail", "imap.gmail.com"),
        ("自定义IMAP", "custom")
    ]

    for i, (text, value) in enumerate(providers):
        rb = ttk.Radiobutton(left_panel, text=text, variable=email_provider, value=value)
        rb.grid(row=2 + i // 2, column=i % 2, padx=5, pady=2, sticky='w')

    tk.Label(left_panel, text="自定义IMAP服务器:").grid(row=8, column=0, padx=5, pady=5, sticky='e')
    entry_custom_imap = tk.Entry(left_panel, width=25)
    entry_custom_imap.grid(row=8, column=1, padx=5, pady=5)
    entry_custom_imap.config(state=tk.DISABLED)
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

    def toggle_custom_imap(*args):
        entry_custom_imap.config(state=tk.NORMAL if email_provider.get() == "custom" else tk.DISABLED)

    email_provider.trace("w", toggle_custom_imap)

    # 右侧Outlook组件
    outlook_btn = ttk.Button(right_panel, text="Outlook登录获取",
                             command=fetch_outlook_button_click,
                             width=20)
    outlook_btn.pack(pady=15)

    ttk.Separator(right_panel, orient='horizontal').pack(fill=tk.X, pady=10)

    start_btn = ttk.Button(right_panel, text="开始IMAP处理",
                           command=start_imap_processing,
                           width=20)
    start_btn.pack(pady=15)

    email_count_label = ttk.Label(right_panel, text="当前匹配的邮件数: 0")
    email_count_label.pack(pady=10)

    # 日志区域
    log_text = tk.Text(left_panel, height=12, width=50)
    log_text.grid(row=10, column=0, columnspan=2, pady=10)

    # ==================== EML导入界面 ====================
    eml_instructions = ttk.Label(folder_frame,
                                 text="请选择包含EML文件的文件夹：\n\n1. 在邮箱客户端中导出邮件为EML格式\n2. 将所有文件放入同一文件夹\n3. 点击下方按钮选择该文件夹",
                                 justify="center")
    eml_instructions.pack(pady=30)

    select_btn = ttk.Button(folder_frame, text="选择文件夹", command=select_folder)
    select_btn.pack(pady=10)
    log_text_eml = tk.Text(folder_frame, height=12, width=50)
    log_text_eml.pack(pady=10)

    # ==================== 关于界面 ====================
    about_text = """
    邮箱处理工具 v1.0

    功能说明：
    - 支持主流邮箱的IMAP登录（QQ/163/126/139/Gmail）
    - 支持Outlook账号登录（Microsoft Graph API）
    - 支持批量处理EML文件
    - 自动提取包含指定关键词的邮件内容

    技术特性：
    - 使用Python 3开发
    - 基于Tkinter的GUI界面
    - 多线程处理防止界面卡顿
    - 支持跨平台运行

    开发者：Your Name
    发布日期：2023-10-01
    License：MIT
    """
    about_label = ttk.Label(about_frame, text=about_text, justify=tk.CENTER)
    about_label.pack(expand=True, padx=20, pady=20)

    root.mainloop()


if __name__ == "__main__":
    main()
