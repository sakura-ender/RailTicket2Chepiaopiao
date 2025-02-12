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
import requests
from azure.identity import InteractiveBrowserCredential

# Outlook 登录配置
CLIENT_ID = "1a7a3bf9-b338-4b17-b10e-10149e48df97"
TENANT_ID = "common"
SCOPES = ["https://graph.microsoft.com/Mail.Read"]

# 全局控制变量
is_paused = False
matching_emails = []
is_processing = False
outlook_button_pressed = False

# GUI全局变量
root = None
log_text = None
log_text_eml = None
email_count_label = None
entry_username = None
entry_password = None
email_provider = None
entry_custom_imap = None
progress_bar = None
start_btn = None
cancel_btn = None


# ---------------------- 辅助的线程安全日志方法 ---------------------- #
def safe_log_message(message):
    print(message)
    if root is not None and log_text is not None:
        root.after(0, lambda: update_log(message))


def update_log(message):
    log_text.insert(tk.END, message + '\n')
    log_text.see(tk.END)
    email_count_label.config(text=f"当前匹配的邮件数: {len(matching_emails)}")


def safe_log_eml_message(message):
    print(message)
    if root is not None and log_text_eml is not None:
        root.after(0, lambda: update_log_eml(message))


def update_log_eml(message):
    log_text_eml.insert(tk.END, message + '\n')
    log_text_eml.see(tk.END)


def update_progress(value):
    if root is not None and progress_bar is not None:
        root.after(0, lambda: progress_bar.config(value=value))


def login_email(username, password, imap_server):
    try:
        imaplib.Commands['ID'] = ('AUTH')  # 针对 163/126 的特殊操作
        if imap_server == 'imap.139.com':
            ctx = ssl.create_default_context()
            ctx.set_ciphers('DEFAULT')
            imap = imaplib.IMAP4_SSL(imap_server, ssl_context=ctx)
            imap.login(username, password)
            safe_log_message(f"Logged in as {username}")
            return imap

        imap = imaplib.IMAP4_SSL(imap_server)
        imap.login(username, password)
        # 163/126 邮箱需要发送 ID 命令标识客户端信息
        if imap_server in ['imap.163.com', 'imap.126.com']:
            args = ("name", "myclient", "contact", username, "version", "1.0.0", "vendor", "myclient")
            typ, dat = imap._simple_command('ID', '("' + '" "'.join(args) + '")')
            imap._untagged_response(typ, dat, 'ID')
        safe_log_message(f"Logged in as {username}")
        return imap
    except imaplib.IMAP4.error as e:
        safe_log_message(f"登录失败: {e}")
        root.after(0, lambda: messagebox.showerror("登录失败", f"登录失败: {e}"))
        return None
    except Exception as e:
        safe_log_message(f"连接错误: {e}")
        root.after(0, lambda: messagebox.showerror("连接错误", f"连接到IMAP服务器时发生错误: {e}"))
        return None


def decode_mime_words(s):
    decoded_fragments = decode_header(s)
    return ''.join([
        fragment.decode(encoding if encoding else 'utf-8') if isinstance(fragment, bytes) else fragment
        for fragment, encoding in decoded_fragments
    ])


def decode_payload(payload):
    for charset in ['utf-8', 'gbk', 'gb2312', 'iso-8859-1']:
        try:
            return payload.decode(charset.strip()).strip()
        except UnicodeDecodeError:
            continue
    return None


def extract_body_from_msg(msg):
    # 从邮件中提取纯文本正文内容
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
    global matching_emails
    try:
        imap.select("inbox")
    except Exception as e:
        safe_log_message(f"无法选择收件箱: {e}")
        return []

    try:
        status, messages = imap.search(None, 'ALL')
    except Exception as e:
        safe_log_message(f"搜索邮件时出错: {e}")
        return []

    if status != "OK":
        safe_log_message("获取邮件失败")
        return []

    email_ids = messages[0].split()
    safe_log_message(f"收件箱中共有 {len(email_ids)} 封邮件。")
    matching_emails = []
    total_emails = len(email_ids)

    for index, email_id in enumerate(email_ids, start=1):
        if is_paused:
            break
        result = process_email(email_id, imap, keyword)
        if result:
            matching_emails.append(result)
            safe_log_message(f"找到匹配的邮件(ID: {email_id.decode()}), 包含关键词 '{keyword}'。")
        else:
            safe_log_message(f"未找到匹配的邮件(ID: {email_id.decode()})。")
        # 更新进度条
        progress = int(index / total_emails * 100)
        update_progress(progress)

    safe_log_message(f"总共找到 {len(matching_emails)} 封正文中包含 '{keyword}' 的邮件。")
    return matching_emails


def save_all_emails_to_single_txt(emails, save_path):
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
        safe_log_message(f"所有邮件已导出到 {save_path}")
        root.after(0, lambda: messagebox.showinfo("完成", f"所有邮件已导出到 {save_path}"))
        folder_path = os.path.dirname(save_path)
        if platform.system() == "Windows":
            os.startfile(folder_path)
        elif platform.system() == "Darwin":  # macOS
            subprocess.call(["open", folder_path])
        else:  # Linux 或其他系统
            subprocess.call(["xdg-open", folder_path])
    except Exception as e:
        safe_log_message(f"保存邮件时出错: {e}")
        root.after(0, lambda: messagebox.showerror("保存错误", f"保存邮件到文件时出错: {e}"))


def fetch_outlook_emails_with_keyword(keyword="12306"):
    """
    使用 InteractiveBrowserCredential 登录 Microsoft 账号获取令牌，
    并通过 Graph API 分页获取所有邮件，然后筛选出包含 keyword 的邮件。
    """
    global matching_emails
    matching_emails = []  # 重置列表

    try:
        safe_log_message("正在使用 Outlook 模式登录，请稍候...")
        credential = InteractiveBrowserCredential(
            client_id=CLIENT_ID,
            tenant_id=TENANT_ID,
            redirect_uri="http://127.0.0.1:8400"
        )
        token = credential.get_token(*SCOPES).token
        safe_log_message("获取 Outlook 令牌成功，正在获取邮件...")

        headers = {
            'Authorization': f'Bearer {token}',
            'Content-Type': 'application/json'
        }

        # 分页获取所有邮件，单次最多获取1000封
        messages = []
        url = "https://graph.microsoft.com/v1.0/me/messages?$select=id,subject,body,from&$top=1000"
        while url:
            response = requests.get(url, headers=headers)
            if response.status_code != 200:
                safe_log_message(f"获取邮件失败，状态码：{response.status_code}")
                break
            data = response.json()
            current_page_messages = data.get("value", [])
            messages.extend(current_page_messages)
            safe_log_message(f"已获取 {len(current_page_messages)} 封邮件，累计总数：{len(messages)}")
            url = data.get('@odata.nextLink', None)

        safe_log_message(f"从 Outlook 共获取到 {len(messages)} 封邮件，开始筛选包含 '{keyword}' 的邮件...")

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

        safe_log_message(f"总共找到 {len(matching_emails)} 封正文中包含 '{keyword}' 的 Outlook 邮件。")
        return matching_emails

    except Exception as ex:
        safe_log_message(f"获取 Outlook 邮件时出错: {ex}")
        return []


def process_eml_files(folder_path):
    # 遍历文件夹中所有 .eml 文件，并解析为邮件对象后导出
    eml_files = [os.path.join(folder_path, f) for f in os.listdir(folder_path) if f.endswith('.eml')]
    if not eml_files:
        root.after(0, lambda: messagebox.showwarning("无文件", "所选文件夹中没有EML文件！"))
        return

    safe_log_eml_message(f"发现 {len(eml_files)} 个EML文件，开始解析...")
    emails = []
    for eml_file in eml_files:
        try:
            with open(eml_file, 'rb') as f:
                msg = BytesParser(policy=policy.default).parse(f)
                emails.append(msg)
                safe_log_eml_message(f"已解析: {os.path.basename(eml_file)}")
        except Exception as e:
            safe_log_eml_message(f"解析失败: {eml_file} - {str(e)}")

    if emails:
        save_path = filedialog.asksaveasfilename(
            defaultextension=".txt",
            filetypes=[("Text files", "*.txt")],
            title="保存解析结果"
        )
        if save_path:
            save_all_emails_to_single_txt(emails, save_path)
    else:
        root.after(0, lambda: messagebox.showinfo("结果", "未找到有效的邮件内容。"))


def fetch_outlook_button_click():
    global matching_emails, outlook_button_pressed
    if outlook_button_pressed:
        safe_log_message("您已经使用过 Outlook 登录。如需重新操作，请重启程序。")
        return
    outlook_button_pressed = True

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
            safe_log_message("没有找到包含 '12306' 的 Outlook 邮件。")

    threading.Thread(target=process).start()


def cancel_process():
    global is_paused
    is_paused = True
    safe_log_message("操作已取消。")


def start_imap_processing():
    global is_processing, is_paused, matching_emails, entry_username, entry_password, email_provider, entry_custom_imap
    if is_processing:
        # 如果正在处理，则认为用户点击了取消
        cancel_process()
        if matching_emails:
            save_path = filedialog.asksaveasfilename(defaultextension=".txt",
                                                     filetypes=[("Text files", "*.txt")])
            if save_path:
                safe_log_message("正在导出当前匹配到的邮件...")
                save_all_emails_to_single_txt(matching_emails, save_path)
        root.quit()
    else:
        def process():
            global is_paused, matching_emails, is_processing
            try:
                is_paused = False
                username = entry_username.get().strip()
                password = entry_password.get().strip()
                keyword = "12306"
                selected_provider = email_provider.get()
                if selected_provider == "custom":
                    imap_server = entry_custom_imap.get().strip()
                else:
                    imap_server = selected_provider

                if not imap_server:
                    root.after(0, lambda: messagebox.showwarning("输入错误", "请填写自定义 IMAP 服务器地址。"))
                    return

                if not username or not password:
                    root.after(0, lambda: messagebox.showwarning("输入错误", "请填写邮箱账号和授权码。"))
                    return

                safe_log_message(f"正在登录至 {imap_server} ...")
                imap = login_email(username, password, imap_server)
                if not imap:
                    return

                safe_log_message("正在搜索邮件中，请稍候...")
                emails = fetch_emails_with_keyword_in_body(imap, keyword)
                if len(emails) > 0 and not is_paused:
                    save_path = filedialog.asksaveasfilename(defaultextension=".txt",
                                                             filetypes=[("Text files", "*.txt")])
                    if save_path and not is_paused:
                        safe_log_message("正在导出匹配到的邮件...")
                        save_all_emails_to_single_txt(emails, save_path)
                else:
                    safe_log_message(f"未找到包含 '{keyword}' 的邮件。")
                    root.after(0, lambda: messagebox.showinfo("结果", f"未找到包含 '{keyword}' 的邮件。"))

                try:
                    imap.logout()
                    safe_log_message("已成功登出 IMAP。")
                except Exception as e:
                    safe_log_message(f"登出时发生错误: {e}")
                    root.after(0, lambda: messagebox.showerror("登出错误", f"登出时发生错误: {e}"))
            finally:
                is_processing = False
                update_progress(0)
                root.after(0, lambda: (start_btn.config(state=tk.NORMAL), cancel_btn.config(state=tk.DISABLED)))

        threading.Thread(target=process).start()
        is_processing = True
        start_btn.config(state=tk.DISABLED)
        cancel_btn.config(state=tk.NORMAL)


def main():
    global is_processing, is_paused, matching_emails, email_count_label, log_text, log_text_eml, root
    global entry_username, entry_password, email_provider, entry_custom_imap, outlook_button_pressed, progress_bar, start_btn, cancel_btn

    is_processing = False
    outlook_button_pressed = False
    root = tk.Tk()
    root.title("RailTicket2Chepiaopiao 邮件提取工具")

    window_width = 800
    window_height = 700
    screen_width = root.winfo_screenwidth()
    screen_height = root.winfo_screenheight()
    position_top = int(screen_height / 2 - window_height / 2)
    position_right = int(screen_width / 2 - window_width / 2)
    root.geometry(f"{window_width}x{window_height}+{position_right}+{position_top}")

    notebook = ttk.Notebook(root)
    notebook.pack(expand=True, fill="both", padx=10, pady=10)

    # 创建三个标签页：邮箱登录、 EML 文件导入、使用帮助
    login_frame = ttk.Frame(notebook)
    eml_frame = ttk.Frame(notebook)
    help_frame = ttk.Frame(notebook)

    notebook.add(login_frame, text="邮箱登录")
    notebook.add(eml_frame, text="EML 文件导入")
    notebook.add(help_frame, text="使用帮助")

    # --- 邮箱登录标签页 ---
    login_container = ttk.Frame(login_frame)
    login_container.pack(padx=20, pady=20, fill="both", expand=True)

    instruction_text = (
        "请输入您的邮箱账号和授权码，选择相应的邮箱服务。\n"
        "如果您使用自定义 IMAP 服务器，请选择“自定义IMAP”并填写服务器地址。\n"
        "点击“登录导出邮件”按钮后，系统将自动搜索导出12306相关的邮件。\n"
        "若要使用Outlook邮件提取，请点击“Outlook 登录获取”按钮, 在弹出网页中进行账号密码输入，无需在此处输入账号密码。(受限于Outlook认证方式）"
    )
    instruction_label = ttk.Label(login_container, text=instruction_text, wraplength=650, justify="left",
                                  foreground="blue")
    instruction_label.grid(row=0, column=0, columnspan=3, pady=(0, 15), sticky="w")

    ttk.Label(login_container, text="邮箱账号:").grid(row=1, column=0, padx=5, pady=5, sticky="e")
    entry_username = ttk.Entry(login_container, width=30)
    entry_username.grid(row=1, column=1, padx=5, pady=5, sticky="w")

    ttk.Label(login_container, text="授权码:").grid(row=2, column=0, padx=5, pady=5, sticky="e")
    entry_password = ttk.Entry(login_container, show="*", width=30)
    entry_password.grid(row=2, column=1, padx=5, pady=5, sticky="w")

    ttk.Label(login_container, text="请选择邮箱服务:").grid(row=3, column=0, padx=5, pady=5, sticky="e")
    email_provider = tk.StringVar(value="imap.qq.com")
    provider_frame = ttk.Frame(login_container)
    provider_frame.grid(row=3, column=1, padx=5, pady=5, sticky="w")

    providers = [
        ("QQ邮箱", "imap.qq.com"),
        ("163邮箱", "imap.163.com"),
        ("126邮箱", "imap.126.com"),
        ("139邮箱", "imap.139.com"),
        ("Gmail", "imap.gmail.com"),
        ("自定义IMAP", "custom")
    ]
    col = 0
    for text, value in providers:
        rb = ttk.Radiobutton(provider_frame, text=text, variable=email_provider, value=value)
        rb.grid(row=0, column=col, padx=5, pady=2, sticky="w")
        col += 1

    ttk.Label(login_container, text="自定义IMAP服务器:").grid(row=4, column=0, padx=5, pady=5, sticky="e")
    entry_custom_imap = ttk.Entry(login_container, width=30)
    entry_custom_imap.grid(row=4, column=1, padx=5, pady=5, sticky="w")
    entry_custom_imap.config(state=tk.DISABLED)

    def toggle_custom_imap(*args):
        state = tk.NORMAL if email_provider.get() == "custom" else tk.DISABLED
        entry_custom_imap.config(state=state)

    email_provider.trace_add("write", toggle_custom_imap)

    action_frame = ttk.Frame(login_container)
    action_frame.grid(row=5, column=0, columnspan=3, pady=15)
    start_btn = ttk.Button(action_frame, text="登录导出邮件", command=start_imap_processing, width=20)
    start_btn.grid(row=0, column=0, padx=10)
    outlook_btn = ttk.Button(action_frame, text="Outlook 登录获取", command=fetch_outlook_button_click, width=20)
    outlook_btn.grid(row=0, column=1, padx=10)
    cancel_btn = ttk.Button(action_frame, text="取消操作", command=cancel_process, width=20)
    cancel_btn.grid(row=0, column=2, padx=10)
    cancel_btn.config(state=tk.DISABLED)

    email_count_label = ttk.Label(login_container, text="当前匹配的邮件数: 0")
    email_count_label.grid(row=6, column=0, columnspan=3, pady=5, sticky="w")

    progress_bar = ttk.Progressbar(login_container, orient="horizontal", length=400, mode="determinate")
    progress_bar.grid(row=7, column=0, columnspan=3, pady=10)

    log_text = tk.Text(login_container, height=12, width=80)
    log_text.grid(row=8, column=0, columnspan=3, pady=10)
    log_scroll = ttk.Scrollbar(login_container, orient="vertical", command=log_text.yview)
    log_scroll.grid(row=8, column=3, sticky="ns", pady=10)
    log_text.config(yscrollcommand=log_scroll.set)



    # --- EML 文件导入标签页 ---
    eml_container = ttk.Frame(eml_frame)
    eml_container.pack(padx=20, pady=20, fill="both", expand=True)

    eml_instructions = (
        "请选择包含 EML 文件的文件夹：\n\n"
        "1. 在邮箱客户端中导出邮件为 EML 格式\n"
        "2. 将所有 EML 文件放入同一文件夹\n"
        "3. 点击下方按钮选择该文件夹，程序将自动解析并导出邮件内容。"
    )
    eml_label = ttk.Label(eml_container, text=eml_instructions, wraplength=650, justify="left", foreground="blue")
    eml_label.pack(pady=10)
    select_btn = ttk.Button(eml_container, text="选择文件夹", command=lambda: select_folder(), width=20)
    select_btn.pack(pady=10)
    log_text_eml = tk.Text(eml_container, height=12, width=80)
    log_text_eml.pack(pady=10)
    eml_scroll = ttk.Scrollbar(eml_container, orient="vertical", command=log_text_eml.yview)
    eml_scroll.pack(side="right", fill="y", pady=10)
    log_text_eml.config(yscrollcommand=eml_scroll.set)

    # --- 使用帮助标签页 ---
    help_container = ttk.Frame(help_frame)
    help_container.pack(padx=20, pady=20, fill="both", expand=True)
    help_text = (
        "【工具简介】\n\n"
        "本工具支持通过 IMAP 登录和 Outlook API 一键批量导出12306相关的邮件。\n"
        "目前支持的邮箱包括 QQ、163、126、139、Gmail 和 Outlook。\n"
        "常规登录模式：请填写邮箱账号和授权码，并选择对应的邮箱服务。若您的邮箱服务不在列表中，请选择 '自定义IMAP' 并手动输入服务器地址。\n\n"
        "Outlook 登录模式：点击 'Outlook 登录获取' 按钮，按提示在浏览器中完成微软账号验证，程序会自动通过 Microsoft Graph API 获取邮件。\n\n"
        "EML 文件导入模式：如果您已手动导出邮件为 EML 格式，请使用此功能。选择包含 EML 文件的文件夹后，程序将解析并导出邮件内容。\n\n"
        "【注意事项】\n"
        "1. 部分邮箱登录需要使用授权码而非密码，请提前获取对应授权码。\n"
        "2. 程序在导出邮件时，会将所有匹配到的邮件整合到一个 TXT 文件中，方便查阅。\n"
        "3. 若遇到任何问题，请联系：yukiender@foxmail.com 或在 GitHub 上提交 issue\n\n"
        "你可以在这里获取到最新版本：https://github.com/sakura-ender/RailTicket2Chepiaopiao"
    )
    help_label = ttk.Label(help_container, text=help_text, wraplength=650, justify="left")
    help_label.pack(pady=10)

    root.mainloop()


def select_folder():
    folder_path = filedialog.askdirectory()
    if folder_path:
        safe_log_eml_message(f"已选择文件夹: {folder_path}")
        process_eml_files(folder_path)


if __name__ == "__main__":
    main()
