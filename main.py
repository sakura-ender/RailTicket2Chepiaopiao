import imaplib
import email
from email.header import decode_header
import os
import time
from bs4 import BeautifulSoup
import tkinter as tk
from tkinter import filedialog, messagebox
import threading
import re
import subprocess
import platform

is_paused = False
matching_emails = []


def login_email(username, password, imap_server):
    try:
        imap = imaplib.IMAP4_SSL(imap_server)
        imap.login(username, password)
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
        except imaplib.IMAP4.abort as e:
            time.sleep(delay)
        except Exception as e:
            time.sleep(delay)
    else:
        return None

    if not msg_data:
        return None

    for response_part in msg_data:
        if isinstance(response_part, tuple):
            try:
                msg = email.message_from_bytes(response_part[1])
            except Exception as e:
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


def log_message(message):
    global log_text, email_count_label
    log_text.insert(tk.END, message + '\n')
    log_text.see(tk.END)
    email_count_label.config(text=f"当前匹配的邮件数: {len(matching_emails)}")


def main():
    global is_processing, is_paused, matching_emails, email_count_label
    is_processing = False

    def start_processing():
        global is_processing, is_paused, matching_emails
        if is_processing:
            is_paused = True
            if matching_emails:
                save_path = filedialog.asksaveasfilename(defaultextension=".txt",
                                                         filetypes=[("Text files", "*.txt")])
                if save_path:
                    log_message("Saving emails...")
                    save_all_emails_to_single_txt(matching_emails, save_path)
            messagebox.showinfo("暂停", "程序已暂停并导出当前匹配的邮件。")
            root.quit()
        else:
            def process():
                global is_paused, matching_emails
                is_paused = False
                username = entry_username.get()
                password = entry_password.get()
                keyword = "12306"

                if not username or not password:
                    messagebox.showwarning("输入错误", "请填写所有字段。")
                    return

                selected_provider = email_provider.get()
                if selected_provider == "custom":
                    imap_server = entry_custom_imap.get()
                else:
                    imap_server = selected_provider

                log_message("Logging in...")
                imap = login_email(username, password, imap_server)

                if not imap:
                    return

                log_message("Fetching emails...")
                emails = fetch_emails_with_keyword_in_body(imap, keyword)
                if len(emails) > 0:
                    save_path = filedialog.asksaveasfilename(defaultextension=".txt",
                                                             filetypes=[("Text files", "*.txt")])
                    if save_path:
                        log_message("Saving emails...")
                        save_all_emails_to_single_txt(emails, save_path)
                else:
                    log_message(f"No emails found containing '{keyword}'")
                    messagebox.showinfo("结果", f"未找到包含 '{keyword}' 的邮件。")

                try:
                    imap.logout()
                    log_message("Logged out.")
                except Exception as e:
                    log_message(f"Error logging out: {e}")
                    messagebox.showerror("登出错误", f"登出时出错: {e}")

            threading.Thread(target=process).start()
            is_processing = True
            start_button.config(text="中途暂停并导出目前的邮件(不推荐)")

    root = tk.Tk()
    root.title("邮箱登录")

    window_width = 400
    window_height = 400
    screen_width = root.winfo_screenwidth()
    screen_height = root.winfo_screenheight()
    position_top = int(screen_height / 2 - window_height / 2)
    position_right = int(screen_width / 2 - window_width / 2)
    root.geometry(f'{window_width}x{window_height}+{position_right}+{position_top}')

    tk.Label(root, text="邮箱账号:").grid(row=0, column=0, padx=10, pady=5)
    entry_username = tk.Entry(root)
    entry_username.grid(row=0, column=1, padx=10, pady=5)

    tk.Label(root, text="授权码:").grid(row=1, column=0, padx=10, pady=5)
    entry_password = tk.Entry(root, show="*")
    entry_password.grid(row=1, column=1, padx=10, pady=5)

    email_provider = tk.StringVar(value="imap.qq.com")
    tk.Radiobutton(root, text="QQ邮箱", variable=email_provider, value="imap.qq.com").grid(row=2, column=0, padx=10,
                                                                                           pady=5)
    tk.Radiobutton(root, text="163邮箱", variable=email_provider, value="imap.163.com").grid(row=2, column=1, padx=10,
                                                                                             pady=5)
    tk.Radiobutton(root, text="Outlook（不支持）", variable=email_provider, value="imap-mail.outlook.com").grid(row=3,
                                                                                                              column=0,
                                                                                                              padx=10,
                                                                                                              pady=5)
    tk.Radiobutton(root, text="Gmail", variable=email_provider, value="imap.gmail.com").grid(row=3, column=1, padx=10,
                                                                                             pady=5)
    tk.Radiobutton(root, text="自定义", variable=email_provider, value="custom").grid(row=4, column=0, padx=10, pady=5)

    tk.Label(root, text="自定义IMAP服务器:").grid(row=5, column=0, padx=10, pady=5)
    entry_custom_imap = tk.Entry(root)
    entry_custom_imap.grid(row=5, column=1, padx=10, pady=5)
    entry_custom_imap.config(state=tk.DISABLED)

    def toggle_custom_imap(*args):
        if email_provider.get() == "custom":
            entry_custom_imap.config(state=tk.NORMAL)
        else:
            entry_custom_imap.config(state=tk.DISABLED)

    email_provider.trace("w", toggle_custom_imap)

    start_button = tk.Button(root, text="开始处理", command=start_processing)
    start_button.grid(row=6, columnspan=2, pady=10)

    email_count_label = tk.Label(root, text="当前匹配的邮件数: 0")
    email_count_label.grid(row=7, columnspan=2, pady=5)

    global log_text
    log_text = tk.Text(root, height=10, width=50)
    log_text.grid(row=8, columnspan=2, padx=10, pady=5)

    root.mainloop()


if __name__ == "__main__":
    main()
