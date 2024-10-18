import imaplib
import email
from email.header import decode_header
import os
import time
from bs4 import BeautifulSoup


def login_qq_email(username, password):
    imap_server = 'imap.qq.com'
    try:
        imap = imaplib.IMAP4_SSL(imap_server)
        imap.login(username, password)
        print("登录成功！")
        return imap
    except imaplib.IMAP4.error as e:
        print(f"登录失败: {e}")
        return None
    except Exception as e:
        print(f"连接到IMAP服务器时发生错误: {e}")
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
    print("无法解码邮件正文。")
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
                    if body and '<' in body and '>' in body:  # Check if body contains HTML tags
                        text = BeautifulSoup(body, 'html.parser').get_text(separator='\n')
                        return '\n'.join([line.strip() for line in text.splitlines() if line.strip()])
    else:
        payload = msg.get_payload(decode=True)
        if payload:
            body = decode_payload(payload)
            if body and '<' in body and '>' in body:  # Check if body contains HTML tags
                text = BeautifulSoup(body, 'html.parser').get_text(separator='\n')
                return '\n'.join([line.strip() for line in text.splitlines() if line.strip()])
    return None


def process_email(email_id, imap, keyword, retries=3, delay=5):
    for attempt in range(1, retries + 1):
        try:
            status, msg_data = imap.fetch(email_id, '(RFC822)')
            if status == 'OK':
                break
            else:
                print(f"意外的状态: {status}, 正在重试 ({attempt}/{retries})...")
        except imaplib.IMAP4.abort as e:
            print(f"Fetch命令中止: {e}, 正在重试 ({attempt}/{retries})...")
            time.sleep(delay)
        except Exception as e:
            print(f"发生意外错误: {e}, 正在重试 ({attempt}/{retries})...")
            time.sleep(delay)
    else:
        print(f"在 {retries} 次尝试后无法获取邮件 {email_id}。")
        return None

    if not msg_data:
        print(f"邮件 {email_id} 没有消息数据。")
        return None

    for response_part in msg_data:
        if isinstance(response_part, tuple):
            try:
                msg = email.message_from_bytes(response_part[1])
            except Exception as e:
                print(f"解析邮件时出错: {e}")
                continue

            body = extract_body_from_msg(msg)

            if body and keyword in body:
                return msg
    return None


def fetch_emails_with_keyword_in_body(imap, keyword):
    try:
        imap.select("inbox")
    except Exception as e:
        print(f"无法选择收件箱: {e}")
        return []

    try:
        status, messages = imap.search(None, 'ALL')
    except Exception as e:
        print(f"搜索邮件时出错: {e}")
        return []

    if status != "OK":
        print("获取邮件失败")
        return []

    email_ids = messages[0].split()
    print(f"收件箱中共有 {len(email_ids)} 封邮件。")

    matching_emails = []

    for email_id in email_ids:
        result = process_email(email_id, imap, keyword)
        if result:
            matching_emails.append(result)
            print(f"找到匹配的邮件(ID: {email_id.decode()}), 包含关键词 '{keyword}'。")

    print(f"总共找到 {len(matching_emails)} 封正文中包含 '{keyword}' 的邮件。")
    return matching_emails


def save_all_emails_to_single_txt(emails, save_path):
    try:
        with open(save_path, 'w', encoding='utf-8') as f:
            for i, msg in enumerate(emails, 1):
                subject = msg.get("Subject", "")
                decoded_subject = decode_mime_words(subject)
                f.write(f"\n###\n")
                # f.write(f"标题: {decoded_subject}\n")

                body = extract_body_from_msg(msg)

                if body:
                    f.write(f"正文:\n{body}\n")
                else:
                    f.write("正文: 无法解码或为空。\n")
                f.write("\n###\n")
                print(f"正在导出第 {i} 封邮件到 {save_path}。")
        print(f"所有邮件已导出到 {save_path}")
    except Exception as e:
        print(f"保存邮件到文件时出错: {e}")


def main():
    username = '570802322@qq.com'
    password = 'cbymezsqlqiybbjh'

    if username == 'your_email@qq.com' or password == 'your_authorization_code':
        print("请在代码中填写您的QQ邮箱地址和授权码。")
        return

    imap = login_qq_email(username, password)

    if not imap:
        print("登录失败，程序终止。")
        return

    keyword = '12306'
    print(f"正在搜索包含关键词 '{keyword}' 的邮件...")
    emails = fetch_emails_with_keyword_in_body(imap, keyword)

    print(f"总共有 {len(emails)} 封邮件包含关键词 '{keyword}'。")

    if len(emails) > 0:
        save_choice = input("是否将这些邮件保存到一个txt文件中？（y/n）：").strip().lower()
        if save_choice == 'y':
            save_path = './all_emails.txt'
            save_all_emails_to_single_txt(emails, save_path)
        else:
            print("不保存邮件。")
    else:
        print(f"未找到包含 '{keyword}' 的邮件。")

    try:
        imap.logout()
        print("已成功登出邮箱。")
    except Exception as e:
        print(f"登出时出错: {e}")


if __name__ == "__main__":
    main()
