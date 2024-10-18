import imaplib
import email
from email.header import decode_header
import os
import concurrent.futures
import time

# 登录QQ邮箱
def login_qq_email(username, password):
    imap_server = 'imap.qq.com'
    imap = imaplib.IMAP4_SSL(imap_server)
    try:
        imap.login(username, password)
        print("登录成功！")
    except imaplib.IMAP4.error as e:
        print(f"登录失败: {e}")
        return None
    return imap





# 处理每封邮件的函数
def process_email(email_id, imap, keyword, retries=3, delay=5):
    for attempt in range(retries):
        try:
            status, msg_data = imap.fetch(email_id, '(RFC822)')
            if status == 'OK':
                break
            else:
                print(f"Unexpected status: {status}, retrying ({attempt + 1}/{retries})...")
        except imaplib.IMAP4.abort as e:
            print(f"Fetch command aborted: {e}, retrying ({attempt + 1}/{retries})...")
            time.sleep(delay)
        except Exception as e:
            print(f"Unexpected error: {e}, retrying ({attempt + 1}/{retries})...")
            time.sleep(delay)
    else:
        print(f"Failed to fetch email {email_id} after {retries} attempts.")
        return None

    if not msg_data:
        print(f"No message data for email {email_id}.")
        return None

    for response_part in msg_data:
        if isinstance(response_part, tuple):
            try:
                msg = email.message_from_bytes(response_part[1])
            except Exception as e:
                print(f"Error parsing email: {e}")
                continue

            # 获取邮件正文
            body = None
            if msg.is_multipart():
                for part in msg.walk():
                    content_type = part.get_content_type()
                    content_disposition = str(part.get("Content-Disposition"))

                    if "attachment" not in content_disposition:
                        try:
                            charset = part.get_content_charset() or 'utf-8'  # 获取邮件的编码格式
                            body = part.get_payload(decode=True).decode(charset)
                        except UnicodeDecodeError as e:
                            print(f"解码错误: {e}，尝试使用不同编码解码")
                            continue
            else:
                try:
                    charset = msg.get_content_charset() or 'utf-8'
                    body = msg.get_payload(decode=True).decode(charset)
                except UnicodeDecodeError as e:
                    print(f"解码错误: {e}，尝试使用不同编码解码")
                    continue

            # 检查正文是否包含关键词
            if body and keyword in body:
                return msg
    return None
# 获取邮件列表并并行处理
def fetch_emails_with_keyword_in_body(imap, keyword):
    imap.select("inbox")
    status, messages = imap.search(None, 'ALL')

    if status != "OK":
        print("获取邮件失败")
        return []

    email_ids = messages[0].split()
    print(f"收件箱中共有 {len(email_ids)} 封邮件。")

    matching_emails = []

    # 使用线程池并行处理邮件
    with concurrent.futures.ThreadPoolExecutor() as executor:
        futures = [executor.submit(process_email, email_id, imap, keyword) for email_id in email_ids]

        for future in concurrent.futures.as_completed(futures):
            result = future.result()
            if result:
                matching_emails.append(result)
                print("找到匹配的邮件，包含关键词")

    print(f"总共找到 {len(matching_emails)} 封正文中包含 '{keyword}' 的邮件。")
    return matching_emails


# 保存所有符合条件的邮件到一个txt文件
def save_all_emails_to_single_txt(emails, save_path):
    with open(save_path, 'w', encoding='utf-8') as f:
        for i, msg in enumerate(emails, 1):
            subject = msg["Subject"]
            f.write(f"\n--- 第 {i} 封邮件 ---\n")
            f.write(f"标题: {subject}\n")

            body = None
            if msg.is_multipart():
                for part in msg.walk():
                    content_type = part.get_content_type()
                    content_disposition = str(part.get("Content-Disposition"))
                    if "attachment" not in content_disposition:
                        try:
                            charset = part.get_content_charset() or 'utf-8'
                            body = part.get_payload(decode=True).decode(charset)
                        except UnicodeDecodeError as e:
                            print(f"解码错误: {e}，跳过此部分")
                            continue
            else:
                try:
                    charset = msg.get_content_charset() or 'utf-8'
                    body = msg.get_payload(decode=True).decode(charset)
                except UnicodeDecodeError as e:
                    print(f"解码错误: {e}，跳过此部分")
                    continue

            f.write(f"正文:\n{body}\n")
            f.write("\n--- 结束 ---\n")
            print(f"正在导出第 {i} 封邮件到总文件。")
    print(f"所有邮件已导出到 {save_path}")


# 主程序
def main():
    username = '570802322@qq.com'
    password = 'kkgtbajebeclbgaa'

    imap = login_qq_email(username, password)

    if not imap:
        print("登录失败，程序终止")
        return

    keyword = '12306'
    emails = fetch_emails_with_keyword_in_body(imap, keyword)

    if emails:
        save_path = './all_emails.txt'
        save_all_emails_to_single_txt(emails, save_path)
    else:
        print(f"未找到包含 '{keyword}' 的邮件。")

    imap.logout()


if __name__ == "__main__":
    main()
