import imaplib
import email
from email.header import decode_header
import os
import time
from bs4 import BeautifulSoup  # 导入BeautifulSoup库

# 登录QQ邮箱
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

# 解码邮件标题
def decode_mime_words(s):
    decoded_fragments = decode_header(s)
    return ''.join([
        fragment.decode(encoding if encoding else 'utf-8') if isinstance(fragment, bytes) else fragment
        for fragment, encoding in decoded_fragments
    ])

# 处理每封邮件的函数
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

            # 获取邮件正文
            body = None
            if msg.is_multipart():
                for part in msg.walk():
                    content_type = part.get_content_type()
                    content_disposition = str(part.get("Content-Disposition"))

                    if "attachment" not in content_disposition and content_type.startswith("text/"):
                        payload = part.get_payload(decode=True)
                        if payload:
                            # 尝试多种编码
                            for charset in ['utf-8', 'gbk', 'gb2312', 'iso-8859-1']:
                                try:
                                    decoded_payload = payload.decode(charset).strip()
                                    # 使用BeautifulSoup解析HTML并提取纯文本
                                    soup = BeautifulSoup(decoded_payload, 'html.parser')
                                    body = soup.get_text(separator='\n')  # 使用换行符分隔不同的文本块
                                    break
                                except UnicodeDecodeError:
                                    continue
                            else:
                                print("无法解码邮件正文。")
            else:
                payload = msg.get_payload(decode=True)
                if payload:
                    for charset in ['utf-8', 'gbk', 'gb2312', 'iso-8859-1']:
                        try:
                            decoded_payload = payload.decode(charset).strip()
                            soup = BeautifulSoup(decoded_payload, 'html.parser')
                            body = soup.get_text(separator='\n')
                            break
                        except UnicodeDecodeError:
                            continue
                    else:
                        print("无法解码邮件正文。")

            # 检查正文是否包含关键词
            if body and keyword in body:
                return msg
    return None

# 获取邮件列表并处理
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

# 保存所有符合条件的邮件到一个txt文件
def save_all_emails_to_single_txt(emails, save_path):
    try:
        with open(save_path, 'w', encoding='utf-8') as f:
            for i, msg in enumerate(emails, 1):
                subject = msg.get("Subject", "")
                decoded_subject = decode_mime_words(subject)
                f.write(f"\n--- 第 {i} 封邮件 ---\n")
                f.write(f"标题: {decoded_subject}\n")

                body = None
                if msg.is_multipart():
                    for part in msg.walk():
                        content_type = part.get_content_type()
                        content_disposition = str(part.get("Content-Disposition"))
                        if "attachment" not in content_disposition and content_type.startswith("text/"):
                            payload = part.get_payload(decode=True)
                            if payload:
                                for charset in ['utf-8', 'gbk', 'gb2312', 'iso-8859-1']:
                                    try:
                                        decoded_payload = payload.decode(charset).strip()
                                        soup = BeautifulSoup(decoded_payload, 'html.parser')
                                        body = soup.get_text(separator='\n')
                                        break
                                    except UnicodeDecodeError:
                                        continue
                                else:
                                    print("无法解码邮件正文，跳过此部分。")
                else:
                    payload = msg.get_payload(decode=True)
                    if payload:
                        for charset in ['utf-8', 'gbk', 'gb2312', 'iso-8859-1']:
                            try:
                                decoded_payload = payload.decode(charset).strip()
                                soup = BeautifulSoup(decoded_payload, 'html.parser')
                                body = soup.get_text(separator='\n')
                                break
                            except UnicodeDecodeError:
                                continue
                        else:
                            print("无法解码邮件正文，跳过此部分。")

                if body:
                    f.write(f"正文:\n{body}\n")
                else:
                    f.write("正文: 无法解码或为空。\n")
                f.write("\n--- 结束 ---\n")
                print(f"正在导出第 {i} 封邮件到 {save_path}。")
        print(f"所有邮件已导出到 {save_path}")
    except Exception as e:
        print(f"保存邮件到文件时出错: {e}")

# 主程序
def main():
    # 请填入您的完整QQ邮箱地址和授权码
    username = '570802322@qq.com'  # 替换为你的QQ邮箱地址
    password = 'cbymezsqlqiybbjh'  # 替换为你的QQ邮箱授权码

    if username == 'your_email@qq.com' or password == 'your_authorization_code':
        print("请在代码中填写您的QQ邮箱地址和授权码。")
        return

    imap = login_qq_email(username, password)

    if not imap:
        print("登录失败，程序终止。")
        return

    keyword = '12306'  # 需要搜索的关键词
    print(f"正在搜索包含关键词 '{keyword}' 的邮件...")
    emails = fetch_emails_with_keyword_in_body(imap, keyword)

    print(f"总共有 {len(emails)} 封邮件包含关键词 '{keyword}'。")

    if len(emails) > 0:
        # 询问用户是否需要保存这些邮件
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