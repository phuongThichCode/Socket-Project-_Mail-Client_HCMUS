import re
import os
from os import path
from os.path import basename
from socket import socket, AF_INET, SOCK_STREAM
import email
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email.utils import formatdate
from string import printable
from typing import List
from email import encoders
from email import message_from_string
from email import message_from_bytes
from email.policy import default
import configparser
import base64
import codecs
import quopri
import json
import threading
import time
from email.parser import BytesParser
from email import policy
from email.header import decode_header
from email.iterators import walk
import sys

exit_thread_flag = True
#Hàm để kết nối server bằng giao thức TCP với socket
def initiate(mailServer, port) -> socket:
    client_socket = socket(AF_INET, SOCK_STREAM)
    client_socket.connect((mailServer, port))
    client_socket.recv(1024)
    return client_socket


def send_email_with_attachments(client_socket, from_user, to_user, cc_users, bcc_users, subject, message, attachment_paths) -> None:
    client_socket.send('EHLO\r\n'.encode())
    client_socket.recv(1024)

    client_socket.send(f'MAIL FROM: {from_user}\r\n'.encode())
    client_socket.recv(1024)
    print(len(cc_users))
    receivers = []
    receivers.extend(to_user)
    if cc_users[0] != '':
         receivers.extend(cc_users)
    if bcc_users[0] != '':
         receivers.extend(bcc_users)

    print(receivers) # print to check if emails are correctly separated
    for receiver in receivers:
        if receiver != "['']":
             client_socket.send(f'RCPT TO: {receiver}\r\n'.encode())
             client_socket.recv(1024)

    client_socket.send('DATA\r\n'.encode())
    client_socket.recv(1024)

    msg = MIMEMultipart()
    msg['From'] = from_user
    msg['To'] = ", ".join(to_user)
    msg['Subject'] = subject
    if cc_users[0] != '':
        msg['Cc'] = ", ".join(cc_users)
    
    msg.attach(MIMEText(f'{message}'))


    for attachment_path in attachment_paths:
        print(attachment_path)
        part = MIMEBase('application', 'octet-stream')
        part.set_payload(open(attachment_path, "rb").read())
        encoders.encode_base64(part)
        part.add_header('Content-Disposition', f'attachment;filename={os.path.basename(attachment_path)}')
        msg.attach(part)

    client_socket.sendall(msg.as_string().encode())
    end_message = f'\r\n.\r\n'
    client_socket.send(end_message.encode())
    client_socket.recv(1024)

    print("Đã gửi email thành công")

    quit_command = 'QUIT\r\n'
    client_socket.send(quit_command.encode())
    client_socket.recv(1024)
    client_socket.close()
    
def login(ssl_socket, username, password):
    # Gửi lệnh USER để đăng nhập
    ssl_socket.sendall(f'USER {username}\r\n'.encode())
    print("USER RESPONSE: ")
    print(ssl_socket.recv(1024).decode())
    ssl_socket.sendall(f'PASS {password}\r\n'.encode())
    print("PASS RESPONSE: ")
    print(ssl_socket.recv(1024).decode())

#Hàm để phân loại dựa vào cấu hình(trả về tên Folder sẽ tạo ra để đưa emails tương ứng vào)
def classify(sender, subject, content):
    if ('ahihi@testing.com' in sender or 'ahuu@testing.com' in sender):
        return 'Project'
    elif ('urgent' in subject or 'ASAP' in subject):
        return 'Important'
    elif ('report' in content or 'meeting' in content):
        return 'Work'
    else:
        spams = ["virus","hack","crack"]
        for spam in spams:
            if (spam in subject or spam in content):
                return 'Spam'
    return 'Inbox'

BUFFER_SIZE = 1024
#Tạo vòng lặp để nhận liên tục đến khi hết content
def receive_all(socket):
    data = b""
    while True:
        part = socket.recv(BUFFER_SIZE)
        data += part
        if len(part) < BUFFER_SIZE:
            break
    return data

#Đếm số lượng emails đã lưu về máy
def count_mail_on_local(user):
    path = f"C:/Users/Thuc Do Huu/OneDrive - VNU-HCMUS/Desktop/{user}"
    try:
        file_list = []
        for each_folder in os.listdir(path):
            file_list += os.listdir(f'{path}/{each_folder}') # lấy toàn bộ files
        num = len(list(set(file_list))) + 1
        return num
    except FileNotFoundError:
        print(" ")
        return 0

#Đếm số lượng emails có trên server
def count_mail_on_server(ssl_socket, user):
    ssl_socket.sendall(f'USER {user}\r\n'.encode())
    recv = ssl_socket.recv(1024).decode()
    ssl_socket.sendall(f'STAT\r\n'.encode())
    recv = ssl_socket.recv(1024).decode().split(' ')
    return int(recv[1]) + 1

#Nhận emails, sau đó phân loại( bằng cách đưa emails vào các thư mục như Inbox,...)
def retrieve_email_with_attachment(ssl_socket, save_directory, user):  

    count = count_mail_on_local(user)
    more = count_mail_on_server(ssl_socket, user)
    
    for i in range(count, more):    
        ssl_socket.sendall(f'RETR {i}\r\n'.encode())  
        retr_response = receive_all(ssl_socket).decode()
   
        sender=""
        subject=""
        content=""
        to=""
        cc=""
        bcc=""

        sender_p = re.compile(r'^From:\s*(.+)$', re.MULTILINE)
        to_p = re.compile(r'^To:\s*(.+)$', re.MULTILINE)
        subject_p = re.compile(r'^Subject:\s*(.+)$', re.MULTILINE)
      
        cc_p = re.compile(r'^Cc: (.+)$', re.MULTILINE)
        bcc_p = re.compile(r'^Bcc: (.+)$', re.MULTILINE)

        sender_m = sender_p.search(retr_response)
        to_m = to_p.search(retr_response)
        subject_m = subject_p.search(retr_response)
        bcc_m = bcc_p.search(retr_response)
      
        cc_m = cc_p.search(retr_response)
        if sender_m:
            sender = sender_m.group(1)
        #else: print('error sender')
        if subject_m:
            subject = subject_m.group(1)
           
        #else: print('error subject')
        if to_m:
            to = to_m.group(1)
        #else: print('error to')
        
        lines = retr_response.split('\n')

        # Tìm lần xuất hiện cuối cùng của một dòng bắt đầu bằng mẫu tiêu đề
        last_header_index = max((i for i, line in enumerate(lines) if re.match(r'^[A-Za-z-]+:', line)), default=-1)

        # Xác dịnh nội dung
        content_lines = lines[last_header_index + 1:]
        content = '\n'.join(content_lines).strip()

        
        # Đưa emails vào folder tương ứng
        name = classify(sender, subject, content)
        folder_path = os.path.join(save_directory, name)
        os.makedirs(folder_path, exist_ok=True)
  
        email_path = os.path.join(folder_path, f'email_{i}.eml')
        with open(email_path, 'wb') as email_file:
            email_file.write(retr_response.encode())
    
 
# Hàm để hiển thị danh sách email trong folder được chọn
def show_emails_in_folder(selected_folder, email_read_status, path_to_save, folders):
    folder_path = os.path.join(path_to_save, selected_folder)
    print(f"Đây là danh sách email trong {folders[selected_folder]}:")
    
    emails = os.listdir(folder_path)
    while len(email_read_status) < len(emails):
        email_read_status.append(False)

    if not emails:
        print("Không có email trong thư mục.")
    else:
        for index, (email, read_status) in enumerate(zip(emails, email_read_status), start=1):
            read_status_str = "" if read_status else "(Chưa đọc)"
            sender = ''
            subject = ''
            
            # Trích xuất thông tin người gửi và chủ đề từ tên tệp email
            email_path = os.path.join(folder_path, email)
            with open(email_path, 'r') as email_file:
                email_content = email_file.read()
                sender_match = re.search(r'^From:\s*(.+)$', email_content, re.MULTILINE)
                subject_match = re.search(r'^Subject:\s*(.+)$', email_content, re.MULTILINE)
                
                if sender_match:
                    sender = sender_match.group(1).strip()
                if subject_match:
                    subject = subject_match.group(1).strip()

            print(f"{index}.{read_status_str} <{sender}> , <{subject}>")

# Hàm load_email_read_status(username) sẽ kiểm tra xem có tệp tin lưu trạng thái email đọc của người dùng không.
# Nếu có, nó sẽ đọc nó và trả về, nếu không, nó sẽ trả về một trạng thái mới.
def load_email_read_status(username, password):
    status_file_path = f'email_read_status_{username}_{password}.json'
    if os.path.exists(status_file_path):
        with open(status_file_path, 'r') as status_file:
            return json.load(status_file)
    else:
        return {
            'Inbox': [],
            'Project': [],
            'Important': [],
            'Work': [],
            'Spam': []
        }

# Hàm save_email_read_status(username,password, email_read_status) sẽ lưu trạng thái email đọc của người dùng vào một tệp tin.
def save_email_read_status(username, password, email_read_status):
    status_file_path = f'email_read_status_{username}_{password}.json'
    with open(status_file_path, 'w') as status_file:
        json.dump(email_read_status, status_file)


# Hàm lưu thông tin người dùng vào tệp tin cấu hình
def save_user_config(username, password, mail_server, smtp_port, pop3_port):
    user_config = {
        'General': {
            'Username': username,
            'Password': password,
            'MailServer': mail_server,
            'SMTP': smtp_port,
            'POP3': pop3_port,
            'AUTOLOAD' :10
        }
    }

    # Lưu thông tin người dùng vào tệp tin cấu hình với tên đặc biệt cho mỗi người dùng
    with open(f'user_config_{username}.json', 'w') as configfile:
        json.dump(user_config, configfile)

# Hàm load thông tin người dùng từ tệp tin cấu hình
def load_user_config(username):
    config_file_path = f'user_config_{username}.json'
    if os.path.exists(config_file_path):
        with open(config_file_path, 'r') as configfile:
            user_config = json.load(configfile)
            return user_config.get('General', {})
    else:
        return {}

# Hàm kiểm tra xem thông tin người dùng đã tồn tại hay chưa
def check_user_exist(username):
    config_file_path = f'user_config_{username}.json'
    return os.path.exists(config_file_path)

#Hàm kiếm tra có tồn tại tệp đính kèm, sau đó tải về nếu muốn
def check_and_download_attachments(email_content, selected_email_index, save_path):
    msg = message_from_string(email_content)

    if msg.is_multipart():
        for part in walk(msg):
            if part.get_filename():
                attachment_name = part.get_filename()
                print(f"Email thứ {selected_email_index + 1} có đính kèm: {attachment_name}")

                # Hỏi người dùng có muốn tải về không
                user_response = input("Bạn có muốn tải về không? (y/n): ")

                if user_response.lower() == 'y':
                    # Lấy dữ liệu của đính kèm và giải mã base64
                    attachment_data = part.get_payload(decode=True)
                    attachment_data_base64 = base64.b64encode(attachment_data).decode('utf-8')

                    # Ghi dữ liệu vào tệp mới
                    attachment_path = os.path.join(save_path, attachment_name)
                    with open(attachment_path, 'wb') as attachment_file:
                        attachment_file.write(base64.b64decode(attachment_data_base64))

                    print(f"Đã tải về đính kèm và lưu vào tệp {attachment_path}")

# Hàm để tự động tải emails về sau mỗi khoảng thời gian
def auto_load(user):
    ssl_socket= initiate("127.0.0.1",3335)
    save_directory = r'C:/Users/Thuc Do Huu/OneDrive - VNU-HCMUS/Desktop/{user}'
    #Tạo đường dẫn tới folder của người dùng
    
    while exit_thread_flag:
        retrieve_email_with_attachment(ssl_socket, save_directory, user)
        time.sleep(10)

#Hàm chạy menu
def main_menu(username, password):
    global exit_thread_flag
    ssl_socket = initiate("127.0.0.1", 3335)
    save_directory = r'C:\Users\Thuc Do Huu\OneDrive - VNU-HCMUS\Desktop'
    
    
    # Tạo thư mục mang tên người dùng ngoài desktop
    user_folder_name = username
    user_folder_path = os.path.join(save_directory, user_folder_name)
    
    # Kiểm tra xem thư mục đã tồn tại chưa, nếu chưa thì tạo mới
    if not os.path.exists(user_folder_path):
        os.makedirs(user_folder_path)
        print(f'Thư mục cho người dùng {username} đã được tạo tại: {user_folder_path}')
    else:
        print(f'Thư mục cho người dùng {username} đã tồn tại tại: {user_folder_path}')

    # Lưu đường dẫn thư mục(của user) vào biến path_to_save
    path_to_save = user_folder_path
    
    login(ssl_socket, username, password)
    email_read_status = load_email_read_status(username, password)
    # Kiểm tra xem thông tin người dùng đã tồn tại sau khi đăng nhập hay không
    if check_user_exist(username):
        print("Tài khoản đã tạo trước đó, đăng nhập thành công")
        # In ra màn hình cấu hình từ tệp tin config
        user_config = load_user_config(username)
        print("Thông tin người dùng(Lấy từ file config):")
        for key, value in user_config.items():
            print(f"{key}: {value}")
    else:
        # Nếu tài khoản không tồn tại, lưu thông tin người dùng và thông báo thành công
        save_user_config(username, password, "127.0.0.1", 2225, 3335)
        print("Lần đầu đăng nhập")
         
    while exit_thread_flag:
        print("Vui lòng chọn Menu:")
        print("1. Để gửi email")
        print("2. Để xem danh sách các email đã nhận")
        print("3. Thoát")

        choice = input("Bạn chọn: ")

        if choice == '1':

            client_socket = initiate("127.0.0.1",2225)
            # Nhận thông tin đầu vào của người dùng để biết chi tiết email
            to_user = input("To (nếu có nhiều tài khoản thì nhập cách nhau bởi dấu phẩy và 1 khoảng trắng): ")
            if to_user and ',' in to_user:
                to = to_user.split(", ")
            else: to = [to_user]
            cc_users = input("CC (ấn Enter để bỏ qua): ")
            if cc_users and ',' in cc_users:
                cc =cc_users.split(", ")
            else: cc = [cc_users]

            bcc_users = input("BCC (ấn Enter để bỏ qua): ")
            if bcc_users and ',' in bcc_users:
                bcc =bcc_users.split(", ")
            else: bcc = [bcc_users]

            subject = input("Subject: ")
            message = input("Content: ")

            # Nhận thông tin đầu vào của người dùng cho đường dẫn đính kèm
            attachment_paths = []
            send_attachments = input("Bạn có muốn gửi tệp đính kèm không? (y/n): ").lower()

            if send_attachments == 'y':
                base_directory = r'C:\Users\Thuc Do Huu\OneDrive - VNU-HCMUS\Desktop'
                num_attachments = int(input('Số lượng tệp: '))
                      
                for i in range(1, num_attachments + 1):
                    file_name = input(f'Tên tệp thứ {i}: ')
                    if not file_name:
                        print("Tên tập tin trống. Đang thoát.")
                        break
    
                    file_path = os.path.join(base_directory, file_name)
    
                    while os.path.isfile(file_path) and os.path.getsize(file_path) > 3 * 1024 * 1024:
                        print(f"Dung lượng của file {file_name} vượt quá 3 MB. Vui lòng chọn tập tin khác!")
                        file_name = input(f'File name for attachment {i}: ')
                        if not file_name:
                            print("Tên tập tin trống. Đang thoát.")
                            break
                        file_path = os.path.join(base_directory, file_name)

                    if os.path.isfile(file_path):
                        attachment_paths.append(file_path)
                        print(f"Đã gửi file: {file_path}")
                    else:
                        print(f"Không tìm thấy file: {file_name} trong {base_directory}")
                        
                send_email_with_attachments(
                    client_socket,
                    username,
                    to_user=to,
                    cc_users=cc,
                    bcc_users=bcc,
                    subject=subject,
                    message=message,
                    attachment_paths=attachment_paths
                )
            else:
                send_email_with_attachments(
                    client_socket,
                    username,
                    to_user=to,
                    cc_users=cc,
                    bcc_users=bcc,
                    subject=subject,
                    message=message,
                    attachment_paths=attachment_paths
                )
        elif choice == '2':
            retrieve_email_with_attachment(ssl_socket, path_to_save, username)

            folders = {
                'Inbox': 'Thư mục Inbox',
                'Project': 'Thư mục Project',
                'Important': 'Thư mục Important',
                'Work': 'Thư mục Work',
                'Spam': 'Thư mục Spam' 
            }

            while True:
                print("Đây là danh sách các folder trong mailbox của bạn:")
                for index, (folder_key, folder_description) in enumerate(folders.items(), start=1):
                    print(f"Chọn {index} để truy xuất {folder_description}: {folder_key}")

                selected_folder_index = input("Bạn muốn xem email trong folder nào (nhấn enter để thoát): ")

                if selected_folder_index:
                    selected_folder_index = int(selected_folder_index) - 1
                    selected_folder_keys = list(folders.keys())

                    if 0 <= selected_folder_index < len(selected_folder_keys):
                        selected_folder = selected_folder_keys[selected_folder_index]
                        selected_folder_path = os.path.join(path_to_save, selected_folder)
                        
                        # Kiểm tra xem thư mục có tồn tại và có tập tin không
                        if not os.path.exists(selected_folder_path) or len(os.listdir(selected_folder_path)) == 0:
                            print("Thư mục không tồn tại hoặc không có email nào trong thư mục. Vui lòng chọn lại.")
                            continue
                            
                        print("Danh sách các emails trong thư mục đã chọn:")
                        show_emails_in_folder(selected_folder, email_read_status[selected_folder],path_to_save, folders)

                        while True:
                            selected_email_index = input("Bạn muốn đọc Email thứ mấy (nhấn enter để thoát, 0 để xem lại danh sách email): ")
                            selected_emails = os.listdir(selected_folder_path)

                            if selected_email_index:
                                selected_email_index = int(selected_email_index) - 1

                                if selected_email_index == -1:
                                    print("Danh sách các emails trong thư mục đã chọn:")
                                    show_emails_in_folder(selected_folder, email_read_status[selected_folder], path_to_save, folders)
                                    continue

                                if 0 <= selected_email_index < len(selected_emails):
                                    selected_email = selected_emails[selected_email_index]
                                    selected_email_folder = list(folders.keys())[selected_folder_index]
                                    
                                    # Đánh dấu email đã được đọc trong thư mục tương ứng
                                    if selected_email_folder in email_read_status:
                                        email_read_status[selected_email_folder][selected_email_index] = True

                                    with open(os.path.join(selected_folder_path, selected_email), 'r') as file:
                                        email_content = file.read()
                                        print(f"Nội dung email của email thứ {selected_email_index + 1}:")
                                        print(email_content)

                                         # Thêm đoạn mã sau để xử lý nội dung email và lưu tệp đính kèm
                                        boundary_pattern = re.compile(r'boundary="(.+)"', re.MULTILINE)
                                        match_boundary = re.search(boundary_pattern, email_content)
                                        if match_boundary:
                                            boundary = match_boundary.group(1).strip()
                                            parts = email_content.split(boundary)
                                        # Trích xuất văn bản tiêu đề và nội dung
                                        if len(parts) > 2:
                                            
                                            header = str(parts[1].split('MIME-Version: 1.0')[-1]).lstrip()
                                            body_text = str(parts[2].split('Content-Type: text/plain; charset="utf-8"')[-1]).lstrip()

                                            print("Header:")
                                            print(header)
                                            print("Body Text:")
                                            print(body_text)

                                            # Trích xuất và lưu tệp đính kèm
                                            file_names = []
                                            file_data = []

                                            for part in parts[2:]:
                                                if 'Content-Disposition: attachment;' in part:
                                                    file_part = part.split('filename=')[-1].split()
                                                    file_names.append(file_part[0])
                                                    file_data.append(file_part[1:-1])

                                            if file_names:
                                                is_download = input('Trong mail này có attached file, bạn có muốn save không (1. Có; 2. Không): ')
                                                if is_download == '1':
                                                    for i in range(len(file_names)):
                                                        base = os.getcwd()
                                                        os.chdir(path_to_save)
                                                        print(f"Đường dẫn hiện tại của chương trình: {os.getcwd()}")
                                                        save_path = input('Cho biết đường dẫn bạn muốn lưu: ')
                                                        if os.path.exists(save_path):
                                                            with open(os.path.join(save_path, file_names[i]), "wb") as file:
                                                                for data in file_data[i]:
                                                                    data = base64.b64decode(data)
                                                                    file.write(data)
                                                            print(f"File {file_names[i]} đã được lưu tại {save_path}")
                                                        else:
                                                            print(f"Đường dẫn không tồn tại: {save_path}")
                                        else:
                                            print("Emails không có tệp đính kèm")
                                         
                                else:
                                    print("Lựa chọn email không hợp lệ. Vui lòng chọn lại.")
                            else:
                                break
                    else:
                        print("Lựa chọn thư mục không hợp lệ. Vui lòng chọn lại.")
                else:
                    break
            save_email_read_status(username,password, email_read_status)

        elif choice == '3':
            
            exit_thread_flag = False
            ssl_socket.close()
            break

if __name__== "__main__":
    username = input('Nhập email của bạn: ')
    password = input('Nhập mật khẩu: ')
    thread_1 = threading.Thread (target=main_menu, args=(username, password))
    thread_2 = threading.Thread (target=auto_load, args=(username,))
    thread_1.start()
    thread_2.start()
    thread_1.join()
    thread_2.join()
