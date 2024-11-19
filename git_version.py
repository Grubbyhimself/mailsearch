#通过IMAP协议与邮件服务器进行通信，以获取邮件和附件信息
#目标是从指定的邮箱中检索特定名称的邮件，并下载其附件。应用场景包括自动化邮件处理、邮件监控等。
import logging
import imaplib
import email
import os
import configparser
import sqlite3
from datetime import datetime
import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, simpledialog

# 配置文件路径
CONFIG_FILE = './config.ini'

# 读取配置文件
con = configparser.ConfigParser()
con.read(CONFIG_FILE, encoding='utf-8')
mail_box = dict(con.items('mail_box'))  # 读取邮件相关的配置
box_list = dict(con.items('box_list'))  # 读取邮箱列表的配置
principal = dict(con.items('principal'))  # 读取主要用户的配置

# 配置日志
log_level = logging.INFO if mail_box['log_level'] == 'INFO' else logging.DEBUG
logger = logging.getLogger('my_logger')
logger.setLevel(log_level)

# 创建一个handler，用于写入日志文件，使用UTF-8编码
file_handler = logging.FileHandler('./run.log', 'a', 'utf-8')
file_handler.setLevel(log_level)

# 创建一个formatter
formatter = logging.Formatter('%(asctime)s - line:[%(lineno)d] - %(levelname)s: %(message)s')

# 给handler设置formatter
file_handler.setFormatter(formatter)

# 给logger添加handler
logger.addHandler(file_handler)

# 如果是DEBUG模式，打印一条调试信息
if log_level == logging.DEBUG:
    logger.debug('您已进入测试环境中，请留意')


# 初始化数据库
def init_db():
    """
    初始化数据库，创建表（如果表不存在）。
    """
    conn = sqlite3.connect('attachments.db')  # 连接数据库
    cursor = conn.cursor()
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS attachments (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            filename TEXT NOT NULL,
            download_time TEXT NOT NULL,
            email_title TEXT NOT NULL
        )
    ''')  # 创建表，如果表不存在
    conn.commit()  # 提交事务
    conn.close()  # 关闭数据库连接


import threading

# 创建一个全局的锁
db_lock = threading.Lock()

def save_attachment_info(filename, email_title):
    """
    将附件信息保存到数据库中。
    """
    if not email_title:  # 如果 email_title 为空，则提供一个默认值
        email_title = "无标题"

    conn = sqlite3.connect('attachments.db')  # 连接数据库
    cursor = conn.cursor()
    download_time = datetime.now().strftime('%Y-%m-%d %H:%M:%S')  # 获取当前时间

    # 使用锁确保同一时刻只有一个线程可以写入数据库
    with db_lock:
        cursor.execute('''INSERT INTO attachments (filename, download_time, email_title)
                          VALUES (?, ?, ?)''', (filename, download_time, email_title))  # 插入数据
        conn.commit()  # 提交事务
    conn.close()  # 关闭数据库连接




# 获取邮件标题的解码信息
def Get_title(message):
    """
    解码邮件标题，返回解码后的标题。
    """
    subject_encoded, charset = email.header.decode_header(message.get('Subject'))[0]
    if isinstance(subject_encoded, bytes):
        # 如果subject_encoded是字节序列，尝试使用指定的编码进行解码，如果失败则忽略错误
        title = subject_encoded.decode(charset if charset else 'utf-8', errors='ignore')
    else:
        # 如果subject_encoded已经是字符串，则直接使用
        title = subject_encoded
    return title


# 判断目录是否存在，如果不存在则创建
def Judge_folder(dir):
    """
    检查目录是否存在，如果不存在则创建该目录。
    """
    if os.path.exists(dir):
        return
    logging.info(f'【{dir}】目录不存在，已进行创建')
    os.mkdir(dir)


# 生成唯一文件名，避免文件覆盖
def get_unique_filename(directory, filename):
    """
    生成唯一的文件名，避免文件覆盖。
    """
    base, ext = os.path.splitext(filename)
    counter = 1
    unique_filename = filename
    while os.path.exists(os.path.join(directory, unique_filename)):
        unique_filename = f"{base}_{counter}{ext}"
        counter += 1
    return unique_filename


# 下载单个附件的函数
def download_attachment(attachment, dir, message):
    """
    下载单个附件并保存到指定目录。
    """
    filename_unchar = attachment.get_filename()
    if not filename_unchar:
        return  # 如果没有文件名，跳过
    fne = email.header.decode_header(filename_unchar)
    filename = fne[0][0].decode(fne[0][1]) if fne[0][1] else fne[0][0]
    # 生成唯一文件名
    unique_filename = get_unique_filename(dir, filename)
    logging.info(f'获取到附件[{filename}]，正在尝试下载')
    attach_data = attachment.get_payload(decode=True)
    # 保存附件到文件
    with open(os.path.join(dir, unique_filename), 'wb') as f:
        f.write(attach_data)
    logging.info(f'[{unique_filename}]附件下载成功!')

    # 获取邮件标题
    email_title = Get_title(message)

    # 保存附件信息到数据库
    save_attachment_info(unique_filename, email_title)

    # 输出附件信息
    print(f"文件名: {unique_filename}, 邮件标题: {email_title}, "
          f"下载时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}, 存储目录: {dir}")


# 修改 Get_file 函数，使用线程下载附件
def Get_file(message, locate):
    """
    下载邮件中的附件，并使用多线程并行下载。
    """
    dir = "./" + locate + "/"
    Judge_folder(dir)  # 确保目录存在
    attachments_downloaded = 0  # 用于计数下载的附件数量
    threads = []  # 用于存储线程
    for part in message.walk():
        if part.get_content_maintype() == 'multipart':
            continue
        if part.get('Content-Disposition') is None:
            continue
        # 为每个附件创建并启动线程
        thread = threading.Thread(target=download_attachment, args=(part, dir, message))
        threads.append(thread)
        thread.start()
        attachments_downloaded += 1
    # 等待所有下载线程完成
    for thread in threads:
        thread.join()
    logging.info(f'邮件中共找到{attachments_downloaded}个附件')
    if attachments_downloaded == 0:
        logging.info('未获取到附件')



# 设置邮件的已读/未读状态
def Set_flags(uid, conn, status, title):
    """
    设置邮件的已读或未读状态。
    """
    if status == '已读':
        s = '+'
    else:
        s = '-'

    typ, _ = conn.store(uid, s + 'FLAGS', '\\Seen')
    if typ == 'OK':
        logging.debug(f'已经将邮件【{title}】标记为{status}')


# 登录邮箱
def Login():
    """
    连接到IMAP服务器并登录邮箱。
    """
    imaplib.Commands['ID'] = ('AUTH')
    conn = imaplib.IMAP4_SSL(mail_box['mail_ssl'], mail_box['mail_ssl_port'])  # 连接IMAP服务器
    logging.info(f'已连接服务器{mail_box["mail_ssl"]}')
    conn.login(mail_box['mail_user'], mail_box['mail_password'])  # 登录邮箱
    logging.info(f'已登陆{mail_box["mail_user"]}账户')
    args = ("name", "huangjiajun", "contact", mail_box['mail_user'], "version", "1.0.0", "vendor", "myclient")
    conn._simple_command('ID', '("' + '" "'.join(args) + '")')
    return conn


# 获取邮箱列表
def BoxList(conn):
    """
    获取邮箱列表并记录日志。
    """
    for i in conn.list()[1]:
        logging.debug(i)


# 处理邮件
def handle_mail(mail_boxs, conn, search_filename, log_text):
    """
    处理指定邮箱中的邮件，搜索并下载符合条件的附件。
    """
    conn.select(mailbox=mail_boxs, readonly=False)
    typ, num = conn.search(None, mail_box['read_mail'])
    if typ == 'OK':
        for uid in num[0].split():
            typ, data = conn.fetch(uid, "(RFC822)")
            if typ == 'OK':
                text = data[0][1].decode("utf-8", errors='ignore')
                message = email.message_from_string(text)
                title = Get_title(message)
                log_text.insert(tk.END, f'获取到主题为【{title}】的邮件\n')
                log_text.see(tk.END)

                if "信息通报" in title:
                    log_text.insert(tk.END, '找到日报附件，进行尝试下载附件\n')
                    log_text.see(tk.END)
                    if "滔博" in title:
                        locate = "TB"
                        Get_file(message, locate)
                        Set_flags(uid, conn, '已读', title)
                    elif "皇族" in title:
                        locate = "UZI"
                        Get_file(message, locate)
                        Set_flags(uid, conn, '已读', title)
                    else:
                        locate = "Cool"
                        Get_file(message, locate)
                        Set_flags(uid, conn, '已读', title)

                elif "正午报" in title and "信息通报" not in title:
                    log_text.insert(tk.END, '找到日报附件，进行尝试下载附件\n')
                    log_text.see(tk.END)
                    if "宙斯2" in title:
                        locate = "ZS2"
                        Get_file(message, locate)
                        Set_flags(uid, conn, '已读', title)
                    elif "宙斯" in title:
                        locate = "ZS"
                        Get_file(message, locate)
                        Set_flags(uid, conn, '已读', title)
                    else:
                        locate = "Others"
                        Get_file(message, locate)
                        Set_flags(uid, conn, '已读', title)

                elif search_filename in title:
                    locate = "Others"
                    Get_file(message, locate)
                    Set_flags(uid, conn, '已读', title)

                else:
                    log_text.insert(tk.END, '邮件非日报邮件，将其重新标记为未读\n')
                    log_text.see(tk.END)
                    Set_flags(uid, conn, '未读', title)
            else:
                log_text.insert(tk.END, f"邮件编号'{uid}'信息获取失败，请重试！\n")
                log_text.see(tk.END)


# 周期性处理邮件
def handle_mail_periodically(mail_boxs, conn, search_filename, log_text, stop_event):
    """
    周期性地处理邮件，直到stop_event被设置。
    """
    while not stop_event.is_set():  # 当stop_event未设置时，持续运行
        handle_mail(mail_boxs, conn, search_filename, log_text)
        log_text.insert(tk.END, "等待10s后进行下一次检索...\n")
        log_text.see(tk.END)
        stop_event.wait(10)  # 等待10s或直到stop_event被设置


# 创建主窗口
def create_main_window():
    """
    创建并配置主窗口，包括所有控件。
    """
    root = tk.Tk()
    root.title("邮件附件管理器")
    root.geometry("600x400")

    # 创建一个Frame来容纳所有的控件
    main_frame = ttk.Frame(root)
    main_frame.grid(row=0, column=0, sticky="nsew")

    # 配置行和列权重，使控件在窗口大小变化时居中
    root.grid_rowconfigure(0, weight=1)
    root.grid_columnconfigure(0, weight=1)

    # 添加必要的控件
    search_label = ttk.Label(main_frame, text="请输入你想要检索的附件邮件名：")
    search_label.grid(row=0, column=0, columnspan=2, pady=10, sticky="ew")

    search_entry = ttk.Entry(main_frame, width=50)
    search_entry.grid(row=1, column=0, columnspan=2, pady=10, sticky="ew")

    start_button = ttk.Button(main_frame, text="开始检索",
                              command=lambda: start_search(root, search_entry.get(), start_button, stop_button))
    start_button.grid(row=2, column=0, pady=10, sticky="ew")

    stop_button = ttk.Button(main_frame, text="停止检索", command=lambda: stop_event.set())
    stop_button.grid(row=2, column=1, pady=10, sticky="ew")

    delete_button = ttk.Button(main_frame, text="删除数据", command=lambda: delete_data(log_text))
    delete_button.grid(row=3, column=0, pady=10, sticky="ew")

    delete_db_button = ttk.Button(main_frame, text="删除整个数据库", command=delete_entire_database)
    delete_db_button.grid(row=3, column=1, pady=10, sticky="ew")

    query_db_button = ttk.Button(main_frame, text="查询数据库", command=lambda: display_query_result(log_text))
    query_db_button.grid(row=4, column=0, pady=10, sticky="ew")

    quit_button = ttk.Button(main_frame, text="退出", command=root.quit)
    quit_button.grid(row=4, column=1, pady=10, sticky="ew")

    # 添加一个文本框用于显示日志信息
    log_text = tk.Text(main_frame, height=10, width=70)
    log_text.grid(row=5, column=0, columnspan=2, pady=10, sticky="nsew")

    # 配置Frame的行和列权重，使日志文本框在窗口大小变化时扩展
    main_frame.grid_rowconfigure(5, weight=1)
    main_frame.grid_columnconfigure((0, 1), weight=1)

    # 初始化 stop_event
    stop_event = threading.Event()

    return root, log_text, stop_event, start_button, stop_button


# 开始检索邮件
def start_search(root, search_filename, start_button, stop_button):
    """
    开始检索邮件，启动多线程处理邮件。
    """
    if not search_filename.strip():
        messagebox.showerror("输入错误", "请提供有效的邮件名！")
        return

    # 更新GUI状态，例如禁用按钮
    start_button.config(state=tk.DISABLED)
    stop_button.config(state=tk.NORMAL)

    # 重置 stop_event
    stop_event.clear()

    # 开始邮件处理流程
    def process_emails():
        try:
            init_db()
            conn = Login()
            BoxList(conn)

            # 创建线程列表
            threads = []

            # 为每个邮箱创建并启动线程
            for mail_boxs in box_list.values():
                thread = threading.Thread(target=handle_mail_periodically,
                                          args=(mail_boxs, conn, search_filename, log_text, stop_event))
                threads.append(thread)
                thread.start()

            # 等待所有线程完成
            for thread in threads:
                thread.join()

            # 检索完成后，询问用户是否查询数据库
            user_input = messagebox.askyesno("检索完成", "是否查询数据库？")
            if user_input:
                display_query_result(log_text)

            # 更新GUI状态，例如启用按钮
            start_button.config(state=tk.NORMAL)
            stop_button.config(state=tk.DISABLED)
        except Exception as e:
            messagebox.showerror("错误", f"发生错误: {e}")
            start_button.config(state=tk.NORMAL)
            stop_button.config(state=tk.DISABLED)

    # 在新的线程中执行邮件处理，防止阻塞UI
    threading.Thread(target=process_emails).start()


# 删除数据
def delete_data(log_text):
    """
    删除数据库中的指定记录。
    """
    # 显示一个输入框，让用户输入要删除的记录 ID 或邮件标题
    input_value = simpledialog.askstring("删除数据", "请输入要删除的记录 ID 或邮件标题：")
    if not input_value:
        messagebox.showerror("输入错误", "请输入有效的记录 ID 或邮件标题！")
        return

    conn = sqlite3.connect('attachments.db')
    cursor = conn.cursor()

    # 尝试按记录 ID 删除
    cursor.execute('DELETE FROM attachments WHERE id = ?', (input_value,))
    if cursor.rowcount == 0:
        # 如果按记录 ID 删除失败，尝试按邮件标题删除
        cursor.execute('DELETE FROM attachments WHERE email_title = ?', (input_value,))
        if cursor.rowcount == 0:
            messagebox.showerror("删除失败", "没有找到匹配的记录！")
        else:
            messagebox.showinfo("删除成功", f"已删除 {cursor.rowcount} 条记录。")
    else:
        messagebox.showinfo("删除成功", f"已删除 {cursor.rowcount} 条记录。")

    conn.commit()
    conn.close()

    log_text.insert(tk.END, f"已删除记录：{input_value}\n")
    log_text.see(tk.END)


# 删除整个数据库
def delete_entire_database():
    """
    删除整个数据库文件并重新初始化。
    """
    # 确认用户是否真的要删除整个数据库
    user_input = messagebox.askyesno("确认删除", "确定要删除整个数据库吗？此操作无法撤销！")
    if user_input:
        # 删除数据库文件
        db_path = 'attachments.db'
        if os.path.exists(db_path):
            os.remove(db_path)
            messagebox.showinfo("删除成功", "数据库已成功删除。")
        else:
            messagebox.showwarning("警告", "数据库文件不存在。")

        # 重新初始化数据库
        init_db()
        messagebox.showinfo("初始化成功", "数据库已重新初始化。")


# 显示查询结果
def display_query_result(log_text):
    """
    查询数据库并显示结果。
    """
    result = query_db()
    log_text.insert(tk.END, result)
    log_text.see(tk.END)


# 查询并显示数据库内容
def query_db():
    """
    查询数据库中的所有记录并返回结果。
    """
    conn = sqlite3.connect('attachments.db')
    cursor = conn.cursor()
    cursor.execute('SELECT * FROM attachments')
    rows = cursor.fetchall()
    if rows:
        result = "查询结果：\n"
        for row in rows:
            result += f"ID: {row[0]}, Filename: {row[1]}, Download Time: {row[2]}, Email Title: {row[3]}\n"
    else:
        result = "数据库中没有记录。\n"
    conn.close()
    return result


if __name__ == '__main__':
    root, log_text, stop_event, start_button, stop_button = create_main_window()
    root.mainloop()



"""
#配置多个邮箱样例config
[mail_box1]
mail_ssl = imap.example1.com
mail_ssl_port = 993
mail_user = user1@example1.com
mail_password = password1
log_level = INFO
read_mail = (UNSEEN)

[mail_box2]
mail_ssl = imap.example2.com
mail_ssl_port = 993
mail_user = user2@example2.com
mail_password = password2
log_level = INFO
read_mail = (UNSEEN)

[mail_box3]
mail_ssl = imap.example3.com
mail_ssl_port = 993
mail_user = user3@example3.com
mail_password = password3
log_level = INFO
read_mail = (UNSEEN)

[box_list]
inbox1 = inbox
inbox2 = inbox2
inbox3 = inbox3


# 读取多个邮箱的配置python
mail_boxes = {
    'mail_box1': dict(con.items('mail_box1')),
    'mail_box2': dict(con.items('mail_box2')),
    'mail_box3': dict(con.items('mail_box3'))
}

"""