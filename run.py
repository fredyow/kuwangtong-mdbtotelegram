import ctypes
import pypyodbc
import requests
import tkinter as tk
from tkinter import filedialog, simpledialog
import configparser

CONFIG_FILE = "config.ini" #用以记录telebot id ,mdf文件

def read_latest_sms(cursor, last_processed_id):
    query = f"SELECT TOP 1 id, pcui, simnum, content FROM L_SMS WHERE id > {last_processed_id} ORDER BY id DESC" #获取MDB table里面的pcui, simnum, content最后一行
    cursor.execute(query)
    row = cursor.fetchall()  # 使用 fetchall() 获取所有结果

    return row[-1] if row else None  # 返回最后一行记录或 None

def send_to_telegram(message, bot_token, chat_id): #发送至telegram
    api_url = f'https://api.telegram.org/bot{bot_token}/sendMessage'
    params = {'chat_id': chat_id, 'text': message}

    response = requests.post(api_url, params=params)

def open_database():
    root = tk.Tk()
    root.withdraw()  # 隐藏主窗口

    mdb_path = filedialog.askopenfilename(title="选择MDB文件", filetypes=[("MDB files", "*.mdb")])
    password = simpledialog.askstring("密码", "请输入密码", show='*')

    return mdb_path, password

def get_bot_info(): #提示输入bottoken , Chat ID
    bot_token = simpledialog.askstring("Bot Token", "请输入Telegram Bot Token：")
    chat_id = simpledialog.askstring("Chat ID", "请输入Telegram Chat ID：")

    return bot_token, chat_id

def save_config(mdb_path, password, bot_token, chat_id): #储存config
    config = configparser.ConfigParser()
    config['DATABASE'] = {'MDBPath': mdb_path, 'Password': password}
    config['TELEGRAM'] = {'BotToken': bot_token, 'ChatID': chat_id}

    with open(CONFIG_FILE, 'w') as configfile:
        config.write(configfile)

def load_config():
    config = configparser.ConfigParser()
    config.read(CONFIG_FILE)

    mdb_path = config.get('DATABASE', 'MDBPath', fallback='')
    password = config.get('DATABASE', 'Password', fallback='')
    bot_token = config.get('TELEGRAM', 'BotToken', fallback='')
    chat_id = config.get('TELEGRAM', 'ChatID', fallback='')

    return mdb_path, password, bot_token, chat_id

if __name__ == "__main__":
    mdb_path, password, bot_token, chat_id = load_config()

    if not mdb_path or not password:
        mdb_path, password = open_database()
        save_config(mdb_path, password, bot_token, chat_id)

    if not bot_token or not chat_id:
        bot_token, chat_id = get_bot_info()
        save_config(mdb_path, password, bot_token, chat_id)
        
    custom_text = "by @fred_mb" #定制窗口文字
    ctypes.windll.kernel32.SetConsoleTitleW(f"{bot_token[:10]}-{custom_text}")  # 设置窗口标题
            
        
    conn = pypyodbc.connect('DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=' + mdb_path + ';PWD=' + password)
    cursor = conn.cursor()  # 创建 cursor 对象
    
    last_processed_id = 0
    
    while True:
        try:
            row = read_latest_sms(cursor, last_processed_id) 
            
            if row:
                record_id, pcui, simnum, content = row
                hidden_simnum = simnum[:len(simnum)-8] + '****' + simnum[len(simnum)-4:]
                message = f"PORT:{pcui}\nNumber:{hidden_simnum}\nSMS:{content}"
                print(message)
                send_to_telegram(message, bot_token, chat_id)

                last_processed_id = record_id  # 更新已处理的记录ID
                
        except KeyboardInterrupt:
            print("Ctrl+C pressed. Exiting...")
            break
        except Exception as e:
            print(f"An error occurred: {e}")
