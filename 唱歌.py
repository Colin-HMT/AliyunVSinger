import base64
import csv
import hashlib
import hmac
import time
import uuid
from urllib import parse
import requests
import threading
import nls
import tkinter as tk
from tkinter import filedialog
import pandas as pd
from pydub import AudioSegment
from pydub.silence import detect_nonsilent
import shutil
import os


def encode_text(text):
    encoded_text = parse.quote_plus(text)
    return encoded_text.replace('+', '%20').replace('*', '%2A').replace('%7E', '~')


def encode_dict(dic):
    keys = dic.keys()
    dic_sorted = [(key, dic[key]) for key in sorted(keys)]
    encoded_text = parse.urlencode(dic_sorted)
    return encoded_text.replace('+', '%20').replace('*', '%2A').replace('%7E', '~')


def GetTokenFromFile(FileFullName):
    with open(FileFullName, 'r') as file:
        csv_reader = csv.reader(file)
        data = list(csv_reader)
        AccessKeyID = data[1][0]
        AccessKeySecret = data[1][1]
    parameters = {
        'AccessKeyId': AccessKeyID,
        'Action': 'CreateToken',
        'Format': 'JSON',
        'RegionId': 'cn-shanghai',
        'SignatureMethod': 'HMAC-SHA1',
        'SignatureNonce': str(uuid.uuid1()),
        'SignatureVersion': '1.0',
        'Timestamp': time.strftime("%Y-%m-%dT%H:%M:%SZ", time.gmtime()),
        'Version': '2019-02-28'
    }

    query_string = encode_dict(parameters)
    string_to_sign = 'GET' + '&' + encode_text('/') + '&' + encode_text(query_string)
    secreted_string = hmac.new(bytes(AccessKeySecret + '&', encoding='utf-8'),
                               bytes(string_to_sign, encoding='utf-8'),
                               hashlib.sha1).digest()
    signature = base64.b64encode(secreted_string)
    signature = encode_text(signature)
    full_url = 'http://nls-meta.cn-shanghai.aliyuncs.com/?Signature=%s&%s' % (signature, query_string)
    print('URL: %s' % full_url)
    response = requests.get(full_url)
    if response.ok:
        root_obj = response.json()
        key = 'Token'
        if key in root_obj:
            token = root_obj[key]['Id']
            return token
    return None, None


# 以下代码会根据上述TEXT文本反复进行语音合成
def test_run(tid, filefullname, appkey, final_folder, text, tune, length, voice, fmt):
    with open(final_folder+text+"."+fmt, "wb") as f:
        tts = nls.NlsSpeechSynthesizer(
            url="wss://nls-gateway-cn-shanghai.aliyuncs.com/ws/v1",
            token=GetTokenFromFile(FileFullName=filefullname),
            appkey=appkey,
            on_metainfo=lambda message, *args: print("on_metainfo message=>{}".format(message)),
            on_data=lambda data, *args: f.write(data),
            on_completed=lambda message, *args: print("on_completed:args=>{} message=>{}".format(args, message)),
            on_error=lambda message, *args: print("on_error args=>{}".format(args)),
            on_close=lambda *args: print("on_close: args=>{}".format(args)),
            callback_args=[tid]
        )
        r = tts.start(
            text=text,
            voice=voice,
            aformat=fmt,
            pitch_rate=tune,
            speech_rate=length
        )
    f.close()
    sound = AudioSegment.from_file(final_folder+text+"."+fmt, format=fmt)
    nonsilent_chunks = detect_nonsilent(sound, silence_thresh=-32)
    start_time = nonsilent_chunks[0][0]
    end_time = nonsilent_chunks[-1][1]
    trimmed_sound = sound[start_time:end_time]
    trimmed_sound.export(final_folder+text+"."+fmt, format=fmt)


def multiruntest(num, filefullname, appkey, final_folder, text, tune, length, voice, fmt):
    threads = []
    for i in range(0, num):
        name = "thread" + str(i)
        t = threading.Thread(target=test_run, args=(name, filefullname, appkey, final_folder, text, tune, length, voice, fmt))
        t.start()
        threads.append(t)
    for t in threads:
        t.join()


def ffmpeg_file_dialog():
    file_path = filedialog.askopenfilename()
    if file_path:
        # 将选中的文件路径显示在文本框中
        ffmpeg_path_entry.delete(0, tk.END)
        ffmpeg_path_entry.insert(0, file_path)


def open_file_dialog():
    file_path = filedialog.askopenfilename()
    if file_path:
        # 将选中的文件路径显示在文本框中
        file_path_entry.delete(0, tk.END)
        file_path_entry.insert(0, file_path)


def open_folder_dialog():
    folder_path = filedialog.askdirectory()
    if folder_path:
        # 将选中的文件夹路径显示在文本框中
        output_entry.delete(0, tk.END)
        output_entry.insert(0, folder_path)


def open_notes_dialog():
    notes_path = filedialog.askopenfilename()
    if notes_path:
        # 将选中的文件路径显示在文本框中
        notes_entry.delete(0, tk.END)
        notes_entry.insert(0, notes_path)


def VoiceSyne():
    ffmpeg_file_ = ffmpeg_path_entry.get()
    source_file = ffmpeg_file_
    destination_dir = '.'
    destination_file = os.path.join(destination_dir, os.path.basename(source_file))
    shutil.copy(source_file, destination_file)
    filefullname = file_path_entry.get()
    appkey = appkey_entry.get()
    test_folder = output_entry.get()
    notes = notes_entry.get()
    voice = voice_entry.get()
    fmt = fmt_entry.get()
    df = pd.read_excel(notes, sheet_name='notes')
    for index, row in df.iterrows():
        # 读取每行中的列
        character = row['character']
        pitch = row['note_pitch']
        length = row['note_length']
        multiruntest(num=1, filefullname=filefullname, appkey=appkey, final_folder=test_folder+"/", text=character, tune=pitch, length=length, voice=voice, fmt=fmt)


window = tk.Tk()

# 设置窗口标题
window.title("我的Python窗口")

# 设置窗口大小
window.geometry("400x400")

tk.messagebox.showinfo("提示", "请阅读并删除所有文本框提示文字后再输入")
tk.messagebox.showinfo("提示", "如果当前文件夹已存在同名文件，会覆盖之前的文件")

ffmpeg_path_entry = tk.Entry(window)
ffmpeg_path_entry.insert(0, "请选择ffmpeg路径")
ffmpeg_path_entry.place(x=110, y=20, width=210, height=30)

ffmpeg_button = tk.Button(window, text="选择文件", command=ffmpeg_file_dialog)
ffmpeg_button.place(x=260, y=20, width=100, height=30)

file_path_entry = tk.Entry(window)
file_path_entry.insert(0, "请选择Accesskey路径")
file_path_entry.place(x=110, y=60, width=210, height=30)

open_file_button = tk.Button(window, text="选择文件", command=open_file_dialog)
open_file_button.place(x=260, y=60, width=100, height=30)

appkey_entry = tk.Entry(window)
appkey_entry.insert(0, "请输入AppKey")
appkey_entry.place(x=110, y=100, width=210, height=30)

output_entry = tk.Entry(window)
output_entry.insert(0, "请选择输出文件夹")
output_entry.place(x=110, y=140, width=210, height=30)

open_folder_button = tk.Button(window, text="选择文件夹", command=open_folder_dialog)
open_folder_button.place(x=260, y=140, width=100, height=30)

notes_entry = tk.Entry(window)
notes_entry.insert(0, "请选择Excel曲谱")
notes_entry.place(x=110, y=180, width=210, height=30)

notes_file_button = tk.Button(window, text="选择文件", command=open_notes_dialog)
notes_file_button.place(x=260, y=180, width=100, height=30)

voice_entry = tk.Entry(window)
voice_entry.insert(0, "请输入发音人名字")
voice_entry.place(x=110, y=220, width=210, height=30)

fmt_entry = tk.Entry(window)
fmt_entry.insert(0, "请输入输出文件格式,仅可以输出pcm/wav/mp3格式")
fmt_entry.place(x=110, y=260, width=210, height=30)

btn = tk.Button(window, text="开始合成", command=VoiceSyne)
btn.place(x=110, y=300, width=210, height=30)

window.mainloop()
