from datetime import datetime
from time import sleep
import tkinter as tk
import time
import cv2
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from PIL import Image, ImageTk
from pyzbar.pyzbar import decode


# 2つのAPIを記述しないとリフレッシュトークンを3600秒毎に発行し続けなければならない
scope = ['https://www.googleapis.com/auth/spreadsheets',
         'https://www.googleapis.com/auth/drive']
# ダウンロードしたjsonファイル名をクレデンシャル変数に設定。
credentials = ServiceAccountCredentials.from_json_keyfile_name(
    ".\\reservesheetvaccine-354009-23bb619165dc.json", scopes=scope)
# OAuth2の資格情報を使用してGoogle APIにログイン。
gc = gspread.authorize(credentials)

# スプレッドシートIDを変数に格納する。
SPREADSHEET_KEY = '1XDuxieFGagg8jkvIqL7Rk1Gc-XFwXDklUn9afwXjQKU'
# スプレッドシート（ブック）を開く
worksheet = gc.open_by_key(SPREADSHEET_KEY).worksheet('data3')

cap = cv2.VideoCapture(1, cv2.CAP_DSHOW)
font = cv2.FONT_HERSHEY_SIMPLEX
# BUF_FILE_PATH = r'.\buf.txt'
barcodes = []

main_win = tk.Tk()
# 最前面表示
main_win.attributes("-topmost", True)
main_win.title("集団接種会場 受付事務 QR読取")
main_win.geometry("480x400")
# メインフレーム
main_frm = tk.Frame(main_win)
main_frm.pack(anchor='center', padx=0, pady=20)

CANVAS_X = 420
CANVAS_Y = 280
# Canvas作成
canvas = tk.Canvas(main_frm, width=CANVAS_X, height=CANVAS_Y)
canvas.pack(anchor='center', expand=1)

# global_value = tk.StringVar()
# global_value.set('ここに通番と予約者氏名が表示されます')

txRsvName = tk.Label(font=('MS Gothic', 14),
                     text=u'ここに通番と予約者氏名が表示されます', foreground='#000000')
txRsvName.place(x=20, y=320)


def cam_reload():
    cv2.destroyAllWindows()
    main_win.quit()
    time.sleep(1)
    txRsvName['text'] = 'リロード完了'
    main_win.mainloop()


cmdReload = tk.Button(main_win, text="Reload", command=cam_reload)
cmdReload.place(x=420, y=360)
# VideoCaptureの引数にカメラ番号を入れる。
# デフォルトでは0、ノートPCの内臓Webカメラは0、別にUSBカメラを接続した場合は1を入れる。
# cap = cv2.VideoCapture(0)


def change_label_text():
    txRsvName['font'] = ('MS Gothic', 14)
    txRsvName['text'] = '次の方の読み込みができます。'
    txRsvName['foreground'] = '#000000'


def show_frame():
    global CANVAS_X, CANVAS_Y

    ret, frame = cap.read()

    image_rgb = cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)  # BGRなのでRGBに変換
    image_pil = Image.fromarray(image_rgb)  # RGBからPILフォーマットへ変換
    image_tk = ImageTk.PhotoImage(image_pil)  # ImageTkフォーマットへ変換
    # image_tkがどこからも参照されないとすぐ破棄される。
    # そのために下のようにインスタンスを作っておくかグローバル変数にしておく
    canvas.image_tk = image_tk
    # global image_tk

    # ImageTk 画像配置　画像の中心が指定した座標x,yになる
    canvas.create_image(CANVAS_X / 2, CANVAS_Y / 2, image=image_tk)
    # Canvasに現在の日時を表示
    # now_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    # canvas.create_text(CANVAS_X / 2, 30, text=now_time,
    #                   font=("Helvetica", 14, "bold"), fill="yellow")
    if ret:
        # 画像からQRコードを読み取る
        decoded_objs = decode(frame)

        # pyzbar.decode(frame)の返値
        # ('Decoded', ['data', 'type', 'rect', 'polygon'])
        # [0][0]->.data, [0][1]->.type, [0][2]->rect, [0][3]->polygon

        # 配列要素がある場合
        if decoded_objs != []:
            # [0][0]->.data, [0][1]->.type, [0][2]->rect
            # example
            # for obj in decoded_objs:
            #     print('Type: ', obj.type)
            str_dec_obj = decoded_objs[0][0].decode('utf-8', 'ignore')
            # 接種券番号のみにする
            codeData = str(int(str_dec_obj[9:]))
            # print('QR cord: {}'.format(str_dec_obj))
            left, top, width, height = decoded_objs[0][2]
            # 取得したQRコードの範囲を描画
            canvas.create_rectangle(left-110, top-100, left-110 + width,
                                    top-100 + height, outline="green", width=5)
    ###########Time()を導入するかどうか###########################
            if codeData not in barcodes:
                barcodes.append(codeData)
                # 取得したQRの内容を表示
                canvas.create_text(left-110 + (width / 2), top-120,
                                   text=codeData, font=("Helvetica", 18, "bold"))

                columns = worksheet.col_values(4)
                for i, column in enumerate(columns):
                    global global_value
                    if column == codeData:
                        # print(f'値:{column}、列のindex:{4}、行のindex:{i+1}')
                        dt = datetime.now()
                        stime = f'{dt.hour}:{dt.minute}:{dt.second}'
                        worksheet.update_cell(i+1, 13, stime)
                        worksheet.update_cell(i+1, 12, '受付済')
                        global_value = "通番:" + worksheet.cell(i+1, 1).value + "   券No:" + worksheet.cell(i+1, 4).value + "\n 氏名:" + worksheet.cell(
                            i+1, 5).value
                        txRsvName['font'] = ('MS Gothic', 20)
                        txRsvName['text'] = global_value
                        txRsvName['foreground'] = '#0000ff'
                        main_win.after(5000, change_label_text,)
            #         else:
            #             txRsvName['text'] = 'リストにありません'
            #             txRsvName['foreground'] = '#800000'
            # else:
            #     canvas.create_text(left-110 + (width / 2), top-120,
            #                        text="既に読み込まれています " + codeData, font=("Helvetica", 18, "bold"))
            #     txRsvName['font'] = ('MS Gothic', 14)
            #     txRsvName['text'] = '既に読み込まれています'
            #     txRsvName['foreground'] = '#800000'

        # decoded_objs.clear()
        # QRコードを取得して、その内容をTextに書き出し、そのままTKのプログラムを終了するコード
        # with open('QR_read_data.txt', 'w') as exportFile:
        #    exportFile.write(str_dec_obj)
        # sleep(1)
        # cap.release()
        # root.quit()
    # else:
    #     txRsvName['font'] = ('MS Gothic', 14)
    #     txRsvName['text'] = 'カメラから画像を取得できませんでした'
    #     txRsvName['foreground'] = '#800000'
    # 10msごとにこの関数を呼び出す
    canvas.after(5, show_frame)


show_frame()
main_win.mainloop()
