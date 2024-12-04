import base64
import json
import os
import queue
import time

import websocket
import threading
from datetime import datetime
from wsgiref.handlers import format_date_time
import hmac
import hashlib
from time import mktime
from urllib.parse import urlencode


class AssembleHeaderException(Exception):
    def __init__(self, msg):
        self.message = msg


class Url:
    def __init__(this, host, path, schema):
        this.host = host
        this.path = path
        this.schema = schema
        pass


def parse_url(requset_url):
    stidx = requset_url.index("://")
    host = requset_url[stidx + 3:]
    schema = requset_url[:stidx + 3]
    edidx = host.index("/")
    if edidx <= 0:
        raise AssembleHeaderException("invalid request url:" + requset_url)
    path = host[edidx:]
    host = host[:edidx]
    u = Url(host, path, schema)
    return u


def assemble_ws_auth_url(requset_url, method="GET", api_key="", api_secret=""):
    u = parse_url(requset_url)
    host = u.host
    path = u.path
    now = datetime.now()
    date = format_date_time(mktime(now.timetuple()))
    print(date)
    signature_origin = "host: {}\ndate: {}\n{} {} HTTP/1.1".format(host, date, method, path)
    signature_sha = hmac.new(api_secret.encode('utf-8'), signature_origin.encode('utf-8'),
                             digestmod=hashlib.sha256).digest()
    signature_sha = base64.b64encode(signature_sha).decode(encoding='utf-8')
    authorization_origin = "api_key=\"%s\", algorithm=\"%s\", headers=\"%s\", signature=\"%s\"" % (
        api_key, "hmac-sha256", "host date request-line", signature_sha)
    authorization = base64.b64encode(authorization_origin.encode('utf-8')).decode(encoding='utf-8')
    values = {
        "host": host,
        "date": date,
        "authorization": authorization
    }
    return requset_url + "?" + urlencode(values)


class WebsocketDemo:

    def __init__(self, appId, apiKey, apiSecret):
        self.requestUrl = "wss://ws-api.xf-yun.com/v1/private/ma008db16"
        self.appId = appId
        self.apiSecret = apiSecret
        self.queue = queue.Queue()
        self.result_type = result_type
        onOpen = lambda ws: self.__onOpen(ws)
        onMessage = lambda ws, msg: self.__onMessage(ws, msg)
        onError = lambda ws, err: self.__onFail(ws, err)
        onClose = lambda ws: self.__onClose(ws)
        self.requestUrl = assemble_ws_auth_url(self.requestUrl, api_key=apiKey, api_secret=apiSecret)
        ws = websocket.WebSocketApp(self.requestUrl, on_message=onMessage, on_error=onError, on_close=onClose,
                                    on_open=onOpen)
        self.ws = ws
        self.t = threading.Thread(target=self.ws.run_forever)
        self.t.start()
        self.all_texts = []

    def startSendMessage(self, file_path, output_file_path):
        with open(file_path, 'rb') as file:
            buf = file.read()
        if not buf:
            print(f"File {file_path} is empty.")
            return

        body = {
            "header": {
                "app_id": self.appId,
                "status": 2,
            },
            "parameter": {
                "s15282f39": {
                    "category": "ch_en_public_cloud",
                    "result": {
                        "encoding": "utf8",
                        "compress": "raw",
                        "format": "plain"
                    }
                },
                "s5eac762f": {
                    "result_type": self.result_type,
                    "result": {
                        "encoding": "utf8",
                        "compress": "raw",
                        "format": "plain"
                    }
                }
            },
            "payload": {
                "test": {
                    "encoding": "png",
                    "image": str(base64.b64encode(buf), 'utf-8'),
                    "status": 3
                }
            }
        }
        paramStr = json.dumps(body)
        self.queue.put((paramStr, output_file_path))

    def __onOpen(self, ws):
        print("onOpen")
        run = lambda: self.start()
        t = threading.Thread(target=run)
        t.start()

    def __onMessage(self, ws, message):
        print("onMessage", message)
        message = json.loads(message)
        if message["header"]["status"] == 1:
            text = message["payload"]["result"]["text"]
            text_de = base64.b64decode(text)

            if result_type == "0":
                file_extension = ".xls"
            elif result_type == "1":
                file_extension = ".docx"
            elif result_type == "2":
                file_extension = ".pptx"

            with open(self.current_output_file_path + file_extension, 'ab') as file:
                file.write(text_de)
            print(f"文件已保存至：{self.current_output_file_path}{file_extension}")

    def __onFail(self, ws, err):
        pass

    def __onClose(self, ws):
        print("***onClose***")
        pass

    def start(self):
        while True:
            item = self.queue.get()
            if item == 4:
                self.ws.close()
                break
            frame, self.current_output_file_path = item
            self.ws.send(frame)
            print("start send message...")


if __name__ == '__main__':
    appId = "9462f3a0"
    apiSecret = "YzJhNWM4M2Q3YzcxODI3MWEyMTkxMGRh"
    apiKey = "d01daae6308f021ac9b5282ca802fdf2"
    root_folder_path = r"D:/Desktop/test1"  # 一级文件夹路径
    output_root_path = r"D:/Desktop/test2"  # 自定义输出路径
    result_type = "1"

    for folder_name in os.listdir(root_folder_path):
        folder_path = os.path.join(root_folder_path, folder_name)
        if os.path.isdir(folder_path):
            output_dir = os.path.join(output_root_path, folder_name)
            os.makedirs(output_dir, exist_ok=True)

            for file_name in os.listdir(folder_path):
                if file_name.endswith('.png'):
                    file_path = os.path.join(folder_path, file_name)
                    output_file_path = os.path.join(output_dir, file_name.split('.')[0])
                    demo = WebsocketDemo(appId, apiKey, apiSecret)
                    demo.startSendMessage(file_path, output_file_path)
                    time.sleep(1)
