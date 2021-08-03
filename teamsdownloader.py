# -*- coding: utf-8 -*-
import asyncio
from typing import Dict
from pyppeteer import launch, page
import json
from pyppeteer.browser import Browser
from pyppeteer.element_handle import ElementHandle
import requests
import shutil
import os
import datetime
import time
import tkinter as tk
from concurrent.futures import ThreadPoolExecutor

_executor = ThreadPoolExecutor(10)


class Application(tk.Frame):
    chats: Dict
    cookie: Dict
    token: str

    def __init__(self, master=None, in_cookie: Dict = None, in_token: str = None, in_chats: Dict = None):
        super().__init__(master)
        self.chats = in_chats
        self.cookie = in_cookie
        self.token = in_token
        self.master = master
        self.pack(fill=tk.BOTH, expand=tk.YES)
        self.create_widgets()

    def create_widgets(self):
        scrollData = tk.StringVar()
        self.download_btn = tk.Button(self)
        self.chat_list = tk.Listbox(self, listvariable=scrollData)
        for c in self.chats:
            self.chat_list.insert('end', str(
                c) + ': ' + self.chats[c]["topic"])
        self.download_btn["text"] = "Download Selected"
        self.download_btn["command"] = self.download
        self.download_btn.pack(side="top")
        self.chat_list.pack(side="left", fill='y')

        self.quit = tk.Button(self, text="QUIT", fg="red",
                              command=self.master.destroy)
        self.quit.pack(side="bottom")

    def on_resize(self, event):
        # determine the ratio of old width/height to new width/height
        wscale = event.width/self.width
        hscale = event.height/self.height
        self.width = event.width
        self.height = event.height
        # rescale all the objects
        self.scale("all", 0, 0, wscale, hscale)

    def download(self):
        selected = self.chat_list.selection_get()
        print("Downloading Chat: " + selected)
        download_chat(cookie=self.cookie, token=self.token,
                      chat=self.chats[int(selected[0:1])])
# Save cookie


async def save_cookie(cookie):
    with open("cookie.json", 'w+', encoding="utf-8") as file:
        json.dump(cookie, file, ensure_ascii=False)


async def save_token(token):
    with open("token.txt", 'w+', encoding="utf-8") as file:
        file.write(token)


def download_file(url, folder, cookie):
    local_filename = url.split('/')[-1]
    with requests.get(url, stream=True, cookies=cookie) as r:
        with open(folder + '/' + local_filename, 'wb') as f:
            shutil.copyfileobj(r.raw, f)

    return local_filename

 # Read cookie


async def load_cookie():
    with open("cookie.json", 'r', encoding="utf-8") as file:
        cookie = json.load(file)
    return cookie


async def load_token():
    with open('token.txt', 'r', encoding="utf-8") as file:
        token = file.read()
    return token


async def sharepoint(page, url):
    await page.setViewport({'width': 1366, 'height': 768})
    await page.setUserAgent('Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) '
                            'Chrome/58.0.3029.110 Safari/537.36 Edge/16.16299')
    await page.goto(url)
    await page.waitForSelector('#O365_SuiteBranding_container', {'timeout': 300000})
    print('found selector')
    await asyncio.sleep(2)
    cookies = await page.cookies()
    await save_cookie(cookies)
    print(cookies)


async def graph(page: page.Page, url):
    token_element: ElementHandle
    token: str
    await page.setViewport({'width': 1366, 'height': 768})
    await page.setUserAgent('Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) '
                            'Chrome/58.0.3029.110 Safari/537.36 Edge/16.16299')
    await page.goto(url)
    await page.waitForSelector('.ms-Persona', options={'timeout': 0})
    await page.click('button[name^=Access')
    await asyncio.sleep(1)
    await page.waitForSelector('label.ms-Label:nth-child(2)', {'timeout': 300000})
    token_element = await page.querySelector('label.ms-Label:nth-child(2)')
    token = await page.evaluate('(element) => element.textContent', token_element)
    print('found token' + token)
    await save_token(token)


async def launch_browser():
    return await launch({'headless': False,
                         'dumpio': True,
                         'args': [
                             '--disable-dev-shm-usage',
                             '--shm-size=1gb'
                             '--disable-gpu',
                         ],
                         'executablePath': 'C:\Program Files\Google\Chrome\Application\chrome.exe'
                         })


async def load():
    browser: Browser = None

    minutes_15 = 900
    minutes_45 = 2700
    try:
        if time.time() - os.stat('token.txt').st_mtime <= minutes_45:
            pass
        else:
            print("the token has timed out, refreshing now")
            raise Exception
    except Exception as e:
        if not browser:
            browser = await launch_browser()
        page1 = await browser.pages()
        page1 = page1[0]
        await graph(page=page1, url='https://developer.microsoft.com/en-us/graph/graph-explorer')

    try:
        if time.time() - os.stat('cookie.json').st_mtime <= minutes_45:
            pass
        else:
            print("The cookie has timed out, refreshing now")
            raise Exception
    except Exception as e:
        if not browser:
            browser = await launch_browser()
        page2 = await browser.newPage()
        await sharepoint(page2, 'https://inoffice.sharepoint.com/')

    cookie = await load_cookie()
    req_cookies = {}
    for entry in cookie:
        req_cookies[entry['name']] = entry['value']
    token = await load_token()

    try:
        await browser.close()
    except:
        pass
    return [req_cookies, token]


async def load_chats(token):
    chats = {}
    _headers = {'Authorization': 'Bearer ' + token}
    data = requests.get(
        'https://graph.microsoft.com/beta/me/chats', headers=_headers).json()
    i = 1
    for v in data["value"]:
        #print(str(i) + ': ' + (v['chatType'] or "No Chat type") + ' ::: ' + (v['topic'] or "No Topic") + ' - ' + (v['id'] or "No ID"))
        chats[i] = {'id': v['id'], 'topic': (v['topic'] or "No_Topic"), 'chat_type': (
            v['chatType'] or "No Chat type"), 'folder': "default"}
        chats[i]['folder'] = chats[i]['topic'] + \
            '_'+chats[i]['id'].replace(':', '')
        print(str(i) + ': ' + chats[i]['topic'] + ' ::: ' + chats[i]['id'])
        if not os.path.exists(chats[i]['folder']):
            os.mkdir(chats[i]['folder'])
        i += 1
    return chats


def download_chat(token: str, cookie: Dict, chat: Dict):
    _headers = {'Authorization': 'Bearer ' + token}
    chatDetailFull = []
    reqHost = "https://graph.microsoft.com/beta/me/chats/" + \
        chat['id'] + "/messages"
    outFile = open(chat['folder']+'/' +
                   chat['topic'] + '.log', 'w')
    while True:
        chatDetail = requests.get(reqHost, headers=_headers).json()
        time.sleep(0.05)
        if "value" in chatDetail:
            chatDetailFull.extend(chatDetail["value"])
            for val in chatDetail["value"]:
                for attach in val["attachments"]:
                    print(attach["contentUrl"])
                    if attach["contentType"] == "reference":
                        download_file(
                            attach["contentUrl"], chat['folder'], cookie=cookie)
                    else:
                        print("not a file attachment")
        else:
            print(chatDetail)

        if "@odata.nextLink" in chatDetail:
            reqHost = chatDetail["@odata.nextLink"]
        else:
            outFile.write(json.dumps(chatDetailFull))
            outFile.flush()
            print("Done, Output can be found here: " + chat["folder"])
            break
    return


async def main():
    (cookie, token) = await load()
    chats = await load_chats(token=token)
    root = tk.Tk()
    app = Application(master=root, in_cookie=cookie,
                      in_token=token, in_chats=chats)
    app.pack(fill=tk.BOTH, expand=tk.YES)
    app.mainloop()

 # Entrance run
if __name__ == '__main__':
    asyncio.get_event_loop().run_until_complete(main())
