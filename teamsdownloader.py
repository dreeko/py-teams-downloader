# -*- coding: utf-8 -*-
import asyncio
from datetime import datetime
from typing import Dict
from pyppeteer import launch, page
import json
from pyppeteer.browser import Browser
from pyppeteer.element_handle import ElementHandle
import requests
import shutil
import os
import time
import tkinter as tk
from tkinter import messagebox

from html.parser import HTMLParser


class MyHTMLParser(HTMLParser):
    image_srcs: list = []

    def handle_starttag(self, tag, attrs):
        #print("Start tag:", tag)
        for attribute in attrs:
            if attribute[0] == 'src':
                self.image_srcs.append(attribute[1])

    def close(self):
        self.image_srcs = []
        super().close()


HTML_PARSER = MyHTMLParser()


class Application(tk.Frame):
    chats: Dict
    cookie: Dict
    token: str
    label_text: tk.StringVar

    def __init__(self, master=None, in_cookie: Dict = None, in_token: str = None, in_chats: Dict = None):
        super().__init__(master)
        self.chats = in_chats
        self.cookie = in_cookie
        self.token = in_token
        self.master = master
        self.label_text = tk.StringVar(self, "Select an item")
        self.pack(fill=tk.BOTH, expand=tk.YES)
        self.create_widgets()

    def create_widgets(self):
        scrollData = tk.StringVar()
        self.download_btn = tk.Button(self)
        self.open_folder_btn = tk.Button(self)
        self.chat_list = tk.Listbox(
            self, listvariable=scrollData, selectmode="multiple")
        self.chat_list['width'] = 48
        self.chat_list['height'] = 32
        for c in self.chats:
            c = int(c)
            self.chat_list.insert('end', str(
                c) + ': ' + self.chats[c]["topic"])
        self.download_btn["text"] = "Download Selected"
        self.open_folder_btn["text"] = "Open Download Folder"
        self.open_folder_btn["command"] = self.open_folder
        self.open_folder_btn.pack(side=tk.TOP)
        self.download_btn["command"] = self.download
        self.download_btn.pack(side=tk.TOP)
        self.chat_list.bind("<<ListboxSelect>>", self.on_lb_select)
        self.chat_list.pack(side="left", fill='y')

        self.lbl_chat_info = tk.Label(self, anchor='w')
        self.lbl_chat_info['width'] = 64
        self.lbl_chat_info['height'] = 12
#        self.lbl_chat_info['wraplength'] = 128
        self.lbl_chat_info['justify'] = tk.LEFT
        self.lbl_chat_info['textvariable'] = self.label_text
        self.lbl_chat_info.pack(side=tk.LEFT, fill='both')

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

    def on_lb_select(self, event):
        w = event.widget
        index = int(w.curselection()[0])
        value = w.get(index)
        chat = self.chats[int(value.split(':')[0])]
        self.label_text.set("Chat Id: " + '\n' + chat['id'] + '\n\n' + "Chat Type: " + '\n' + chat['chat_type'] +
                            '\n\n' + "Chat Topic" + '\n' + chat['topic'] + '\n\n' + "Chat Participants" + '\n' + '\n'.join(chat['members']))

    def open_folder(self):
        selected: str = self.chat_list.selection_get()
        chat = self.chats[int(selected.split(':')[0])]
        os.startfile(chat['folder'])

    def download(self):
        chat: Dict = {}
        folders = []
        for selected in self.chat_list.selection_get().split('\n'):
            print("Downloading Chat: " + selected)
            chat = self.chats[int(selected.split(':')[0])]
            download_chat(cookie=self.cookie, token=self.token,
                          chat=chat)
            folders.append(chat['folder'])
        print("Done, Output can be found here: " +
              '\n'.join(folder for folder in folders))
        if len(self.chat_list.selection_get().split('\n')) == 1:
            res = messagebox.askquestion(
                'Open dl folder', 'Would you like to open the download folder?')
            if res == 'yes':
                os.startfile(chat["folder"])


async def save_cookie(cookie):
    with open("cookie.json", 'w+', encoding="utf-8") as file:
        json.dump(cookie, file, ensure_ascii=False)


async def save_chat_cache(chats):
    with open("chats.json", 'w+', encoding="utf-8") as file:
        json.dump(chats, file)


async def load_chat_cache():
    with open("chats.json", 'r', encoding="utf-8") as file:
        payload = json.load(file)
        chats = {}
        for k, v in payload.items():
            print(str(k))
            chats[int(k)] = v
        return chats


async def save_token(token):
    with open("token.txt", 'w+', encoding="utf-8") as file:
        file.write(token)


def download_file(url, folder, cookie, file_override: str = None, header=None):
    local_filename: str
    if file_override is None:
        local_filename = url.split('/')[-1]
    else:
        print(file_override)
        local_filename = file_override

    with requests.get(url, stream=True, cookies=cookie, headers=header) as r:
        with open(normalize_str(folder) + '/' + local_filename, 'wb') as f:
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
    await page.focus('#TextField18')
    await page.keyboard.type('/chats')
    await page.keyboard.press('Enter')
    await page.click('button[name^=Modify')
    await page.waitForXPath("//button[contains(., 'Consent')][1]")
    btn: ElementHandle = await page.xpath("//button[contains(., 'Consent')][1]")
    await btn[0].click()
    # await page.waitForXPath("//label[@class='ms-Label consented-300']")
    await page.waitForXPath("//span[text()='Consented']")
    await page.click('button[name^=Access')
    await page.waitForSelector('label.ms-Label:nth-child(2)', {'timeout': 300000})
    token_element = await page.querySelector('label.ms-Label:nth-child(2)')
    token = await page.evaluate('(element) => element.textContent', token_element)
    print('found token' + token)
    await save_token(token)


async def launch_browser():
    browser: Browser = None
    browser_path: str = ""
    try:
        if os.path.isfile("C:\Program Files\Google\Chrome\Application\chrome.exe"):
            print("64bit chrome discovered")
            browser_path = "C:\Program Files\Google\Chrome\Application\chrome.exe"
        elif os.path.isfile("C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"):
            print("32bit chrome discovered")
            browser_path = "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"
        else:
            print("Chrome not discovered, exiting")
            exit()

        browser = await launch({'headless': False,
                                'dumpio': True,
                                'args': [
                                    '--disable-dev-shm-usage',
                                    '--shm-size=1gb'
                                    '--disable-gpu',
                                ],
                                'executablePath': browser_path
                                })
    except Exception as e:
        print("could not launch chrome !")
    return browser


async def load():
    browser: Browser = None

    minutes_15 = 900
    minutes_45 = 2700
    try:
        if os.path.isfile('token.txt') and time.time() - os.stat('token.txt').st_mtime <= minutes_45:
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
        if os.path.isfile('cookie.json') and time.time() - os.stat('cookie.json').st_mtime <= minutes_45:
            pass
        else:
            print("The cookie has timed out, refreshing now")
            raise Exception
    except Exception as e:
        if not browser:
            browser = await launch_browser()
        page2 = await browser.newPage()
        # await sharepoint(page2, 'https://wapol-my.sharepoint.com/')
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
    try:
        if os.path.isfile('./chats.json'):
            print("Chat Cache Exists, utlizing it")
            return await load_chat_cache()
        else:
            print("Chat Cache Doesn't Exist, Refreshing")
    except Exception as e:
        print("Chat Cache Doesn't Exist, Refreshing -- ")
        print(e)

    chats = {}
    chats_data = []
    chaturl = 'https://graph.microsoft.com/beta/me/chats?$top=50'
    _headers = {'Authorization': 'Bearer ' + token}
    while True:
        data = requests.get(chaturl, headers=_headers).json()
        if "value" in data:
            chats_data.extend(data["value"])
            if "@odata.nextLink" in data and data["@odata.nextLink"] != chaturl:
                chaturl = data["@odata.nextLink"]
            else:
                break
        else:
            break
    i = 1
    for v in chats_data:
        #print(str(i) + ': ' + (v['chatType'] or "No Chat type") + ' ::: ' + (v['topic'] or "No Topic") + ' - ' + (v['id'] or "No ID"))

        chats[i] = {'id': v['id'], 'topic': "No_Topic", 'chat_type': (
            v['chatType'] or "No Chat type"), 'folder': "default"}
        chats[i]['members'] = await load_chat_members(token, chats[i]['id'])
        if chats[i]['chat_type'] == "oneOnOne":
            print("oneOnOne" + str(chats[i]['members']))
        chats[i]["topic"] = chats[i]['members'][0] + "_" + (chats[i]['members'][1] if len(
            chats[i]['members']) > 1 else "?????") if v['chatType'] == 'oneOnOne' else "No_Topic" if not v['topic'] else v['topic']
        chats[i]['folder'] = normalize_str(
            chats[i]['topic'] + '_'+chats[i]['id'])
        chats[i]['members'] = await load_chat_members(token, chats[i]['id'])
        print(str(i) + ': ' + chats[i]['topic'] + ' ::: ' + chats[i]['id'])
        for m in chats[i]['members']:
            print(m)
        i += 1
    await save_chat_cache(chats)
    return chats


async def load_chat_members(token: str, chatID: str):
    _headers = {'Authorization': 'Bearer ' + token}
    members = []
    membersURL = 'https://graph.microsoft.com/beta/me/chats/' + chatID + '/members'
    members_resp = requests.get(membersURL, headers=_headers).json()
    if "value" in members_resp:
        for v in members_resp["value"]:
            members.append(v['displayName'])
    else:
        print(json.dumps(members_resp))
    return members


def download_chat(token: str, cookie: Dict, chat: Dict):
    _headers = {'Authorization': 'Bearer ' + token}
    chatDetailFull = []
    reqHost = "https://graph.microsoft.com/beta/me/chats/" + \
        chat['id'] + "/messages" + '?$top=50'
    if not os.path.exists(normalize_str(chat['folder'])):
        os.mkdir(normalize_str(chat['folder']))

    outFile = open(normalize_str(chat['folder'])+'/' +
                   normalize_str(chat['topic'], False) + '_chat_log.json', 'w')

    while True:
        chatDetail = []
        chatDetail = requests.get(reqHost, headers=_headers).json()
        time.sleep(0.05)
        if "value" in chatDetail:
            chatDetailFull.extend(chatDetail["value"])
            for val in chatDetail["value"]:
                if 'graph' in val["body"]["content"]:
                    HTML_PARSER.feed(val["body"]["content"])
                    img_loop_c = 0
                    for src in HTML_PARSER.image_srcs:
                        msg_time: datetime = datetime.strptime(
                            val["lastModifiedDateTime"], "%Y-%m-%dT%H:%M:%S.%f%z")
                        download_file(src, chat['folder'], cookie=cookie, file_override=msg_time.strftime(
                            "%Y%m%d%H%M") + "_" + val["from"]["user"]["displayName"] + "_" + str(img_loop_c) + ".jpg", header=_headers)
                        img_loop_c = img_loop_c + 1
                    HTML_PARSER.reset()
                    HTML_PARSER.close()
                for attach in val["attachments"]:
                    print(attach["contentUrl"])
                    if attach["contentType"] == "reference":
                        download_file(
                            attach["contentUrl"], chat['folder'], cookie=cookie)
                    else:
                        print("not a file attachment")
        else:
            print(chatDetail)

        if "@odata.nextLink" in chatDetail and chatDetail["@odata.nextLink"] != reqHost:
            print(chatDetail["@odata.nextLink"])
            reqHost = chatDetail["@odata.nextLink"]

        else:
            outFile.write(json.dumps(chatDetailFull, indent=2))

            break
    return


def normalize_str(in_str: str, path: bool = True):
    ret: str = in_str
    for x in '<>:"/\|?* ':
        ret = ret.replace(x, '_')
    return os.path.normpath(ret)


async def main():
    print("Initializing App")
    (cookie, token) = await load()
    print("Auth has loaded")
    chats = await load_chats(token=token)
    print("Chats have been loaded")
    root = tk.Tk()
    app = Application(master=root, in_cookie=cookie,
                      in_token=token, in_chats=chats)
    app.pack(fill=tk.BOTH, expand=tk.YES)
    print("Entering GUI Main loop")
    app.mainloop()

 # Entrance run
if __name__ == '__main__':
    asyncio.get_event_loop().run_until_complete(main())
