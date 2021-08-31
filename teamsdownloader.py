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
from tkinter import messagebox
from concurrent.futures import ThreadPoolExecutor

import wx
from wx.core import ListBox
from wxasync import AsyncBind, WxAsyncApp, StartCoroutine
import asyncio
from asyncio.events import get_event_loop

_executor = ThreadPoolExecutor(10)

class MainFrame(wx.Frame):
    cookies : dict
    token: str
    chats: dict
    chatList: ListBox

    def __init__(self, parent=None):
        super(MainFrame, self).__init__(parent)
        self.cookies = {}
        self.token = ""
        self.chatList = None

        StartCoroutine(self.load(), self)
        while self.token == None:
            time.sleep(1)
        
        StartCoroutine(self.load_chats(self.token), self)
        hbox = wx.BoxSizer(wx.HORIZONTAL)
        self.chatList = wx.ListBox(self, style=wx.LB_SINGLE, choices=[])
        #button1 =  wx.Button(self, label="Submit")
        self.status_text =  wx.StaticText(self, style=wx.ALIGN_CENTRE_HORIZONTAL|wx.ST_NO_AUTORESIZE)

        hbox.Add(self.chatList, 2, wx.EXPAND|wx.ALL)
        hbox.AddStretchSpacer(1)
        hbox.Add(self.status_text, 1, wx.EXPAND|wx.ALL)

        self.SetSizer(hbox)
        self.Layout()
        AsyncBind(wx.EVT_LISTBOX , self.lb_select, self.chatList)
        #AsyncBind(wx.EVT_BUTTON, self.async_callback, button1)

    async def lb_select(self, evt):
        print("wat")
        sel_chat: Dict = self.chats[self.chatList.GetSelection() +1]
        topic: str = sel_chat["topic"]
        self.status_text.SetLabel(self.status_text.LabelText + '\n' + topic)
        self.status_text.SetLabel(self.status_text.LabelText + '\n'.join(sel_chat["members"]))

    async def update_clock(self):
        while True:
            self.edit_timer.SetLabel(time.strftime('%H:%M:%S'))
            await asyncio.sleep(0.5)

    async def load(self):
        browser: Browser = None
        minutes_15 = 900
        minutes_45 = 2700
        try:
            if time.time() - os.stat('token.txt').st_mtime <= minutes_45:
                pass
            else:
                self.status_text.SetLabel(self.status_text.LabelText + '\n' + "the token has timed out, refreshing now")
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
                self.status_text.SetLabel(self.status_text.LabelText + '\n' + "The cookie has timed out, refreshing now")
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
        self.status_text
        self.cookies, self.token = [req_cookies, token]

    async def load_chats(self, token):
        try:
            if os.path.isfile('./chats.json'):
                self.status_text.SetLabel(self.status_text.LabelText + '\n' + "Chat Cache Exists, utlizing it")
                await load_chat_cache(self)
                return
            else:
                self.status_text.SetLabel(self.status_text.LabelText + '\n' + "Chat Cache Doesn't Exist, Refreshing")
        except Exception as e:
            self.status_text.SetLabel(self.status_text.LabelText + '\n' + "Chat Cache Doesn't Exist, Refreshing -- ")
            self.status_text.SetLabel(self.status_text.LabelText + '\n' + str(e))

        chats = {}
        chats_data = []
        chaturl = 'https://graph.microsoft.com/beta/me/chats'
        _headers = {'Authorization': 'Bearer ' + self.token}
        while True:
            data = requests.get(chaturl, headers=_headers).json()
            print(data)
            chats_data.extend(data["value"])
            if "@odata.nextLink" in data:
                chaturl = data["@odata.nextLink"]
            else:
                break
        i = 1
        for v in chats_data:
            #print(str(i) + ': ' + (v['chatType'] or "No Chat type") + ' ::: ' + (v['topic'] or "No Topic") + ' - ' + (v['id'] or "No ID"))
            chats[i] = {'id': v['id'], 'topic': (v['topic'] or "No_Topic"), 'chat_type': (
                v['chatType'] or "No Chat type"), 'folder': "default"}
            chats[i]['folder'] = chats[i]['topic'] + \
                '_'+chats[i]['id'].replace(':', '')

            chats[i]['members'] = await load_chat_members(token, chats[i]['id'])
            print(str(i) + ': ' + chats[i]['topic'] + ' ::: ' + chats[i]['id'])
            for m in chats[i]['members']:
                print(m)
            i += 1
        await save_chat_cache(chats)
        self.chats = chats
        self.chatList.InsertItems( [c["topic"] for k,c in chats.items()])

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
        self.label_text.set(chat['id'] + '\n' + chat['chat_type'] +
                            '\n' + chat['topic'] + '\n' + '\n'.join(chat['members']))

    def open_folder(self):
        selected: str = self.chat_list.selection_get()
        chat = self.chats[int(selected.split(':')[0])]
        os.startfile(chat['folder'])

    async def download(self):
        selected: str = self.chat_list.selection_get()
        print("Downloading Chat: " + selected)
        await download_chat(cookie=self.cookie, token=self.token,
                      chat=self.chats[int(selected.split(':')[0])])
# Save cookie


async def save_cookie(cookie):
    with open("cookie.json", 'w+', encoding="utf-8") as file:
        json.dump(cookie, file, ensure_ascii=False)


async def save_chat_cache(chats):
    with open("chats.json", 'w+', encoding="utf-8") as file:
        json.dump(chats, file)


async def load_chat_cache(self):
    with open("chats.json", 'r', encoding="utf-8") as file:
        payload = json.load(file)
        chats = {}
        for k, v in payload.items():
            print(str(k))
            chats[int(k)] = v
        self.chatList.InsertItems([c["topic"] for c in chats.values()], 0)
        self.chats = chats


async def save_token(token):
    with open("token.txt", 'w+', encoding="utf-8") as file:
        file.write(token)


async def download_file(url, folder, cookie):
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
    await page.focus('#TextField18')
    await page.keyboard.type('/chats')
    await page.keyboard.press('Enter')
    await page.click('button[name^=Modify')
    await asyncio.sleep(3)
    btn : ElementHandle = await page.xpath("//button[contains(., 'Consent')][1]")
    input()
    await btn[0].click()
    input()
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


async def download_chat(token: str, cookie: Dict, chat: Dict):
    _headers = {'Authorization': 'Bearer ' + token}
    chatDetailFull = []
    reqHost = "https://graph.microsoft.com/beta/me/chats/" + \
        chat['id'] + "/messages"
    if not os.path.exists(chat['folder']):
            os.mkdir(chat['folder'])
    outFile = open(chat['folder']+'/' +
                   chat['topic'] + '_chat_log.json', 'w')
    while True:
        chatDetail = []
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

        if "@odata.nextLink" in chatDetail and chatDetail["@odata.nextLink"] != reqHost:
            reqHost = chatDetail["@odata.nextLink"]
        else:
            outFile.write(json.dumps(chatDetailFull, indent=2))
            outFile.flush()
            print("Done, Output can be found here: " + chat["folder"])
            res = messagebox.askquestion('Open dl folder', 'Would you like to open the download folder?')
            if res == 'yes':
                os.startfile(chat["folder"])
            break
    return


async def main():
    print("Initializing App")
    (cookie, token) = await load()
    print("Auth has loaded")
    chats = await load_chats(token=token)
    print("Chats have been loaded")
    # root = tk.Tk()
    # app = Application(master=root, in_cookie=cookie,
    #                   in_token=token, in_chats=chats)
    # app.pack(fill=tk.BOTH, expand=tk.YES)
    # print("Entering GUI Main loop")
    # app.mainloop()

 # Entrance run
if __name__ == '__main__':
    app = WxAsyncApp()
    frame = MainFrame()
    frame.Show()
    app.SetTopWindow(frame)
    loop = get_event_loop()
    loop.run_until_complete(app.MainLoop())