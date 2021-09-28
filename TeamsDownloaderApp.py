# -*- coding: utf-8 -*-
import asyncio
from datetime import datetime
from enum import Enum
import pathlib
from typing import Dict, MutableSet
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

from websockets import uri




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





async def main():
    print("Initializing App")
    teams_downloader = TeamsDownloader()
    await teams_downloader.init()
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
