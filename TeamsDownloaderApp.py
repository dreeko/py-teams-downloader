# -*- coding: utf-8 -*-
import asyncio
from datetime import datetime
from enum import Enum
import pathlib
from TeamsDownloader import TeamsChat, TeamsDownloader
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

from TeamsDownloader import TeamsDownloader
from TeamsDownloader import TeamsChat

import wx
from wx.core import ListBox
from wxasync import AsyncBind, WxAsyncApp, StartCoroutine


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


class MainFrame(wx.Frame):
    cookies: dict
    token: str
    chats: dict
    chatList: ListBox

    def __init__(self, parent=None):
        super(MainFrame, self).__init__(parent)
        StartCoroutine(self.init, self)
        
    async def init(self):
        self.cookies = {}
        self.token = ""
        self.chatList = None
        self.downloader = TeamsDownloader()
        button1 =  wx.Button(self, label="load chats")
        
        #AsyncBind(wx.EVT_WINDOW_CREATE, self.downloader.init, self)
        #StartCoroutine(self.downloader.init(self.populate_chat_list), self)
        
        hbox = wx.BoxSizer(wx.HORIZONTAL)
        #vbox = wx.BoxSizer(wx.VERTICAL)
        #vbox.Add(button1, 1, wx.EXPAND|wx.ALL)
        #hbox.Add(hbox, wx.EXPAND|wx.ALL)
        self.chatList = wx.ListBox(self, style=wx.LB_MULTIPLE, choices=[])
        self.status_text = wx.StaticText(
            self, style=wx.ALIGN_CENTRE_HORIZONTAL | wx.ST_NO_AUTORESIZE)

        hbox.Add(self.chatList, 2, wx.EXPAND | wx.ALL)
        hbox.AddStretchSpacer(1)
        hbox.Add(self.status_text, 1, wx.EXPAND | wx.ALL)

        self.SetSizer(hbox)
        self.Layout()
        AsyncBind(wx.EVT_LISTBOX, self.lb_select, self.chatList)
        await self.downloader.init()
        await self.populate_chat_list()
        #AsyncBind(wx.EVT_BUTTON, self.downloader.init(callback=self.populate_chat_list), button1)
        #AsyncBind(wx.EVT_BUTTON, self.async_callback, button1)


    async def lb_select(self, evt):
        print("wat")
        sel_chat: TeamsChat = self.downloader.chats[self.chatList.GetSelection()]
        topic: str = sel_chat.topic
        self.status_text.SetLabel("")
        self.status_text.SetLabel(self.status_text.LabelText + '\n' + topic)
        self.status_text.SetLabel(
            self.status_text.LabelText + '\n'.join([x.name for x in sel_chat.members]))

    async def update_clock(self):
        while True:
            self.edit_timer.SetLabel(time.strftime('%H:%M:%S'))
            await asyncio.sleep(0.5)

    async def populate_chat_list(self):
        topics = [x.topic for k,x in self.downloader.chats.items()]
        self.chatList.InsertItems([x.topic for k,x in self.downloader.chats.items()], 0)

    def open_folder(self):
        selected: str = self.chat_list.selection_get()
        chat = self.downloader.chats[int(selected.split(':')[0])]
        os.startfile(chat['folder'])

    def download(self):
        chat: Dict = {}
        folders = []
        for selected in self.chat_list.selection_get().split('\n'):
            print("Downloading Chat: " + selected)
            chat = self.downloader.chats[int(selected.split(':')[0])]
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