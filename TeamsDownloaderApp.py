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
from tkinter import Label, messagebox

from websockets import uri

from TeamsDownloader import TeamsDownloader
from TeamsDownloader import TeamsChat

import wx
from wx.core import DefaultSize, ListBox
from wxasync import AsyncBind, WxAsyncApp, StartCoroutine




class MainFrame(wx.Frame):
    cookies: dict
    token: str
    chats: dict
    chatList: ListBox

    def __init__(self, parent=None):
        super(MainFrame, self).__init__(parent)
        panel = wx.Panel(self)
        
        btn_download =  wx.Button(panel, label="download chat")
        btn_open = wx.Button(panel, label="open selected folders")
        btn_sizer = wx.BoxSizer(wx.HORIZONTAL)
        btn_sizer.Add(btn_download)
        btn_sizer.Add(btn_open)

        
        
        self.status_text = wx.StaticText(panel, label="Select a chat to begin")
        vbox = wx.BoxSizer(wx.VERTICAL)
        vbox.Add(btn_sizer)
        hbox = wx.BoxSizer(wx.HORIZONTAL)
        self.chatList = wx.ListBox(panel, style=wx.LB_MULTIPLE, choices=[])
        AsyncBind(wx.EVT_LISTBOX, self.lb_select, self.chatList)

        vbox.Add(self.chatList, proportion=1, flag=wx.EXPAND)
        hbox.Add(vbox, 1, wx.EXPAND )
        
        hbox.Add(self.status_text, flag=wx.EXPAND | wx.LEFT, border=10)
        
        panel.SetSizer(hbox)

        AsyncBind(wx.EVT_BUTTON, self.download, btn_download)
        AsyncBind(wx.EVT_BUTTON, self.open_folder, btn_open)
        StartCoroutine(self.init, self)
        
    async def init(self):
        self.downloader = TeamsDownloader()
        
        await self.downloader.init()
        await self.populate_chat_list()
        #AsyncBind(wx.EVT_BUTTON, self.downloader.init(callback=self.populate_chat_list), button1)
        #AsyncBind(wx.EVT_BUTTON, self.async_callback, button1)


    async def lb_select(self, evt):
        self.status_text.SetLabel("")

        selections = self.chatList.GetSelections()
        for sel_idx in selections:
            selected_chat = self.downloader.chats[sel_idx]        
            self.status_text.SetLabel(self.status_text.LabelText + '\n' + selected_chat.topic + '\n')
            self.status_text.SetLabel(
                self.status_text.LabelText + '\n'.join([x.name for x in selected_chat.members]))
            self.status_text.SetLabel(self.status_text.LabelText + '\n-----\n')

    async def update_clock(self):
        while True:
            self.edit_timer.SetLabel(time.strftime('%H:%M:%S'))
            await asyncio.sleep(0.5)

    async def populate_chat_list(self):
        topics = [x.topic for k,x in self.downloader.chats.items()]
        self.chatList.InsertItems([x.topic for k,x in self.downloader.chats.items()], 0)
        self.Layout()

    async def open_folder(self, evt):
        selection = self.chatList.GetSelections()
        if len(selection) >= 1:
            for idx in selection:
                os.startfile(self.downloader.chats[idx].folder)

    async def download(self, evt):
        selection = self.chatList.GetSelections()
        if len(selection) >= 1:
            await self.downloader.download_chats(selection)
        else:
            self.status_text.SetLabel("Select a chat to download first")