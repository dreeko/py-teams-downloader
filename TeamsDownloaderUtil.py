# -*- coding: utf-8 -*-
import asyncio
import pathlib
from typing import Dict, List, MutableSet
from wsgiref import headers
from pyppeteer import launch, page
import json
from pyppeteer.browser import Browser
from pyppeteer.element_handle import ElementHandle

import os
import time
import tkinter as tk
from tkinter import messagebox
import aiohttp
import aiofiles
import asyncio
import jsonpickle as jp
import shutil


from html.parser import HTMLParser

from websockets import uri


class TeamsDownloaderUtil():
    http_client: aiohttp.ClientSession

    async def init_http(self, in_cookies: Dict, in_headers: Dict):
        jar = aiohttp.CookieJar(quote_cookie=False, unsafe=True)
        self.http_client = aiohttp.ClientSession(
            cookies=in_cookies, headers=in_headers, cookie_jar=jar)

    async def launch_browser(self):
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

    async def save_file(self, payload, file_path: str, is_json: bool = False, ignore_fields: List[str] = None):
        data: str = ""
        if (is_json):
            if(ignore_fields):
                for k,ent in payload.items():
                    for field in ignore_fields:
                        setattr(ent,field, None)
            data = jp.encode(payload, unpicklable=False, max_depth=4)
        else:
            data = payload
        async with aiofiles.open(file_path, 'w+', encoding="utf-8") as file:
            await file.write(data)

    async def load_file(self, file_path: str, is_json: bool = True) -> Dict:
        async with aiofiles.open(file_path, 'r', encoding="utf-8") as file:
            if (is_json):
                data = jp.decode(await file.read())
            else:
                data = await file.read()
        return data

    async def download_file(self, url, folder, file_override: str = None):
        local_filename: str
        if file_override is None:
             local_filename = url.split('/')[-1]
        else:
             print(file_override)
             local_filename = file_override
        async with self.http_client.get(url) as r:
            async with aiofiles.open(await self.normalize_str(folder) + '/' + local_filename, 'wb') as f:
                await f.write(await r.read())

        return local_filename

    async def normalize_str(self, in_str: str, path: bool = True):
        ret: str = in_str
        for x in '<>:"/\|?* ':
            ret = ret.replace(x, '_')
        return os.path.normpath(ret)

    async def file_within_age_threshold(self, in_file_path: str, in_max_age_seconds: int) -> bool:
        try:
            if os.path.isfile(in_file_path) and time.time() - os.stat(in_file_path).st_mtime <= in_max_age_seconds:
                return True
            else:
                return False
        except Exception as e:
            return False
