from typing import Coroutine, MutableSet, Dict
from enum import Enum
import aiofiles
import aiohttp
from pyppeteer import launch, page
import json
from pyppeteer.browser import Browser
from pyppeteer.element_handle import ElementHandle
import requests
import shutil
import os
import time
import asyncio
from datetime import datetime
from TeamsDownloaderUtil import TeamsDownloaderUtil as TeamsDownloaderUtil


class ChatType(Enum):
    CHANNEL = 1
    ROOM = 2


class RoomType(Enum):
    ONE_ON_ONE = 1  # between two people
    GROUP = 2  # more than two people
    CHANNEL = 3  # a channel


class ChatMember():
    name: str
    id: str

    def __init__(self, in_id: str, in_displayname: str) -> None:
        self.id = in_id
        self.name = in_displayname


class TeamsChat():
    id: str
    topic: str
    room_type: ChatType  # channel or group
    chatType: RoomType
    folder: str
    members: 'MutableSet[ChatMember]'
    _http_client: aiohttp.ClientSession 

    def __init__(self) -> None:
        pass
    def __init__(self, in_client: aiohttp.ClientSession) -> None:
        self.http_client = in_client

    async def create_chat(self, v: Dict):
        self.id = v["id"]
        self.members = await self.load_chat_members(self.id)
        self.topic = self.members[0].name + "_" + (self.members[1].name if len(
            self.members) > 1 else "?????") if v['chatType'] == 'oneOnOne' else "No_Topic" if not v['topic'] else v['topic']
        self.chatType = v['chatType'] or "No Chat type"
        self.folder = await TeamsDownloaderUtil.normalize_str(
            self.topic + '_' + self.id)
        return self

    async def load_chat_members(self, chatID: str):
        members = []
        membersURL = 'https://graph.microsoft.com/beta/me/chats/' + chatID + '/members'
        async with self.http_client.get(membersURL) as resp:
            members_resp = await resp.json()
        if "value" in members_resp:
            for v in members_resp["value"]:
                members.append(ChatMember(v["id"], v['displayName']))
        else:
            print(json.dumps(members_resp))
        return members


class TeamsDownloader():
    chats: 'dict[int, TeamsChat]' = {}
    _sharepoint_cookie: Dict = {}
    _graph_token: str = ""
    _teams_util = TeamsDownloaderUtil()

    def __init__(self) -> None:
        pass

    async def init(self, callback: Coroutine=None) -> None:
        await self.load_auth()
        await self._teams_util.init_http(in_cookies=self.sharepoint_cookie, in_headers={'Authorization': 'Bearer ' + self.graph_token})
        await self.load_chats()

        for k, v in self.chats.items():
            print(v.topic)


        


    async def load_graph_explorer_token(self, page: page.Page, url):
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
        await page.waitForXPath("//span[text()='Consented']")
        await page.click('button[name^=Access')
        await page.waitForSelector('label.ms-Label:nth-child(2)', {'timeout': 300000})
        token_element = await page.querySelector('label.ms-Label:nth-child(2)')
        token = await page.evaluate('(element) => element.textContent', token_element)
        print('found token' + token)
        await self._teams_util.save_file(token, "token.txt")

    async def load_sharepoint_cookies(self, page, url):
        await page.setViewport({'width': 1366, 'height': 768})
        await page.setUserAgent('Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) '
                                'Chrome/58.0.3029.110 Safari/537.36 Edge/16.16299')
        await page.goto(url)
        await page.waitForSelector('#O365_SuiteBranding_container', {'timeout': 300000})
        print('found selector')
        await asyncio.sleep(2)
        cookies = await page.cookies()
        await self._teams_util.save_file(cookies, "cookie.json", True)
        print('found cookies: ' + str(cookies))

    async def load_auth(self):
        browser: Browser = None
        if(not(await self._teams_util.file_within_age_threshold("token.txt", 2700))):
            if (not browser):
                browser = await self._teams_util.launch_browser()
            page1 = await browser.pages()
            page1 = page1[0]
            await self.load_graph_explorer_token(page=page1, url='https://developer.microsoft.com/en-us/graph/graph-explorer')

        if(not(await self._teams_util.file_within_age_threshold("cookie.json", 2700))):
            if (not browser):
                browser = await self._teams_util.launch_browser()
            page2 = await browser.newPage()
            # await self.load_sharepoint_cookies(page2, 'https://wapol-my.sharepoint.com/')
            await self.load_sharepoint_cookies(page2, 'https://inoffice.sharepoint.com/')

        cookie = await self._teams_util.load_file("cookie.json", is_json=True)
        req_cookies = {}
        for entry in cookie:
            req_cookies[entry['name']] = entry['value']
        token = await self._teams_util.load_file("token.txt", is_json=False)

        try:
            await browser.close()
        except:
            pass
        self.sharepoint_cookie = req_cookies
        self.graph_token = token
        return [req_cookies, token]

    async def load_chats(self):
        chats_data = []
        try:
            if os.path.isfile('./chats.json'):
                print("Chat Cache Exists, utlizing it")
                await self.load_chat_cache()
                return
            else:
                print("Chat Cache Doesn't Exist, Refreshing")
                
                chaturl = 'https://graph.microsoft.com/beta/me/chats?$top=50'

                while True:
                    async with self._teams_util.http_client.get(chaturl) as resp:
                        data = await resp.json()
                    if "value" in data:
                        chats_data.extend(data["value"])
                        if "@odata.nextLink" in data and data["@odata.nextLink"] != chaturl:
                            chaturl = data["@odata.nextLink"]
                        else:
                            break
                    else:
                        break
        except Exception as e:
            print("Exception loading chats -- ")
            print(e)

        

        for i, v in enumerate(chats_data):
            new_chat = TeamsChat(self._teams_util.http_client)
            self.chats[i] = await new_chat.create_chat(v)

        await self._teams_util.save_file(self.chats, "chats.json", True)

    async def load_chat_cache(self):
        print('loading chats from cached chats.json')
        f = await self._teams_util.load_file("chats.json", is_json=True)
        for k, v in f.items():
            new_chat = TeamsChat(self._teams_util.http_client)
            self.chats[int(k)] = await new_chat.create_chat(v)

    
