from typing import Coroutine, MutableSet, Dict, List
from enum import Enum
import aiofiles
from pyppeteer import launch, page
import json
from pyppeteer.browser import Browser
from pyppeteer.element_handle import ElementHandle
import os
import time
import asyncio
from datetime import datetime
from TeamsDownloaderUtil import TeamsDownloaderUtil as TeamsDownloaderUtil
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
    _util: TeamsDownloaderUtil
    base_msg_url: str

    def __init__(self) -> None:
        pass

    def __init__(self, in_util: TeamsDownloaderUtil) -> None:
        self._util = in_util

    async def create_chat(self, v: Dict):

        self.id = v["id"]
        self.base_msg_url = f"https://graph.microsoft.com/beta/me/chats/{self.id}/messages?$top=50"
        self.members = await self.load_chat_members(self.id)
        self.topic = self.members[0].name + "_" + (self.members[1].name if len(
            self.members) > 1 else "?????") if v['chatType'] == 'oneOnOne' else "No_Topic" if not v['topic'] else v['topic']
        self.chatType = v['chatType'] or "No Chat type"
        self.folder = await self._util.normalize_str(
            self.topic + '_' + self.id)
        return self

    async def load_chat_members(self, chatID: str):
        members = []
        membersURL = 'https://graph.microsoft.com/beta/me/chats/' + chatID + '/members'
        async with self._util.http_client.get(membersURL) as resp:
            members_resp = await resp.json()
        if "value" in members_resp:
            for v in members_resp["value"]:
                members.append(ChatMember(v["id"], v['displayName']))
        else:
            print(json.dumps(members_resp))
        return members

    async def download(self):
        chatDetailFull = []
        reqHost = self.base_msg_url
        if not os.path.exists(await self._util.normalize_str(self.folder)):
            os.mkdir(await self._util.normalize_str(self.folder))

        HTML_PARSER = MyHTMLParser()

        while True:
            chatDetail = []

            chatDetail = await self._util.http_client.get(reqHost)
            chatDetail = await chatDetail.json()

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
                            asyncio.create_task(self._util.download_file(url=src, folder=self.folder, file_override=msg_time.strftime(
                                "%Y%m%d%H%M") + "_" + val["from"]["user"]["displayName"] + "_" + str(img_loop_c) + ".jpg"))
                            img_loop_c = img_loop_c + 1
                        HTML_PARSER.reset()
                        HTML_PARSER.close()
                    for attach in val["attachments"]:
                        print(attach["contentUrl"])
                        if attach["contentType"] == "reference":
                            asyncio.create_task(self._util.download_file(
                                attach["contentUrl"], self.folder))
                        else:
                            print("not a file attachment")
            else:
                print(chatDetail)

            if "@odata.nextLink" in chatDetail and chatDetail["@odata.nextLink"] != reqHost:
                print(chatDetail["@odata.nextLink"])
                reqHost = chatDetail["@odata.nextLink"]

            else:
                async with aiofiles.open(await self._util.normalize_str(self.folder)+'/' + await self._util.normalize_str(self.topic, False) + '_chat_log.json', 'w') as f:
                    await f.write(json.dumps(chatDetailFull, indent=2))

                break
        return


class TeamsChannel(TeamsChat):
    id: str
    display_name: str
    description: str
    team_id: str
    team_name: str

    def __init__(self, obj, in_util: TeamsDownloaderUtil) -> None:
        super(TeamsChannel, self).__init__(in_util=in_util)
        self.id = obj["id"]
        self.display_name = obj["displayName"]
        self.description = obj["description"]
        self.team_name = obj["team_name"]
        self.team_id = obj["team_id"]
        self.base_msg_url = f'https://graph.microsoft.com/beta/teams/{self.team_id}/channels/{self.id}/messages?$top=50'

    async def create_chat(self, v: Dict):
        self.members = []
        self.topic = self.display_name
        self.chatType = "channel"
        self.folder = await self._util.normalize_str(
            self.topic + '_' + self.id)
        return self


class Team():
    id: str
    createdDateTime: str
    display_name: str
    description: str
    channels: List[TeamsChannel]

    def __init__(self, obj) -> None:
        self.id = obj["id"]
        self.display_name = obj["displayName"]
        self.description = obj["description"]
        self.channels = []


class TeamsDownloader():
    chats: 'dict[int, TeamsChat]' = {}
    channels: 'dict[int, TeamsChannel]' = {}
    teams: 'dict[int, Team]' = {}
    _sharepoint_cookie: Dict = {}
    _graph_token: str = ""
    _teams_util = TeamsDownloaderUtil()

    def __init__(self) -> None:
        pass

    def __getstate__(self):
        state = self.__dict__.copy()
        del state['_teams_util']
        return state

    def __setstate__(self, state):
        self.__dict__.update(state)

    async def init(self, callback: Coroutine = None) -> None:
        await self.load_auth()
        await self._teams_util.init_http(in_cookies=self.sharepoint_cookie, in_headers={'Authorization': 'Bearer ' + self.graph_token})
        await self.load_chats()
        await self.load_teams()
        await self.load_channels(self.teams)

        for t_idx, team in self.teams.items():
            for chan in team.channels:
                print(f'{chan.id} : {chan.description}')

        for k, v in self.chats.items():
            print(v.topic)

    async def download_chats(self, chat_indexes: List[int], channel_indexes: List[int]):
        for chat_idx in chat_indexes:
            if (not self.chats[chat_idx]._util):
                print(f'Http client not initialized passing in')
                self.chats[chat_idx]._util = self._teams_util
            print(f'Downloading: {self.chats[chat_idx].topic}')

        for chan_idx in chat_indexes:
            if (not self.channels[chan_idx]._util):
                print(f'Http client not initialized passing in')
                self.channels[chan_idx]._util = self._teams_util

        await asyncio.gather(*[self.chats[x].download() for x in chat_indexes], *[self.channels[x].download() for x in channel_indexes])
        print("Done for all")

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

    async def load_graph_data(self, base_url: str):
        chats_data = []
        try:
            while True:
                async with self._teams_util.http_client.get(base_url) as resp:
                    data = await resp.json()
                if "value" in data:
                    chats_data.extend(data["value"])
                    if "@odata.nextLink" in data and data["@odata.nextLink"] != base_url:
                        base_url = data["@odata.nextLink"]
                    else:
                        break
                else:
                    break
        except Exception as e:
            print("Exception loading data -- ")
            print(e)
        return chats_data

    async def load_chats(self):
        if os.path.isfile('./chats.json'):
            print("Chat Cache Exists, utlizing it")
            await self.load_chat_cache()
            return
        else:
            print("Chat Cache Doesn't Exist, Refreshing")

        for i, v in enumerate(await self.load_graph_data(base_url='https://graph.microsoft.com/beta/me/chats?$top=50')):
            new_chat = TeamsChat(self._teams_util)
            self.chats[i] = await new_chat.create_chat(v)
        await self._teams_util.save_file(self.chats, "chats.json", True, ignore_fields=["http_client", "_util"])

    async def load_teams(self):
        for i, v in enumerate(await self.load_graph_data(base_url='https://graph.microsoft.com/beta/me/joinedTeams')):
            self.teams[i] = Team(v)
        print("Loaded Teams")

    async def load_channels(self, teams: List[Team]):
        for t_idx, team in teams.items():
            base_url = f'https://graph.microsoft.com/beta/teams/{team.id}/channels'
            for c_idx, v in enumerate(await self.load_graph_data(base_url=base_url)):
                v["team_name"] = team.display_name
                v["team_id"] = team.id
                tmp_chan = TeamsChannel(v, self._teams_util)
                tmp_chan = await tmp_chan.create_chat(v)
                team.channels.append(tmp_chan)
                self.channels[c_idx] = tmp_chan

    async def load_chat_cache(self):
        print('loading chats from cached chats.json')
        f = await self._teams_util.load_file("chats.json", is_json=True)
        for k, v in f.items():
            new_chat = TeamsChat(self._teams_util)
            self.chats[int(k)] = await new_chat.create_chat(v)
