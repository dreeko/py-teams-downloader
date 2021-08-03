# -*- coding: utf-8 -*-
import asyncio
from pyppeteer import launch
import json
import requests
import shutil
import os


# Save cookie
async def save_cookie(cookie):
    with open("cookie.json", 'w+', encoding="utf-8") as file:
        json.dump(cookie, file, ensure_ascii=False)


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

 # Load Home


async def index(page, url):
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

 # Main function


async def main():
    chats = {}
    chatDetail = {}

    # Initialize the browser
    browser = await launch({'headless': False,
                            'dumpio': True,
                            'args': [
                                # '--disable-extensions',
                                # '--disable-bundled-ppapi-flash',
                                # '--mute-audio',
                                # '--no-sandbox',
                                # '--disable-setuid-sandbox',
                                '--disable-dev-shm-usage',
                                '--shm-size=1gb'
                                '--disable-gpu',
                            ],
                            'executablePath': 'C:\Program Files\Google\Chrome\Application\chrome.exe'
                            })
    # Open a new tab
    new_cookie = False
    page1 = await browser.pages()
    page1 = page1[0]
    await page1.setViewport({'width': 1366, 'height': 768})
    await page1.setUserAgent('Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) '
                             'Chrome/58.0.3029.110 Safari/537.36 Edge/16.16299')
    await page1.goto('https://developer.microsoft.com/en-us/graph/graph-explorer')
    if new_cookie:
        page2 = await browser.newPage()
        await index(page2, 'https://inoffice.sharepoint.com/')

    cookie = await load_cookie()
    req_cookies = {}
    for entry in cookie:
        req_cookies[entry['name']] = entry['value']

    token = input(
        "please log in on the graph explorer from your browser, copy the access token and paste it here then press enter")
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

    while True:
        chatDetailFull = []
        choice = int(input("choose a chat to load, 999 quits "))
        if choice == 999:
            break
        elif (choice in chats):
            reqHost = "https://graph.microsoft.com/beta/me/chats/" + \
                chats[choice]['id'] + "/messages"
            outFile = open(chats[choice]['folder']+'/' +
                           chats[choice]['topic'] + '.log', 'w')
            while True:
                chatDetail = requests.get(reqHost, headers=_headers).json()
                await asyncio.sleep(0.1)
                if "value" in chatDetail:
                    chatDetailFull.extend(chatDetail["value"])
                    for val in chatDetail["value"]:
                        for attach in val["attachments"]:
                            print(attach["contentUrl"])
                            if attach["contentType"] == "reference":
                                await download_file(
                                    attach["contentUrl"], chats[choice]['folder'], cookie=req_cookies)
                            else:
                                print("not a file attachment")
                else:
                    print(chatDetail)

                if "@odata.nextLink" in chatDetail:
                    reqHost = chatDetail["@odata.nextLink"]
                else:
                    outFile.write(json.dumps(chatDetailFull))
                    outFile.flush()
                    break
        else:
            print("enter a valid choice")

 # Entrance run
if __name__ == '__main__':
    asyncio.get_event_loop().run_until_complete(main())
