# py-teams-downloader

a basic Python GUI that piggy backs on top of the Microsoft Graph Explorer in order to download chat logs and attached items

## Requirements

+ Windows 10
+ An updated version of Chrome Installed

## Building

while inside the repository:
 ```pyinstaller.exe --onefile --copy-metadata pyppeteer .\teamsdownloader.py```

## Running

1. Put py-teams-downloader.exe in a directory you dont mind getting cluttered
2. run py-teams-downloader
3. On the first run it will open up a browser window to [Graph Explorer](https://developer.microsoft.com/en-us/graph/graph-explorer)
![image](https://user-images.githubusercontent.com/26187585/128124774-764b03a2-e9c3-4739-b830-2dab0c9185a4.png)
4. Once signed in it will load up a sharepoint page, no action is required here
![image](https://user-images.githubusercontent.com/26187585/128125007-910a71e5-d580-4964-b42d-4672429a58d7.png)

5. Once the Authorization automation has been completed the browser should close and a console should appear, this console will do
an initial load of chats to choose from, no action required
![image](https://user-images.githubusercontent.com/26187585/128125641-72257bf8-c560-44e7-8187-f54768b36580.png)
6. Once the initialization has completed the User interface will load with a list of available chats and the chat's topic (if available)
7. Select an item and see that it loads the technical id of the chat, the type of chat (onOnOne, Group, Meeting) and the participants in that chat on the right hand side
![image](https://user-images.githubusercontent.com/26187585/128125489-3143dd39-be75-459b-8575-5794b8dbdadb.png)
8. The "Download Selected" Button will initiate the download of the chat logs and attached files into a folder next to where the tool is, large chats can take a while to download and will freeze the user interface, just be patient and wait for it to complete which is indicated by the console output ```Done, Output can be found here: XXXXXXX``` where XXX is the folder.
9. the button "Open Download" folder will launch a file explorer at the selected item's directory

## Notes

The Authorization only happens once and is valid for 45 minutes, after that time it will be initiated again, in order to force this process delete the ```token.txt``` and ```cookie.json``` files
