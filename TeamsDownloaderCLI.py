import asyncio
import TeamsDownloader
from TeamsDownloaderApp import MainFrame
import tkinter as tk
from wxasync  import StartCoroutine, WxAsyncApp
 

if __name__ == '__main__':
    print("Initializing App")
    app = WxAsyncApp()
    frame = MainFrame()
    frame.Show()
    app.SetTopWindow(frame)
    loop = asyncio.get_event_loop()
    loop.run_until_complete(app.MainLoop())

    