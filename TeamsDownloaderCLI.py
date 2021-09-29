import asyncio
import TeamsDownloader

async def main():
    td = TeamsDownloader.TeamsDownloader()
    await td.init()

if __name__ == '__main__':
    loop = asyncio.get_event_loop()
    loop.run_until_complete(main())
    loop.run_until_complete(asyncio.sleep(0.250))
    loop.close()