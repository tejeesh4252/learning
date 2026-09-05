# automation/test_downloader.py

import asyncio
from playwright.async_api import async_playwright
from downloader_updated import FundDownloader


async def main():

    # ── 1. ONE fund only for testing ──────────────────────────────────────
    test_funds = [
        {"name": "AREEF GE 3.1", "code": "157257"},
    ]

    # ── 2. Date range ─────────────────────────────────────────────────────
    start_date = "01/07/2026"
    end_date   = "31/07/2026"

    # ── 3. Where to save the downloaded PDF ───────────────────────────────
    download_folder = r"C:\Users\wj596\Downloads\RE_Bank_Test"  # 🔧 change as needed

    async with async_playwright() as p:

        # ── 4. Launch visible browser MAXIMIZED ──────────────────────────────────
        # ✅ FIXED — full screen browser, hit-testing intact
        # ✅ SIMPLEST — just make it work first
        browser = await p.chromium.launch(headless=False, slow_mo=300)
        context = await browser.new_context()
        page = await context.new_page()


        # ── 5. Go to portal and wait for manual login ─────────────────────
        await page.goto(
            "https://www.olisnet.com/OlisAuthenticate/JSP/login.jsp"
        )
        print("─" * 60)
        print("👉  Log in manually in the browser window.")
        print("👉  Once you see the DASHBOARD, press ENTER here.")
        print("─" * 60)
        input()

        # ── 6. Create downloader and connect signals ───────────────────────
        downloader = FundDownloader(
            page            = page,
            funds           = test_funds,
            start_date      = start_date,
            end_date        = end_date,
            download_folder = download_folder,  # ✅ now passed in
        )

        downloader.status.connect(  lambda msg:       print(f"  STATUS  : {msg}"))
        downloader.progress.connect(lambda i, t, msg: print(f"  PROGRESS: [{i}/{t}] {msg}"))
        downloader.error.connect(   lambda msg:       print(f"  ❌ ERROR : {msg}"))
        downloader.finished.connect(lambda:           print("  ✅ FINISHED!"))

        # ── 7. Run ────────────────────────────────────────────────────────
        await downloader.run()

        # ── 8. Keep open to inspect result ────────────────────────────────
        print("\n👉  Done! Press ENTER to close the browser...")
        input()
        await browser.close()


if __name__ == "__main__":
    import os
    os.makedirs(r"C:\Users\wj596\Downloads\RE_Bank_Test", exist_ok=True)
    asyncio.run(main())
