import asyncio
from playwright.async_api import async_playwright
import os

async def verify_drawer():
    async with async_playwright() as p:
        browser = await p.chromium.launch()
        context = await browser.new_context(viewport={'width': 1280, 'height': 800})
        page = await context.new_page()

        # Load the local HTML file
        curr_dir = os.getcwd()
        file_path = f"file://{curr_dir}/VLIndex.html"
        await page.goto(file_path)

        # 1. Open User Drawer
        await page.click('#userAvatar')
        await asyncio.sleep(1) # wait for animation
        await page.screenshot(path='verification/screenshots/user_drawer_my.png')
        print("Captured user_drawer_my.png")

        # 2. Mock some team data and switch to Team tab
        await page.evaluate("""
            serverEvents.push({
                id: 'mock-1',
                ldap: 'testagent',
                email: 'testagent@google.com',
                date: '2026-05-15',
                channel: 'Chat',
                workGroup: 'Play NA',
                reason: 'Personal',
                status: 'Pending',
                teamLead: 'Team Al',
                timestamp: '2026-05-01 10:00:00'
            });
            isSupervisor = true;
            myTeamName = 'Team Al';
            document.getElementById('drawerSupTabs').style.display = 'flex';
            renderUserRequests();
        """)

        await page.click('#tabTeamReq')
        await asyncio.sleep(0.5)
        await page.screenshot(path='verification/screenshots/user_drawer_team.png')
        print("Captured user_drawer_team.png")

        # 3. Expand a card
        await page.click('.req-card')
        await asyncio.sleep(0.5)
        await page.screenshot(path='verification/screenshots/user_drawer_expanded.png')
        print("Captured user_drawer_expanded.png")

        await browser.close()

if __name__ == "__main__":
    if not os.path.exists('verification/screenshots'):
        os.makedirs('verification/screenshots')
    asyncio.run(verify_drawer())
