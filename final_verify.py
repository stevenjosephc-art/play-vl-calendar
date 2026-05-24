import asyncio
from playwright.async_api import async_playwright
import os

async def final_verify():
    async with async_playwright() as p:
        browser = await p.chromium.launch()
        context = await browser.new_context(viewport={'width': 1400, 'height': 900})
        page = await context.new_page()

        file_path = "file://" + os.path.abspath("VLIndex.html")
        await page.goto(file_path)

        # Mock data
        await page.evaluate("""
          serverEvents = [
            { id: '1', ldap: 'user1', channel: 'Chat', workGroup: 'Play NA', reason: 'Personal', status: 'Pending', date: '2026-05-01', email: 'user1@google.com', timestamp: '2026-04-01 10:00:00', accruals: '10', attendance: '95%' },
            { id: '2', ldap: 'user2', channel: 'Phone', workGroup: 'Play GTV', reason: 'Birthday', status: 'Approved', date: '2026-05-15', email: 'user2@google.com', timestamp: '2026-04-01 11:00:00', accruals: '5', attendance: '98%' },
            { id: '3', ldap: 'user3', channel: 'Email', workGroup: 'Play HVU', reason: 'Family Gathering', status: 'Pending', date: '2026-05-20', email: 'user3@google.com', timestamp: '2026-04-02 09:00:00', accruals: '2', attendance: '90%' }
          ];
          currentUserEmail = 'user1@google.com';
          isSupervisor = true;
          isAdmin = true;
          myTeamName = 'Team Al';
          renderCalendar();
          updateSlotBadge();
        """)

        # 1. Main Calendar with Holidays
        await page.screenshot(path="final_calendar_holidays.png")

        # 2. User Drawer with Staggered Cards
        await page.click("#userAvatar")
        await asyncio.sleep(0.5)
        await page.screenshot(path="final_user_drawer.png")
        await page.click("#drawerBackdrop") # Close drawer

        # 3. Admin Console - Pending Queue
        await page.evaluate("document.getElementById('adminBtnContainer').style.display='flex'")
        await page.click("text=👑 Admin Console")
        await asyncio.sleep(0.5)
        await page.screenshot(path="final_admin_queue.png")

        # 4. Admin Console - Insights
        await page.click("text=Insights & Analytics")
        await asyncio.sleep(1) # Wait for chart animations
        await page.screenshot(path="final_admin_insights.png")

        # 5. Request Detail with Holiday Info
        await page.evaluate("document.getElementById('adminOv').classList.remove('on')")
        await page.evaluate("openViewModal(null, '2')") # Open user2's approved birthday request
        await asyncio.sleep(0.5)
        await page.screenshot(path="final_request_detail.png")

        await browser.close()

if __name__ == "__main__":
    asyncio.run(final_verify())
