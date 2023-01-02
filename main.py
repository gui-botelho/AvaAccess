import asyncio
import datetime
import time

import pyppeteer.errors
from bs4 import BeautifulSoup
from pyppeteer import launch

user = input('What is your username?')
password = input('And your password?')

# .
async def main():
    # Gets current date's ISOWEEK so I can access only the associated week's class. So it doesn't have to access
    # all of them runtime. It also adjusts them to account for first and second semester. The math looks weird but works
    current_week = datetime.date.isocalendar(datetime.date.today())[1]
    if current_week > 6:
        if current_week > 26:
            week_index = current_week - 32
        else:
            week_index = current_week - 6
    else:
        week_index = 0

    # Launches an internet browser and accesses the main page for login.
    browser = await launch()
    page = await browser.newPage()
    page_path = 'https://www.avaeduc.com.br/theme/kroton/login/loginmanualunidade.php'
    await page.goto(page_path, {'waitUntil': 'load', 'timeout': 0})

    # Logs into page with my credentials and clicks the button. Then it sleeps for 7 seconds so the page loads.
    # Couldn't make {'waitUntil': 'load'} work on this. Need to figure it out so I don't hardcode the time.
    await page.type('[id = units]', 'Unic')
    await page.type('[id = username]', user)
    await page.type('[name = password]', password)
    await page.click('[type = submit]')
    print('Accessing AVA')
    time.sleep(7)
    try:
        await page.click('[id = drawer-toggle-button]')
    except pyppeteer.errors.PageError:
        print('Timed out. Trying again.')
        time.sleep(20)
        await page.click('[id = drawer-toggle-button]')
    # Gets page html content for bs4 parsing, which allows me to get the links for the courses I teach.
    # Then, from this html and links, I get the unit I need to access and ask the program to access them.
    # The last one is a website side bug, so I just have it removed.
    html = await page.content()
    soup = BeautifulSoup(html, 'html.parser')
    courses = soup.select('[data-key]')
    keys = [course.get('data-key', None) for course in courses if course.get('data-key', None).isnumeric()]
    keys.remove(keys[-1])
    counter = 1
    units = {}
    for key in keys:
        #
        await page.goto(f'https://www.avaeduc.com.br/course/view.php?id={key}', {'waitUntil': 'load', 'timeout': 0})
        unit_soup = BeautifulSoup(await page.content(), 'html.parser')
        unit = unit_soup.select('.timeline-menu > ul > li > a')

        unit_urls = [section.get('href', None) for section in unit]
        units[f'subject{counter}'] = [url for url in unit_urls]

        for course_to_access in units:
            new_page = await browser.newPage()

            # Some courses don't have as many links. Estágios, mainly. Thus, I told the program to try and run the
            # current week index associated link, if that is not possible., to access the last link in course list.
            try:
                week_link = units[course_to_access][week_index]

            except IndexError:
                week_link = units[course_to_access][-1]
                print(f'Failed to access unit. Defaulted to link {week_link}')

            finally:
                await new_page.goto(week_link, {'waitUntil': 'load', 'timeout': 0})

        print(f'Accessed page {counter}')
        counter += 1

    print('All your pages have been accessed! Enjoy your week. :)')
    # This is here mainly so it keeps the command line open so I can see that it is done.
    input('Press enter to exit.')


asyncio.get_event_loop().run_until_complete(main())
