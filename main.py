import asyncio
from pyppeteer import launch
import openpyxl

async def main():
    browser = await launch()
    page = await browser.newPage()

    await page.goto('https://www.google.com/maps/search/auto+repair+shop+in+San+Francisco+Bay+Area,+CA,+USA')
    await page.waitForSelector('div[role="article"]', timeout=5000)
    # while (True):
    #     ele = page.querySelector('span.HlvSq')
    #     print(ele)
    #     if ele:
    #         break
    #     else:
    #         await page.evaluate('window.scrollTo(0, document.body.scrollHeight)')

    result_names = await page.evaluate('''() => {
        return Array.from(document.querySelectorAll('div.lI9IFe')).map(ele => {
            type = ele.querySelector('div.UaQhfb > :last-child > :first-child > :first-child').textContent;
            
            name = ele.querySelector('div.NrDZNb').textContent;
            rate = ele.querySelector('span.MW4etd').textContent;
            reviews = ele.querySelector('span.UY7F9').textContent.replace('(', '').replace(')', '');
            street = ele.querySelector('div.UaQhfb > :last-child > :first-child > :last-child > :last-child').textContent;
            phone = ele.querySelector('div.UaQhfb > :last-child > :last-child > :last-child > :last-child').textContent;
            city = 'San Francisco Bay Area';
            state = 'CA';
            zip = '';
            email = '';
            try {
                website = ele.querySelector('div.Rwjeuc > div > a').getAttribute('href').replace('/url?q=', '');
            } catch (error) {
                website = '';
            }
            
            return [
                name,
                street,
                city,
                state,
                zip,
                phone,
                email,
                website,
                rate,
                reviews,
            ];
            
        });
    }''')

    #print(len(result_names))

    await browser.close()

    # Create a new workbook and select the active worksheet
    workbook = openpyxl.Workbook()
    worksheet = workbook.active
    # sheet = workbook['Sheet1']

    # Add some data to the worksheet
    worksheet['A1'] = 'Name'
    worksheet['B1'] = 'Street Address'
    worksheet['C1'] = 'City'
    worksheet['D1'] = 'State'
    worksheet['E1'] = 'ZIP Code'
    worksheet['F1'] = 'Phone Number'
    worksheet['G1'] = 'Email'
    worksheet['H1'] = 'Website'
    worksheet['I1'] = 'Average Rating'
    worksheet['J1'] = 'Number of Reviews'

    for res in result_names:
        worksheet.append(res)
    # worksheet.append(['Jane', 30, 'London'])
    # worksheet.append(['Bob', 40, 'Paris'])

    # Save the workbook
    workbook.save('sheet.xlsx')

asyncio.run(main())
