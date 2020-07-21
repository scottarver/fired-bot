const puppeteer = require('puppeteer');

const jobsurl = '';
console.log('starting');

module.exports = async function getJobs(){


    function extractItems(){
        const extractedElements = document.querySelectorAll('.oracletaleocwsv2-accordion-head');
        const items             = [];
        for (let element of extractedElements) {
            const [title, division, location, id] = element.innerText.split('\n')
            items.push({title, division, location, id});
        }
        return items;
    }


    async function scrapeInfiniteScrollItems(
        page,
        extractItems,
        itemTargetCount,
        scrollDelay = 1000,
    ){
        let items = [];
        try {
            let previousHeight;
            while (items.length < itemTargetCount) {
                items = await page.evaluate(extractItems);
                previousHeight = await page.evaluate('document.body.scrollHeight');
                await page.evaluate('window.scrollTo(0, document.body.scrollHeight)');
                await page.waitForFunction(`document.body.scrollHeight > ${previousHeight}`);
                await page.waitFor(scrollDelay);
            }
        } catch (e) {
            //nothing here, waitForFunction has to timeout
        }
        return items;
    }


    const browser = await puppeteer.launch({
        // executablePath: '/Applications/Google Chrome.app/Contents/MacOS/Google Chrome',
        // userDataDir   : '/Users/jtarver/Library/Application Support/Google/Chrome',
        headless: true
    });
    const page    = await browser.newPage();
    page.setViewport({width: 1280, height: 926});
    await page.goto(jobsurl);
    console.log('opened page');
    await page.waitFor(5000);

    const items = await scrapeInfiniteScrollItems(page, extractItems, 100);
    console.log('done');

    await browser.close();
    return items;


};

module.exports();

