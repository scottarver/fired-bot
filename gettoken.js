const puppeteer = require('puppeteer');
const dotenv = require('dotenv');

dotenv.config();

const username = process.env.USERNAME;
const password = process.env.PASSWORD;

console.log('starting');

module.exports = async function getToken(){
    const browser = await puppeteer.launch({
        // executablePath: '/Applications/Google Chrome.app/Contents/MacOS/Google Chrome',
        // userDataDir   : '/Users/usernamehere/Library/Application Support/Google/Chrome',
        headless: false
    });
    const page    = await browser.newPage();
    await page.goto('https://developer.microsoft.com/graph/graph-explorer/');
    console.log('opened page');
    await page.waitForSelector('#ms-signin-button');
    console.log('found login button');
    await page.screenshot({path: 'example1.png'});
    await page.click('#ms-signin-button');
    await page.screenshot({path: 'example2.png'});
    console.log('clicked login button');

    // const pages = await browser.pages();
    // const popup = pages[pages.length - 1];
    // console.log('got popup');

    const login = page;
    await login.waitForSelector('#i0116');
    console.log('found username box');
    await login.click('#i0116');
    await login.keyboard.type(username);
    // await login.screenshot({path: 'example3.png'});

    await login.waitForSelector('#idSIButton9');
    await login.click('#idSIButton9');

    await login.waitFor(1000);
    await login.waitForSelector('#i0118', {visible: true});
    console.log('found password box');
    await login.click('#i0118');
    await login.keyboard.type(password);
    await login.click('#idSIButton9');
    console.log('submitted');
    await login.waitFor(10000);

    // await page.waitForSelector('input[value="Yes"]');
    // await page.click('input[value="Yes"]');
    // console.log('remember please');

    await page.waitForSelector('#userDisplayName');
    console.log('found username');

    await login.waitFor(10000);

    // let token;
    // page.on('console', msg => {
    //     if(msg._type === 'log' && msg._text && msg._text.length >2000 ){
    //         console.log('got token from console');
    //         token = msg._text;
    //     }
    // });

    const token = await page.evaluate('tokenPlease()');

    // await login.waitFor(10000);
    console.log(token);
    console.log('got token');
    console.log('closing, done');


    await browser.close();
    return token;
};

// module.exports();

