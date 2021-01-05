const {
    GoogleSpreadsheet
} = require('google-spreadsheet');
const {
    promisify
} = require('util');
const creds = require('./kouenkai-b0c912fdac26.json');
require("chromedriver");
var webdriver = require('selenium-webdriver'),
    By = webdriver.By,
    until = webdriver.until;
map = webdriver.promise.map;
const {
    Driver
} = require("selenium-webdriver/chrome");
const {
    stringify
} = require('querystring');
var promise = require('selenium-webdriver').promise;

var driver = new webdriver.Builder()
    .forBrowser('chrome')
    .build();

const ACTION = {
    CLICK: "クリック",
    GET_ALL: "すべて内容取得",
    GET_PDF_LINK: "PDFリンク取得",
    GET_TITLE: "タイトル取得",
    GET_TIME: "時間取得",
    GET_IMAGE: "画像取得"
}
var insertList = [];

async function accessSpreadSheet() {
    const doc = new GoogleSpreadsheet('1g-DBN4YyHrnAMnkWR3MY6r8datq5-Ar6c3wOfpDo4iA');
    await doc.useServiceAccountAuth(creds);
    await doc.loadInfo(); // loads document properties and worksheets
    const sheet = doc.sheetsByTitle['Sheet1'];
    // const sheet2 = await doc.addSheet({ headerValues: ['name', 'email'] });
    // const larryRow = await sheet.addRow({ 会社名: 'Larry Page', タイトル: 'larry@google.com' });
    // const moreRows = await sheet.addRows([
    // { 会社名: 'Sergey Brin', タイトル: 'sergey@google.com' },
    // { 会社名: 'Eric Schmidt', タイトル: 'eric@google.com' },
    // ]);
    const sheet2 = doc.sheetsByTitle['Sheet2'];
    await sheet2.clear();
    sheet2.setHeaderRow(["companyName", "title", "time", "content", "pdfLink", "imageLink"]);

    const robotList = await sheet.getRows();

    for (const key in robotList) {
        let stepCount = 1;
        let titleList = [];
        let timeList = [];
        let contentList = [];
        let pdfLinkList = [];
        let imageLinkList = [];
        await driver.get(robotList[key].URL);
        for (let index = 1; index <= 10; index++) {
            let xpath = `Xpath${stepCount}`;
            let action = `Action${stepCount}`;
            if (robotList[key][xpath] && robotList[key][action]) {
                switch (robotList[key][action]) {
                    case ACTION.CLICK:
                        await driver.findElement(By.xpath(robotList[key][xpath])).click();
                        break;
                    case ACTION.GET_ALL:
                        contentList = await getValue(robotList[key][xpath]);
                        break;
                    case ACTION.GET_TITLE:
                        titleList = await getValue(robotList[key][xpath]);
                        break;
                    case ACTION.GET_TIME:
                        timeList = await getValue(robotList[key][xpath]);
                        break;
                    case ACTION.GET_PDF_LINK:
                        pdfLinkList = await getLink(robotList[key][xpath]);
                        break;
                        case ACTION.GET_IMAGE:
                            imageLinkList = await getImage(robotList[key][xpath]);
                            break;
                    default:
                        break;
                }
                ++stepCount;
            } else {
                break;
            }
        }
        if(titleList.length > 0) {
            for (let i = 0; i < titleList.length; i++) {
                insertList.push({
                    companyName: robotList[key].Name,
                    title: titleList[i],
                    time: timeList[i],
                    content: contentList[i],
                    pdfLink: pdfLinkList[i],
                    imageLink: imageLinkList[i]
                })
            }
            await sheet2.addRows(insertList);
        }
        
    }
}

// await driver.get(robotList[0].URL);
// await driver.findElement(By.xpath(robotList[0].Xpath1)).click();


// let titleList = await driver.findElements(By.xpath(robotList[0].Xpath2));

// let titles = await map(titleList, e => e.getText())
//     .then(function (values) {
//         return values;
//     });

// let otherList = await driver.findElements(By.xpath(robotList[0].Xpath3));

// let others = await map(otherList, e => e.getText())
//     .then(function (values) {
//         return values;
//     });

// let sheet2 = doc.sheetsByTitle['Sheet2'];
// for (let index = 0; index < titles.length; index++) {
//     const addRow = await sheet2.addRow({
//         タイトル: titles[index],
//         内容: others[index]
//     })
// }


// rows[1].会社名 = 'hehehe';
// await rows[1].save();
// await rows[1].delete();
// const sheet = info.worksheets[0];
// console.log(`Title: ${sheet.title}, Rows: ${sheet.rowCount} `)
// }

async function getValue(xpath) {
    let resultList = await driver.findElements(By.xpath(xpath));
    let results = await map(resultList, e => e.getText())
        .then(function (values) {
            return values;
        });

    return results;
}

async function getLink(xpath) {
    let resultList = await driver.findElements(By.xpath(xpath));
    let results = await map(resultList, e => e.getAttribute('href'))
        .then(function (values) {
            return values;
        });

    return results;
}

async function getImage(xpath) {
    let resultList = await driver.findElements(By.xpath(xpath));
    let results = await map(resultList, e => e.getAttribute('src'))
        .then(function (values) {
            return values;
        });

    return results;
}


accessSpreadSheet();