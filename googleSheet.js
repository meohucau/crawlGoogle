const { GoogleSpreadsheet } = require('google-spreadsheet');
const { promisify } = require('util');
const creds = require('./kouenkai-b0c912fdac26.json');
require("chromedriver");
var webdriver = require('selenium-webdriver'),
    By = webdriver.By,
    until = webdriver.until;
map = webdriver.promise.map;
const {
    Driver
} = require("selenium-webdriver/chrome");
var promise = require('selenium-webdriver').promise;

var driver = new webdriver.Builder()
    .forBrowser('chrome')
    .build();


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

    const stepList = await sheet.getRows();
    //console.log(stepList[0]);
    await driver.get(stepList[0].URL);
    await driver.findElement(By.xpath(stepList[0].Xpath1)).click();


    let titleList = await driver.findElements(By.xpath(stepList[0].Xpath2));

    let titles = await map(titleList, e => e.getText())
        .then(function (values) {
            return values;
        });

    let otherList = await driver.findElements(By.xpath(stepList[0].Xpath3));

    let others = await map(otherList, e => e.getText())
        .then(function (values) {
            return values;
        });
        
    let sheet2 = doc.sheetsByTitle['Sheet2'];
    for (let index = 0; index < titles.length; index++) {
        const addRow = await sheet2.addRow({
            タイトル: titles[index],
            内容: others[index]
        })
    }
    

    // rows[1].会社名 = 'hehehe';
    // await rows[1].save();
    // await rows[1].delete();
    // const sheet = info.worksheets[0];
    // console.log(`Title: ${sheet.title}, Rows: ${sheet.rowCount} `)
}

accessSpreadSheet();