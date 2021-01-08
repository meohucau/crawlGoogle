const {GoogleSpreadsheet} = require('google-spreadsheet');
require('dotenv').config();
require("chromedriver");
var webdriver = require('selenium-webdriver'),
    By = webdriver.By
    map = webdriver.promise.map;

const chromeCapabilities = webdriver.Capabilities.chrome();
const chromeArgs = [
    '--disable-infobars',
    '--ignore-ssl-errors=yes',    
    '--ignore-certificate-errors',
    '--headless'
];
const chromeOptions = {
    args: chromeArgs,
    excludeSwitches: ['enable-logging'],
};
chromeCapabilities.set('chromeOptions', chromeOptions);

var driver = new webdriver.Builder()
    .forBrowser('chrome')
    .withCapabilities(chromeCapabilities)
    .build();

const creds = require(process.env.CREDS_PATH);
const ACTION = {
    CLICK: "Click",
    GET_ALL: "Get All Content",
    GET_PDF_LINK: "Get Link PDF",
    GET_TITLE: "Get Title",
    GET_TIME: "Get Time",
    GET_IMAGE: "Get Image"
}
var insertList = [];
async function accessSpreadSheet() {
    const doc = new GoogleSpreadsheet(process.env.SHEET_ID);
    await doc.useServiceAccountAuth(creds);
    console.log("Load sheet done!");
    await doc.loadInfo(); // loads document properties and worksheets
    const sheet = doc.sheetsByTitle[process.env.ROBOT_SHEET_NAME];
    console.log("Get robot sheet done!");
    const sheet2 = doc.sheetsByTitle[process.env.RESULT_SHEET_NAME];
    console.log("Get result sheet done!");
    await sheet2.clear();
    console.log("Clear result sheet done!");
    sheet2.setHeaderRow(["companyName", "title", "time", "content", "pdfLink", "imageLink"]);

    var robotCount = 0;
    const robotList = await sheet.getRows();

    for (const key in robotList) {
        ++robotCount;
        console.log(`Robot ${robotCount}: ${robotList[key].Name}`);
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
        if (titleList.length > 0) {
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
            console.log(`Robot ${robotCount}: ${robotList[key].Name} done!`, "\n");
        }else {
            consol.log(`Robot ${robotCount}: ${robotList[key].Name} doesn't has new data`, "\n")
        }

    }
    console.log("All robot done!")
}


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