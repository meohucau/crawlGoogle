const {
    GoogleSpreadsheet
} = require('google-spreadsheet');
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
driver.manage().window().maximize();
const creds = require(process.env.CREDS_PATH);
const ACTION = {
    CLICK: "Click",
    GET_ALL: "Get All Content",
    GET_PDF_LINK: "Get PDF Link",
    GET_TITLE: "Get Title",
    GET_TIME: "Get Time",
    GET_IMAGE: "Get Image",
    GET_DETAIL_LINK: "Get Detail Link",
    GET_OTHER: "Get Other",
    DISPLAY_HIDE_ELEMENT: "Display Hide Element",
    NEXT_PAGE: "Next Page",
    CLOSE_FRAME: "Close Frame",
    SWITCH_TO_FRAME: "Switch To Frame",
    SWITH_TO_MAIN: "Switch To Main",
    LOOP: "Loop"
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
    sheet2.setHeaderRow(["companyName", "title", "time", "content", "pdfLink", "imageLink", "detailLink", "other"]);

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
        let detailLinkList = [];
        let otherList = [];
        await driver.get(robotList[key].URL);
        for (let index = 1; index <= 15; index++) {
            let xpath = `Xpath${stepCount}`;
            let action = `Action${stepCount}`;
            let temp = [];
            if (robotList[key][xpath] && robotList[key][action]) {
                if (robotList[key][action] == ACTION.LOOP) {
                    let arrStep = robotList[key][xpath].split('-');
                    for (let loop = 0; loop < Number(arrStep[2]); loop++) {
                        for (let i = Number(arrStep[0]); i <= Number(arrStep[1]); i++) {
                             xpath = `Xpath${i}`;
                             action = `Action${i}`;
                             await doAction();
                        }
                    }
                } else {
                    await doAction();
                }
                ++stepCount;
            } else {
                break;
            }
            async function doAction() {
                switch (robotList[key][action]) {
                    case ACTION.CLICK:
                        await driver.findElement(By.xpath(robotList[key][xpath])).click();
                        await driver.sleep(3000);
                        break;
                    case ACTION.NEXT_PAGE:
                        await driver.findElement(By.xpath(robotList[key][xpath])).click();
                        await driver.sleep(3000);
                        break;
                    case ACTION.GET_ALL:
                        temp = await getValue(robotList[key][xpath]);
                        contentList = contentList.concat(temp);
                        break;
                    case ACTION.GET_TITLE:
                        temp = await getValue(robotList[key][xpath]);
                        titleList = titleList.concat(temp);
                        break;
                    case ACTION.GET_TIME:
                        temp = await getValue(robotList[key][xpath]);
                        timeList = timeList.concat(temp);
                        break;
                    case ACTION.GET_PDF_LINK:
                        temp = await getLink(robotList[key][xpath]);
                        pdfLinkList = pdfLinkList.concat(temp);
                        break;
                    case ACTION.GET_IMAGE:
                        temp = await getImage(robotList[key][xpath]);
                        imageLinkList = imageLinkList.concat(temp);
                        break;
                    case ACTION.GET_DETAIL_LINK:
                        temp = await getLink(robotList[key][xpath]);
                        detailLinkList = detailLinkList.concat(temp);
                        break;
                    case ACTION.GET_OTHER:
                        temp = await getValue(robotList[key][xpath]);
                        otherList = otherList.concat(temp);
                        break;
                    case ACTION.DISPLAY_HIDE_ELEMENT:
                        await displayHideElement(robotList[key][xpath]);
                        break;
                    case ACTION.CLOSE_FRAME:
                        await closeFrame(robotList[key][xpath]);
                        break;
                    case ACTION.SWITCH_TO_FRAME:
                        let frame = await driver.findElement(By.xpath(robotList[key][xpath]));
                        await driver.switchTo().frame(frame);
                        break;
                    case ACTION.SWITH_TO_MAIN:
                        await driver.switchTo().defaultContent();
                        break;
                    default:
                        break;
                }
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
                    imageLink: imageLinkList[i],
                    detailLink: detailLinkList[i],
                    other: otherList[i]
                })
            }
            await sheet2.addRows(insertList);
            console.log(`Robot ${robotCount}: ${robotList[key].Name} done!`, "\n");
        } else {
            console.log(`Robot ${robotCount}: ${robotList[key].Name} doesn't has new data`, "\n")
        }

    }
    console.log("All robot done!");
    // await driver.quit();
}


async function getValue(xpath) {
    let resultList = await driver.findElements(By.xpath(xpath));
    let results = await map(resultList, async e => {
            await driver.executeScript("arguments[0].style.display = 'block';", e)
            return e.getText();
        })
        .then(function (values) {
            return values;
        });

    return results;
}

async function displayHideElement(xpath) {
    let resultList = await driver.findElements(By.xpath(xpath));
    resultList.forEach(async element => {
        await driver.executeScript("arguments[0].style.display = 'block';", element)
    })
}

async function closeFrame(xpath) {
    let element = await driver.findElement(By.xpath(xpath));
    await driver.executeScript("arguments[0].style.display = 'none';", element)
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