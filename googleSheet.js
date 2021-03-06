const {
    GoogleSpreadsheet
} = require('google-spreadsheet');
var webdriver = require('selenium-webdriver'),
    By = webdriver.By
map = webdriver.promise.map;
require('dotenv').config();
const {
    ACTION, STEP
} = require('./constants/action');
const {
    HEADER
} = require('./constants/header');
const seleniumAction = require('./function/seleniumAction');
const driver = seleniumAction.driver;

const creds = require(process.env.CREDS_PATH);
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
    sheet2.setHeaderRow(HEADER);

    var robotCount = 0;
    const robotList = await sheet.getRows();
    for (const key in robotList) {
        let stepCount = 1;
        let titleList = [];
        let timeList = [];
        let timeList2 = [];
        let contentList = [];
        let pdfLinkList = [];
        let imageLinkList = [];
        let detailLinkList = [];
        let typeList = [];
        let otherList = [];
        let otherList1 = [];
        let otherList2 = [];
        let elementList = [];
        if (robotList[key].URL !== undefined && robotList[key].URL !== null && robotList[key].URL !== '') {
            ++robotCount;
            console.log(`Robot ${robotCount}: ${robotList[key].Name}`);
            await driver.get(robotList[key].URL);
            await seleniumAction.waitPageLoad();
            for (let index = 1; index <= 15; index++) {
                let xpath = `${STEP.XPATH}${stepCount}`;
                let action = `${STEP.ACTION}${stepCount}`;
                let temp = [];
                if (robotList[key][xpath] && robotList[key][action]) {
                    if (robotList[key][action] === ACTION.LOOP) {
                        let arrStep = robotList[key][xpath].split('-');
                        for (let loop = 0; loop < Number(arrStep[2]); loop++) {
                            for (let i = Number(arrStep[0]); i <= Number(arrStep[1]); i++) {
                                xpath = `${STEP.XPATH}${i}`;
                                action = `${STEP.ACTION}${i}`;
                                await doAction();
                            }
                        }
                        index = arrStep[1];
                    } else if (robotList[key][action] === ACTION.FOR_EACH_ELEMENT) {
                        elementList = await driver.findElements(By.xpath(robotList[key][xpath]));
                        let arrStep = robotList[key][STEP.STEP_FOR_EACH].split('-');
                        for (let loop = 1; loop <= elementList.length; loop++) {
                            for (let i = Number(arrStep[0]); i <= Number(arrStep[1]); i++) {
                                xpath = `${STEP.XPATH}${i}`;
                                let xpathText = robotList[key][xpath].replace(STEP.VARIABLE, String(loop));
                                action = `${STEP.ACTION}${i}`;
                                await doEachAction(xpathText, elementList[loop - 1]);
                            }
                        }
                        index = arrStep[1];
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
                            await driver.sleep(2000);
                            try {
                                await driver.findElement(By.xpath(robotList[key][xpath])).click();
                            } catch (error) {

                            }
                            await seleniumAction.waitPageLoad();
                            await driver.sleep(2000);
                            break;
                        case ACTION.NEXT_PAGE:
                            await driver.sleep(2000);
                            try {
                                await driver.findElement(By.xpath(robotList[key][xpath])).click();
                            } catch (error) {

                            }
                            await driver.sleep(2000);
                            await seleniumAction.waitPageLoad();
                            break;
                        case ACTION.GET_ALL:
                            temp = await seleniumAction.getValue(robotList[key][xpath]);
                            contentList = contentList.concat(temp);
                            break;
                        case ACTION.GET_TITLE:
                            temp = await seleniumAction.getValue(robotList[key][xpath]);
                            titleList = titleList.concat(temp);
                            break;
                        case ACTION.GET_TIME:
                            temp = await seleniumAction.getValue(robotList[key][xpath]);
                            timeList = timeList.concat(temp);
                            break;
                        case ACTION.GET_TIME_2:
                            temp = await seleniumAction.getValue(robotList[key][xpath]);
                            timeList2 = timeList2.concat(temp);
                            break;
                        case ACTION.GET_PDF_LINK:
                            temp = await seleniumAction.getLink(robotList[key][xpath]);
                            pdfLinkList = pdfLinkList.concat(temp);
                            break;
                        case ACTION.GET_IMAGE:
                            temp = await seleniumAction.getImage(robotList[key][xpath]);
                            imageLinkList = imageLinkList.concat(temp);
                            break;
                        case ACTION.GET_DETAIL_LINK:
                            temp = await seleniumAction.getLink(robotList[key][xpath]);
                            detailLinkList = detailLinkList.concat(temp);
                            break;
                        case ACTION.GET_TYPE:
                            temp = await seleniumAction.getLink(robotList[key][xpath]);
                            typeList = typeList.concat(temp);
                            break;
                        case ACTION.GET_OTHER:
                            temp = await seleniumAction.getValue(robotList[key][xpath]);
                            otherList = otherList.concat(temp);
                            break;
                        case ACTION.GET_OTHER_1:
                            temp = await seleniumAction.getValue(robotList[key][xpath]);
                            otherList1 = otherList1.concat(temp);
                            break;
                        case ACTION.GET_OTHER_2:
                            temp = await seleniumAction.getValue(robotList[key][xpath]);
                            otherList2 = otherList2.concat(temp);
                            break;
                        case ACTION.DISPLAY_HIDE_ELEMENT:
                            await seleniumAction.displayHideElement(robotList[key][xpath]);
                            break;
                        case ACTION.CLOSE_FRAME:
                            await seleniumAction.closeFrame(robotList[key][xpath]);
                            break;
                        case ACTION.SWITCH_TO_FRAME:
                            let frame = await driver.findElement(By.xpath(robotList[key][xpath]));
                            await driver.switchTo().frame(frame);
                            break;
                        case ACTION.SWITH_TO_MAIN:
                            await driver.switchTo().defaultContent();
                            break;
                        case ACTION.SWITCH_TAB:
                            let windowHandles = await driver.getAllWindowHandles();
                            await driver.switchTo().window(windowHandles[1]);
                            break;
                        case ACTION.SCROLL_TO_BOTTOM:
                            await seleniumAction.scrollToBottom();
                        default:
                            break;
                    }
                }
                async function doEachAction(xpathText, element) {
                    switch (robotList[key][action]) {
                        case ACTION.CLICK:
                            try {    
                                await element.click();
                                await driver.sleep(1000);
                            } catch (error) {
                                console.log(error);
                            }
                        case ACTION.GET_TITLE:
                            let title = await seleniumAction.getSingleValue(xpathText);
                            titleList.push(title);
                            break;
                        case ACTION.GET_PDF_LINK:
                            let pdfLink = await seleniumAction.getSingleLink(xpathText);
                            pdfLinkList.push(pdfLink);
                            break;
                        case ACTION.GET_IMAGE:
                            let image = await seleniumAction.getSingleImage(xpathText);
                            imageLinkList.push(image);
                            break;
                        case ACTION.GET_DETAIL_LINK:
                            let detailLink = await seleniumAction.getSingleLink(xpathText);
                            detailLinkList.push(detailLink);
                            break;
                        case ACTION.GET_TIME:
                            let time = await seleniumAction.getSingleValue(xpathText);
                            timeList.push(time);
                            break;
                        case ACTION.GET_TIME_2:
                            let time2 = await seleniumAction.getSingleValue(xpathText);
                            timeList2.push(time2);
                            break;
                        case ACTION.GET_ALL:
                            let allContent = await seleniumAction.getSingleValue(xpathText);
                            contentList.push(allContent);
                            break;
                        case ACTION.GET_TYPE:
                            let type = await seleniumAction.getSingleValue(xpathText);
                            typeList.push(type);
                            break;
                        case ACTION.GET_OTHER:
                            let other = await seleniumAction.getSingleValue(xpathText);
                            otherList.push(other);
                            break;
                        case ACTION.GET_OTHER_1:
                            let other1 = await seleniumAction.getSingleValue(xpathText);
                            otherList1.push(other1);
                            break;
                        case ACTION.GET_OTHER_2:
                            let other2 = await seleniumAction.getSingleValue(xpathText);
                            otherList2.push(other2);
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
                        time2: timeList2[i],
                        content: contentList[i],
                        pdfLink: pdfLinkList[i],
                        imageLink: imageLinkList[i],
                        detailLink: detailLinkList[i],
                        type: typeList[i],
                        other: otherList[i],
                        other1: otherList1[i],
                        other2: otherList2[i]
                    })
                }
                await sheet2.addRows(insertList);
                insertList = [];
                console.log(`Robot ${robotCount}: ${robotList[key].Name} done!`, "\n");
            } else {
                console.log(`Robot ${robotCount}: ${robotList[key].Name} doesn't has new data`, "\n")
            }
        }
    }
    console.log("All robot done!");
    // await driver.quit();
}

accessSpreadSheet();