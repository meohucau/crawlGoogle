// index.js

/**
 * Required External Modules
 */
const express = require("express");
const path = require("path");
const bodyParser = require("body-parser");
const serp = require("serp");
const excel = require("node-excel-export");
require("chromedriver");
var webdriver = require('selenium-webdriver'),
    By = webdriver.By,
    until = webdriver.until;
map = webdriver.promise.map;
const {
    Driver
} = require("selenium-webdriver/chrome");
var promise = require('selenium-webdriver').promise;

/**
 * App Variables
 */
const app = express();
const port = process.env.PORT || "8000";
var driver = new webdriver.Builder()
    .forBrowser('chrome')
    .build();

/**
 *  App Configuration
 */
app.set("views", path.join(__dirname, "views"));
app.set("view engine", "pug");
var urlencodedParser = bodyParser.urlencoded({
    extended: false
})
app.use(express.static(__dirname + '/public'));
/**
 * Routes Definitions
 */
app.get("/", (req, res) => {
    res.render("index", {
        title: "Home"
    });
})

app.get("/resultList", (req, res) => {
    res.render("resultList", {
        title: "Home"
    });
})

app.post("/", urlencodedParser, async (req, res) => {
    //let companyName = req.body.companyName;
    let drugNameList = req.body.drugNameList.split("\n");
    let companyNameList = req.body.companyNameList.split("\n");
    let numberResult = req.body.numberResult;

    let pageCount;
    if (numberResult % 10 !== 0) {
        pageCount = Math.ceil(numberResult / 10);
    } else {
        pageCount = numberResult / 10;
    }

    let resultList = await getData();

    //createExcel(resultList);


    async function getData() {
        let getResultProcess = [];
        for (let i = 0; i < companyNameList.length; i++) {  
            let infoList = [];
            let linksList = [];
            let titlesList = [];
            let queryString = companyNameList[i] + " " + drugNameList[i];
            await driver.sleep(2000);
            await driver.get('https://google.jp');
            await driver.findElement(By.xpath("//*/input[@type='text']")).sendKeys(queryString, webdriver.Key.ENTER);
            let googlePageNum = await driver.findElements(By.xpath("//a[contains(@aria-label, 'Page')]"));
            loop2: for (let index = 1; index <= pageCount; index++) {
                let titleList = await driver.findElements(By.xpath("//*[@id='rso']//a/h3/span"));
                let linkList = await driver.findElements(By.xpath("//*[@id='rso']//a/h3/span/parent::h3/parent::a"));

                let titles = await map(titleList, e => e.getText())
                    .then(function (values) {
                        return values;
                    });
                let links = await map(linkList, e => e.getAttribute('href'))
                    .then(function (values) {
                        return values;
                    });

                linksList = linksList.concat(links);
                titlesList = titlesList.concat(titles);


                if (index < pageCount && index <= googlePageNum.length) {
                    await driver.sleep(2000);
                    await driver.get('https://google.jp');
                    await driver.findElement(By.xpath("//*/input[@type='text']")).sendKeys(queryString, webdriver.Key.ENTER);
                    await driver.sleep(2000);
                    await driver.findElement(By.xpath("//a[@aria-label='Page " + (index + 1) + "']")).click();
                } else {
                    break loop2;
                }
            }

            // let titleList = driver.findElements(By.xpath("//*[@id='rso']//a/h3/span"));
            // let linkList = driver.findElements(By.xpath("//*[@id='rso']//a/h3/span/parent::h3/parent::a"));

            linksList = await linksList.slice(0, numberResult);
            titlesList = await titlesList.slice(0, numberResult);

            infoList = await getInfoList(titlesList, linksList);

            getResultProcess.push({
                "companyName": companyNameList[i],
                "drugName": drugNameList[i],
                "infoList": infoList
            })
            //console.log(links);
        }
        return getResultProcess;
    }

    async function getInfoList(titles, links) {
        let result = []
        for (let i = 0; i < titles.length; i++) {
            result.push({
                "title": titles[i],
                "link": links[i]
            })
        }
        return result;
    }


    // function createExcel(resultList) {
    //     // You can define styles as json object
    //     const styles = {
    //         headerDark: {
    //             fill: {
    //                 fgColor: {
    //                     rgb: '00FE92'
    //                 }
    //             },
    //             font: {
    //                 color: {
    //                     rgb: '000000'
    //                 },
    //                 sz: 12
    //             }
    //         },
    //         linkStyle: {
    //             alignment: {
    //                 wrapText: true,
    //                 vertical: "center"
    //             }
    //         }
    //     };

    //     //Here you specify the export structure
    //     const specification = {
    //         companyName: { // <- the key should match the actual data key
    //             displayName: 'Company Name', // <- Here you specify the column header
    //             headerStyle: styles.headerDark, // <- Header style
    //             cellStyle: styles.linkStyle,
    //             width: 220 // <- width in pixels
    //         },
    //         drugName: {
    //             displayName: 'Drug Name',
    //             headerStyle: styles.headerDark,
    //             cellStyle: styles.linkStyle,
    //             // cellFormat: function (value, row) { // <- Renderer function, you can access also any row.property
    //             //     return (value == 1) ? 'Active' : 'Inactive';
    //             // },
    //             width: 200 // <- width in chars (when the number is passed as string)
    //         },
    //         title: {
    //             displayName: 'Title',
    //             headerStyle: styles.headerDark,
    //             cellStyle: styles.linkStyle, // <- Cell style
    //             width: 600 // <- width in pixels
    //         },
    //         link: {
    //             displayName: 'Link',
    //             headerStyle: styles.headerDark,
    //             cellStyle: styles.linkStyle, // <- Cell style
    //             width: 600 // <- width in pixels
    //         }
    //     }

    //     // Create the excel report.
    //     // This function will return Buffer
    //     const report = excel.buildExport(
    //         [ // <- Notice that this is an array. Pass multiple sheets to create multi sheet report
    //             {
    //                 name: 'Report', // <- Specify sheet name (optional)
    //                 specification: specification, // <- Report specification
    //                 data: resultList // <-- Report data
    //             }
    //         ]
    //     );

    //     // You can then return this straight
    //     res.attachment('report.xlsx'); 
    //return res.send(report);
    return res.render("result", {
        resultList: resultList
    }); //ADD HERE redirect o render
    //}


});

app.listen(port, () => {
    console.log(`Listening to ${port}`);
})