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
const { Driver } = require("selenium-webdriver/chrome");
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

/**
 * Routes Definitions
 */
app.get("/", (req, res) => {
    res.render("index", {
        title: "Home"
    });
})

app.post("/", urlencodedParser, async (req, res) => {
    let companyNameList = req.body.companyNameList.split("\n");
    let drugNameList = req.body.drugNameList.split("\n");

    let resultList = await getData();
    createExcel(resultList);

    async function getData() {
        let getResultProcess = [];
        for (let i = 0; i < companyNameList.length; i++) {
            let queryString = companyNameList[i] + " " + drugNameList[i];
            await driver.sleep(2000);
            await driver.get('https://google.jp');
            await driver.findElement(By.xpath("//*[@id='tsf']/div[2]/div[1]/div[1]/div/div[2]/input")).sendKeys(queryString, webdriver.Key.ENTER);
            let titleList = driver.findElements(By.xpath("//*[@id='rso']/div/div/div[1]/a/h3/span"));
            let linkList = driver.findElements(By.xpath("//*[@id='rso']/div/div/div[1]/a"));
            let titles = await map(titleList, e => e.getText())
                .then(function (values) {
                    return values;
                });

            let title = titles.join("\n");

            let links = await map(linkList, e => e.getAttribute('href'))
                .then(function (values) {
                    return values;
                });
            let link = links.join("\n");

            getResultProcess.push({
                "companyName": companyNameList[i],
                "drugName": drugNameList[i],
                "title": title,
                "link": link,
            })
            //console.log(links);
        }
        return getResultProcess;
    }


    function createExcel(resultList) {
        // You can define styles as json object
        const styles = {
            headerDark: {
                fill: {
                    fgColor: {
                        rgb: '00FE92'
                    }
                },
                font: {
                    color: {
                        rgb: '000000'
                    },
                    sz: 12
                }
            },
            linkStyle: {
                alignment: {
                    wrapText: true,
                    vertical: "center"
                }
            }
        };

        //Here you specify the export structure
        const specification = {
            companyName: { // <- the key should match the actual data key
                displayName: 'Company Name', // <- Here you specify the column header
                headerStyle: styles.headerDark, // <- Header style
                cellStyle: styles.linkStyle,
                width: 220 // <- width in pixels
            },
            drugName: {
                displayName: 'Drug Name',
                headerStyle: styles.headerDark,
                cellStyle: styles.linkStyle,
                // cellFormat: function (value, row) { // <- Renderer function, you can access also any row.property
                //     return (value == 1) ? 'Active' : 'Inactive';
                // },
                width: 200 // <- width in chars (when the number is passed as string)
            },
            title: {
                displayName: 'Title',
                headerStyle: styles.headerDark,
                cellStyle: styles.linkStyle, // <- Cell style
                width: 600 // <- width in pixels
            },
            link: {
                displayName: 'Link',
                headerStyle: styles.headerDark,
                cellStyle: styles.linkStyle, // <- Cell style
                width: 600 // <- width in pixels
            }
        }

        // Create the excel report.
        // This function will return Buffer
        const report = excel.buildExport(
            [ // <- Notice that this is an array. Pass multiple sheets to create multi sheet report
                {
                    name: 'Report', // <- Specify sheet name (optional)
                    specification: specification, // <- Report specification
                    data: resultList // <-- Report data
                }
            ]
        );

        // You can then return this straight
        res.attachment('report.xlsx'); 
        return res.send(report);
    }



});

app.listen(port, () => {
    console.log(`Listening to ${port}`);
})