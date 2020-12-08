// index.js

/**
 * Required External Modules
 */
const express = require("express");
const path = require("path");
const bodyParser = require("body-parser");
const googleIt = require("google-it");
const serp = require("serp");
const excel = require("node-excel-export");
require("chromedriver");
var webdriver = require('selenium-webdriver');
const { Driver } = require("selenium-webdriver/chrome");
var promise = require('selenium-webdriver').promise;

/**
 * App Variables
 */
const app = express();
const port = process.env.PORT || "8000";
// var driver = new webdriver.Builder()
//   .forBrowser('chrome')
//   .build();

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
    //await driver.get('https://google.jp');
    //await driver.findElement(webdriver.By.xpath("//*[@id='tsf']/div[2]/div[1]/div[1]/div/div[2]/input")).sendKeys("クリノリル錠50 日医工", webdriver.Key.ENTER);
    // let linkText = await driver.findElements(webdriver.By.xpath("//*[@id='rso']/div/div/div[1]/a"));
    // linkText.then(function (elements) {
    //     var pendingHtml = elements.map(function (elem) {
    //         return elem.getInnerHtml();
    //     });
    
    //     promise.all(pendingHtml).then(function (allHtml) {
    //         // `allHtml` will be an `Array` of strings
    //         console.log(allHtml);
    //     });
    // });
    // var pendingElements = driver.findElements(webdriver.By.xpath("//*[@id='rso']/div/div/div[1]/a"))

    // pendingElements.then(function (elements) {
    //     var pendingHtml = elements.map(function (elem) {
    //         return elem.getInnerHtml();
    //     });

    //     promise.all(pendingHtml).then(function (allHtml) {
    //         // `allHtml` will be an `Array` of strings
            
    // console.log(allHtml);
    //     });
    // });
    // linkText.forEach(element => {
    //     console.log(element.getText());
    // })
    
    let companyNameList = req.body.companyNameList.split("\n");
    let drugNameList = req.body.drugNameList.split("\n");
    let getResultProcess = await (async () => {
        let lastResult = [];
        for (let i = 0; i < companyNameList.length; i++) {

            let queryString = companyNameList[i] + drugNameList[i];
         

                // with request options
                let result = await getSearchResult(queryString);

                let linkList = "";

                if (result.length > 0) {
                    result.forEach(value => {
                        linkList = linkList + value.url + "\n";
                    });
                }
                lastResult.push({
                    companyName: companyNameList[i],
                    drugName: drugNameList[i],
                    link: linkList
                });;
        }
        return lastResult;
    })();

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
        link: {
            displayName: 'Link',
            headerStyle: styles.headerDark,
            cellStyle: styles.linkStyle, // <- Cell style
            width: 600 // <- width in pixels
        }
    }

    // The data set should have the following shape (Array of Objects)
    // The order of the keys is irrelevant, it is also irrelevant if the
    // dataset contains more fields as the report is build based on the
    // specification provided above. But you should have all the fields
    // that are listed in the report specification
    const dataset = [{
            companyName: 'IBM',
            drugName: 1,
            link: 'some note'
        },
        {
            companyName: 'IBM',
            drugName: 1,
            link: 'some note'
        },
        {
            companyName: 'IBM',
            drugName: 1,
            link: 'some note'
        }
    ]


    // Create the excel report.
    // This function will return Buffer
    const report = excel.buildExport(
        [ // <- Notice that this is an array. Pass multiple sheets to create multi sheet report
            {
                name: 'Report', // <- Specify sheet name (optional)
                specification: specification, // <- Report specification
                data: getResultProcess // <-- Report data
            }
        ]
    );

    // You can then return this straight
    res.attachment('report.xlsx'); // This is sails.js specific (in general you need to set headers)
    return res.send(report);

    // OR you can save this buffer to the disk by creating a file.



});
/**
 * Server Activation
 */

function getSearchResult(queryString) {
    var options = {
        host: "google.jp",
        qs: {
            q: queryString,
            filter: 0,
            pws: 0
        },
        num: 10
    };
    const links = serp.search(options)
    return links;
}
app.listen(port, () => {
    console.log(`Listening to ${port}`);
})