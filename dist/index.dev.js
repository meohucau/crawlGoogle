"use strict";

// index.js

/**
 * Required External Modules
 */
var express = require("express");

var path = require("path");

var bodyParser = require("body-parser");

var googleIt = require("google-it");

var serp = require("serp");

var excel = require("node-excel-export");

require("chromedriver");

var webdriver = require('selenium-webdriver');

var _require = require("selenium-webdriver/chrome"),
    Driver = _require.Driver;

var promise = require('selenium-webdriver').promise;
/**
 * App Variables
 */


var app = express();
var port = process.env.PORT || "8000"; // var driver = new webdriver.Builder()
//   .forBrowser('chrome')
//   .build();

/**
 *  App Configuration
 */

app.set("views", path.join(__dirname, "views"));
app.set("view engine", "pug");
var urlencodedParser = bodyParser.urlencoded({
  extended: false
});
/**
 * Routes Definitions
 */

app.get("/", function (req, res) {
  res.render("index", {
    title: "Home"
  });
});
app.post("/", urlencodedParser, function _callee2(req, res) {
  var companyNameList, drugNameList, getResultProcess, styles, specification, dataset, report;
  return regeneratorRuntime.async(function _callee2$(_context3) {
    while (1) {
      switch (_context3.prev = _context3.next) {
        case 0:
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
          companyNameList = req.body.companyNameList.split("\n");
          drugNameList = req.body.drugNameList.split("\n");
          _context3.next = 4;
          return regeneratorRuntime.awrap(function _callee() {
            var lastResult, _loop, i;

            return regeneratorRuntime.async(function _callee$(_context2) {
              while (1) {
                switch (_context2.prev = _context2.next) {
                  case 0:
                    lastResult = [];

                    _loop = function _loop(i) {
                      var queryString, result, linkList;
                      return regeneratorRuntime.async(function _loop$(_context) {
                        while (1) {
                          switch (_context.prev = _context.next) {
                            case 0:
                              queryString = companyNameList[i] + drugNameList[i]; // with request options

                              _context.next = 3;
                              return regeneratorRuntime.awrap(getSearchResult(queryString));

                            case 3:
                              result = _context.sent;
                              linkList = "";

                              if (result.length > 0) {
                                result.forEach(function (value) {
                                  linkList = linkList + value.url + "\n";
                                });
                              }

                              lastResult.push({
                                companyName: companyNameList[i],
                                drugName: drugNameList[i],
                                link: linkList
                              });
                              ;

                            case 8:
                            case "end":
                              return _context.stop();
                          }
                        }
                      });
                    };

                    i = 0;

                  case 3:
                    if (!(i < companyNameList.length)) {
                      _context2.next = 9;
                      break;
                    }

                    _context2.next = 6;
                    return regeneratorRuntime.awrap(_loop(i));

                  case 6:
                    i++;
                    _context2.next = 3;
                    break;

                  case 9:
                    return _context2.abrupt("return", lastResult);

                  case 10:
                  case "end":
                    return _context2.stop();
                }
              }
            });
          }());

        case 4:
          getResultProcess = _context3.sent;
          // You can define styles as json object
          styles = {
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
          }; //Here you specify the export structure

          specification = {
            companyName: {
              // <- the key should match the actual data key
              displayName: 'Company Name',
              // <- Here you specify the column header
              headerStyle: styles.headerDark,
              // <- Header style
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
              cellStyle: styles.linkStyle,
              // <- Cell style
              width: 600 // <- width in pixels

            }
          }; // The data set should have the following shape (Array of Objects)
          // The order of the keys is irrelevant, it is also irrelevant if the
          // dataset contains more fields as the report is build based on the
          // specification provided above. But you should have all the fields
          // that are listed in the report specification

          dataset = [{
            companyName: 'IBM',
            drugName: 1,
            link: 'some note'
          }, {
            companyName: 'IBM',
            drugName: 1,
            link: 'some note'
          }, {
            companyName: 'IBM',
            drugName: 1,
            link: 'some note'
          }]; // Create the excel report.
          // This function will return Buffer

          report = excel.buildExport([// <- Notice that this is an array. Pass multiple sheets to create multi sheet report
          {
            name: 'Report',
            // <- Specify sheet name (optional)
            specification: specification,
            // <- Report specification
            data: getResultProcess // <-- Report data

          }]); // You can then return this straight

          res.attachment('report.xlsx'); // This is sails.js specific (in general you need to set headers)

          return _context3.abrupt("return", res.send(report));

        case 11:
        case "end":
          return _context3.stop();
      }
    }
  });
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
  var links = serp.search(options);
  return links;
}

app.listen(port, function () {
  console.log("Listening to ".concat(port));
});