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
module.exports.driver = driver;

module.exports.waitPageLoad = async function () {
    try {
        driver.wait(function () {
            return driver.executeScript('return document.readyState').then(function (readyState) {
                return readyState === 'complete';
            }, 5 * 1000);
        });
    } catch (error) {
        console.log(error);
    }
}

module.exports.getValue = async function (xpath) {
    try {
        let resultList = await driver.findElements(By.xpath(xpath));
        let results = await map(resultList, e => e.getText())
            .then(function (values) {
                return values;
            });
        return results;
    } catch (error) {
        return [];
    }
}

module.exports.getSingleValue = async function (xpathText) {
    try {
        let element = await driver.findElement(By.xpath(xpathText));
        return element.getText().then(function (text) {
            return text === null ? "" : text;
        });
    } catch (error) {
        return "";
    }
}

module.exports.displayHideElement = async function (xpath) {
    try {

        let resultList = await driver.findElements(By.xpath(xpath));
        resultList.forEach(async element => {
            await driver.executeScript("arguments[0].style.display = 'block';", element)
        })
    } catch (error) {
        console.log(error);
    }
}

module.exports.closeFrame = async function (xpath) {
    try {

        let element = await driver.findElement(By.xpath(xpath));
        await driver.executeScript("arguments[0].style.display = 'none';", element)
    } catch (error) {
        console.log(error);
    }
}

module.exports.getLink = async function (xpath) {
    try {

        let resultList = await driver.findElements(By.xpath(xpath));
        let results = await map(resultList, e => e.getAttribute('href'))
            .then(function (values) {
                return values;
            });

        return results;
    } catch (error) {
        return [];
    }
}

module.exports.getSingleLink = async function (xpathText) {
    try {
        let element = await driver.findElement(By.xpath(xpathText));
        return element.getAttribute('href').then(function (text) {
            return text === null ? "" : text;
        })
    } catch (error) {
        return "";
    }
}

module.exports.getImage = async function (xpath) {
    try {

        let resultList = await driver.findElements(By.xpath(xpath));
        let results = await map(resultList, e => e.getAttribute('src'))
            .then(function (values) {
                return values;
            });

        return results;
    } catch (error) {
        return [];
    }
}

module.exports.getSingleImage = async function (xpath) {
    try {
        let element = await driver.findElement(By.xpath(xpath));
        return element.getAttribute('src').then(function (text) {
            return text === null ? "" : text;
        })
    } catch (error) {
        return "";
    }
}


module.exports.waitPageLoad = async function () {
    try {

        driver.wait(function () {
            return driver.executeScript('return document.readyState').then(function (readyState) {
                return readyState === 'complete';
            });
        });
    } catch (error) {

    }
}

module.exports.displayHideElement = async function (xpath) {
    try {

        let resultList = await driver.findElements(By.xpath(xpath));
        resultList.forEach(async element => {
            await driver.executeScript("arguments[0].style.display = 'block';", element)
        })
    } catch (error) {

    }
}

module.exports.closeFrame = async function (xpath) {
    try {

        let element = await driver.findElement(By.xpath(xpath));
        await driver.executeScript("arguments[0].style.display = 'none';", element)
    } catch (error) {
        console.log(error);
    }
}