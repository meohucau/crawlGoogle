const https = require("https");
const http = require("http");
process.env['NODE_TLS_REJECT_UNAUTHORIZED'] = '0';


const {
    GoogleSpreadsheet
} = require('google-spreadsheet');

const creds = require('./kouenkai-b0c912fdac26.json');

let httpsPatern = /https:\/\/www/gmi;
let errPattern = /not found|404|みつかりません|見つかりません|302|found|error/gmi;

async function accessSpreadSheet() {
    const doc = new GoogleSpreadsheet('1EAlo8VhX2-zJHxueSWnJeWEbwUM2JE5-VNs-chir7R8');
    await doc.useServiceAccountAuth(creds);
    await doc.loadInfo(); // loads document properties and worksheets
    const sheet = doc.sheetsByTitle['Sheet1'];
    const linkList = await sheet.getRows();
    for (const key in linkList) {
        let options = new URL(linkList[key].LinkPDF);
        console.log(linkList[key].LinkPDF);
        let a = linkList[key].LinkPDF;
        if (linkList[key].LinkPDF.match(httpsPatern)) {
            let myRequest = https.request(options, res => {
                // Same as previos example
                res.on('data', d => {
                    if(d.toString().match(errPattern)){
                        linkList[key].CheckResult = "Died";
                        linkList[key].save();
                    }
                })
            })
            myRequest.on("error", console.error)
            myRequest.end()
        } else {
            let myRequest = http.request(options, res => {
                // Same as previos example
                res.on('data', d => {
                    //...
                    if(d.toString().match(errPattern)){
                        linkList[key].CheckResult = "Died";
                        linkList[key].save();
                    }
                })
                //... etc
            })

            myRequest.on("error", console.error)
            myRequest.end()
        }
    }
}



accessSpreadSheet();