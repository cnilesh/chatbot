var restify = require('restify');
var builder = require('botbuilder');
//var parseXlsx = require('excel');
//var Excel = require('exceljs');
var XLSX = require('xlsx');

// Setup Restify Server
var server = restify.createServer();
server.listen(process.env.port || process.env.PORT || 3978, function () {
   console.log('%s listening to %s', server.name, server.url); 
});

// Create chat connector for communicating with the Bot Framework Service
var connector = new builder.ChatConnector({
    appId: '02b6fc65-b6d9-4965-b279-5af051073019',
    appPassword: '=e|V-47@KZL@r=_c'
});

// Listen for messages from users 
server.post('/api/messages', connector.listen());

// Receive messages from the user and respond by echoing each message back (prefixed with 'You said:')
var bot = new builder.UniversalBot(connector, function (session) {
    //session.send('Hi');
    read(session.message.text,function (responseData) {
        session.send("Your WBS Code for the current period is: %s", responseData);
    });
    
 });

/*var readExcel = function (eid, callback) {
    console.log(eid);
    parseXlsx('https://pmochatbot.scm.azurewebsites.net/dev/api/files/wwwroot/wbs_format.xlsx', function (err, data) {
        if (err)
            return console.log(err);
        if (data.length === 0) {
            callback('No data found');
        }
        console.log(data.length);
        for (var i = 0; i < data.length; i++) {
            console.log(data[i][2]);
            if (data[i][2] === eid) {
                callback(data[i][7] + ' for the project:' + data[i][5]);
                break;
            }
        }
    });
}*/

var read = function (eid,callback) {
    var workbook = XLSX.readFile('wbs_format.xlsx');
    var sheet_name_list = workbook.SheetNames;
    //console.log(sheet_name_list);
    var text = XLSX.utils.sheet_to_csv(workbook.Sheets[sheet_name_list[0]]);
    var textArr = [];
    textArr = text.split(/\r?\n/);
    for (var i = 0; i < textArr.length; i++) {
        var tempArr = textArr[i].split(',');
        //console.log(tempArr);
        if (tempArr[2] == eid) {
            callback(tempArr[7] + ' for the project:' + tempArr[5]);
        }

    }

    // pipe from stream
    //var workbook = new Excel.Workbook();
    //stream.pipe(workbook.xlsx.createInputStream());
}