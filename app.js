var restify = require('restify');
var builder = require('botbuilder');
var parseXlsx = require('excel');

// Setup Restify Server
var server = restify.createServer();
server.listen(process.env.port || process.env.PORT || 3978, function () {
   console.log('%s listening to %s', server.name, server.url); 
});

// Create chat connector for communicating with the Bot Framework Service
var connector = new builder.ChatConnector({
    appId: process.env.MICROSOFT_APP_ID,
    appPassword: process.env.MICROSOFT_APP_PASSWORD
});

// Listen for messages from users 
server.post('/api/messages', connector.listen());

// Receive messages from the user and respond by echoing each message back (prefixed with 'You said:')
var bot = new builder.UniversalBot(connector, function (session) {
    readExcel(session.message.text,function (responseData) {
        session.send("Your WBS Code for the current period is: %s", responseData);
    });
 });

var readExcel = function (eid, callback) {
    console.log(eid);
    parseXlsx('./wbs_format.xlsx', function (err, data) {
        console.log(data);
        if (err)
            return console.log(err);
        if (data.length === 0) {
            callback('No data found');
        }
        for (var i = 0; i < data.length; i++) {
            if (data[i][2] === eid) {
                callback(data[i][7] + ' for the project:' + data[i][5]);
            }
        }
    });
}