/*-----------------------------------------------------------------------------
This template demonstrates how to use an IntentDialog with a LuisRecognizer to add 
natural language support to a bot. 
For a complete walkthrough of creating this type of bot see the article at
http://docs.botframework.com/builder/node/guides/understanding-natural-language/
-----------------------------------------------------------------------------*/
"use strict";
var builder = require("botbuilder");
var botbuilder_azure = require("botbuilder-azure");

var useEmulator = (process.env.NODE_ENV == 'development');

var connector = useEmulator ? new builder.ChatConnector() : new botbuilder_azure.BotServiceConnector({
    appId: process.env['MicrosoftAppId'],
    appPassword: process.env['MicrosoftAppPassword'],
    stateEndpoint: process.env['BotStateEndpoint'],
    openIdMetadata: process.env['BotOpenIdMetadata']
});

var bot = new builder.UniversalBot(connector);

// Make sure you add code to validate these fields
var luisAppId = process.env.LuisAppId;
var luisAPIKey = process.env.LuisAPIKey;
var luisAPIHostName = process.env.LuisAPIHostName || 'westus.api.cognitive.microsoft.com';

const LuisModelUrl = 'https://' + luisAPIHostName + '/luis/v1/application?id=' + luisAppId + '&subscription-key=' + luisAPIKey;

// Main dialog with LUIS
var recognizer = new builder.LuisRecognizer(LuisModelUrl);
var intents = new builder.IntentDialog({ recognizers: [recognizer] })
/*
.matches('<yourIntent>')... See details at http://docs.botframework.com/builder/node/guides/understanding-natural-language/
*/
.onDefault((session) => {
    session.send('Sorry, We did not understand \'%s\'.', session.message.text);
});

bot.dialog('/', intents);    

// intents.matches('Help',  (session) => {session.send('you need help');});

intents.matches('Help', [
    function(session) { 
        builder.Prompts.text(session, 'What do you want help for?');
    },
    function(session, results) {
        session.send('You need help for ' + results.response );
    }
    ]);

// This is the intents section 
//

intents.matches('Understand', [
    function(session, args) { 
        session.dialogData.entities = args.entities;
        
        var service = builder.EntityRecognizer.findEntity(args.entities, 'Service');
        var activity = builder.EntityRecognizer.findEntity(args.entities, 'Activity');
        
        if (service) {
            session.send( 'You want to understand the service: ' + service.entity + ' - Cool !');
            if (service.entity == 'onedrive') {
                session.send('OneDrive is your personal store! I provides a lot of important features like external sharing.')
            } else if  (service.entity == 'sharepoint') {
                 session.send('SharePoint is the place for teams.')
            } else if (service.entity == 'Office 365') {
                session.send('Office 365 is a set of Online Services for better collaboration')
            }
        };
        
        if (activity) {
            session.send( 'You want to understand the activity' + activity.entity + ' - Cool !');
        };
            if (activity.entity == 'Sharing') {
                session.send('Sharing enables you to easily give others access to a document or folder')
            } else if  (activity.entity == 'Co-authoring') {
                 session.send('With this feature you can jointly edit a document. In the Online Version of Office even in real-time.')
            } else if (activity.entity == 'Version History') {
                session.send('Whenever a document is stored on OneDrive or SharePoint, the old version is stored in the version history.')
            }
    }
    ]);


intents.matches('Search', (session) => {session.send('you want to search for ');});

intents.matches('Greeting', (session) => {session.send('Hallo, I am your Digital Workplace Bot for Office 365! Tell me what I can do for you');});

intents.matches('Learning', (session) => {session.send('The best place to learn is to go to the Intranet');});

if (useEmulator) {
    var restify = require('restify');
    var server = restify.createServer();
    server.listen(3978, function() {
        console.log('test bot endpont at http://localhost:3978/api/messages');
    });
    server.post('/api/messages', connector.listen());    
} else {
    module.exports = { default: connector.listen() }
}

