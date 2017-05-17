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


// Information lists

var ServiceLabels = {
    SharePoint: 'SharePoint Online',
    OneDrive: 'OneDrive for Business',
    OfficeOnline: 'Office Online',
};

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

// let's create cards

function createThumbnailCard(session) {
    return new builder.ThumbnailCard(session)
        .title('Help Card')
        .subtitle('There are different areas we can help you with')
        .text('Which service do you need help for?')
        .images([
            builder.CardImage.create(session, 'https://sec.ch9.ms/ch9/7ff5/e07cfef0-aa3b-40bb-9baa-7c9ef8ff7ff5/buildreactionbotframework_960.jpg')
        ])
        .buttons([
            builder.CardAction.openUrl(session, 'https://docs.microsoft.com/bot-framework', 'OneDrive'),
            builder.CardAction.openUrl(session, 'https://docs.microsoft.com/bot-framework', 'SharePoint'),
            builder.CardAction.openUrl(session, 'https://docs.microsoft.com/bot-framework', 'Office Online')
        ]);
}

bot.dialog('/', intents);    

// intents.matches('Help',  (session) => {session.send('you need help');});

intents.matches('Help', [

    function (session) {
        // prompt for helpl options
    
          /* builder.Prompts.choice(
            session,
            'Which service do you need help for? ',
            [ServiceLabels.SharePoint, ServiceLabels.OneDrive, ServiceLabels.OfficeOnline],
            {
                maxRetries: 3,
                retryPrompt: 'You selected a wrong option'
            }) */
        session.send('you need help - fine lets see');
        // create the card based on selection
        console.log('Aufruf create card');
        var card = createCardThumbnailCard(session);
        console.log('Nach create card');
        // attach the card to the reply message
        var msg = new builder.Message(session).addAttachment(card);
        console.log('message created');
        session.send(msg);
        console.log('Message sent');



        }, /*,
    function (session, result) {
        if (!result.response) {
            // exhausted attemps and no selection, start over
            session.send('Ooops! Too many attemps :( But don\'t worry, I\'m handling that exception and you can try again!')
        }
        else {
            var selection = result.response.entity;
            switch (selection) {
                case ServiceLabels.SharePoint:
                    session.send('You selected SharePoint.');
                    break;
                case ServiceLabels.OfficeOnline:
                    session.send('You selected Office Online.');
                    break;
                case ServiceLabels.OneDrive:
                    session.send('You selected OneDrive.');
                    break;
            }
        }

        
    }*/
    function (session) {
        session.send('Done')
    }
        ]);

// This is the intents section 
//

intents.matches('Understand', [
    function(session, args) { 
        session.dialogData.entities = args.entities;
        
        var service = builder.EntityRecognizer.findEntity(args.entities, 'Service');

        
        if (service) {
            // session.send( 'You want to understand the service: "' + service.entity + '" - Cool !');
            if (service.entity == 'onedrive') {
                session.send('OneDrive is your personal store! It provides a lot of important features like external sharing.')
            } else if  (service.entity == 'sharepoint') {
                 session.beginDialog('/u_spo');
                 // session.send('now we are back from the spo dialog')
            } else if (service.entity == 'office 365') {
                session.send('Office 365 is a set of Online Services for better collaboration')
            }
        };

        var activity = builder.EntityRecognizer.findEntity(args.entities, 'Activity');
        
        if (activity) {
            // session.send( 'You want to understand the activity: "' + activity.entity + '" - Cool !');
       
            if (activity.entity == 'Sharing') {
                session.send('Sharing enables you to easily give others access to a document or folder.')
            } else if  (activity.entity == 'Co-Authoring') {
                 session.send('With this feature you can jointly edit a document. In the Online Version of Office even in real-time.')
            } else if (activity.entity == 'Version History') {
                session.send('Whenever a document is stored on OneDrive or SharePoint, the old version is stored in the version history.')
            }
         };
    }
    ]);


intents.matches('Search', (session) => {session.send('you want to search for ');});

intents.matches('Greeting', (session) => {session.send('Hallo, I am your Digital Workplace Bot for Office 365! Tell me what I can do for you');});

intents.matches('Learning', (session) => {session.send('The best place to learn is to go to the Intranet');});

// now there are the bot dialogs

bot.dialog('/u_spo', [
    function(session,args, next) {
        session.send('SharePoint Online is a service that supports collaboration in larger teams.');
        next();
    },
    function(session,args, next) {
        // session.send('now we leave the spo dialog.');
        session.endDialog();
    }

]);

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

