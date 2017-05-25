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

var ServiceFeatures = {
    coauthoring: 'Co-Authoring',
    sharing: 'Sharing Documents',
    controlaccess: 'Controll access to files and folders',
    workoffline: 'Working offline',
    exit: 'Exit',
}

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
       /* .images([
            builder.CardImage.create(session, 'https://sec.ch9.ms/ch9/7ff5/e07cfef0-aa3b-40bb-9baa-7c9ef8ff7ff5/buildreactionbotframework_960.jpg')
        ])*/
        .buttons([
            builder.CardAction.openUrl(session, 'http://microsoft.com/sharepoint', 'SharePoint'),
            builder.CardAction.openUrl(session, 'http://microsoft.com/office', 'Office Online'),
            builder.CardAction.openUrl(session, 'http://microsoft.com/onedrive', 'OneDrive')
        ]);
}

bot.dialog('/', intents);    

// intents.matches('Help',  (session) => {session.send('you need help');});

intents.matches('Help', [

    function (session) {
      
        session.send('you need help - fine lets see');
        // create the card based on selection
        console.log('Aufruf create card');
        var card = createThumbnailCard(session);
        console.log('Nach create card');
        // attach the card to the reply message
        var msg = new builder.Message(session).addAttachment(card);
        console.log('message created');
        session.send(msg);
        console.log('Message sent')
    }, 
    
    function (session) {
        session.send('Done')
    }
        ]);

// This is the intents section 
//

intents.matches('Understand', [
    function(session, args) { 
        session.dialogData.entities = args.entities;
        // if (!args.entities) {session.send('Sorry, I did not understand. :-(')};
      
        var service = builder.EntityRecognizer.findEntity(args.entities, 'Service');
        session.dialogData.service = service;
        var activity = builder.EntityRecognizer.findEntity(args.entities, 'Activity');
        session.dialogData.activity = activity;
        var scope = builder.EntityRecognizer.findEntity(args.entities, 'Scope');
        session.dialogData.scope = scope;
        if (scope) {session.send('Scope found'+ scope.entity)}
        else {
            session.send('no scope found');
        };

        if (service) {
            // session.send( 'You want to understand the service: "' + service.entity + '" - Cool !');
            if (service.entity == 'onedrive') {
                 session.beginDialog('/u_od');
            } else if  (service.entity == 'sharepoint') {
                 session.beginDialog('/u_spo');

            } else if ((service.entity == 'office 365') || (service.entity == 'o365' )) {
                session.send('Office 365 is a set of Online Services for better collaboration')
            }
        };


        
        if (activity) {
            // session.send( 'You want to understand the activity: "' + activity.entity + '" - Cool !');
       
            if  ((activity.entity == 'share') || (activity.entity == "Sharing")) {
                //session.send('Sharing enables you to easily give others access to a document or folder.')
                session.beginDialog('/u_share');
            } else if  ((activity.entity == 'co - author') || (activity.entity == 'co - authoring') || (activity.entity == 'joint editing'))  {
                 session.send('With this feature you can jointly edit a document. In the Online Version of Office even in real-time.')
            } else if (activity.entity == 'versioning') {
                session.send('Whenever a document is stored on OneDrive or SharePoint, the old version is stored in the version history.')
            } else if (activity.entity == 'migrate') {
                session.send('You don\'t need to migrate all your files to OneDrive. Just do it in a step-by-step approach.')
            } else if (activity.entity == 'synchronize') {
                session.send('In OneDrive click the sync button. In Sharepoint Online navigate the item and click the sync button.')
            } else {
                session.send('I did not get that, sorry!')
            }
         };
    }
    ]);


intents.matches('Search', (session) => {session.send('you want to search for something.... go to google :-)');});

intents.matches('None', (session) => {session.send('No intent found ... this if for debugging purposes')});

intents.matches('Greeting', (session) => {session.send('Hallo, I am your Digital Workplace Bot for Office 365! <br> <br>Tell me what I can do for you');});

intents.matches('Learning', (session) => {session.send('The best place to learn is to go to the Intranet');});

// now there are the bot dialogs

// SharePoint dialog

bot.dialog('/u_spo', [
    
    function (session, args, next) {
        builder.Prompts.choice(
            session,
            'SharePoint Online is a service that supports collaboration in larger teams.<br>Which SharePoint Online feature would you like to know? ',
            [ServiceFeatures.sharing, ServiceFeatures.controlaccess, ServiceFeatures.coauthoring, ServiceFeatures.workoffline],
            {
                maxRetries: 3,
                retryPrompt: 'You selected a wrong option! Try again.'
            }) ;
        
    },
    function (session, result, next) {
        if (!result.response){
            // exhausted attemps and no selection, start over
            session.send('Ooops! Too many attemps :( But don\'t worry, I\'m handling that exception and you can try again!')
        } else {
            var selection = result.response.entity;
            switch (selection) {
                case ServiceFeatures.sharing:
                    session.send('In SharePoint Online sharing information with others is done by just moving data to a document library and making sure that all relevant persons have access to the site.');
                    break;
                case ServiceFeatures.controlaccess:
                    session.send('In SharePoint Online access to elements is controlled by so-called SharePoint groups. The admin of a site can change access rights.');
                
                    break;
                case ServiceFeatures.coauthoring:
                    session.send('Co-authoring allows many users to work on the same document at the same time. In Office Online even in real-time');
                    
                    break;
                case ServiceFeatures.workoffline:
                    session.send('In the browser navigate to htts://portal.office.com. After log-in select SharePoint and navigate to the site and document library. There you click the Sync button and connect your PC to this SharePoint document library.');
                    
                    break;

            };
        };
        next();
    },

    function (session,args, next) {
        // session.send('now we leave the spo dialog.');
        // session.beginDialog('/u_spo');
        session.endDialog();
    }

]);

// OneDrive Dialog

bot.dialog('/u_od', [
    
       function (session, args, next) {
        builder.Prompts.choice(
            session,
            'OneDrive for Business is you personal store with the key features of Office 365.<br>Which OneDrive feature would you like to know? ',
            [ServiceFeatures.sharing, ServiceFeatures.controlaccess, ServiceFeatures.coauthoring, ServiceFeatures.workoffline],
            {
                maxRetries: 3,
                retryPrompt: 'You selected a wrong option! Try again.'
            }) ;
        
    },
    function (session, result, next) {
        if (!result.response){
            // exhausted attemps and no selection, start over
            session.send('Ooops! Too many attemps :( But don\'t worry, I\'m handling that exception and you can try again!')
        } else {
            var selection = result.response.entity;
            switch (selection) {
                case ServiceFeatures.sharing:
                    session.send('In OneDrive you share with others via the Share feature. You simply add the e-mail of you business partner and his or her access rights.');
                    break;
                case ServiceFeatures.controlaccess:
                    session.send('In OneDrive this is done with the Share feature, which also allows you to change access rights');
                
                    break;
                case ServiceFeatures.coauthoring:
                    session.send('Co-authoring allows many users to work on the same document at the same time. In Office Online even in real-time');
                    
                    break;
                case ServiceFeatures.workoffline:
                    session.send('In the browser navigate to htts://portal.office.com. After log-in select OneDrive. There you click the Sync button and connect your PC to OneDrive.');
                    
                    break;

            };
        };
        next();
    },

    function (session,args, next) {
        // session.send('now we leave the spo dialog.');
        // session.beginDialog('/u_spo');
        session.endDialog();
    }
]);

// Understanding Sharing Dialog

bot.dialog('/u_share', [
    
    function (session, args, next) {
       builder.Prompts.confirm(session,'Sharing enables you to easily give others access to a document or folder. <br>Do you want to know more about sharing documents?');
        
    },
    function (session, result, next) {
        if (!result.response){
            session.send('OK, what\'s next?');
            session.endDialog();
        } else {
            next();
        };
    },

    function (session, result, next) {
        var scopelocal = session.dialogData.scope;
        if (scopelocal.entity == "") {
        builder.Prompts.confirm(session, 'Do you want to share externally?');
        next();
        } else {
            session.send('found dialogdata scope' );
            if ((scopelocal.entity == 'external') || (scopelocal.entity == "externally")) {
                session.send('If you share externally, you need to look at the classification before you use the "Share" function');
            } else { if ((scopelocal.entity == 'internal') || (scopelocal.entity == 'internally')) {
                session.send('Fine, then you simply use the "Share" feature in either the browser or in Windows explorer');
            } };
            next();
        };
    },

    function (session, result, next) {
        // session.send('test: start check for session dialog data.')
        if (!result.response) { 
            session.send('Fine, then you simply use the "Share" feature in either the browser or in Windows explorer.');
            session.endDialog();
        } else {
            session.send('If you share externally, you need to look at the classification before you use the "Share" function.');
            next();
        };
    },

    function (session,args, next) {
        session.send('Go to the intranet to find more about classification.');
        session.endDialog();
    }

]);



// special Stuff

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

