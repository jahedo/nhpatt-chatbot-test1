'use strict';

const winston = require('winston');

const logging = winston.createLogger({
    level: 'debug',
    format: winston.format.json(),
    transports: [
        new winston.transports.File({filename: 'D:/home/site/wwwroot/debug.log', level: 'debug'}),
    ]
});

logging.log({level: 'debug', message: 'Starting log...'});

const builder = require('botbuilder');
const botBuilderAzure = require('botbuilder-azure');
const requestPromise = require('request-promise');
const curl = require('request-to-curl');
const promise = require('bluebird');
const locationDialog = require('botbuilder-location');
const path = require('path');

const LOCALE = 'en_US';

const USE_DEFAULT_PASSWORD = process.env.LIFERAY_USE_DEFAULT_PASSWORD || false;
const DEFAULT_USERNAME = process.env.LIFERAY_USER;
const DEFAULT_PASSWORD = process.env.LIFERAY_PASSWORD;
const LIFERAY_USER_PASSWORD = process.env.LIFERAY_USER_PASSWORD || process.env.USER_PASSWORD;
const USE_EMULATOR = process.env.NODE_ENV === 'localhost';
const SERVER_URL = (USE_EMULATOR ? 'http://localhost:8080/' : (process.env.LIFERAY_SERVER_URL || process.env.URL)) + '/api/jsonws/';

const LIFERAY_BING_KEY = process.env.LIFERAY_BING_KEY || process.env.BING_MAP;

const LIFERAY_STRUCTURE_ID = process.env.LIFERAY_STRUCTURE_ID || 157436;
const LIFERAY_GROUP_ID = process.env.LIFERAY_GROUP_ID || 20152;
const LIFERAY_RECORD_SET_ID = process.env.LIFERAY_RECORD_SET_ID || 271054;
const LIFERAY_REPOSITORY_ID = process.env.LIFERAY_REPOSITORY_ID || (LIFERAY_GROUP_ID || 20152);
const LIFERAY_FOLDER_ID = process.env.LIFERAY_FOLDER_ID || 184570;

logging.log({
    level: 'debug', message: `Environment variables: ${
        JSON.stringify({
            LIFERAY_STRUCTURE_ID,
            LIFERAY_GROUP_ID,
            LIFERAY_RECORD_SET_ID,
            LIFERAY_REPOSITORY_ID,
            LIFERAY_FOLDER_ID
        })
        }`
});

try {

    logging.log({
        level: 'debug',
        message: `Dependencies ready with host ${SERVER_URL} and using emulator ${USE_EMULATOR}`
    });

    const connector = USE_EMULATOR ? new builder.ChatConnector() : new botBuilderAzure.BotServiceConnector({
        appId: process.env['MicrosoftAppId'],
        appPassword: process.env['MicrosoftAppPassword'],
        openIdMetadata: process.env['BotOpenIdMetadata']
    });

    logging.log({level: 'debug', message: `Connector initialized...`});

    const bot = new builder.UniversalBot(connector,
        {
            localizerSettings: {
                defaultLocale: 'en',
                botLocalePath: './locale'
            }
        }
    );

    logging.log({level: 'debug', message: `Bot initialized...`});

    const tableName = 'botdata';
    const azureTableClient = new botBuilderAzure.AzureTableClient(tableName, process.env['AzureWebJobsStorage']);
    const tableStorage = new botBuilderAzure.AzureBotStorage({gzipData: false}, azureTableClient);
    bot.set('storage', tableStorage);

    bot.localePath(path.join(__dirname, './locale'));

    const mapLibrary = locationDialog.createLibrary(LIFERAY_BING_KEY || '');
    mapLibrary.dialog('confirm-dialog', createDialog(), true);
    bot.library(mapLibrary);

    bot.dialog('survey', [
        session => {

            logging.log({level: 'debug', message: 'Survey...'});

            setTimeout(() => builder.Prompts.number(session, 'I would not like them to make me scrap! 😯 ' +
                'Can you help me with a good assessment? ' +
                'From 1 to 5, 1 being very little satisfied 😞 and 5 sooo satisfied 😊'), 3000);
        },
        (session, results, next) => {

            logging.log({level: 'debug', message: 'Thanks...'});

            session.userData.valoration = results.response;
            let review = results.response < 3 ? '😞' : '😊';
            session.send(review + ' Thank you very much!');
            next();
        }
    ]);

    logging.log({level: 'debug', message: `First dialog...`});

    const LUIS_APP_ID = process.env.LuisAppId;
    const LUIS_API_KEY = process.env.LuisAPIKey;
    const LUIS_API_HOSTNAME = process.env.LuisAPIHostName || 'westus.api.cognitive.microsoft.com';
    const LUIS_API_URL = 'https://' + LUIS_API_HOSTNAME + '/luis/v2.0/apps/' + LUIS_APP_ID + '?subscription-key=' + LUIS_API_KEY;

    const recognizer = new builder.LuisRecognizer(LUIS_API_URL);

    logging.log({level: 'debug', message: `LUIS settings... ${LUIS_API_URL}`});

    const intents = new builder.IntentDialog({recognizers: [recognizer]}).onBegin(session => {

        logging.log({level: 'debug', message: 'Hi!...'});

        session.conversationData.name = '';

        tryToLogin(session);

        session.send(['Welcome to Liferay Mutual! How can I help you?', 'Hello! How can I help you?',]);

        session.preferredLocale('en', err => logging.log({level: 'debug', message: JSON.stringify(err)}));
    }).matches('Greeting', [
        (session, results, next) => {

            logging.log({level: 'debug', message: 'Greeting!...'});

            if (session.conversationData.name) {
                next();
            } else {
                builder.Prompts.text(session, 'Hello, what is your name?');
            }
        },
        (session) => {

            tryToLogin(session);

            if (!session.conversationData.name) {
                session.conversationData.name = session.message.text;
            }
            session.send([
                'Nice to meet you %s, can I help you? 😊',
                'Hello %s, welcome to Liferay Mutual. How can I help you? 😊'
            ], session.conversationData.name);

            session.send('Today, I can tell you what Insurance you can Hire or make a CLAIM');

        }]
    ).matches('Help', session => session.send('You have asked for help... \'%s\'.', session.message.text)
    ).matches('Issue', [
        (session, results, next) => {

            logging.log({level: 'debug', message: 'Parte...'});

            if (results.entities && results.entities.length) {
                session.send('Ok, I understand, a part about %s', results.entities[0].entity);
                next();
            } else {
                builder.Prompts.text(session, 'Can you tell me about what kind of insurance do you want to register a claim?');
            }
        },
        (session) => {
            builder.Prompts.confirm(session, 'Have you had a traffic accident?');
        },
        (session, results) => {

            logging.log({level: 'debug', message: 'Encuesta...'});

            session.send('Ok, do not worry about anything, in a couple of minutes we wil have finished. 😉');
            session.send('We are going to ask you a series of questions to help you better');

            session.userData.type = results.response;

            post(session, 'ddm.ddmstructure/get-structure', {'structureId': LIFERAY_STRUCTURE_ID}).then(response => {

                logging.log({level: 'debug', message: `... response ...`});

                const message = JSON.parse(response);
                return JSON.parse(message.definition);
            }).then(function (result) {

                logging.log({level: 'debug', message: 'Result of getting structure...'});
                logging.log({level: 'debug', message: JSON.stringify(result)});

                let random = '' + Math.random();
                let numberOfFields = result.fields.length;

                session.userData.form = {};

                let dialogs = result.fields.map(field =>
                    (session, results, next) => createAndProcessFields(session, results, next, numberOfFields, field)
                );

                bot.dialog(random, dialogs);

                session.beginDialog(random);
            }).catch(err =>
                logging.log({level: 'debug', message: JSON.stringify(err)})
            )
        },
        (session, results, next) => {

            processResults(session, results).then(() => {
                logging.log({level: 'debug', message: 'Sending record...'});
                logging.log({level: 'debug', message: JSON.stringify(session.userData.form)});
                return post(session, 'ddl.ddlrecord/add-record',
                    {
                        groupId: LIFERAY_GROUP_ID,
                        recordSetId: LIFERAY_RECORD_SET_ID,
                        displayIndex: 0,
                        fieldsMap: JSON.stringify(session.userData.form)
                    }
                )
            }).then(() => {

                logging.log({level: 'debug', message: 'Fin!'});

                session.send('We have finished %s, I hope it was fast.', session.conversationData.name);

                timeout(session,
                    'Thank you for the pacience! Shortly you will receive an email with the acknowledgment of ' +
                    'receipt of the party. You can also check your status from the website or from the app, ' +
                    'in the "Issues" section.', 2000);

                timeout(session, [
                    'Remember that for any questions we are available at 666999999.',
                    'If you need to communicate with us during the spara we are available at 666999999 for any consultation you require.',
                    'Remember to install our app!'
                ], 4000);

                let random = Math.random();

                if (random > 0.5) {
                    setTimeout(() => session.beginDialog('survey'), 5000);
                } else {
                    next();
                }
            });
        },
        (session, results, next) => {

            timeout(session, 'Thank you for the pacience!', 2000);
            setTimeout(() => session.send('See you soon! 😊'), 4000);

            next();
        }
    ]).matches('Insurances', [
        (session) => {

            logging.log({level: 'debug', message: 'Seguros...'});

            timeout(session, 'I am glad you ask me that question, we have the best car insurance in the market.', 2000);
            timeout(session, 'We have four types of car insurance: All risk, third parties, franchise and classic cars.', 3000);
            timeout(session, 'This is the page where you can find all the information: https://liferay-insurances-demo.liferay.org.es//web/liferay-mutual/car-insurance/third-party-insurance', 5000);

            setTimeout(() => builder.Prompts.choice(session, 'Have you found something that matches what you are looking for?', ['Yes', 'No']), 7000);
        },
        (session) => {

            session.sendTyping();
            setTimeout(() => {
                session.send('Pleased to have helped you, %s! :-D', session.conversationData.name || '');
                session.sendTyping();
            }, 1000);

            setTimeout(() => session.beginDialog('survey'), 2000);
        },
    ]).matches('Cancel', (session) => {
        session.send('You reached Cancel intent, you said \'%s\'.', session.message.text);
        session.conversationData = {};
    }).onDefault(session => session.send('Sorry, I did not understand \'%s\'.', session.message.text));

    logging.log({level: 'debug', message: `Dialogs initialized...`});

    bot.dialog('/', intents);

    bot.on('conversationUpdate', function (message) {
        if (message.membersAdded) {
            message.membersAdded.forEach(function (identity) {
                if (identity.id === message.address.bot.id) {
                    bot.beginDialog(message.address, '/');
                }
            });
        }
    });

    logging.log({level: 'debug', message: `Ready...`});

    if (USE_EMULATOR) {
        const restify = require('restify');
        const server = restify.createServer();
        server.listen(3978, function () {
            console.log('test bot endpoint at http://localhost:3978/api/messages');
        });
        server.post('/api/messages', connector.listen());
    } else {
        module.exports = connector.listen();
    }

    logging.log({level: 'debug', message: `Go!...`});

} catch (e) {
    logging.log({level: 'debug', message: 'Error :(' + JSON.stringify(e)});
}

function createAndProcessFields(session, results, next, numberOfFields, field) {

    processResults(session, results).then(() => {

        const userData = session.userData;
        
        logging.log({level: 'debug', message: JSON.stringify(userData) });

        const dialogDatum = session.dialogData['BotBuilder.Data.WaterfallStep'] + 1;
             
        logging.log({level: 'debug', message: JSON.stringify(dialogDatum) });        
        logging.log({level: 'debug', message: JSON.stringify(dialogDatum) });
        logging.log({level: 'debug', message: JSON.stringify(numberOfFields) });
        logging.log({level: 'debug', message: JSON.stringify(field.label[LOCALE]) });

        const label = dialogDatum + '/' + numberOfFields + ' - ' + field.label[LOCALE];
        
        logging.log({level: 'debug', message: JSON.stringify(label) });
        
        writeEncouragingMessages(dialogDatum, session);
                     
        logging.log({level: 'debug', message: JSON.stringify("kk") });

        userData.lastField = field;

        createPrompts(session, label, field);

    }).catch(err => logging.log({level: 'debug', message: JSON.stringify(err)}))
}

function processResults(session, results) {

    const userData = session.userData;
    if (!results || !results.response || !userData.lastField) {
        return promise.resolve();
    }

    const lastField = userData.lastField.name;

    const response = results.response;

    logging.log({level: 'debug', message: 'Parsing fields...'});
    logging.log({level: 'debug', message: JSON.stringify(response)});

    if (response.geo) {
        userData.form[lastField] = '{\"latitude\":' + response.geo.latitude + ', \"longitude\":' + response.geo.longitude + '}';
    } else if (response.resolution) {
        const d = response.resolution.start;
        userData.form[lastField] = `${d.getFullYear()}-${d.getMonth() + 1}-${d.getDate()}`;
    } else if (response.entity) {
        userData.form[lastField] = '[\"' + userData.lastField.options.filter(x => x.label[LOCALE] === response.entity)[0].value + '\"]';
    } else if (Array.isArray(response)) {

        const file = response[0];

        return requestPromise({encoding: null, uri: file.contentUrl}).then(function (response) {

            logging.log({level: 'debug', message: 'Retrieving file...'});
            logging.log({level: 'debug', message: JSON.stringify(file)});

            const randomNumber = ('' + Math.random()).substr(2);

            let extension = file.contentType === 'image/png' ? '.png' :
                file.contentType === 'image/jpg' ? 'jpg' : 'application/octet-stream';

            let fileName = file.name || (randomNumber + extension);

            const form = {
                repositoryId: LIFERAY_REPOSITORY_ID,
                folderId: LIFERAY_FOLDER_ID,
                sourceFileName: fileName,
                mimeType: file.contentType,
                title: fileName,
                description: '-',
                changeLog: '-',
                bytes: '[' + [...response].toString() + ']',
                'serviceContext.scopeGroupId': LIFERAY_GROUP_ID,
                'serviceContext.addGuestPermissions': true,
            };

            logging.log({level: 'debug', message: 'Adding file...'});
            logging.log({level: 'debug', message: JSON.stringify(form)});

            return post(session, 'dlapp/add-file-entry', form)
        }).then(function (response) {

            const obj = JSON.parse(response);

            logging.log({level: 'debug', message: 'Parsing file response...'});
            logging.log({level: 'debug', message: JSON.stringify(obj)});

            userData.form[userData.lastField.name] = JSON.stringify({
                groupId: LIFERAY_GROUP_ID,
                uuid: obj.uuid,
                version: '1.0',
                folderId: LIFERAY_FOLDER_ID,
                title: obj.fileName
            });

            logging.log({level: 'debug', message: 'Linking file...'});
            logging.log({level: 'debug', message: JSON.stringify(userData)});
        }).catch(
            err => logging.log({level: 'debug', message: JSON.stringify(err)})
        );
    } else {
        userData.form[lastField] = response;
    }
    return promise.resolve();
}

function writeEncouragingMessages(dialogDatum, session) {
    if (dialogDatum === 2) {
        session.send('Perfect! Without that I would not have been able to register the part 😊');
    } else if (dialogDatum === 7) {
        session.send('Thanks, we are about to finish.');
    } else if (session.userData.lastField && session.userData.lastField.dataType === 'date' && session.message.text) {
        if (session.message.text.toLowerCase() === 'hoy') {
            session.send('Shortly, technical assistance will come to help you. ' +
                'You will receive a notification to the mobile phone where you can see the path that the crane follows until it is with you.');
        } else {
            session.send('Shortly you will receive an email with the acknowledgment of receipt of the party. ' +
                'You can also check your status from the website or from the app, in the "Issues" section.');
        }
    }
}

function createPrompts(session, label, field) {
    if ('select' === (field.type)) {
        let choices = field.options.map(x => x.label[LOCALE]);
        const choiceSynonyms = [
            {value: 'Sí', synonyms: ['Si', 'Sí', 'Yes']},
            {value: 'No', synonyms: ['No', 'Nop']}
        ];
        builder.Prompts.choice(session, label, choices.indexOf('Yes') !== -1 ? choiceSynonyms : choices);
    } else if ('date' === (field.dataType)) {
        builder.Prompts.time(session, label);
    } else if ('document-library' === (field.dataType)) {
        builder.Prompts.attachment(session, label, {maxRetries: 0})
    } else if ('geolocation' === (field.dataType)) {
        locationDialog.getLocation(session, {
            prompt: label,
            requiredFields:
            locationDialog.LocationRequiredFields.locality
        });
    } else {
        builder.Prompts.text(session, label);
    }
}

function post(session, url, form) {

    const uri = SERVER_URL + url;
    const user = (!USE_DEFAULT_PASSWORD && session.userData && session.userData.username) || DEFAULT_USERNAME;
    const pass = (!USE_DEFAULT_PASSWORD && session.userData && session.userData.password) || DEFAULT_PASSWORD;

    logging.log({level: 'debug', message: `post... ${uri} with authentication ${user} and password ${pass}`});
    logging.log({level: 'debug', message: `... request ...`});

    const options = {
        method: 'POST',
        uri, form,
        auth: {user, pass, sendImmediately: true}
    };

    logging.log({level: 'debug', message: `... options ... ${JSON.stringify(options)}`});

    return requestPromise(options);
}

function tryToLogin(session) {
    const message = session.message;

    logging.log({level: 'debug', message: 'Trying login...'});
    logging.log({level: 'debug', message: JSON.stringify(message)});

    if (message && message.text && message.text.indexOf('start') !== -1) {
        session.userData.username = message.text.replace('/start ', '');
        session.userData.password = LIFERAY_USER_PASSWORD;
        session.sendTyping();
    }
}

function timeout(session, message, delay) {

    logging.log({level: 'debug', message: `delay`});

    session.sendTyping();
    setTimeout(() => {
        session.send(message);
        session.sendTyping();
    }, delay);
}

function createDialog() {
    return createBaseDialog().onBegin(function (session, args) {
        const confirmationPrompt = args.confirmationPrompt;
        session.send(confirmationPrompt).sendBatch();
    }).onDefault(function (session) {
        const message = parseBoolean(session.message.text);
        if (typeof message === 'boolean') {
            session.endDialogWithResult({response: {confirmed: message}});
            return;
        }
        session.send('InvalidYesNo').sendBatch();
    });
}

function createBaseDialog(options) {
    return new builder.IntentDialog(options).matches(/^cancel$/i, function (session) {
        session.send(consts_1.Strings.CancelPrompt);
        session.endDialogWithResult({response: {cancel: true}});
    }).matches(/^help$/i, function (session) {
        session.send(consts_1.Strings.HelpMessage).sendBatch();
    }).matches(/^reset$/i, function (session) {
        session.endDialogWithResult({response: {reset: true}});
    });
}

function parseBoolean(input) {
    input = input.trim();
    const yesExp = /^(y|si|sí|yes|yep|sure|ok|true)/i;
    const noExp = /^(n|no|nope|not|false)/i;
    if (yesExp.test(input)) {
        return true;
    } else if (noExp.test(input)) {
        return false;
    }
    return undefined;
}
