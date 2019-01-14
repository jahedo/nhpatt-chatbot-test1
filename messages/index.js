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

console.log(curl);

const LOCALE = 'es_ES';
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
                defaultLocale: 'es',
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

            setTimeout(() => builder.Prompts.number(session, 'No me gustaría que me hiciesen chatarra! 😯 ' +
                '¿me ayudas con una buena valoración? ' +
                'Del 1 al 5, siendo 1 muy poco satisfecho 😞 y 5 muuuuy satisfecho 😊'), 3000);
        },
        (session, results, next) => {

            logging.log({level: 'debug', message: 'Thanks...'});

            session.userData.valoration = results.response;
            let review = results.response < 3 ? '😞' : '😊';
            session.send(review + ' Muchas gracias!');
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

        session.send(['Te damos la bienvenida a Liferay Mutual! ¿Cómo puedo ayudarte?', 'Hola! ¿Cómo puedo ayudarte?',]);

        session.preferredLocale('es', err => logging.log({level: 'debug', message: JSON.stringify(err)}));
    }).matches('Greeting', [
        (session, results, next) => {

            logging.log({level: 'debug', message: 'Greeting!...'});

            if (session.conversationData.name) {
                next();
            } else {
                builder.Prompts.text(session, 'Hola, te puedo preguntar cómo te llamas?');
            }
        },
        (session) => {

            tryToLogin(session);

            if (!session.conversationData.name) {
                session.conversationData.name = session.message.text;
            }
            session.send([
                'Encantado de conocerte %s, ¿en qué puedo ayudarte? 😊',
                'Hola %s, bienvenido a Liferay Mutual. ¿En qué puedo ayudarte? 😊'
            ], session.conversationData.name);

            session.send('A día de hoy, te puedo decir que seguros puedes contratar o dar un parte');

        }]
    ).matches('Help', session => session.send('Has pedido ayuda... \'%s\'.', session.message.text)
    ).matches('Parte', [
        (session, results, next) => {

            logging.log({level: 'debug', message: 'Parte...'});

            if (results.entities && results.entities.length) {
                session.send('Ok, entendido, un parte de %s', results.entities[0].entity);
                next();
            } else {
                builder.Prompts.text(session, '¿Me puedes decir sobre qué tipo de seguro quieres dar de alta un parte?');
            }
        },
        session => {
            builder.Prompts.confirm(session, '¿Has tenido un accidente de tráfico?');
        },
        (session, results) => {

            logging.log({level: 'debug', message: 'Encuesta...'});

            session.send('Ok, no te preocupes de nada, en un par de minutos habremos acabado. 😉');
            session.send('Vamos a hacerte una serie de preguntas para poder ayudarte mejor');

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

                session.send('Ya hemos terminado %s, espero que haya sido rápido.', session.conversationData.name);

                timeout(session,
                    'Muchas gracias por la paciencia! En breve recibirás un correo electrónico con el ' +
                    'acuse de recibo del alta del parte. Además podrás consultar su estado desde la página web' +
                    ' o desde app, en el apartado de "Incidences".', 2000);

                timeout(session, [
                    'Recuerda que para cualquier duda estamos disponibles en el teléfono 666999999.',
                    'Si necesitas comunicar con nosotros durante la espara estamos disponibles en el teléfono 666999999 para cualquier consulta que requieras.',
                    'Recuerda instalarte nuestra app!'
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

            timeout(session, 'Muchas gracias por la paciencia!', 2000);
            setTimeout(() => session.send('Nos vemos pronto! 😊'), 4000);

            next();
        }
    ]).matches('Seguros', [
        (session) => {

            logging.log({level: 'debug', message: 'Seguros...'});

            timeout(session, 'Me alegra que me hagas esa pregunta, tenemos los mejores seguros de coches del mercado.', 1000);
            timeout(session, 'Disponemos de cuatro tipos de seguro de coche: Todo riesgo, a terceros, con franquicia y para coches clásicos.', 3000);
            timeout(session, 'Esta es la página donde podrás encontrar toda la información: https://liferay-insurances-demo.liferay.org.es/web/liferay-mutual/car-insurance/third-party-insurance', 5000);

            setTimeout(() => builder.Prompts.choice(session, 'Has encontrado algo que cuadre con lo que buscas?', ['Si', 'No']), 7000);
        },
        (session) => {

            session.sendTyping();
            setTimeout(() => {
                session.send('Encantado de haberte ayudado %s! :-D', session.conversationData.name || '');
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

        const dialogDatum = session.dialogData['BotBuilder.Data.WaterfallStep'] + 1;

        const label = dialogDatum + '/' + numberOfFields + ' - ' + field.label[LOCALE];
        writeEncouragingMessages(dialogDatum, session);

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
        session.send('Perfecto! Sin eso no habría podido darte de alta el parte :-J');
    } else if (dialogDatum === 7) {
        session.send('Gracias, ya estamos a punto de terminar.');
    } else if (session.userData.lastField && session.userData.lastField.dataType === 'date' && session.message.text) {
        if (session.message.text.toLowerCase() === 'hoy') {
            session.send('En breve llegará la asistencia técnica a ayudarte. ' +
                'Recibirás una notificación al teléfono móvil en el que podrás ver el camino que sigue la grúa hasta que se encuentre contigo.');
        } else {
            session.send('En breve recibirás un correo electrónico con el acuse de recibo del alta del parte. ' +
                'Además podrás consultar su estado desde la página web o desde app, en el apartado de "Incidences"');
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
        builder.Prompts.choice(session, label, choices.indexOf('Sí') !== -1 ? choiceSynonyms : choices);
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
