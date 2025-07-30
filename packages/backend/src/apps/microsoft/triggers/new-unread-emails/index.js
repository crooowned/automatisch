import defineTrigger from '../../../../helpers/define-trigger.js';
import Crypto from 'node:crypto';

export default defineTrigger({
    name: 'Neue ungelesene E-Mails',
    key: 'newUnreadEmails',
    pollInterval: 15,
    description: 'Wird ausgelöst, wenn neue ungelesene E-Mails empfangen werden.',
    arguments: [
        {
            label: 'Shared Mailbox',
            key: 'mailboxId',
            type: 'dropdown',
            required: false,
            description: 'Wählen Sie optional eine Shared Mailbox aus. Wenn keine ausgewählt wird, wird Ihre persönliche Mailbox verwendet.',
            variables: true,
            source: {
                type: 'query',
                name: 'getDynamicData',
                arguments: [
                    {
                        name: 'key',
                        value: 'listMailboxes'
                    }
                ]
            }
        },
        {
            label: 'Ordner',
            key: 'folderId',
            type: 'dropdown',
            required: false,
            description: 'Wählen Sie optional einen Ordner aus. Wenn keiner ausgewählt wird, wird der Posteingang verwendet.',
            variables: true,
            dependsOn: ['parameters.mailboxId'],
            source: {
                type: 'query',
                name: 'getDynamicData',
                arguments: [
                    {
                        name: 'key',
                        value: 'listFolders'
                    },
                    {
                        name: 'parameters.mailboxId',
                        value: '{parameters.mailboxId}'
                    }
                ]
            }
        }
    ],

    async run($) {
        const { mailboxId, folderId } = $.step.parameters;

        // Hole die verarbeiteten Mail-IDs aus dem Datastore
        const processedMailsStore = await $.datastore.get({
            key: 'processedMails'
        });
        const processedIds = processedMailsStore?.value || [];

        const params = new URLSearchParams({
            '$filter': 'isRead eq false',
            '$orderby': 'receivedDateTime desc',
            '$top': '50',
        });

        // Basis-URL für die API-Anfrage erstellen
        let baseUrl = 'https://graph.microsoft.com/v1.0';
        if (mailboxId) {
            baseUrl += `/users/${mailboxId}`;
        } else {
            baseUrl += '/me';
        }

        // Wenn ein Ordner ausgewählt wurde, suche in diesem Ordner
        if (folderId) {
            baseUrl += `/mailFolders/${folderId}`;
        }
        baseUrl += '/messages';

        const response = await $.http.get(`${baseUrl}?${params.toString()}`, {
            headers: {
                'Content-Type': 'application/json',
            },
            additionalProperties: {
                skipAddingBaseUrl: true,
            },
        });

        if (response.data.value?.length) {
            const newEmails = [];
            for (const mail of response.data.value) {
                // Überprüfe, ob mail und mail.id definiert sind
                if (mail && mail.id && !processedIds.includes(mail.id)) {
                    newEmails.push(mail);
                    // Mail als gelesen markieren
                    await $.http.patch(`${baseUrl}/${mail.id}`, {
                        isRead: true
                    }, {
                        headers: {
                            'Content-Type': 'application/json',
                        },
                        additionalProperties: {
                            skipAddingBaseUrl: true,
                        },
                    });
                    $.pushTriggerItem({
                        raw: mail,
                        meta: {
                            internalId: mail.id,
                        },
                    });
                }
            }

            if (newEmails.length > 0) {
                // Aktualisiere die Liste der verarbeiteten Mails
                await $.datastore.set({
                    key: 'processedMails',
                    value: [...processedIds, ...newEmails.map(mail => mail.id)]
                });
            }
        }
    },

    async testRun($) {
        const lastExecutionStep = await $.getLastExecutionStep();

        if (lastExecutionStep?.dataOut) {
            $.pushTriggerItem({
                raw: lastExecutionStep.dataOut,
                meta: {
                    internalId: Crypto.randomUUID(),
                },
            });
            return;
        }

        const sampleEmail = {
            id: Crypto.randomUUID(),
            subject: 'Test E-Mail',
            receivedDateTime: new Date().toISOString(),
            from: {
                emailAddress: {
                    name: 'Test Absender',
                    address: 'absender@beispiel.de'
                }
            },
            bodyPreview: 'Dies ist eine Test-E-Mail für den Microsoft Graph E-Mail-Trigger.',
            isRead: false,
            body: {
                content: 'Dies ist der vollständige Inhalt der Test-E-Mail.',
                contentType: 'text'
            }
        };

        $.pushTriggerItem({
            raw: sampleEmail,
            meta: {
                internalId: sampleEmail.id,
            },
        });
    }
});
