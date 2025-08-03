import defineTrigger from '../../../../helpers/define-trigger.js';
import Crypto from 'node:crypto';

export default defineTrigger({
    name: 'Neue ungelesene E-Mails',
    key: 'newUnreadEmails',
    pollInterval: 5,
    description: 'Wird ausgelöst, wenn neue ungelesene E-Mails empfangen werden.',
    arguments: [
        {
            label: 'Shared Mailbox',
            key: 'sharedMailbox',
            type: 'string',
            required: false,
            description: 'Tippen Sie ihre Shared Mailbox ein. Wenn keine angegeben wird, wird Ihre persönliche Mailbox verwendet.',
            variables: true
        },
        {
            label: 'Ordner',
            key: 'folderId',
            type: 'dropdown',
            required: false,
            description: 'Wählen Sie optional einen Ordner aus. Wenn keiner ausgewählt wird, wird der Posteingang verwendet.',
            variables: true,
            dependsOn: ['parameters.sharedMailbox'],
            source: {
                type: 'query',
                name: 'getDynamicData',
                arguments: [
                    {
                        name: 'key',
                        value: 'listFolders'
                    },
                    {
                        name: 'parameters.sharedMailbox',
                        value: '{parameters.sharedMailbox}'
                    }
                ]
            }
        },
        {
            label: 'Betreff enthält',
            key: 'subjectContains',
            type: 'string',
            required: false,
            description: 'Filtert E-Mails, deren Betreff den angegebenen Text enthält. Wenn leer gelassen, werden alle E-Mails verarbeitet.',
            variables: true
        }
    ],

    async run($) {
        const { sharedMailbox, folderId, subjectContains } = $.step.parameters;

        // Hole die verarbeiteten Mail-IDs aus dem Datastore
        const processedMailsStore = await $.datastore.get({
            key: 'processedMails'
        });
        const processedIds = processedMailsStore?.value || [];

        // Basis-URL für die API-Anfrage erstellen
        let baseUrl = 'https://graph.microsoft.com/v1.0';
        if (sharedMailbox) {
            baseUrl += `/users/${sharedMailbox}`;
        } else {
            baseUrl += '/me';
        }

        // Wenn ein Ordner ausgewählt wurde, suche in diesem Ordner
        if (folderId) {
            baseUrl += `/mailFolders/${folderId}`;
        }
        baseUrl += '/messages';

        const newEmails = [];
        let nextLink = null;
        const maxEmailsPerRequest = 20; // Reduziert von 50 auf 20
        const maxTotalEmails = 50; // Maximale Anzahl von E-Mails insgesamt

        do {
            // URL für aktuelle Anfrage erstellen
            let currentUrl;
            if (nextLink) {
                currentUrl = nextLink;
            } else {
                // Basis-Filter für ungelesene E-Mails
                const filter = 'isRead eq false';
                
                // Betreff-Filter wird clientseitig angewendet, um API-Komplexität zu vermeiden
                
                const params = new URLSearchParams({
                    '$filter': filter,
                    '$orderby': 'receivedDateTime desc',
                    '$top': maxEmailsPerRequest.toString(),
                });
                currentUrl = `${baseUrl}?${params.toString()}`;
            }

            const response = await $.http.get(currentUrl, {
                headers: {
                    'Content-Type': 'application/json',
                },
                additionalProperties: {
                    skipAddingBaseUrl: true,
                },
            });
            console.log('currentUrl', currentUrl);
            console.log('response', response);

            if (response.data.value?.length) {
                for (const mail of response.data.value) {
                    // Überprüfe, ob mail und mail.id definiert sind
                    if (mail && mail.id && !processedIds.includes(mail.id)) {
                        // Clientseitige Filterung nach Betreff
                        if (subjectContains && subjectContains.trim() !== '') {
                            const subject = mail.subject || '';
                            if (!subject.toLowerCase().includes(subjectContains.toLowerCase())) {
                                console.log('Betreff nicht übereinstimmt', subject, subjectContains);
                                continue; // E-Mail überspringen, wenn Betreff nicht übereinstimmt
                            }
                        }
                        
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
                        
                        // Fetch attachments if the email has any
                        if (mail.hasAttachments) {
                            try {
                                let attachmentsUrl = 'https://graph.microsoft.com/v1.0';
                                if (sharedMailbox) {
                                    attachmentsUrl += `/users/${sharedMailbox}`;
                                } else {
                                    attachmentsUrl += '/me';
                                }
                                attachmentsUrl += `/messages/${mail.id}/attachments`;
                                
                                const attachmentsResponse = await $.http.get(attachmentsUrl, {
                                    headers: {
                                        'Content-Type': 'application/json',
                                    },
                                    additionalProperties: {
                                        skipAddingBaseUrl: true,
                                    },
                                });
                                
                                if (attachmentsResponse.data.value?.length) {
                                    mail.attachments = attachmentsResponse.data.value;
                                }
                            } catch (error) {
                                // Log error but don't fail the entire process
                                console.error(`Failed to fetch attachments for email ${mail.id}:`, error);
                                mail.attachments = [];
                            }
                        }
                        
                        $.pushTriggerItem({
                            raw: mail,
                            meta: {
                                internalId: mail.id,
                            },
                        });

                        // Stoppe, wenn wir das Maximum erreicht haben
                        if (newEmails.length >= maxTotalEmails) {
                            break;
                        }
                    } else {
                        console.log('Mail nicht verarbeitet', mail.id, processedIds.includes(mail.id));
                    }
                }
            }

            // Prüfe auf weitere Seiten (Pagination)
            nextLink = response.data['@odata.nextLink'] || null;
            
            // Stoppe die Schleife, wenn wir das Maximum erreicht haben
            if (newEmails.length >= maxTotalEmails) {
                break;
            }
            
        } while (nextLink);

        if (newEmails.length > 0) {
            // Aktualisiere die Liste der verarbeiteten Mails
            await $.datastore.set({
                key: 'processedMails',
                value: [...processedIds, ...newEmails.map(mail => mail.id)]
            });
        }
    },

    async testRun($) {
        const { subjectContains } = $.step.parameters;
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

        // Erstelle einen Betreff basierend auf dem Filter
        let testSubject = 'Test E-Mail';
        if (subjectContains && subjectContains.trim() !== '') {
            testSubject = `Test E-Mail - ${subjectContains}`;
        }

        const sampleEmail = {
            id: Crypto.randomUUID(),
            subject: testSubject,
            receivedDateTime: new Date().toISOString(),
            from: {
                emailAddress: {
                    name: 'Test Absender',
                    address: 'absender@beispiel.de'
                }
            },
            bodyPreview: 'Dies ist eine Test-E-Mail für den Microsoft Graph E-Mail-Trigger.',
            isRead: false,
            hasAttachments: true,
            body: {
                content: 'Dies ist der vollständige Inhalt der Test-E-Mail.',
                contentType: 'text'
            },
            attachments: [
                {
                    id: 'attachment-1',
                    name: 'test-document.pdf',
                    contentType: 'application/pdf',
                    size: 1024,
                    isInline: false,
                    lastModifiedDateTime: new Date().toISOString()
                },
                {
                    id: 'attachment-2',
                    name: 'image.jpg',
                    contentType: 'image/jpeg',
                    size: 2048,
                    isInline: true,
                    lastModifiedDateTime: new Date().toISOString()
                }
            ]
        };

        $.pushTriggerItem({
            raw: sampleEmail,
            meta: {
                internalId: sampleEmail.id,
            },
        });
    }
});
