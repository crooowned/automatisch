import defineAction from '../../../../helpers/define-action.js';

export default defineAction({
    name: 'E-Mail senden',
    key: 'sendEmail',
    description: 'Sendet eine E-Mail über Microsoft Graph.',
    arguments: [
        {
            label: 'Von (optional)',
            key: 'from',
            type: 'dropdown',
            required: false,
            description: 'Wählen Sie optional eine andere Absenderadresse aus (z.B. Shared Mailbox). Wenn keine ausgewählt wird, wird Ihre primäre E-Mail-Adresse verwendet.',
            variables: true,
            source: {
                type: 'query',
                name: 'getDynamicData',
                arguments: [
                    {
                        name: 'key',
                        value: 'listSendAsAddresses'
                    }
                ]
            }
        },
        {
            label: 'An',
            key: 'to',
            type: 'string',
            required: true,
            description: 'Die E-Mail-Adresse des Empfängers. Mehrere Adressen können durch Komma getrennt werden.',
            variables: true
        },
        {
            label: 'Betreff',
            key: 'subject',
            type: 'string',
            required: true,
            description: 'Der Betreff der E-Mail',
            variables: true
        },
        {
            label: 'Inhalt',
            key: 'body',
            type: 'string',
            required: true,
            description: 'Der Inhalt der E-Mail. HTML ist erlaubt.',
            variables: true
        }
    ],

    async run($) {
        const { from, to, subject, body } = $.step.parameters;

        // Basis-URL für die API-Anfrage erstellen
        let baseUrl = 'https://graph.microsoft.com/v1.0';
        if (from) {
            baseUrl += `/users/${from}`;
        } else {
            baseUrl += '/me';
        }
        baseUrl += '/sendMail';

        if (!to) {
            throw new Error('Die "An"-Adresse ist erforderlich.');
        }

        if (!subject) {
            throw new Error('Der Betreff der E-Mail ist erforderlich.');
        }

        if (!body) {
            throw new Error('Der Inhalt der E-Mail ist erforderlich.');
        }


        // E-Mail-Adressen aufbereiten
        const toRecipients = to.split(',').map(email => ({
            emailAddress: {
                address: email.trim()
            }
        }));

        const messageData = {
            message: {
                subject,
                body: {
                    contentType: 'HTML',
                    content: body
                },
                toRecipients
            },
            saveToSentItems: true
        };

        await $.http.post(baseUrl, messageData, {
            headers: {
                'Content-Type': 'application/json'
            },
            additionalProperties: {
                skipAddingBaseUrl: true
            }
        });

        // Rückgabedaten für nachfolgende Schritte
        $.setActionItem({
            raw: {
                subject,
                to: toRecipients.map(r => r.emailAddress.address).join(', '),
                from: from || 'Standard-Absender',
                body
            }
        });
        return true;
    }
});
