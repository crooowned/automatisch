export default {
    name: 'List mailboxes',
    key: 'listMailboxes',

    async run($) {
        const mailboxes = {
            data: [],
            error: null,
        };

        try {
            const response = await $.http.get(`https://graph.microsoft.com/v1.0/users`, {
                headers: {
                    'Content-Type': 'application/json',
                },
                additionalProperties: {
                    skipAddingBaseUrl: true,
                }
            });

            if (!response.data.value)
                return [];

            mailboxes.data = response.data.value.map((mailbox) => ({
                value: mailbox.id,
                name: mailbox.displayName || mailbox.userPrincipalName
            }));
        } catch (error) {
            mailboxes.error = error;
            console.error("Error fetching mailboxes:", error); // Error logging
        }

        return mailboxes;
    },
};
