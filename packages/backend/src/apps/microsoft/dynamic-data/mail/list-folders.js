export default {
    name: 'List folders',
    key: 'listFolders',

    async run($) {
        const folders = {
            data: [],
            error: null,
        };

        try {
            const { mailboxId } = $.step.parameters;

            let nextLink = 'https://graph.microsoft.com/v1.0';
            if (mailboxId) {
                nextLink += `/users/${mailboxId}`;
            } else {
                nextLink += '/me';
            }
            nextLink += '/mailFolders';

            do {
                const response = await $.http.get(nextLink, {
                    headers: {
                        'Content-Type': 'application/json',
                    },
                    additionalProperties: {
                        skipAddingBaseUrl: true,
                    }
                });

                const data = response.data;

                if (data.value) {
                    folders.data.push(...data.value.map((folder) => ({
                        value: folder.id,
                        name: folder.displayName
                    })));
                }

                nextLink = data['@odata.nextLink'];
            } while (nextLink);
        } catch (error) {
            folders.error = error;
            console.error("Error fetching folders:", error);
        }

        return folders;
    },
};
