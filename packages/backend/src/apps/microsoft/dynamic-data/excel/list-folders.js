export default {
    name: 'OneDrive-Ordner auflisten',
    key: 'listExcelFolders',

    async run($) {
        const folders = {
            data: [],
            error: null,
        };

        try {
            let nextLink = 'https://graph.microsoft.com/v1.0/me/drive/root/children';

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
                    folders.data.push(...data.value
                        .filter(item => item.folder) // Nur Ordner
                        .map((folder) => ({
                            value: folder.id,
                            name: folder.name
                        })));
                }

                nextLink = data['@odata.nextLink'];
            } while (nextLink);
        } catch (error) {
            folders.error = error;
            console.error('Fehler beim Abrufen der Ordner:', error);
        }

        return folders;
    },
};
