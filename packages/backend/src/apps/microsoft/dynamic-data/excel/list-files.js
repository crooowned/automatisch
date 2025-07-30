export default {
    name: 'Excel-Dateien auflisten',
    key: 'listExcelFiles',

    async run($) {
        const files = {
            data: [],
            error: null,
        };

        try {
            const { folder } = $.step.parameters;

            // Wenn kein Ordner ausgewählt ist, Root-Ordner verwenden
            let nextLink = folder
                ? `https://graph.microsoft.com/v1.0/me/drive/items/${folder}/children`
                : 'https://graph.microsoft.com/v1.0/me/drive/root/children';

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
                    files.data.push(...data.value
                        .filter(item => item.name.endsWith('.xlsx')) // Nur Excel-Dateien
                        .map((file) => ({
                            value: file.id,
                            name: file.name
                        })));
                }

                nextLink = data['@odata.nextLink'];
            } while (nextLink);
        } catch (error) {
            files.error = error;
            console.error('Fehler beim Abrufen der Excel-Dateien:', error);
        }

        return files;
    },
};
