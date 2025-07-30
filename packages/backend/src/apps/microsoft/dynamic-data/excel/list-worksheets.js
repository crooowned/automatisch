export default {
    name: 'Excel-Arbeitsblätter auflisten',
    key: 'listWorksheets',

    async run($) {
        const worksheets = {
            data: [],
            error: null,
        };

        try {
            const { file } = $.step.parameters;

            if (!file) {
                return worksheets; // Keine Datei ausgewählt, leeres Ergebnis zurückgeben
            }

            const response = await $.http.get(`https://graph.microsoft.com/v1.0/me/drive/items/${file}/workbook/worksheets`, {
                headers: {
                    'Content-Type': 'application/json',
                },
                additionalProperties: {
                    skipAddingBaseUrl: true,
                }
            });

            if (response.data.value) {
                worksheets.data = response.data.value.map((sheet) => ({
                    value: sheet.id,
                    name: sheet.name
                }));
            }
        } catch (error) {
            worksheets.error = error;
            console.error('Fehler beim Abrufen der Arbeitsblätter:', error);
        }

        return worksheets;
    },
};
