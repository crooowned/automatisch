const hasValue = (value) => value !== null && value !== undefined;
export default {
    name: 'Excel-Spalten auflisten',
    key: 'listExcelColumns',

    async run($) {
        const columns = {
            data: [],
            error: null,
        };

        try {
            const { file, worksheet } = $.step.parameters;

            if (!hasValue(file) || !hasValue(worksheet)) {
                return columns; // Keine Datei oder Arbeitsblatt ausgewählt, leeres Ergebnis zurückgeben
            }

            console.log("Listing columns", file, worksheet);

            // Zuerst die Tabelle im Arbeitsblatt finden
            const tablesResponse = await $.http.get(
                `https://graph.microsoft.com/v1.0/me/drive/items/${file}/workbook/worksheets/${worksheet}/tables`,
                {
                    headers: {
                        'Content-Type': 'application/json',
                    },
                    additionalProperties: {
                        skipAddingBaseUrl: true,
                    }
                }
            );

            if (!tablesResponse.data.value || tablesResponse.data.value.length === 0) {
                throw new Error('Keine Tabelle im Arbeitsblatt gefunden');
            }

            const tableId = tablesResponse.data.value[0].id;

            // Dann die Spalten der Tabelle abrufen
            const columnsResponse = await $.http.get(
                `https://graph.microsoft.com/v1.0/me/drive/items/${file}/workbook/tables/${tableId}/columns`,
                {
                    headers: {
                        'Content-Type': 'application/json',
                    },
                    additionalProperties: {
                        skipAddingBaseUrl: true,
                    }
                }
            );

            if (columnsResponse.data.value) {
                columns.data = columnsResponse.data.value.map((column) => ({
                    value: column.name,
                    name: column.name
                }));
            }
        } catch (error) {
            columns.error = error;
            console.error('Fehler beim Abrufen der Spalten (data):', error);
        }

        console.log("Columns", columns);

        return columns;
    },
}; 