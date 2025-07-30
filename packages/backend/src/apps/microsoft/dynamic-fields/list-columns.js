const hasValue = (value) => value !== null && value !== undefined;

export default {
    name: 'Excel-Spalten auflisten',
    key: 'listExcelColumns',

    async run($) {
        if (!hasValue($.step.parameters.file) || !hasValue($.step.parameters.worksheet)) {
            return;
        }

        try {
            // Erst die Tabellen im Arbeitsblatt finden
            const tablesResponse = await $.http.get(
                `https://graph.microsoft.com/v1.0/me/drive/items/${$.step.parameters.file}/workbook/worksheets/${$.step.parameters.worksheet}/tables`,
                {
                    headers: {
                        'Content-Type': 'application/json'
                    },
                    additionalProperties: {
                        skipAddingBaseUrl: true
                    }
                }
            );

            let tableId;
            if (tablesResponse.data.value && tablesResponse.data.value.length > 0) {
                tableId = tablesResponse.data.value[0].id;
            } else {
                // Wenn keine Tabelle existiert, UsedRange verwenden und Tabelle erstellen
                const rangeResponse = await $.http.get(
                    `https://graph.microsoft.com/v1.0/me/drive/items/${$.step.parameters.file}/workbook/worksheets/${$.step.parameters.worksheet}/usedRange`,
                    {
                        headers: {
                            'Content-Type': 'application/json'
                        },
                        additionalProperties: {
                            skipAddingBaseUrl: true
                        }
                    }
                );

                if (!rangeResponse.data.address) {
                    return [];
                }

                // Tabelle aus dem verwendeten Bereich erstellen
                const createTableResponse = await $.http.post(
                    `https://graph.microsoft.com/v1.0/me/drive/items/${$.step.parameters.file}/workbook/worksheets/${$.step.parameters.worksheet}/tables/add`,
                    {
                        address: rangeResponse.data.address,
                        hasHeaders: true
                    },
                    {
                        headers: {
                            'Content-Type': 'application/json'
                        },
                        additionalProperties: {
                            skipAddingBaseUrl: true
                        }
                    }
                );

                tableId = createTableResponse.data.id;
                console.log('Erstellte Tabelle:', tableId);
            }

            // Spalten der Tabelle abrufen
            const columnsResponse = await $.http.get(
                `https://graph.microsoft.com/v1.0/me/drive/items/${$.step.parameters.file}/workbook/tables/${tableId}/columns`,
                {
                    headers: {
                        'Content-Type': 'application/json'
                    },
                    additionalProperties: {
                        skipAddingBaseUrl: true
                    }
                }
            );

            if (!columnsResponse.data.value) {
                return [];
            }

            // Spaltennamen aus der Antwort extrahieren
            return columnsResponse.data.value.map((column, index) => ({
                label: column.name,
                key: `header-${index}`,
                type: 'string',
                required: false,
                value: '',
                variables: true
            }));

        } catch (error) {
            console.error('Fehler beim Abrufen der Spalten (fields):', error);
            return [];
        }
    }
};
