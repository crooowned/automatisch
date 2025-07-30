import defineAction from '../../../../helpers/define-action.js';

export default defineAction({
    name: 'Excel-Zeile löschen',
    key: 'deleteExcelRow',
    description: 'Löscht eine Zeile in einer Excel-Tabelle basierend auf einem Suchkriterium',
    arguments: [
        {
            label: 'Ordner (optional)',
            key: 'folder',
            type: 'dropdown',
            required: false,
            description: 'Wählen Sie den OneDrive-Ordner aus. Leer lassen für Root-Ordner.',
            variables: true,
            source: {
                type: 'query',
                name: 'getDynamicData',
                arguments: [
                    {
                        name: 'key',
                        value: 'listExcelFolders'
                    }
                ]
            }
        },
        {
            label: 'Excel-Datei',
            key: 'file',
            type: 'dropdown',
            required: true,
            description: 'Wählen Sie die Excel-Datei aus',
            variables: true,
            source: {
                type: 'query',
                name: 'getDynamicData',
                arguments: [
                    {
                        name: 'key',
                        value: 'listExcelFiles'
                    }
                ]
            }
        },
        {
            label: 'Arbeitsblatt',
            key: 'worksheet',
            type: 'dropdown',
            required: true,
            description: 'Wählen Sie das Arbeitsblatt aus',
            variables: true,
            dependsOn: ['parameters.file'],
            source: {
                type: 'query',
                name: 'getDynamicData',
                arguments: [
                    {
                        name: 'key',
                        value: 'listWorksheets'
                    },
                    {
                        name: 'parameters.file',
                        value: '{{parameters.file}}'
                    }
                ]
            }
        },
        {
            label: 'Suchspalte',
            key: 'searchColumn',
            type: 'dropdown',
            required: true,
            description: 'Wählen Sie die Spalte aus, in der gesucht werden soll',
            variables: true,
            dependsOn: ['parameters.file', 'parameters.worksheet'],
            source: {
                type: 'query',
                name: 'getDynamicData',
                arguments: [
                    {
                        name: 'key',
                        value: 'listExcelColumns'
                    },
                    {
                        name: 'parameters.file',
                        value: '{{parameters.file}}'
                    },
                    {
                        name: 'parameters.worksheet',
                        value: '{{parameters.worksheet}}'
                    }
                ]
            }
        },
        {
            label: 'Suchwert',
            key: 'searchValue',
            type: 'string',
            required: true,
            description: 'Der Wert in der Suchspalte, der die zu löschende Zeile identifiziert',
            variables: true
        }
    ],

    async run($) {
        const { file, worksheet, searchColumn, searchValue } = $.step.parameters;

        if (!file || !worksheet || !searchColumn || !searchValue) {
            throw new Error('Datei, Arbeitsblatt, Suchspalte und Suchwert sind erforderlich');
        }

        try {
            // Tabelle im Arbeitsblatt finden
            const tablesResponse = await $.http.get(
                `https://graph.microsoft.com/v1.0/me/drive/items/${file}/workbook/worksheets/${worksheet}/tables`,
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
                throw new Error('Keine Tabelle im Arbeitsblatt gefunden');
            }

            // Spalten der Tabelle abrufen
            const columnsResponse = await $.http.get(
                `https://graph.microsoft.com/v1.0/me/drive/items/${file}/workbook/tables/${tableId}/columns`,
                {
                    headers: {
                        'Content-Type': 'application/json'
                    },
                    additionalProperties: {
                        skipAddingBaseUrl: true
                    }
                }
            );

            const columns = columnsResponse.data.value;
            const searchColumnIndex = columns.findIndex(col => col.name === searchColumn);

            if (searchColumnIndex === -1) {
                throw new Error(`Suchspalte "${searchColumn}" nicht gefunden`);
            }

            // Alle Zeilen der Tabelle abrufen mit Paginierung
            let allRows = [];
            let nextLink = `https://graph.microsoft.com/v1.0/me/drive/items/${file}/workbook/tables/${tableId}/rows?$top=1000`;

            while (nextLink) {
                const rowsResponse = await $.http.get(nextLink, {
                    headers: {
                        'Content-Type': 'application/json'
                    },
                    additionalProperties: {
                        skipAddingBaseUrl: true
                    }
                });

                allRows = allRows.concat(rowsResponse.data.value);

                // Prüfen ob es weitere Seiten gibt
                nextLink = rowsResponse.data['@odata.nextLink'] || null;

                // Optional: Abbrechen wenn der Wert gefunden wurde
                const foundInBatch = rowsResponse.data.value.find(row =>
                    String(row.values[0][searchColumnIndex]) === String(searchValue)
                );
                if (foundInBatch) {
                    allRows = rowsResponse.data.value; // Nur die aktuelle Seite behalten
                    break;
                }
            }

            // Zeile finden, die gelöscht werden soll
            const foundRow = allRows.find(row =>
                String(row.values[0][searchColumnIndex]) === String(searchValue)
            );

            if (!foundRow) {
                $.setActionItem({
                    raw: {
                        success: false,
                        message: `Keine Zeile mit dem Wert "${searchValue}" in der Spalte "${searchColumn}" gefunden`
                    }
                });
                return false;
            }

            // Ein Objekt mit den zu löschenden Werten erstellen (für die Rückgabe)
            const deletedValues = columns.reduce((obj, column, index) => {
                obj[column.name] = foundRow.values[0][index];
                return obj;
            }, {});

            // Zeile löschen
            await $.http.delete(
                `https://graph.microsoft.com/v1.0/me/drive/items/${file}/workbook/tables/${tableId}/rows/itemAt(index=${foundRow.index})`,
                {
                    headers: {
                        'Content-Type': 'application/json'
                    },
                    additionalProperties: {
                        skipAddingBaseUrl: true
                    }
                }
            );

            $.setActionItem({
                raw: {
                    success: true,
                    rowIndex: foundRow.index,
                    deletedValues: deletedValues
                }
            });

            return true;
        } catch (error) {
            console.error('Fehler beim Löschen der Excel-Zeile:', error);
            throw error;
        }
    }
});
