import defineAction from '../../../../helpers/define-action.js';

export default defineAction({
    name: 'Excel-Zeile aktualisieren',
    key: 'updateExcelRow',
    description: 'Aktualisiert eine existierende Zeile in einer Excel-Tabelle basierend auf einem Suchkriterium',
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
            description: 'Der Wert in der Suchspalte, der die zu aktualisierende Zeile identifiziert',
            variables: true
        },
        {
            label: 'Zu aktualisierende Spalten',
            key: 'updateColumns',
            type: 'dynamic',
            required: true,
            description: 'Wählen Sie die Spalten und Werte, die aktualisiert werden sollen',
            dependsOn: ['parameters.file', 'parameters.worksheet'],
            fields: [
                {
                    label: 'Spalte',
                    key: 'column',
                    type: 'dropdown',
                    required: true,
                    variables: true,
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
                    label: 'Wert',
                    key: 'value',
                    type: 'string',
                    required: true,
                    variables: true
                }
            ]
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

            // Zeile finden, die aktualisiert werden soll
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

            // Neue Werte vorbereiten
            const updatedValues = [...foundRow.values[0]];
            const updateColumns = $.step.parameters.updateColumns || [];
            
            for (const updateColumn of updateColumns) {
                const { column, value } = updateColumn;
                const columnIndex = columns.findIndex(col => col.name === column);
                if (columnIndex !== -1) {
                    updatedValues[columnIndex] = value;
                }
            }

            // Zeile aktualisieren
            const response = await $.http.patch(
                `https://graph.microsoft.com/v1.0/me/drive/items/${file}/workbook/tables/${tableId}/rows/itemAt(index=${foundRow.index})`,
                {
                    values: [updatedValues]
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

            // Ergebnisobjekt erstellen
            const resultObject = columns.reduce((obj, column, index) => {
                obj[column.name] = updatedValues[index];
                return obj;
            }, {});

            $.setActionItem({
                raw: {
                    success: true,
                    rowIndex: foundRow.index,
                    values: resultObject
                }
            });

            return true;
        } catch (error) {
            console.error('Fehler beim Aktualisieren der Excel-Zeile:', error);
            throw error;
        }
    }
});
