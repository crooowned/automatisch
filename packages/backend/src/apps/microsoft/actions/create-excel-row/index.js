import defineAction from '../../../../helpers/define-action.js';

export default defineAction({
    name: 'Excel-Zeile erstellen',
    key: 'createExcelRow',
    description: 'Erstellt eine neue Zeile in einer Excel-Tabelle',
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
            },
            additionalFields: {
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
        }
    ],

    async run($) {
        const { file, worksheet } = $.step.parameters;

        if (!file || !worksheet) {
            throw new Error('Datei und Arbeitsblatt sind erforderlich');
        }

        try {
            // Tabelle im Arbeitsblatt finden oder erstellen
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
                // Wenn keine Tabelle existiert, UsedRange verwenden und Tabelle erstellen
                const rangeResponse = await $.http.get(
                    `https://graph.microsoft.com/v1.0/me/drive/items/${file}/workbook/worksheets/${worksheet}/usedRange`,
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
                    throw new Error('Keine Daten im Arbeitsblatt gefunden');
                }

                // Tabelle aus dem verwendeten Bereich erstellen
                const createTableResponse = await $.http.post(
                    `https://graph.microsoft.com/v1.0/me/drive/items/${file}/workbook/worksheets/${worksheet}/tables/add`,
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
            }

            // Werte aus den header-X Parametern extrahieren
            const dataValues = Object.entries($.step.parameters)
                .filter((entry) => entry[0].startsWith('header-'))
                .map((value) => value[1]);

            const response = await $.http.post(
                `https://graph.microsoft.com/v1.0/me/drive/items/${file}/workbook/tables/${tableId}/rows/add`,
                {
                    values: [dataValues]
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

            $.setActionItem({
                raw: {
                    success: true,
                    rowIndex: response.data.index,
                    values: response.data.values
                }
            });

            return true;
        } catch (error) {
            console.error('Fehler beim Erstellen der Excel-Zeile:', error);
            throw error;
        }
    }
});
