const vscode = acquireVsCodeApi();
const grid = document.getElementById('config-grid');

// We'll receive the keys from the extension via reflection
let allKeys = [];
let descriptions = {};
let parameterTypes = {};

// Extract spreadsheet + sheet identifiers from a Google Sheets URL
function parseGoogleSheetsUrl(value) {
    if (typeof value !== 'string') {
        return {};
    }

    const trimmed = value.trim();
    if (trimmed.length === 0) {
        return {};
    }

    const idMatch = trimmed.match(/\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/i);
    if (!idMatch) {
        return {};
    }

    const result = { spreadsheetId: idMatch[1] };
    const gidMatch = trimmed.match(/[?#&]gid=(\d+)/i);
    if (gidMatch) {
        result.sheetGid = gidMatch[1];
    }

    return result;
}

// Update the sheet identifier fields when a gid is discovered
function applySheetIdFromUrl(sheetGid) {
    if (sheetGid === undefined) {
        return;
    }

    const sheetIdField = document.getElementById('field-sheet_name-id');
    if (!sheetIdField) {
        return;
    }

    const sheetNameField = document.getElementById('field-sheet_name');
    sheetIdField.value = sheetGid;

    const trimmedName = sheetNameField && typeof sheetNameField.value === 'string'
        ? sheetNameField.value.trim()
        : '';
    const combinedValue = trimmedName ? `${sheetGid} | ${trimmedName}` : sheetGid;

    vscode.postMessage({
        type: 'edit',
        key: 'sheet_name',
        value: combinedValue
    });
}

window.addEventListener('message', event => {
    const message = event.data;
    switch (message.type) {
        case 'update':
            // Update our keys if provided
            if (message.keys && Array.isArray(message.keys)) {
                allKeys = message.keys;
            }

            // Store parameter descriptions for tooltips
            if (message.descriptions) {
                descriptions = message.descriptions;
            }

            // Store parameter types for field rendering
            if (message.types) {
                parameterTypes = message.types;
            }

            updateGrid(message.config);
            break;
    }
});

function updateGrid(config) {
    // Use a form-based approach for better editing experience
    grid.style.display = 'none'; // Hide the grid

    // Create or get our manual form
    let form = document.getElementById('manual-form');
    if (!form) {
        form = document.createElement('form');
        form.id = 'manual-form';
        form.style.display = 'grid';
        form.style.gridTemplateColumns = '150px 1fr';
        form.style.gap = '8px';
        form.style.alignItems = 'center';
        document.body.insertBefore(form, grid);

        // Handle form submission
        form.addEventListener('submit', (e) => {
            e.preventDefault();
        });
    }

    // Clear existing content
    form.innerHTML = '';

    // Ensure config is an object
    const safeConfig = config || {};

    // Add fields for each parameter
    allKeys.forEach(key => {
        const label = document.createElement('label');
        label.textContent = key;
        label.htmlFor = `field-${key}`;

        // Add tooltip if there's a description
        if (descriptions[key]) {
            label.title = descriptions[key];
            label.style.cursor = 'help';
        }

        let field;
        const type = parameterTypes[key];

        if (key === 'spreadsheet_id') {
            field = document.createElement('vscode-text-field');
            field.id = `field-${key}`;
            field.value = safeConfig[key] || '';
            field.placeholder = 'Spreadsheet ID or URL';
            if (descriptions[key]) {
                field.title = descriptions[key];
            }

            field.addEventListener('change', (e) => {
                const rawValue = e.target.value || '';
                const parsed = parseGoogleSheetsUrl(rawValue);
                let valueToPersist = rawValue;

                if (parsed.spreadsheetId) {
                    valueToPersist = parsed.spreadsheetId;
                    e.target.value = parsed.spreadsheetId;
                    applySheetIdFromUrl(parsed.sheetGid);
                }

                vscode.postMessage({
                    type: 'edit',
                    key: key,
                    value: valueToPersist
                });
            });

            form.appendChild(label);
            form.appendChild(field);
            return;
        }

        if (key === 'start_cell') {
            const namedRangeValue = safeConfig['start_named_range'] || '';
            const cellValue = safeConfig[key] || '';

            const container = document.createElement('div');
            container.className = 'start-cell-fields';

            const namedRangeField = document.createElement('vscode-text-field');
            namedRangeField.id = `field-${key}-named-range`;
            namedRangeField.value = namedRangeValue;
            namedRangeField.placeholder = 'Named range';
            if (descriptions['start_named_range']) {
                namedRangeField.title = descriptions['start_named_range'];
            }

            const cellField = document.createElement('vscode-text-field');
            cellField.id = `field-${key}`;
            cellField.value = cellValue;
            cellField.placeholder = 'Cell (e.g., A1)';
            if (descriptions[key]) {
                cellField.title = descriptions[key];
            }

            const emitCombinedUpdate = () => {
                const rangePart = (namedRangeField.value || '').trim();
                const cellPart = (cellField.value || '').trim();
                let combined = '';
                if (rangePart && cellPart) {
                    combined = `${rangePart} | ${cellPart}`;
                } else if (rangePart) {
                    combined = rangePart;
                } else {
                    combined = cellPart;
                }

                vscode.postMessage({
                    type: 'edit',
                    key: key,
                    value: combined
                });
            };

            namedRangeField.addEventListener('change', emitCombinedUpdate);
            cellField.addEventListener('change', emitCombinedUpdate);

            container.appendChild(namedRangeField);
            container.appendChild(cellField);

            form.appendChild(label);
            form.appendChild(container);
            return;
        }

        if (key === 'sheet_name') {
            const sheetIdValue = safeConfig['sheet_id'] !== undefined && safeConfig['sheet_id'] !== null
                ? String(safeConfig['sheet_id'])
                : '';
            const sheetNameValue = safeConfig[key] || '';

            const container = document.createElement('div');
            container.className = 'sheet-name-fields';

            const sheetIdField = document.createElement('vscode-text-field');
            sheetIdField.id = `field-${key}-id`;
            sheetIdField.value = sheetIdValue;
            sheetIdField.placeholder = 'Sheet ID (gid)';
            if (descriptions[key]) {
                sheetIdField.title = descriptions[key];
            }

            const sheetNameField = document.createElement('vscode-text-field');
            sheetNameField.id = `field-${key}`;
            sheetNameField.value = sheetNameValue;
            sheetNameField.placeholder = 'Sheet name';
            if (descriptions[key]) {
                sheetNameField.title = descriptions[key];
            }

            const emitSheetUpdate = () => {
                const idPart = (sheetIdField.value || '').trim();
                const namePart = (sheetNameField.value || '').trim();
                let combined = '';
                if (idPart && namePart) {
                    combined = `${idPart} | ${namePart}`;
                } else if (idPart) {
                    combined = idPart;
                } else {
                    combined = namePart;
                }

                vscode.postMessage({
                    type: 'edit',
                    key: key,
                    value: combined
                });
            };

            sheetIdField.addEventListener('change', emitSheetUpdate);
            sheetNameField.addEventListener('change', emitSheetUpdate);

            container.appendChild(sheetIdField);
            container.appendChild(sheetNameField);

            form.appendChild(label);
            form.appendChild(container);
            return;
        }

        // Create appropriate field based on parameter type
        if (type === 'boolean') {
            field = document.createElement('vscode-dropdown');
            field.id = `field-${key}`;

            const trueOption = document.createElement('vscode-option');
            trueOption.value = 'true';
            trueOption.textContent = 'True';

            const falseOption = document.createElement('vscode-option');
            falseOption.value = 'false';
            falseOption.textContent = 'False';

            field.appendChild(trueOption);
            field.appendChild(falseOption);

            // Set selected value based on any truthy value
            let currentValue = safeConfig[key];
            // Convert various string representations to boolean
            if (typeof currentValue === 'string') {
                currentValue = currentValue.toLowerCase();
                field.value = (currentValue === 'true' || currentValue === '1' || currentValue === 'yes') ? 'true' : 'false';
            } else {
                // Handle actual boolean value
                field.value = currentValue ? 'true' : 'false';
            }

            field.addEventListener('change', (e) => {
                vscode.postMessage({
                    type: 'edit',
                    key: key,
                    value: e.target.value
                });
            });
        } else {
            field = document.createElement('vscode-text-field');
            field.id = `field-${key}`;
            field.value = safeConfig[key] || '';

            // Add tooltip if there's a description
            if (descriptions[key]) {
                field.title = descriptions[key];
            }

            field.addEventListener('change', (e) => {
                vscode.postMessage({
                    type: 'edit',
                    key: key,
                    value: e.target.value
                });
            });
        }

        form.appendChild(label);
        form.appendChild(field);
    });

}
