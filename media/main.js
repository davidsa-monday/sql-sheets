const vscode = acquireVsCodeApi();
const grid = document.getElementById('config-grid');

// We'll receive the keys from the extension via reflection
let allKeys = [];
let descriptions = {};
let parameterTypes = {};

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