const vscode = acquireVsCodeApi();
const grid = document.getElementById('config-grid');

const allKeys = [
    'spreadsheet_id',
    'sheet_name',
    'start_cell',
    'start_named_range',
    'table_name'
];

window.addEventListener('message', event => {
    const message = event.data;
    switch (message.type) {
        case 'update':
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
            debugEl.textContent = 'Form submitted';
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

        const field = document.createElement('vscode-text-field');
        field.id = `field-${key}`;
        field.value = safeConfig[key] || '';

        field.addEventListener('change', (e) => {
            vscode.postMessage({
                type: 'edit',
                key: key,
                value: e.target.value
            });
        });

        form.appendChild(label);
        form.appendChild(field);
    });

}