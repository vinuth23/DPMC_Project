let currentTestCases = [];
let customFields = [];

const addFieldBtn = document.getElementById('addFieldBtn');
const fieldsContainer = document.getElementById('fieldsContainer');
const requirementInput = document.getElementById('requirementInput');
const generateBtn = document.getElementById('generateBtn');
const loadingIndicator = document.getElementById('loadingIndicator');
const errorMessage = document.getElementById('errorMessage');
const resultsSection = document.getElementById('resultsSection');
const statsSection = document.getElementById('statsSection');
const testCasesContainer = document.getElementById('testCasesContainer');
const exportBtn = document.getElementById('exportBtn');
const resetBtn = document.getElementById('resetBtn');

const uploadScreenshotsBtn = document.getElementById('uploadScreenshotsBtn');
const screenshotsInput = document.getElementById('screenshotsInput');
const screenshotsFileName = document.getElementById('screenshotsFileName');

const uploadRequirementDocBtn = document.getElementById('uploadRequirementDocBtn');
const requirementDocInput = document.getElementById('requirementDocInput');
const requirementDocFileName = document.getElementById('requirementDocFileName');

const uploadModificationExcelBtn = document.getElementById('uploadModificationExcelBtn');
const modificationExcelInput = document.getElementById('modificationExcelInput');

// Event Listeners
addFieldBtn.addEventListener('click', addFieldEntry);
generateBtn.addEventListener('click', handleGenerate);
exportBtn.addEventListener('click', handleExport);
resetBtn.addEventListener('click', handleReset);

// Upload buttons (UI only for now)
if (uploadScreenshotsBtn && screenshotsInput) {
    uploadScreenshotsBtn.addEventListener('click', () => screenshotsInput.click());
    screenshotsInput.addEventListener('change', () => {
        if (!screenshotsFileName) return;
        const files = screenshotsInput.files;
        if (!files || files.length === 0) {
            screenshotsFileName.textContent = 'No screenshots selected';
            return;
        }
        screenshotsFileName.textContent = `${files.length} screenshot(s) selected`;
    });
}

if (uploadRequirementDocBtn && requirementDocInput) {
    uploadRequirementDocBtn.addEventListener('click', () => requirementDocInput.click());
    requirementDocInput.addEventListener('change', () => {
        if (!requirementDocFileName) return;
        const file = requirementDocInput.files && requirementDocInput.files[0];
        requirementDocFileName.textContent = file ? file.name : 'No document selected';
    });
}

if (uploadModificationExcelBtn && modificationExcelInput) {
    uploadModificationExcelBtn.addEventListener('click', () => modificationExcelInput.click());
    modificationExcelInput.addEventListener('change', () => {
        // UI only for now: intentionally no behavior beyond selecting the file.
        // We still clear any errors to avoid confusing UX.
        hideError();
    });
}

function addFieldEntry() {
    const fieldId = Date.now(); // Unique ID for this field
    const fieldEntry = {
        id: fieldId,
        type: 'text',  // Default to text
        name: '',
        is_data_field: false  // Default to unchecked
    };
    customFields.push(fieldEntry);
    renderFields();
}

function renderFields() {
    fieldsContainer.innerHTML = '';
    
    customFields.forEach((field) => {
        const fieldDiv = document.createElement('div');
        fieldDiv.className = 'field-entry';
        
        // Build checkbox HTML only for TEXT fields
        const checkboxHTML = field.type === 'text' ? `
            <label class="field-checkbox-label">
                <input 
                    type="checkbox" 
                    class="field-data-checkbox" 
                    data-field-id="${field.id}"
                    ${field.is_data_field ? 'checked' : ''}
                />
                <span>Data Field</span>
            </label>
        ` : '';
        
        fieldDiv.innerHTML = `
            <select class="field-type-select" data-field-id="${field.id}">
                <option value="text" ${field.type === 'text' ? 'selected' : ''}>Text Field</option>
                <option value="amount" ${field.type === 'amount' ? 'selected' : ''}>Amount Field</option>
                <option value="button" ${field.type === 'button' ? 'selected' : ''}>Button</option>
            </select>
            <input 
                type="text" 
                class="field-name-input" 
                data-field-id="${field.id}" 
                placeholder="${field.type === 'button' ? 'Button name (e.g., Add, Delete, Submit)' : 'Field name/description (e.g., Customer Name, Loan Amount)'}"
                value="${field.name || ''}"
            />
            ${checkboxHTML}
            <button class="btn btn-danger btn-small" data-field-id="${field.id}">
                <span class="btn-icon"></span>Delete
            </button>
        `;
        fieldsContainer.appendChild(fieldDiv);
        
        // Add event listeners for this field
        const typeSelect = fieldDiv.querySelector('.field-type-select');
        const nameInput = fieldDiv.querySelector('.field-name-input');
        const dataCheckbox = fieldDiv.querySelector('.field-data-checkbox');
        const deleteBtn = fieldDiv.querySelector('.btn-danger');
        
        typeSelect.addEventListener('change', (e) => {
            updateFieldType(field.id, e.target.value);
            // Re-render to show/hide checkbox based on type
            renderFields();
        });
        
        nameInput.addEventListener('change', (e) => {
            updateFieldName(field.id, e.target.value);
        });
        
        if (dataCheckbox) {
            dataCheckbox.addEventListener('change', (e) => {
                updateFieldDataFlag(field.id, e.target.checked);
            });
        }
        
        deleteBtn.addEventListener('click', (e) => {
            deleteField(field.id);
        });
    });
}

function updateFieldType(fieldId, type) {
    const field = customFields.find(f => f.id === fieldId);
    if (field) {
        field.type = type;
    }
}

function updateFieldName(fieldId, name) {
    const field = customFields.find(f => f.id === fieldId);
    if (field) {
        field.name = name;
    }
}

function updateFieldDataFlag(fieldId, isDataField) {
    const field = customFields.find(f => f.id === fieldId);
    if (field) {
        field.is_data_field = isDataField;
    }
}

function deleteField(fieldId) {
    customFields = customFields.filter(f => f.id !== fieldId);
    renderFields();
}

async function handleGenerate() {
    const requirement = requirementInput.value.trim();
    const hasFields = customFields.some(f => f.name && f.name.trim());
    
    // Check if we have at least something to work with
    if (!requirement && !hasFields) {
        showError('Please either enter a requirement or add at least one custom field/button');
        return;
    }
    
    // Clear previous errors
    hideError();
    
    // Show loading indicator
    showLoading();
    
    try {
        const response = await fetch('/generate', {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json',
            },
            body: JSON.stringify({ 
                requirement: requirement,
                custom_fields: customFields
            })
        });
        
        const data = await response.json();
        
        if (!response.ok) {
            throw new Error(data.error || 'Failed to generate test cases');
        }
        
        // Store test cases
        currentTestCases = data.test_cases;
        
        // Display results
        displayTestCases(currentTestCases);
        
        // Update stats
        updateStats(currentTestCases);
        
    } catch (error) {
        showError(error.message);
    } finally {
        hideLoading();
    }
}

function displayTestCases(testCases) {
    // Clear container
    testCasesContainer.innerHTML = '';
    
    // Create cards for each test case
    testCases.forEach((testCase, index) => {
        const card = document.createElement('div');
        card.className = 'test-case-card';
        card.innerHTML = `
            <div class="test-case-header">
                <div class="test-case-id">${testCase.id || 'TC' + String(index + 1).padStart(2, '0')}</div>
                <div class="test-case-scenario">${testCase.scenario || 'Test Scenario'}</div>
            </div>
            <div class="test-case-content">
                <div class="test-case-field">
                    <div class="test-case-label">Expected Result</div>
                    <div class="test-case-value">${testCase.expected_result || 'N/A'}</div>
                </div>
            </div>
        `;
        testCasesContainer.appendChild(card);
    });
    
    // Show results section
    resultsSection.classList.remove('hidden');
    statsSection.classList.remove('hidden');
    
    // Scroll to results
    resultsSection.scrollIntoView({ behavior: 'smooth', block: 'start' });
}

function updateStats(testCases) {
    const totalCount = testCases.length;
    const positiveCount = testCases.filter(tc => 
        tc.scenario && (tc.scenario.toLowerCase().includes('valid') || 
                       tc.scenario.toLowerCase().includes('success') ||
                       tc.scenario.toLowerCase().includes('positive'))
    ).length;
    const negativeCount = testCases.filter(tc => 
        tc.scenario && (tc.scenario.toLowerCase().includes('invalid') || 
                       tc.scenario.toLowerCase().includes('error') ||
                       tc.scenario.toLowerCase().includes('negative') ||
                       tc.scenario.toLowerCase().includes('missing') ||
                       tc.scenario.toLowerCase().includes('unauthorized'))
    ).length;
    
    document.getElementById('totalTestCases').textContent = totalCount;
    document.getElementById('positiveCount').textContent = positiveCount;
    document.getElementById('negativeCount').textContent = negativeCount;
}

async function handleExport() {
    if (currentTestCases.length === 0) {
        showError('No test cases to export');
        return;
    }
    
    try {
        const response = await fetch('/export', {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json',
            },
            body: JSON.stringify({ test_cases: currentTestCases })
        });
        
        if (!response.ok) {
            throw new Error('Failed to export test cases');
        }
        
        // Create blob from response
        const blob = await response.blob();
        
        // Create download link
        const url = window.URL.createObjectURL(blob);
        const link = document.createElement('a');
        link.href = url;
        link.download = 'generated_test_cases.xlsx';
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
        window.URL.revokeObjectURL(url);
        
    } catch (error) {
        showError(error.message);
    }
}

function handleReset() {
    requirementInput.value = '';
    customFields = [];
    currentTestCases = [];
    testCasesContainer.innerHTML = '';
    resultsSection.classList.add('hidden');
    statsSection.classList.add('hidden');
    renderFields();
    hideError();
    requirementInput.focus();

    // Reset upload selections (UI only)
    if (screenshotsInput) screenshotsInput.value = '';
    if (screenshotsFileName) screenshotsFileName.textContent = 'No screenshots selected';
    if (requirementDocInput) requirementDocInput.value = '';
    if (requirementDocFileName) requirementDocFileName.textContent = 'No document selected';
    if (modificationExcelInput) modificationExcelInput.value = '';
}

function showLoading() {
    loadingIndicator.classList.remove('hidden');
}

function hideLoading() {
    loadingIndicator.classList.add('hidden');
}

function showError(message) {
    errorMessage.textContent = message;
    errorMessage.classList.remove('hidden');
}

function hideError() {
    errorMessage.classList.add('hidden');
}
