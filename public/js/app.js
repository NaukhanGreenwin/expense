// DOM Elements
const expenseForm = document.getElementById('expense-form');
const expensesList = document.getElementById('expenses-list');
const expensesTotal = document.getElementById('expenses-total');
const expenseSearch = document.getElementById('expense-search');
const pdfUploadForm = document.getElementById('pdf-upload-form');
const pdfUploadStatus = document.getElementById('pdf-upload-status');
const pdfResultsContainer = document.getElementById('pdf-results-container');
const loadingOverlay = document.getElementById('loading-overlay');

// State
let expenses = [];
let filteredExpenses = [];
let currentEditingExpense = null;
let currentSessionId = null;

// Initialize the app
document.addEventListener('DOMContentLoaded', initApp);

// Constants
const HST_RATE = 0.13; // 13% HST tax rate

function initApp() {
    // Load expenses from the API
    fetchExpenses();
    
    // Event listeners
    if (pdfUploadForm) {
        pdfUploadForm.addEventListener('submit', handlePdfUpload);
    }
    
    if (expenseSearch) {
        expenseSearch.addEventListener('input', applyFilters);
    }
    
    // Add event listener for export button
    const exportButton = document.getElementById('export-excel-btn');
    if (exportButton) {
        exportButton.addEventListener('click', handleExcelExport);
    }
    
    // Add event listener for PDF export button
    const exportPdfButton = document.getElementById('export-pdf-btn');
    if (exportPdfButton) {
        exportPdfButton.addEventListener('click', handlePdfExport);
    }
    
    // Hide the expense form - we only use PDF upload now
    const expenseFormContainer = document.getElementById('expense-form-container');
    if (expenseFormContainer) {
        expenseFormContainer.style.display = 'none';
    }
    
    // Create edit modal if it doesn't exist
    createEditModal();
}

// Create the edit modal structure
function createEditModal() {
    if (document.getElementById('edit-expense-modal')) return;
    
    const modal = document.createElement('div');
    modal.id = 'edit-expense-modal';
    modal.className = 'modal';
    
    modal.innerHTML = `
        <div class="modal-content">
            <span class="close-modal">&times;</span>
            <h3>Edit Expense</h3>
            <form id="edit-expense-form">
                <div class="form-group">
                    <label for="edit-title">Merchant/Title</label>
                    <input type="text" id="edit-title" required>
                </div>
                <div class="form-row">
                    <div class="form-group half">
                        <label for="edit-name">Name</label>
                        <input type="text" id="edit-name" placeholder="Enter your name">
                    </div>
                    <div class="form-group half">
                        <label for="edit-department">Department</label>
                        <input type="text" id="edit-department" placeholder="Enter your department">
                    </div>
                </div>
                <div class="form-row">
                    <div class="form-group half">
                        <label for="edit-amount">Amount ($)</label>
                        <input type="number" id="edit-amount" step="0.01" min="0" required>
                    </div>
                    <div class="form-group half">
                        <label for="edit-tax">Tax (13% HST)</label>
                        <input type="number" id="edit-tax" step="0.01" min="0" required>
                    </div>
                </div>
                <div class="form-row">
                    <div class="form-group half">
                        <label for="edit-date">Date</label>
                        <input type="date" id="edit-date" required>
                    </div>
                    <div class="form-group half">
                        <label for="edit-gl-code">G/L Code</label>
                        <select id="edit-gl-code" required>
                            <option value="6408-000">6408-000 (Office & General)</option>
                            <option value="6402-000">6402-000 (Membership)</option>
                            <option value="6404-000">6404-000 (Subscriptions)</option>
                            <option value="7335-000">7335-000 (Education)</option>
                            <option value="6026-000">6026-000 (Mileage/ETR)</option>
                            <option value="6010-000">6010-000 (Food & Ent.)</option>
                            <option value="6011-000">6011-000 (Social)</option>
                            <option value="6012-000">6012-000 (Travel)</option>
                            <option value="other">Other</option>
                        </select>
                        <input type="text" id="edit-custom-gl-code" placeholder="Enter custom G/L code" style="display: none; margin-top: 8px;">
                    </div>
                </div>
                <div class="form-group">
                    <label for="edit-description">Description</label>
                    <textarea id="edit-description" rows="3"></textarea>
                </div>
                <div class="form-actions">
                    <button type="submit" class="btn-primary">Save Changes</button>
                    <button type="button" class="btn-secondary" id="cancel-edit">Cancel</button>
                </div>
            </form>
        </div>
    `;
    
    document.body.appendChild(modal);
    
    // Add event listeners for the modal
    const closeBtn = modal.querySelector('.close-modal');
    const cancelBtn = modal.querySelector('#cancel-edit');
    const form = modal.querySelector('#edit-expense-form');
    
    closeBtn.addEventListener('click', closeEditModal);
    cancelBtn.addEventListener('click', closeEditModal);
    form.addEventListener('submit', saveEditedExpense);
    
    // Auto-calculate tax when amount changes
    const amountInput = document.getElementById('edit-amount');
    if (amountInput) {
        amountInput.addEventListener('input', function() {
            const amount = parseFloat(this.value) || 0;
            const taxInput = document.getElementById('edit-tax');
            if (taxInput) {
                taxInput.value = calculateTax(amount).toFixed(2);
            }
        });
    }
    
    // Show/hide custom G/L code input when "Other" is selected
    const glCodeSelect = document.getElementById('edit-gl-code');
    const customGlCodeInput = document.getElementById('edit-custom-gl-code');
    
    glCodeSelect.addEventListener('change', function() {
        if (this.value === 'other') {
            customGlCodeInput.style.display = 'block';
            customGlCodeInput.required = true;
        } else {
            customGlCodeInput.style.display = 'none';
            customGlCodeInput.required = false;
        }
    });
    
    // Close modal when clicking outside
    window.addEventListener('click', function(event) {
        if (event.target === modal) {
            closeEditModal();
        }
    });
}

// Open edit modal and populate with expense data
function openEditModal(expense) {
    currentEditingExpense = expense;
    
    document.getElementById('edit-title').value = expense.title || '';
    document.getElementById('edit-name').value = expense.name || '';
    document.getElementById('edit-department').value = expense.department || '';
    document.getElementById('edit-amount').value = expense.amount || 0;
    document.getElementById('edit-tax').value = expense.tax || calculateTax(expense.amount);
    document.getElementById('edit-date').value = expense.date || '';
    
    const glCodeSelect = document.getElementById('edit-gl-code');
    const customGlCodeInput = document.getElementById('edit-custom-gl-code');
    
    // Check if the G/L code is one of the predefined options
    const predefinedCodes = ['6408-000', '6402-000', '6404-000', '7335-000', '6026-000', '6010-000', '6011-000', '6012-000'];
    const glCode = expense.glCode || '6408-000';
    
    if (predefinedCodes.includes(glCode)) {
        glCodeSelect.value = glCode;
        customGlCodeInput.style.display = 'none';
        customGlCodeInput.required = false;
    } else {
        // If it's a custom code, select "Other" and show the custom input
        glCodeSelect.value = 'other';
        customGlCodeInput.value = glCode;
        customGlCodeInput.style.display = 'block';
        customGlCodeInput.required = true;
    }
    
    document.getElementById('edit-description').value = expense.description || '';
    
    const modal = document.getElementById('edit-expense-modal');
    modal.style.display = 'block';
}

// Close edit modal
function closeEditModal() {
    const modal = document.getElementById('edit-expense-modal');
    modal.style.display = 'none';
    currentEditingExpense = null;
}

// Save edited expense
async function saveEditedExpense(event) {
    event.preventDefault();
    
    if (!currentEditingExpense) {
        return;
    }
    
    // Get form values
    const glCodeSelect = document.getElementById('edit-gl-code');
    let glCode = glCodeSelect.value;
    
    // If "Other" is selected, use the custom G/L code
    if (glCode === 'other') {
        const customGlCode = document.getElementById('edit-custom-gl-code').value.trim();
        if (customGlCode) {
            glCode = customGlCode;
        }
    }
    
    const updatedExpense = {
        id: currentEditingExpense.id,
        title: document.getElementById('edit-title').value,
        name: document.getElementById('edit-name').value,
        department: document.getElementById('edit-department').value,
        amount: parseFloat(document.getElementById('edit-amount').value),
        tax: parseFloat(document.getElementById('edit-tax').value),
        date: document.getElementById('edit-date').value,
        glCode: glCode,
        description: document.getElementById('edit-description').value
    };
    
    try {
        // Show loading overlay
        loadingOverlay.classList.add('active');
        document.querySelector('.loading-overlay p').textContent = 'Saving expense...';
        
        // Update expense on server
        const response = await fetch(`/api/expenses/${updatedExpense.id}`, {
            method: 'PUT',
            headers: {
                'Content-Type': 'application/json'
            },
            body: JSON.stringify(updatedExpense)
        });
        
        if (!response.ok) {
            throw new Error('Failed to update expense');
        }
        
        // Update local state
        const index = expenses.findIndex(e => e.id === updatedExpense.id);
        if (index !== -1) {
            expenses[index] = updatedExpense;
        }
        
        // Update filtered expenses
        const filteredIndex = filteredExpenses.findIndex(e => e.id === updatedExpense.id);
        if (filteredIndex !== -1) {
            filteredExpenses[filteredIndex] = updatedExpense;
        }
        
        // Re-render expenses list
        renderExpenses();
        updateTotalAmount();
        
        // Hide loading overlay
        loadingOverlay.classList.remove('active');
        
        // Close modal
        closeEditModal();
        
        // Show success notification
        showNotification('Expense updated successfully', 'success');
        
    } catch (error) {
        console.error('Error updating expense:', error);
        loadingOverlay.classList.remove('active');
        showNotification('Failed to update expense: ' + error.message, 'error');
    }
}

// Calculate HST tax from amount
function calculateTax(amount) {
    return parseFloat((amount * HST_RATE).toFixed(2));
}

// Fetch expenses from the API
async function fetchExpenses() {
    try {
        const response = await fetch('/api/expenses');
        if (!response.ok) {
            throw new Error('Failed to fetch expenses');
        }
        
        expenses = await response.json();
        
        // Calculate tax for any expenses that don't have it
        expenses.forEach(expense => {
            if (!expense.tax && expense.amount) {
                expense.tax = calculateTax(expense.amount);
            }
        });
        
        // Sort expenses by date (ascending order)
        sortExpensesByDate();
        
        filteredExpenses = [...expenses];
        renderExpenses();
        updateTotalAmount();
    } catch (error) {
        console.error('Error fetching expenses:', error);
        showNotification('Failed to load expenses', 'error');
    }
}

// Sort expenses by date in ascending order
function sortExpensesByDate() {
    expenses.sort((a, b) => {
        const dateA = new Date(a.date);
        const dateB = new Date(b.date);
        return dateA - dateB; // Ascending order (oldest first)
    });
}

// Handle PDF upload
async function handlePdfUpload(event) {
    event.preventDefault();
    
    // Get form values
    const fileInput = document.getElementById('pdf-file');
    const nameInput = document.getElementById('user-name');
    const departmentInput = document.getElementById('user-department');
    const files = fileInput.files;
    
    // Validate name and department are filled
    if (!nameInput.value.trim()) {
        showNotification('Please enter your name', 'error');
        nameInput.focus();
        return;
    }
    
    if (!departmentInput.value.trim()) {
        showNotification('Please enter your department', 'error');
        departmentInput.focus();
        return;
    }
    
    // Check if files are selected and they are PDFs
    if (files.length === 0) {
        showNotification('Please select at least one PDF file', 'error');
        return;
    }
    
    // Show loading overlay
    loadingOverlay.classList.add('active');
    document.querySelector('.loading-overlay p').textContent = `Processing ${files.length} receipt${files.length > 1 ? 's' : ''}...`;
    
    // Create a FormData object
    const formData = new FormData();
    let validFilesCount = 0;
    
    // Add name and department to FormData
    formData.append('userName', nameInput.value.trim());
    formData.append('userDepartment', departmentInput.value.trim());
    
    // Check each file and add valid ones to formData
    for (let i = 0; i < files.length; i++) {
        const file = files[i];
        if (file.type === 'application/pdf') {
            formData.append('pdfFiles', file);
            validFilesCount++;
        }
    }
    
    if (validFilesCount === 0) {
        loadingOverlay.classList.remove('active');
        showNotification('No valid PDF files selected', 'error');
        return;
    }
    
    try {
        // Send the FormData to the API
        const response = await fetch('/api/upload-pdf', {
            method: 'POST',
            body: formData
        });
        
        if (!response.ok) {
            throw new Error('Failed to upload PDF');
        }
        
        // Get the result data
        const result = await response.json();
        console.log('PDF upload result:', result);
        
        if (!result.success) {
            throw new Error(result.error || 'Failed to process PDF');
        }
        
        // Store the session ID
        currentSessionId = result.sessionId;
        console.log('Session ID set:', currentSessionId);
        
        // Check if we have results
        if (result.results && result.results.length > 0) {
            // Process each result and add to expenses
            let expensesAdded = 0;
            
            for (const item of result.results) {
                try {
                    // Ensure we have data to process
                    if (item && (item.data || item)) {
                        // Use a small delay between submissions to allow UI to update
                        setTimeout(() => {
                            document.querySelector('.loading-overlay p').textContent = `Adding expense ${expensesAdded + 1} of ${result.results.length}...`;
                            populateExpenseForm(item);
                            expensesAdded++;
                            
                            // Update UI when all done
                            if (expensesAdded === result.results.length) {
                                // Hide loading overlay
                                loadingOverlay.classList.remove('active');
                                
                                // Show success message
                                const message = result.results.length === 1 
                                    ? '1 expense was added successfully' 
                                    : `${result.results.length} expenses were added successfully`;
                                showNotification(message, 'success');
                                
                                // Reset the file input
                                fileInput.value = '';
                            }
                        }, expensesAdded * 300); // 300ms delay between each submission
                    } else {
                        console.warn('Skipping result item with missing data:', item);
                    }
                } catch (err) {
                    console.error('Error processing PDF result item:', err);
                }
            }
        } else {
            // Hide loading overlay
            loadingOverlay.classList.remove('active');
            showNotification('No data extracted from PDF', 'error');
        }
    } catch (error) {
        console.error('Error processing PDFs:', error);
        loadingOverlay.classList.remove('active');
        showNotification(error.message || 'Failed to process PDFs', 'error');
    }
}

// Display PDF results in the container
function displayPdfResults(results) {
    pdfResultsContainer.innerHTML = '';
    
    results.forEach((result, index) => {
        const resultElement = document.createElement('div');
        resultElement.className = 'pdf-result-item';
        
        const data = result.data;
        const filename = result.filename;
        
        // Create result item header with filename and "Use This Data" button
        const header = document.createElement('h4');
        header.innerHTML = `
            <span>${filename}</span>
            <button class="use-data-btn" data-index="${index}">Use This Data</button>
        `;
        
        // Create data display
        const dataDisplay = document.createElement('div');
        dataDisplay.className = 'pdf-data';
        
        // Format the data for display
        let formattedData = '';
        if (data.merchant || data.vendor) {
            formattedData += `<p><strong>Merchant:</strong> ${data.merchant || data.vendor}</p>`;
        }
        if (data.total_amount || data.amount) {
            formattedData += `<p><strong>Amount:</strong> $${(data.total_amount || data.amount).toString()}</p>`;
        }
        if (data.date) {
            formattedData += `<p><strong>Date:</strong> ${data.date}</p>`;
        }
        
        dataDisplay.innerHTML = formattedData;
        
        // Append elements to the result item
        resultElement.appendChild(header);
        resultElement.appendChild(dataDisplay);
        
        // Append the result item to the container
        pdfResultsContainer.appendChild(resultElement);
        
        // Add event listener to the "Use This Data" button
        const useDataBtn = resultElement.querySelector('.use-data-btn');
        useDataBtn.addEventListener('click', () => {
            populateExpenseForm(data);
        });
    });
}

// Function to populate the expense form from PDF data
function populateExpenseForm(data) {
    if (!data) return;
    
    // Get name and department from form inputs
    const nameInput = document.getElementById('user-name');
    const departmentInput = document.getElementById('user-department');
    
    // Check if we have a direct structure or nested data structure
    const extractedData = data.data || data;
    
    // Create a new expense object with the extracted data, with safe access
    const expense = {
        title: extractedData.merchant || extractedData.title || '',
        amount: parseFloat(extractedData.amount || 0),
        tax: parseFloat(extractedData.tax || 0) || calculateTax(parseFloat(extractedData.amount || 0)),
        date: extractedData.date || new Date().toISOString().split('T')[0],
        description: extractedData.description || '',
        glCode: extractedData.gl_code || extractedData.glCode || determineGLCode(extractedData.merchant, extractedData.description) || '6408-000',
        name: nameInput.value.trim() || extractedData.name || '',  // Prioritize form input
        department: departmentInput.value.trim() || extractedData.department || '' // Prioritize form input
    };
    
    // Log the expense object for debugging
    console.log('Populated expense from PDF:', expense);
    
    // Automatically submit the form with the populated data
    submitExpense(expense);
}

// Function to submit an expense
async function submitExpense(expense) {
    try {
        // Show loading indicator
        loadingOverlay.classList.add('active');
        document.querySelector('.loading-overlay p').textContent = 'Adding expense...';
        
        // Send the data to the server
        const response = await fetch('/api/expenses', {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json'
            },
            body: JSON.stringify(expense)
        });
        
        if (!response.ok) {
            throw new Error('Failed to add expense');
        }
        
        // Get the response data with the new expense ID
        const savedExpense = await response.json();
        
        // Add to local state
        expenses.push(savedExpense);
        
        // Update filtered expenses
        applyFilters();
        
        // Update total amount
        updateTotalAmount();
        
        // Hide loading overlay
        loadingOverlay.classList.remove('active');
        
        // Show success notification
        showNotification('AI automatically added the expense to your list', 'success');
        
    } catch (error) {
        console.error('Error submitting expense:', error);
        loadingOverlay.classList.remove('active');
        showNotification('Failed to add expense: ' + error.message, 'error');
    }
}

// Show upload status with type (success, error, info)
function showUploadStatus(message, type) {
    pdfUploadStatus.textContent = message;
    pdfUploadStatus.className = 'upload-status';
    pdfUploadStatus.classList.add(type);
    
    // Clear status after 10 seconds for success/info messages
    if (type !== 'error') {
        setTimeout(() => {
            pdfUploadStatus.textContent = '';
            pdfUploadStatus.className = 'upload-status';
        }, 10000);
    }
}

// Apply search filter only (removed category filter)
function applyFilters() {
    const searchTerm = expenseSearch.value.toLowerCase();
    
    filteredExpenses = expenses.filter(expense => {
        const matchesSearch = expense.title.toLowerCase().includes(searchTerm);
        return matchesSearch;
    });
    
    // Sort filtered expenses by date
    filteredExpenses.sort((a, b) => {
        const dateA = new Date(a.date);
        const dateB = new Date(b.date);
        return dateA - dateB; // Ascending order (oldest first)
    });
    
    renderExpenses();
    updateTotalAmount();
}

// Render expenses to the table
function renderExpenses() {
    expensesList.innerHTML = '';
    
    if (filteredExpenses.length === 0) {
        const emptyRow = document.createElement('tr');
        emptyRow.innerHTML = `
            <td colspan="7" class="text-center">No expenses found</td>
        `;
        expensesList.appendChild(emptyRow);
        return;
    }
    
    filteredExpenses.forEach(expense => {
        const row = document.createElement('tr');
        
        // Format date for display
        const date = new Date(expense.date);
        const formattedDate = date.toLocaleDateString();
        
        // Get description
        const description = expense.description ? 
            `<div class="expense-description">${expense.description}</div>` : '';
            
        // Display tax amount as static text, similar to the amount field
        const taxDisplay = `<span class="tax-value">$${expense.tax.toFixed(2)}</span>`;
        
        // Create dropdown for G/L Code
        const glCodeDropdown = `
            <select class="table-dropdown gl-dropdown" data-id="${expense.id}">
                <option value="6408-000" ${expense.glCode === '6408-000' ? 'selected' : ''}>6408-000 (Office & General)</option>
                <option value="6402-000" ${expense.glCode === '6402-000' ? 'selected' : ''}>6402-000 (Membership)</option>
                <option value="6404-000" ${expense.glCode === '6404-000' ? 'selected' : ''}>6404-000 (Subscriptions)</option>
                <option value="7335-000" ${expense.glCode === '7335-000' ? 'selected' : ''}>7335-000 (Education)</option>
                <option value="6026-000" ${expense.glCode === '6026-000' ? 'selected' : ''}>6026-000 (Mileage/ETR)</option>
                <option value="6010-000" ${expense.glCode === '6010-000' ? 'selected' : ''}>6010-000 (Food & Ent.)</option>
                <option value="6011-000" ${expense.glCode === '6011-000' ? 'selected' : ''}>6011-000 (Social)</option>
                <option value="6012-000" ${expense.glCode === '6012-000' ? 'selected' : ''}>6012-000 (Travel)</option>
                <option value="other" ${!['6408-000', '6402-000', '6404-000', '7335-000', '6026-000', '6010-000', '6011-000', '6012-000'].includes(expense.glCode) ? 'selected' : ''}>Other</option>
            </select>
            <input type="text" class="custom-gl-input" data-id="${expense.id}" placeholder="Enter custom G/L code" value="${!['6408-000', '6402-000', '6404-000', '7335-000', '6026-000', '6010-000', '6011-000', '6012-000'].includes(expense.glCode) ? expense.glCode : ''}" style="display: ${!['6408-000', '6402-000', '6404-000', '7335-000', '6026-000', '6010-000', '6011-000', '6012-000'].includes(expense.glCode) ? 'block' : 'none'}; margin-top: 5px; width: 100%;">
        `;
        
        row.innerHTML = `
            <td data-label="Date">${formattedDate}</td>
            <td data-label="Merchant">
                <div class="merchant-name">${expense.title}</div>
                ${description}
            </td>
            <td data-label="Amount">$${expense.amount.toFixed(2)}</td>
            <td data-label="Tax">${taxDisplay}</td>
            <td data-label="G/L Code">${glCodeDropdown}</td>
            <td data-label="Actions">
                <div class="action-buttons">
                    <button class="btn-edit" data-id="${expense.id}">
                        <i class="fas fa-edit"></i>
                    </button>
                    <button class="btn-delete" data-id="${expense.id}">
                        <i class="fas fa-trash"></i>
                    </button>
                </div>
            </td>
        `;
        
        // Add event listeners to buttons
        const editBtn = row.querySelector('.btn-edit');
        const deleteBtn = row.querySelector('.btn-delete');
        
        editBtn.addEventListener('click', () => {
            const expenseToEdit = expenses.find(e => e.id === expense.id);
            if (expenseToEdit) {
                openEditModal(expenseToEdit);
            }
        });
        
        deleteBtn.addEventListener('click', () => {
            if (confirm('Are you sure you want to delete this expense?')) {
                // For now, just remove from local state
                expenses = expenses.filter(exp => exp.id !== expense.id);
                filteredExpenses = filteredExpenses.filter(exp => exp.id !== expense.id);
                renderExpenses();
                updateTotalAmount();
                showNotification('Expense deleted successfully', 'success');
            }
        });
        
        // Add event listeners to the dropdowns
        const glCodeDropdownEl = row.querySelector('.gl-dropdown');
        const customGlInputEl = row.querySelector('.custom-gl-input');
        
        glCodeDropdownEl.addEventListener('change', (e) => {
            const selectedValue = e.target.value;
            
            if (selectedValue === 'other') {
                // Show the custom input field
                customGlInputEl.style.display = 'block';
                customGlInputEl.focus();
            } else {
                // Hide the custom input field and update the G/L code
                customGlInputEl.style.display = 'none';
                updateExpenseGLCode(expense.id, selectedValue);
            }
        });
        
        // Add event listener for the custom G/L code input
        customGlInputEl.addEventListener('blur', (e) => {
            const customValue = e.target.value.trim();
            if (customValue) {
                updateExpenseGLCode(expense.id, customValue);
            }
        });
        
        // Also handle Enter key press
        customGlInputEl.addEventListener('keypress', (e) => {
            if (e.key === 'Enter') {
                e.preventDefault();
                const customValue = e.target.value.trim();
                if (customValue) {
                    updateExpenseGLCode(expense.id, customValue);
                    e.target.blur();
                }
            }
        });
        
        expensesList.appendChild(row);
    });
}

// Update expense tax amount
async function updateExpenseTax(expenseId, newTaxAmount) {
    try {
        // Find the expense in our local state
        const expense = expenses.find(exp => exp.id === expenseId);
        if (!expense) return;
        
        // Update local state first for immediate feedback
        expense.tax = newTaxAmount;
        
        // Update the expense on the server
        const response = await fetch(`/api/expenses/${expenseId}`, {
            method: 'PUT',
            headers: {
                'Content-Type': 'application/json'
            },
            body: JSON.stringify(expense)
        });
        
        if (!response.ok) {
            throw new Error('Failed to update expense');
        }
        
        // Re-render the expenses list to reflect the changes
        renderExpenses();
        showNotification('Tax amount updated successfully', 'success');
    } catch (error) {
        console.error('Error updating expense:', error);
        showNotification('Failed to update tax amount', 'error');
        
        // Revert changes in case of error
        fetchExpenses();
    }
}

// Update expense GL code
async function updateExpenseGLCode(expenseId, newGLCode) {
    try {
        // Find the expense in our local state
        const expense = expenses.find(exp => exp.id === expenseId);
        if (!expense) return;
        
        // Update local state first for immediate feedback
        expense.glCode = newGLCode;
        
        // Update the expense on the server
        const response = await fetch(`/api/expenses/${expenseId}`, {
            method: 'PUT',
            headers: {
                'Content-Type': 'application/json'
            },
            body: JSON.stringify(expense)
        });
        
        if (!response.ok) {
            throw new Error('Failed to update expense');
        }
        
        // Re-render the expenses list to reflect the changes
        renderExpenses();
        showNotification('G/L code updated successfully', 'success');
    } catch (error) {
        console.error('Error updating expense:', error);
        showNotification('Failed to update G/L code', 'error');
        
        // Revert changes in case of error
        fetchExpenses();
    }
}

// Update the total amount
function updateTotalAmount() {
    const total = filteredExpenses.reduce((sum, expense) => sum + expense.amount, 0);
    expensesTotal.textContent = `$${total.toFixed(2)}`;
}

// Show notification
function showNotification(message, type = 'info') {
    // Remove any existing notifications
    const existingNotification = document.querySelector('.notification');
    if (existingNotification) {
        existingNotification.remove();
    }
    
    // Create notification element
    const notification = document.createElement('div');
    notification.className = `notification ${type}`;
    
    // Add icon based on type
    let icon = '';
    switch(type) {
        case 'success':
            icon = '<i class="fas fa-check-circle"></i>';
            break;
        case 'error':
            icon = '<i class="fas fa-exclamation-circle"></i>';
            break;
        case 'warning':
            icon = '<i class="fas fa-exclamation-triangle"></i>';
            break;
        case 'info':
        default:
            icon = '<i class="fas fa-info-circle"></i>';
            break;
    }
    
    // Set content
    notification.innerHTML = `
        ${icon}
        <span>${message}</span>
        <button class="close-btn">&times;</button>
    `;
    
    // Add to DOM
    document.body.appendChild(notification);
    
    // Add event listener to close button
    notification.querySelector('.close-btn').addEventListener('click', () => {
        notification.remove();
    });
    
    // Show notification
    setTimeout(() => {
        notification.classList.add('show');
    }, 10);
    
    // Hide after 5 seconds
    setTimeout(() => {
        if (document.body.contains(notification)) {
            notification.classList.remove('show');
            setTimeout(() => {
                if (document.body.contains(notification)) {
                    notification.remove();
                }
            }, 300);
        }
    }, 5000);
}

// Handle Excel export
async function handleExcelExport() {
    if (expenses.length === 0) {
        showNotification('No expenses to export', 'error');
        return;
    }
    
    try {
        // Show loading indicator
        loadingOverlay.classList.add('active');
        document.querySelector('.loading-overlay p').textContent = 'Generating Excel report...';
        
        // Get the currently filtered expenses or all expenses
        const dataToExport = filteredExpenses.length > 0 ? filteredExpenses : expenses;
        
        // Send request to server
        const response = await fetch('/api/export-excel', {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json'
            },
            body: JSON.stringify({ expenses: dataToExport })
        });
        
        // Hide loading indicator
        loadingOverlay.classList.remove('active');
        
        if (!response.ok) {
            throw new Error('Failed to generate Excel report');
        }
        
        // Get the blob from the response
        const blob = await response.blob();
        
        // Create a download link
        const url = window.URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.style.display = 'none';
        a.href = url;
        a.download = 'Greenwin_Expense_Report.xlsx';
        
        // Add to the DOM and trigger the download
        document.body.appendChild(a);
        a.click();
        
        // Clean up
        window.URL.revokeObjectURL(url);
        document.body.removeChild(a);
        
        showNotification('Excel report generated successfully', 'success');
    } catch (error) {
        console.error('Error exporting to Excel:', error);
        loadingOverlay.classList.remove('active');
        showNotification('Failed to generate Excel report: ' + error.message, 'error');
    }
}

// Handle PDF merge and export
async function handlePdfExport() {
    try {
        // Show loading indicator
        loadingOverlay.classList.add('active');
        document.querySelector('.loading-overlay p').textContent = 'Merging PDF invoices...';
        
        if (!currentSessionId) {
            loadingOverlay.classList.remove('active');
            showNotification('No active session found. Please upload PDFs first.', 'error');
            return;
        }
        
        // Call the exportPDF function
        await exportPDF();
        
        // Hide loading indicator
        loadingOverlay.classList.remove('active');
        
        showNotification('PDF invoices merged and downloaded successfully', 'success');
    } catch (error) {
        console.error('Error merging PDFs:', error);
        loadingOverlay.classList.remove('active');
        showNotification('Failed to merge PDF invoices: ' + error.message, 'error');
    }
}

// Function to determine G/L code based on merchant name and description
function determineGLCode(merchant, description) {
    // Convert inputs to lowercase strings for easier matching
    const merchantLower = (merchant || '').toLowerCase();
    const descriptionLower = (description || '').toLowerCase();
    const combinedText = merchantLower + ' ' + descriptionLower;
    
    // Check for subscription-related keywords
    if (combinedText.includes('subscription') || 
        combinedText.includes('license') || 
        combinedText.includes('starlink') ||
        combinedText.includes('zoho') ||
        combinedText.includes('microsoft') ||
        combinedText.includes('adobe') ||
        combinedText.includes('office 365')) {
        return '6404-000'; // Subscriptions
    }
    
    // Check for education/training-related keywords
    if (combinedText.includes('training') || 
        combinedText.includes('course') || 
        combinedText.includes('education') ||
        combinedText.includes('certification') ||
        combinedText.includes('workshop') ||
        combinedText.includes('seminar')) {
        return '7335-000'; // Education
    }
    
    // Check for food & entertainment-related keywords
    if (combinedText.includes('restaurant') || 
        combinedText.includes('cafe') || 
        combinedText.includes('coffee') ||
        combinedText.includes('lunch') ||
        combinedText.includes('dinner') ||
        combinedText.includes('catering')) {
        return '6010-000'; // Food & Entertainment
    }
    
    // Check for travel-related keywords
    if (combinedText.includes('hotel') || 
        combinedText.includes('flight') || 
        combinedText.includes('airfare') ||
        combinedText.includes('taxi') ||
        combinedText.includes('uber') ||
        combinedText.includes('lyft')) {
        return '6012-000'; // Travel Expenses
    }
    
    // Check for mileage-related keywords
    if (combinedText.includes('mileage') || 
        combinedText.includes('toll') || 
        combinedText.includes('etr') ||
        combinedText.includes('highway') ||
        combinedText.includes('km')) {
        return '6026-000'; // Mileage/ETR
    }
    
    // Check for membership-related keywords
    if (combinedText.includes('membership') || 
        combinedText.includes('dues') || 
        combinedText.includes('association') ||
        combinedText.includes('professional fee')) {
        return '6402-000'; // Membership
    }
    
    // Default to Office & General for hardware, equipment, and other expenses
    if (combinedText.includes('hardware') || 
        combinedText.includes('equipment') ||
        combinedText.includes('dell') ||
        combinedText.includes('computer') ||
        combinedText.includes('laptop') ||
        combinedText.includes('office')) {
        return '6408-000'; // Office & General
    }
    
    // Default to Office & General if no specific match
    return '6408-000';
}

// Function to export expenses to PDF
async function exportPDF() {
    try {
        const sessionId = currentSessionId; // Get the current session ID
        if (!sessionId) {
            console.error('No active session found');
            return;
        }

        // Create a temporary link element
        const link = document.createElement('a');
        link.href = `/api/export-pdf?sessionId=${sessionId}`;
        link.download = 'expense_report.pdf';
        
        // Append to body, click, and remove
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
    } catch (error) {
        console.error('Error exporting to PDF:', error);
    }
}

// Function to handle file upload
async function handleFileUpload(event) {
    event.preventDefault();
    
    const formData = new FormData();
    const files = document.getElementById('pdfFiles').files;
    const userName = document.getElementById('userName').value;
    const userDepartment = document.getElementById('userDepartment').value;
    
    if (!userName || !userDepartment) {
        alert('Please enter your name and department');
        return;
    }
    
    if (files.length === 0) {
        alert('Please select at least one PDF file');
        return;
    }
    
    formData.append('userName', userName);
    formData.append('userDepartment', userDepartment);
    
    for (let i = 0; i < files.length; i++) {
        formData.append('pdfFiles', files[i]);
    }
    
    try {
        const response = await fetch('/api/upload-pdf', {
            method: 'POST',
            body: formData
        });
        
        const data = await response.json();
        
        if (data.success) {
            // Store the session ID
            currentSessionId = data.sessionId;
            
            // Process successful results
            data.results.forEach(result => {
                expenses.push(result.data);
            });
            
            // Update the expenses list
            updateExpensesList();
            
            // Clear the form
            event.target.reset();
            
            // Show success message
            alert(`Successfully processed ${data.processedCount} files`);
        } else {
            alert('Error processing files: ' + data.error);
        }
    } catch (error) {
        console.error('Error:', error);
        alert('Error uploading files');
    }
}

// Add mileage handling
function handleGLCodeChange(event) {
    const glCode = event.target.value;
    const amountLabel = document.querySelector('label[for="amount"]');
    const amountInput = document.getElementById('amount');
    const taxContainer = document.querySelector('.tax-container');
    const descriptionLabel = document.querySelector('label[for="description"]');
    
    if (glCode === '6026-000') {
        // Mileage entry mode
        amountLabel.textContent = 'Kilometers';
        amountInput.placeholder = 'Enter number of kilometers';
        taxContainer.style.display = 'none';
        descriptionLabel.textContent = 'From - To';
        
        // Add event listener for km calculation
        amountInput.addEventListener('input', calculateMileageAmount);
    } else {
        // Regular expense mode
        amountLabel.textContent = 'Amount ($)';
        amountInput.placeholder = 'Enter amount';
        taxContainer.style.display = 'block';
        descriptionLabel.textContent = 'Description';
        
        // Remove mileage calculation listener
        amountInput.removeEventListener('input', calculateMileageAmount);
    }
}

function calculateMileageAmount(event) {
    const kilometers = parseFloat(event.target.value) || 0;
    const ratePerKm = 0.72;
    const calculatedAmount = (kilometers * ratePerKm).toFixed(2);
    
    // Show calculated amount below the input
    let calculatedDiv = document.getElementById('calculated-amount');
    if (!calculatedDiv) {
        calculatedDiv = document.createElement('div');
        calculatedDiv.id = 'calculated-amount';
        event.target.parentNode.appendChild(calculatedDiv);
    }
    calculatedDiv.textContent = `Amount to be reimbursed: $${calculatedAmount}`;
    
    // Update hidden amount field for form submission
    const hiddenAmount = document.createElement('input');
    hiddenAmount.type = 'hidden';
    hiddenAmount.name = 'actualAmount';
    hiddenAmount.value = calculatedAmount;
    event.target.parentNode.appendChild(hiddenAmount);
}

// Add event listener to G/L code dropdown
document.addEventListener('DOMContentLoaded', function() {
    const glCodeDropdown = document.querySelector('select[name="glCode"]');
    if (glCodeDropdown) {
        glCodeDropdown.addEventListener('change', handleGLCodeChange);
    }
}); 