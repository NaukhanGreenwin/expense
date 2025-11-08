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

// --- Signature Functionality ---
let signatureCanvas, signatureCtx, isDrawing = false;
let userSignature = null; // Store the user's signature as a base64 string

// Initialize the app
document.addEventListener('DOMContentLoaded', function() {
    initApp();
    initSignature(); // Initialize signature functionality
});

// Constants
const HST_RATE = 0.13; // 13% HST tax rate

// Utility function for debouncing resize events
function debounce(func, wait) {
    let timeout;
    return function() {
        const context = this;
        const args = arguments;
        clearTimeout(timeout);
        timeout = setTimeout(() => func.apply(context, args), wait);
    };
}

// Auto-reset function that runs on page load
function autoResetOnLoad() {
    try {
        // 1. Clear expenses
        expenses = [];
        filteredExpenses = [];
        renderExpenses();
        updateTotalAmount();
        
        // 2. Clear signature
        if (signatureCanvas && signatureCtx) {
            clearSignature();
            localStorage.removeItem('userSignature');
            userSignature = null;
            
            // Reset signature UI
            const drawPanel = document.getElementById('draw-signature-panel');
            const drawTab = document.getElementById('draw-tab');
            const uploadPanel = document.getElementById('upload-signature-panel');
            const uploadTab = document.getElementById('upload-tab');
            const currentContainer = document.getElementById('current-signature-container');
            
            if (drawPanel && uploadPanel && currentContainer) {
                drawPanel.classList.add('active');
                drawTab.classList.add('active');
                uploadPanel.classList.remove('active');
                uploadTab.classList.remove('active');
                currentContainer.style.display = 'none';
            }
        }
        
        // 3. Reset form inputs
        const nameInput = document.getElementById('user-name');
        const departmentInput = document.getElementById('user-department');
        const pdfFileInput = document.getElementById('pdf-file');
        
        if (nameInput) nameInput.value = '';
        if (departmentInput) departmentInput.value = '';
        if (pdfFileInput) pdfFileInput.value = '';
        
        // 4. Clear results container
        const resultsContainer = document.getElementById('pdf-results-container');
        if (resultsContainer) resultsContainer.innerHTML = '';
        
        // 5. Clear upload status
        const uploadStatus = document.getElementById('pdf-upload-status');
        if (uploadStatus) uploadStatus.innerHTML = '';
        
        // 6. Reset session ID
        currentSessionId = null;
        
        console.log('Application automatically reset on page load');
    } catch (error) {
        console.error('Error auto-resetting application:', error);
    }
}

function initApp() {
    // Auto-reset on page load
    autoResetOnLoad();
    
    // We don't need to fetch expenses since we're resetting on load
    // fetchExpenses();
    
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
    
    // Add event listener for Reset All button
    const resetAllButton = document.getElementById('reset-all-btn');
    if (resetAllButton) {
        resetAllButton.addEventListener('click', resetAllData);
    }

    // Add event listener for Add Expense button
    const addExpenseButton = document.getElementById('add-expense-btn');
    if (addExpenseButton) {
        addExpenseButton.addEventListener('click', openAddExpenseForm);
    }
    
    // Add event listener for Travel Expense button
    const travelExpenseButton = document.getElementById('travel-expense-btn');
    if (travelExpenseButton) {
        travelExpenseButton.addEventListener('click', openTravelExpenseForm);
    }
    
    // Hide the expense form - we only use PDF upload now
    const expenseFormContainer = document.getElementById('expense-form-container');
    if (expenseFormContainer) {
        expenseFormContainer.style.display = 'none';
    }
    
    // Create edit modal if it doesn't exist
    createEditModal();

    // Add PDF file count validation
    const pdfFileInput = document.getElementById('pdf-file');
    
    if (pdfFileInput) {
        pdfFileInput.addEventListener('change', function() {
            const maxFiles = 50;
            
            if (this.files.length > maxFiles) {
                // Show error and clear the file input
                showNotification(`You can only upload a maximum of ${maxFiles} PDF files at once.`, 'error');
                this.value = ''; // Clear the file input
                return;
            }
            
            // Show selected file count
            if (this.files.length > 0) {
                showNotification(`${this.files.length} file${this.files.length > 1 ? 's' : ''} selected`, 'info', 2000);
            }
        });
    }
    
    // Location dropdown removed per user request
    // populateLocationDropdown();
}

// Function to open the add expense form
function openAddExpenseForm() {
    // Create a blank expense object
    const newExpense = {
        id: 'new', // This will be replaced with a real ID when saved
        title: '',
        name: '',
        department: '',
        amount: '',
        tax: '',
        date: new Date().toISOString().split('T')[0], // Today's date
        glCode: '6408-000', // Default G/L code
        description: ''
    };
    
    // Open the edit modal with the blank expense
    openEditModal(newExpense);
    
    // Change the modal title to indicate we're adding a new expense
    const modalTitle = document.querySelector('#edit-expense-modal h3');
    if (modalTitle) {
        modalTitle.textContent = 'Add New Expense';
    }
}

// Function to open the travel expense form
function openTravelExpenseForm() {
    // Create a blank travel expense object with default G/L code for mileage
    const newTravelExpense = {
        id: 'new', // This will be replaced with a real ID when saved
        title: 'Travel Expense',
        name: '',
        department: '',
        amount: '', // Will be calculated based on kilometers
        tax: '0', // No tax for mileage
        date: new Date().toISOString().split('T')[0], // Today's date
        glCode: '6026-000', // G/L code for mileage
        description: '',
        fromLocation: '',
        toLocation: '',
        kilometers: ''
    };
    
    // Open the edit modal with the blank travel expense
    openEditModal(newTravelExpense);
    
    // Change the modal title to indicate we're adding a new travel expense
    const modalTitle = document.querySelector('#edit-expense-modal h3');
    if (modalTitle) {
        modalTitle.textContent = 'Add Travel Expense';
    }
    
    // Hide the Merchant/Title field as it's not needed for travel expenses
    const titleInput = document.getElementById('edit-title');
    if (titleInput) {
        const titleGroup = titleInput.closest('.form-group');
        if (titleGroup) {
            titleGroup.style.display = 'none';
        }
    }
    
    // Hide the Name and Department fields
    const nameInput = document.getElementById('edit-name');
    const departmentInput = document.getElementById('edit-department');
    
    if (nameInput) {
        const nameGroup = nameInput.closest('.form-group');
        if (nameGroup) {
            nameGroup.style.display = 'none';
        }
    }
    
    if (departmentInput) {
        const departmentGroup = departmentInput.closest('.form-group');
        if (departmentGroup) {
            departmentGroup.style.display = 'none';
        }
    }
    
    // Ensure the mileage fields are shown and focus on the kilometers field
    const glCodeSelect = document.getElementById('edit-gl-code');
    if (glCodeSelect) {
        glCodeSelect.value = '6026-000';
        
        // Trigger the change event to show mileage fields
        const event = new Event('change');
        glCodeSelect.dispatchEvent(event);
        
        // Update labels to make it clear this is for kilometers
        const amountLabel = document.getElementById('amount-label');
        if (amountLabel) {
            amountLabel.textContent = 'Kilometers';
        }
        
        // Focus on the kilometers field
        const kilometersInput = document.getElementById('edit-amount');
        if (kilometersInput) {
            kilometersInput.focus();
        }
    }
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
                        <label for="edit-amount" id="amount-label">Amount ($)</label>
                        <input type="number" id="edit-amount" step="0.01" min="0" required>
                        <div id="calculated-amount" style="display: none; margin-top: 5px; padding: 5px; background-color: #e8f5e9; color: #2e7d32; border-radius: 4px; font-size: 14px;"></div>
                    </div>
                    <div class="form-group half" id="tax-group">
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
                    <label for="edit-location">Location</label>
                    <select id="edit-location" class="location-dropdown">
                        <option value="">Select a location</option>
                        <!-- Options will be populated via JavaScript -->
                    </select>
                    <input type="hidden" id="property-code" name="property-code">
                </div>
                
                <!-- Split Expense Section -->
                <div class="form-group" id="split-expense-section">
                    <div class="split-header">
                        <label>Split Expense</label>
                        <button type="button" class="btn-add-split"><i class="fas fa-plus-circle"></i> Add Split</button>
                    </div>
                    <div class="split-info">Split this expense across multiple G/L codes. The primary G/L code above will receive the remaining amount.</div>
                    <div id="split-container">
                        <!-- Split rows will be added here dynamically -->
                    </div>
                    <div id="split-summary" style="display: none;">
                        <div class="split-totals">
                            <span>Allocated: <span id="split-allocated-amount">$0.00</span> (<span id="split-allocated-percent">0%</span>)</span>
                            <span>Remaining: <span id="split-remaining-amount">$0.00</span> (<span id="split-remaining-percent">100%</span>)</span>
                        </div>
                    </div>
                </div>
                
                <!-- Mileage specific fields - hidden by default -->
                <div id="mileage-fields" style="display: none; margin-top: 10px; padding: 10px; background-color: #f5f5f5; border-radius: 4px;">
                    <h4 style="margin-top: 0; color: #2e7d32;">Mileage Details</h4>
                    <div class="form-row">
                        <div class="form-group half">
                            <label for="edit-from-location">From</label>
                            <input type="text" id="edit-from-location" placeholder="Starting location">
                        </div>
                        <div class="form-group half">
                            <label for="edit-to-location">To</label>
                            <input type="text" id="edit-to-location" placeholder="Destination">
                        </div>
                    </div>
                    <div class="form-group">
                        <label for="edit-trip-purpose">Trip Purpose</label>
                        <select id="edit-trip-purpose">
                            <option value="client-meeting">Client Meeting</option>
                            <option value="site-visit">Site Visit</option>
                            <option value="training">Training/Education</option>
                            <option value="office-travel">Office Travel</option>
                            <option value="other-purpose">Other</option>
                        </select>
                    </div>
                </div>
                
                <div class="form-group" id="description-group">
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
    
    // Add event listener for the Add Split button
    const addSplitBtn = modal.querySelector('.btn-add-split');
    if (addSplitBtn) {
        addSplitBtn.addEventListener('click', function(e) {
            e.preventDefault();
            addSplitRow();
            updateSplitTotals();
        });
    }
    
    // Auto-calculate tax when amount changes
    const amountInput = document.getElementById('edit-amount');
    if (amountInput) {
        amountInput.addEventListener('input', function() {
            const glCode = document.getElementById('edit-gl-code').value;
            
            if (glCode === '6026-000') {
                // For mileage, calculate reimbursement
                const kilometers = parseFloat(this.value) || 0;
                const reimbursementAmount = calculateMileageReimbursement(kilometers);
                
                // Show calculated amount
                const calculatedAmountDiv = document.getElementById('calculated-amount');
                calculatedAmountDiv.style.display = 'block';
                calculatedAmountDiv.textContent = `Reimbursement: $${reimbursementAmount.toFixed(2)} (${kilometers} km × $0.72)`;
            } else {
                // For regular expenses, calculate tax
                const amount = parseFloat(this.value) || 0;
                const taxInput = document.getElementById('edit-tax');
                if (taxInput) {
                    taxInput.value = calculateTax(amount).toFixed(2);
                }
                
                // Hide calculated amount
                const calculatedAmountDiv = document.getElementById('calculated-amount');
                calculatedAmountDiv.style.display = 'none';
            }
            
            // Update split totals if there are any splits
            updateSplitTotals();
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
        
        // Toggle mileage fields when G/L code changes
        toggleMileageFields(this.value === '6026-000');
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
    
    // Handle expense splits if they exist
    const splitContainer = document.getElementById('split-container');
    if (splitContainer) {
        // Clear any existing splits
        splitContainer.innerHTML = '';
        
        // If expense has splits, populate them
        if (expense.splits && expense.splits.length > 0) {
            expense.splits.forEach((split, index) => {
                addSplitRow(split.glCode, split.amount, split.percentage);
            });
        }
    }
    
    // Check for mileage fields and populate them if this is a mileage entry
    const fromLocationInput = document.getElementById('edit-from-location');
    const toLocationInput = document.getElementById('edit-to-location');
    const tripPurposeSelect = document.getElementById('edit-trip-purpose');
    
    if (fromLocationInput && toLocationInput && tripPurposeSelect) {
        if (expense.glCode === '6026-000') {
            fromLocationInput.value = expense.fromLocation || '';
            toLocationInput.value = expense.toLocation || '';
            tripPurposeSelect.value = expense.tripPurpose || 'client-meeting';
            
            // Parse from description if first time
            if (!expense.fromLocation && !expense.toLocation && expense.description) {
                const parts = expense.description.split(' - ');
                if (parts.length >= 2) {
                    fromLocationInput.value = parts[0] || '';
                    toLocationInput.value = parts[1] || '';
                }
            }
        }
    }
    
    // Toggle the display of mileage specific fields
    toggleMileageFields(glCode === '6026-000');
    
    const modal = document.getElementById('edit-expense-modal');
    modal.style.display = 'block';
}

// Toggle the display of mileage specific fields
function toggleMileageFields(showMileageFields) {
    const mileageFields = document.getElementById('mileage-fields');
    const descriptionGroup = document.getElementById('description-group');
    const amountLabel = document.getElementById('amount-label');
    const taxGroup = document.getElementById('tax-group');
    const calculatedAmount = document.getElementById('calculated-amount');
    const amountInput = document.getElementById('edit-amount');
    
    if (showMileageFields) {
        // Show mileage fields
        mileageFields.style.display = 'block';
        
        // Change amount label to kilometers
        amountLabel.textContent = 'Kilometers';
        amountInput.placeholder = 'Enter kilometers driven';
        
        // Hide tax group as mileage has no tax
        taxGroup.style.display = 'none';
        
        // Hide description as we're using from/to instead
        descriptionGroup.style.display = 'none';
        
        // Show calculated amount if there's a value
        const kilometers = parseFloat(amountInput.value) || 0;
        if (kilometers > 0) {
            const reimbursementAmount = calculateMileageReimbursement(kilometers);
            calculatedAmount.style.display = 'block';
            calculatedAmount.textContent = `Reimbursement: $${reimbursementAmount.toFixed(2)} (${kilometers} km × $0.72)`;
        }
    } else {
        // Hide mileage fields
        mileageFields.style.display = 'none';
        
        // Restore original labels and display
        amountLabel.textContent = 'Amount ($)';
        amountInput.placeholder = 'Enter amount';
        
        // Show tax group for regular expenses
        taxGroup.style.display = 'block';
        
        // Show description for regular expenses
        descriptionGroup.style.display = 'block';
        
        // Hide calculated amount
        calculatedAmount.style.display = 'none';
    }
}

// Calculate mileage reimbursement
function calculateMileageReimbursement(kilometers) {
    const MILEAGE_RATE = 0.72; // $0.72 per kilometer
    return kilometers * MILEAGE_RATE;
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
    
    // Build the base expense object
    const updatedExpense = {
        id: currentEditingExpense.id === 'new' ? Date.now() : currentEditingExpense.id, // Generate new ID if new expense
        title: document.getElementById('edit-title').value,
        name: document.getElementById('edit-name').value,
        department: document.getElementById('edit-department').value,
        amount: parseFloat(document.getElementById('edit-amount').value),
        tax: parseFloat(document.getElementById('edit-tax').value),
        date: document.getElementById('edit-date').value,
        glCode: glCode,
        description: document.getElementById('edit-description').value
    };
    
    // Handle expense splits
    const splitContainer = document.getElementById('split-container');
    if (splitContainer && splitContainer.children.length > 0) {
        const splits = [];
        const splitRows = splitContainer.querySelectorAll('.split-row');
        
        // Process each split row
        splitRows.forEach(row => {
            const glSelect = row.querySelector('.split-gl-select');
            const customGlInput = row.querySelector('.split-custom-gl');
            const amountInput = row.querySelector('.split-amount-input');
            const percentInput = row.querySelector('.split-percent-input');
            
            // Get G/L code
            let splitGlCode = glSelect.value;
            if (splitGlCode === 'other') {
                splitGlCode = customGlInput.value.trim();
            }
            
            // Only add if we have a G/L code and amount
            if (splitGlCode && amountInput.value) {
                splits.push({
                    glCode: splitGlCode,
                    amount: parseFloat(amountInput.value) || 0,
                    percentage: parseFloat(percentInput.value) || 0
                });
            }
        });
        
        // Add splits to the expense object if we have any
        if (splits.length > 0) {
            updatedExpense.splits = splits;
        }
    }
    
    // Handle mileage specific fields
    if (glCode === '6026-000') {
        const fromLocation = document.getElementById('edit-from-location').value;
        const toLocation = document.getElementById('edit-to-location').value;
        const tripPurpose = document.getElementById('edit-trip-purpose').value;
        
        // Add mileage specific fields
        updatedExpense.fromLocation = fromLocation;
        updatedExpense.toLocation = toLocation;
        updatedExpense.tripPurpose = tripPurpose;
        updatedExpense.kilometers = updatedExpense.amount; // Store original kilometers
        
        // Calculate the reimbursement amount
        updatedExpense.amount = calculateMileageReimbursement(updatedExpense.kilometers);
        
        // Set tax to 0 for mileage expenses
        updatedExpense.tax = 0;
        
        // Create a description from the mileage fields
        updatedExpense.description = `${fromLocation} - ${toLocation} (${getTripPurposeText(tripPurpose)})`;
    }
    
    try {
        // Show loading overlay
        loadingOverlay.classList.add('active');
        
        let method, url, successMessage;
        
        if (currentEditingExpense.id === 'new') {
            // It's a new expense
            method = 'POST';
            url = '/api/expenses';
            successMessage = 'Expense added successfully';
            document.querySelector('.loading-overlay p').textContent = 'Adding expense...';
        } else {
            // It's an existing expense
            method = 'PUT';
            url = `/api/expenses/${updatedExpense.id}`;
            successMessage = 'Expense updated successfully';
            document.querySelector('.loading-overlay p').textContent = 'Saving expense...';
        }
        
        // Send request to server
        const response = await fetch(url, {
            method: method,
            headers: {
                'Content-Type': 'application/json'
            },
            body: JSON.stringify(updatedExpense)
        });
        
        if (!response.ok) {
            throw new Error(`Failed to ${currentEditingExpense.id === 'new' ? 'add' : 'update'} expense`);
        }
        
        const savedExpense = await response.json();
        
        // Update the ID if it's a new expense
        if (currentEditingExpense.id === 'new') {
            updatedExpense.id = savedExpense.id;
        }
        
        // Update local state
        if (currentEditingExpense.id === 'new') {
            // Add to expenses array
            expenses.push(updatedExpense);
            filteredExpenses.push(updatedExpense);
        } else {
            // Update existing expense
            const index = expenses.findIndex(e => e.id === updatedExpense.id);
            if (index !== -1) {
                expenses[index] = updatedExpense;
            }
            
            // Update filtered expenses
            const filteredIndex = filteredExpenses.findIndex(e => e.id === updatedExpense.id);
            if (filteredIndex !== -1) {
                filteredExpenses[filteredIndex] = updatedExpense;
            }
        }
        
        // Re-render expenses list
        renderExpenses();
        updateTotalAmount();
        
        // Hide loading overlay
        loadingOverlay.classList.remove('active');
        
        // Close modal
        closeEditModal();
        
        // Show success notification
        showNotification(successMessage, 'success');
        
    } catch (error) {
        console.error('Error saving expense:', error);
        loadingOverlay.classList.remove('active');
        showNotification('Failed to save expense: ' + error.message, 'error');
    }
}

// Get readable text for trip purpose value
function getTripPurposeText(purposeValue) {
    const purposeMap = {
        'client-meeting': 'Client Meeting',
        'site-visit': 'Site Visit',
        'training': 'Training/Education',
        'office-travel': 'Office Travel',
        'other-purpose': 'Other'
    };
    
    return purposeMap[purposeValue] || purposeValue;
}

// Close edit modal
function closeEditModal() {
    const modal = document.getElementById('edit-expense-modal');
    modal.style.display = 'none';
    currentEditingExpense = null;
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
    
    // Office humor messages
    const funnyMessages = [
        "💼 Making your expenses look completely legitimate...",
        "📝 Checking if that $500 'office supply' was really a PS5...",
        "🏝️ Verifying that conference was not actually a vacation...",
        "🧠 Remembering all those business discussions at the bar...",
        "🧾 Convincing accounting this coffee was for a client meeting...",
        "💸 Calculating how many pizza lunches until your bonus is gone...",
        "🔍 Searching for receipts you swore were in your pocket...",
        "🤔 Wondering if your boss will believe this was work-related...",
        "💻 Converting lunch receipts into billable hours...",
        "📊 Creating charts to justify that team-building happy hour..."
    ];
    
    // Set initial message
    let currentMessageIndex = 0;
    document.querySelector('.loading-overlay p').textContent = funnyMessages[currentMessageIndex];
    
    // Setup message rotation interval
    const messageInterval = setInterval(() => {
        currentMessageIndex = (currentMessageIndex + 1) % funnyMessages.length;
        document.querySelector('.loading-overlay p').textContent = funnyMessages[currentMessageIndex];
    }, 4000); // Change message every 4 seconds
    
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
        clearInterval(messageInterval); // Clear the interval when closing overlay
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
                            // Office humor messages for progress
                            const progressMessages = [
                                "🧾 Expense #[COUNT] - Preparing for accounting's approval...",
                                "💸 Expense #[COUNT] - Converting receipts to money in your pocket...",
                                "🔍 Expense #[COUNT] - Finding the best tax category for this one...",
                                "📋 Expense #[COUNT] - Adding legitimate business purpose..."
                            ];
                            
                            // Pick a random message and replace [COUNT] with the actual count
                            let progressMessage = progressMessages[Math.floor(Math.random() * progressMessages.length)];
                            progressMessage = progressMessage.replace('[COUNT]', expensesAdded + 1);
                            
                            document.querySelector('.loading-overlay p').textContent = progressMessage;
                            populateExpenseForm(item);
                            expensesAdded++;
                            
                            // Update UI when all done
                            if (expensesAdded === result.results.length) {
                                // Clear the message rotation interval
                                clearInterval(messageInterval);
                                
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
            // Clear the message rotation interval
            clearInterval(messageInterval);
            
            // Hide loading overlay
            loadingOverlay.classList.remove('active');
            showNotification('No data extracted from PDF', 'error');
        }
    } catch (error) {
        // Clear the message rotation interval when error occurs
        clearInterval(messageInterval);
        
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
// Location functionality removed per user request // This file contains the missing functions that need to be appended to app.js

// Complete the renderExpenses function and add all missing functions
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
    
    // Create each expense row
    filteredExpenses.forEach(expense => {
        const row = document.createElement('tr');
        
        // Format date
        const formattedDate = new Date(expense.date).toLocaleDateString('en-US', {
            year: 'numeric',
            month: 'short',
            day: 'numeric'
        });
        
        // Format amount
        const formattedAmount = `$${parseFloat(expense.amount).toFixed(2)}`;
        
        // Format tax
        const taxDisplay = expense.tax ? `$${parseFloat(expense.tax).toFixed(2)}` : '-';
        
        // Get G/L code display
        const glCodeDisplay = getGLCodeName(expense.glCode);
        
        // Create description with split info if applicable
        let descriptionDisplay = expense.description || '';
        
        // Check if this is a split expense
        if (expense.splits && expense.splits.length > 0) {
            descriptionDisplay += '<div class="expense-split-summary">';
            
            // Calculate remaining amount for primary G/L code
            const totalSplitAmount = expense.splits.reduce((sum, split) => sum + parseFloat(split.amount || 0), 0);
            const remainingAmount = parseFloat(expense.amount) - totalSplitAmount;
            const remainingPercent = Math.round((remainingAmount / parseFloat(expense.amount)) * 100);
            
            if (remainingAmount > 0) {
                descriptionDisplay += `<div class="split-detail primary">
                    <span class="split-label">Primary (${expense.glCode}):</span>
                    <span class="split-value">$${remainingAmount.toFixed(2)} (${remainingPercent}%)</span>
                </div>`;
            }
            
            // Add each split
            expense.splits.forEach(split => {
                descriptionDisplay += `<div class="split-detail">
                    <span class="split-label">${split.glCode}:</span>
                    <span class="split-value">$${parseFloat(split.amount).toFixed(2)} (${Math.round(split.percentage)}%)</span>
                </div>`;
            });
            
            descriptionDisplay += '</div>';
        }
        
        // Create row HTML
        row.innerHTML = `
            <td data-label="Date">${formattedDate}</td>
            <td data-label="Merchant">
                <div class="expense-detail">
                    <div class="merchant-name">${expense.title}</div>
                    <div class="expense-description">${descriptionDisplay}</div>
                </div>
            </td>
            <td data-label="Amount">${formattedAmount}</td>
            <td data-label="Tax">${taxDisplay}</td>
            <td data-label="G/L Code">${glCodeDisplay}</td>
            <td data-label="Actions" class="action-cell">
                <div class="action-buttons">
                <button class="btn-edit" data-id="${expense.id}"><i class="fas fa-edit"></i></button>
                <button class="btn-delete" data-id="${expense.id}"><i class="fas fa-trash"></i></button>
                </div>
            </td>
        `;
        
        // Add event listeners
        const editBtn = row.querySelector('.btn-edit');
        const deleteBtn = row.querySelector('.btn-delete');
        
        editBtn.addEventListener('click', () => openEditModal(expense));
        deleteBtn.addEventListener('click', () => deleteExpense(expense.id));
        
        expensesList.appendChild(row);
    });
}

// Get G/L code display name
function getGLCodeName(code) {
    const codeMap = {
        '6408-000': '6408-000 (Office & General)',
        '6402-000': '6402-000 (Membership)',
        '6404-000': '6404-000 (Subscriptions)',
        '7335-000': '7335-000 (Education)',
        '6026-000': '6026-000 (Mileage/ETR)',
        '6010-000': '6010-000 (Food & Ent.)',
        '6011-000': '6011-000 (Social)',
        '6012-000': '6012-000 (Travel)'
    };
    return codeMap[code] || code;
}

// Delete expense
async function deleteExpense(expenseId) {
    if (!confirm('Are you sure you want to delete this expense?')) {
        return;
    }
    
    try {
        // Remove from local state
        expenses = expenses.filter(e => e.id !== expenseId);
        filteredExpenses = filteredExpenses.filter(e => e.id !== expenseId);
        
        // Re-render
        renderExpenses();
        updateTotalAmount();
        
        showNotification('Expense deleted successfully', 'success');
                } catch (error) {
        console.error('Error deleting expense:', error);
        showNotification('Failed to delete expense', 'error');
    }
}

// Update total amount
function updateTotalAmount() {
    const total = filteredExpenses.reduce((sum, expense) => {
        return sum + parseFloat(expense.amount || 0);
    }, 0);
    
    expensesTotal.textContent = `$${total.toFixed(2)}`;
}

// Show notification
function showNotification(message, type = 'info', duration = 3000) {
    // Remove any existing notifications
    const existingNotification = document.querySelector('.notification');
    if (existingNotification) {
        existingNotification.remove();
    }
    
    // Create notification element
    const notification = document.createElement('div');
    notification.className = `notification ${type}`;
    
    // Add icon based on type
    let icon = 'fa-info-circle';
    if (type === 'success') icon = 'fa-check-circle';
    if (type === 'error') icon = 'fa-exclamation-circle';
    if (type === 'warning') icon = 'fa-exclamation-triangle';
    
    notification.innerHTML = `
        <i class="fas ${icon}"></i>
        <span>${message}</span>
        <button class="close-btn">&times;</button>
    `;
    
    document.body.appendChild(notification);
    
    // Trigger animation
    setTimeout(() => notification.classList.add('show'), 10);
    
    // Add close button listener
    const closeBtn = notification.querySelector('.close-btn');
    closeBtn.addEventListener('click', () => {
        notification.classList.remove('show');
        setTimeout(() => notification.remove(), 300);
    });
    
    // Auto-remove after duration
    if (duration > 0) {
        setTimeout(() => {
            notification.classList.remove('show');
            setTimeout(() => notification.remove(), 300);
        }, duration);
    }
}

// Location dropdown removed per user request

// Determine G/L code based on merchant and description
function determineGLCode(merchant, description) {
    const text = `${merchant} ${description}`.toLowerCase();
    
    // Office & General
    if (text.includes('office') || text.includes('staples') || text.includes('supplies')) {
        return '6408-000';
    }
    
    // Membership
    if (text.includes('membership') || text.includes('association')) {
        return '6402-000';
    }
    
    // Subscriptions
    if (text.includes('subscription') || text.includes('software') || text.includes('cloud') || 
        text.includes('microsoft') || text.includes('adobe') || text.includes('zoom')) {
        return '6404-000';
    }
    
    // Education
    if (text.includes('training') || text.includes('course') || text.includes('education') || 
        text.includes('certification') || text.includes('workshop')) {
        return '7335-000';
    }
    
    // Mileage/ETR
    if (text.includes('parking') || text.includes('toll') || text.includes('mileage') || 
        text.includes('transit')) {
        return '6026-000';
    }
    
    // Food & Entertainment
    if (text.includes('restaurant') || text.includes('coffee') || text.includes('lunch') || 
        text.includes('dinner') || text.includes('meal') || text.includes('starbucks') ||
        text.includes('tim hortons')) {
        return '6010-000';
    }
    
    // Travel
    if (text.includes('hotel') || text.includes('flight') || text.includes('airbnb') || 
        text.includes('taxi') || text.includes('uber') || text.includes('lyft')) {
        return '6012-000';
    }
    
    // Default to Office & General
    return '6408-000';
}

// Add split row to the split container
function addSplitRow(glCode = '', amount = '', percentage = '') {
    const splitContainer = document.getElementById('split-container');
    if (!splitContainer) return;
    
    const splitRow = document.createElement('div');
    splitRow.className = 'split-row';
    
    const splitId = Date.now() + Math.random();
    
    splitRow.innerHTML = `
        <div class="split-row-content">
            <div class="split-gl-group">
                <select class="split-gl-select">
                    <option value="">Select G/L Code</option>
                    <option value="6408-000" ${glCode === '6408-000' ? 'selected' : ''}>6408-000 (Office & General)</option>
                    <option value="6402-000" ${glCode === '6402-000' ? 'selected' : ''}>6402-000 (Membership)</option>
                    <option value="6404-000" ${glCode === '6404-000' ? 'selected' : ''}>6404-000 (Subscriptions)</option>
                    <option value="7335-000" ${glCode === '7335-000' ? 'selected' : ''}>7335-000 (Education)</option>
                    <option value="6026-000" ${glCode === '6026-000' ? 'selected' : ''}>6026-000 (Mileage/ETR)</option>
                    <option value="6010-000" ${glCode === '6010-000' ? 'selected' : ''}>6010-000 (Food & Ent.)</option>
                    <option value="6011-000" ${glCode === '6011-000' ? 'selected' : ''}>6011-000 (Social)</option>
                    <option value="6012-000" ${glCode === '6012-000' ? 'selected' : ''}>6012-000 (Travel)</option>
                    <option value="other">Other</option>
                </select>
                <input type="text" class="split-custom-gl" placeholder="Enter custom G/L code" style="display: ${glCode && !['6408-000', '6402-000', '6404-000', '7335-000', '6026-000', '6010-000', '6011-000', '6012-000'].includes(glCode) ? 'block' : 'none'};" value="${glCode && !['6408-000', '6402-000', '6404-000', '7335-000', '6026-000', '6010-000', '6011-000', '6012-000'].includes(glCode) ? glCode : ''}">
            </div>
            <div class="split-amount-inputs">
                <div class="split-dollar">
                    <span class="dollar-sign">$</span>
                    <input type="number" class="split-amount-input" placeholder="Amount" step="0.01" min="0" value="${amount}">
                </div>
                <div class="split-percent">
                    <input type="number" class="split-percent-input" placeholder="%" step="0.1" min="0" max="100" value="${percentage}">
                    <span class="percent-sign">%</span>
                </div>
            </div>
            <button type="button" class="btn-remove-split"><i class="fas fa-times"></i></button>
        </div>
    `;
    
    splitContainer.appendChild(splitRow);
    
    // Show split summary if not already visible
    const splitSummary = document.getElementById('split-summary');
    if (splitSummary) {
        splitSummary.style.display = 'block';
    }
    
    // Add event listeners
    const glSelect = splitRow.querySelector('.split-gl-select');
    const customGlInput = splitRow.querySelector('.split-custom-gl');
    const amountInput = splitRow.querySelector('.split-amount-input');
    const percentInput = splitRow.querySelector('.split-percent-input');
    const removeBtn = splitRow.querySelector('.btn-remove-split');
    
    // Show/hide custom G/L input
    glSelect.addEventListener('change', function() {
        if (this.value === 'other') {
            customGlInput.style.display = 'block';
        } else {
            customGlInput.style.display = 'none';
        }
        updateSplitTotals();
    });
    
    // Update totals when amount changes
    amountInput.addEventListener('input', function() {
        // Auto-calculate percentage
        const totalAmount = parseFloat(document.getElementById('edit-amount').value) || 0;
        const splitAmount = parseFloat(this.value) || 0;
        if (totalAmount > 0) {
            const percent = (splitAmount / totalAmount) * 100;
            percentInput.value = percent.toFixed(1);
        }
        updateSplitTotals();
    });
    
    // Update totals when percentage changes
    percentInput.addEventListener('input', function() {
        // Auto-calculate amount
        const totalAmount = parseFloat(document.getElementById('edit-amount').value) || 0;
        const percent = parseFloat(this.value) || 0;
        const splitAmount = (totalAmount * percent) / 100;
        amountInput.value = splitAmount.toFixed(2);
        updateSplitTotals();
    });
    
    // Remove split
    removeBtn.addEventListener('click', function() {
        splitRow.remove();
        updateSplitTotals();
        
        // Hide split summary if no more splits
        if (splitContainer.children.length === 0 && splitSummary) {
            splitSummary.style.display = 'none';
        }
    });
}

// Update split totals
function updateSplitTotals() {
    const totalAmount = parseFloat(document.getElementById('edit-amount').value) || 0;
    const splitContainer = document.getElementById('split-container');
    
    if (!splitContainer || splitContainer.children.length === 0) {
        return;
    }
    
    let allocatedAmount = 0;
    
    // Sum up all split amounts
    const splitRows = splitContainer.querySelectorAll('.split-row');
    splitRows.forEach(row => {
        const amountInput = row.querySelector('.split-amount-input');
        allocatedAmount += parseFloat(amountInput.value) || 0;
    });
    
    const remainingAmount = totalAmount - allocatedAmount;
    const allocatedPercent = totalAmount > 0 ? (allocatedAmount / totalAmount) * 100 : 0;
    const remainingPercent = totalAmount > 0 ? (remainingAmount / totalAmount) * 100 : 100;
    
    // Update display
    document.getElementById('split-allocated-amount').textContent = `$${allocatedAmount.toFixed(2)}`;
    document.getElementById('split-allocated-percent').textContent = `${allocatedPercent.toFixed(1)}%`;
    document.getElementById('split-remaining-amount').textContent = `$${remainingAmount.toFixed(2)}`;
    document.getElementById('split-remaining-percent').textContent = `${remainingPercent.toFixed(1)}%`;
}

// Initialize signature functionality
function initSignature() {
    signatureCanvas = document.getElementById('signature-canvas');
    if (!signatureCanvas) return;
    
    signatureCtx = signatureCanvas.getContext('2d');
    
    // Check for saved signature
    const savedSignature = localStorage.getItem('userSignature');
    if (savedSignature) {
        userSignature = savedSignature;
        displayCurrentSignature();
    }
    
    // Set up canvas
    setupCanvas();
    
    // Tab switching
    document.getElementById('draw-tab')?.addEventListener('click', () => switchTab('draw'));
    document.getElementById('upload-tab')?.addEventListener('click', () => switchTab('upload'));
    
    // Drawing events
    signatureCanvas.addEventListener('mousedown', startDrawing);
    signatureCanvas.addEventListener('mousemove', draw);
    signatureCanvas.addEventListener('mouseup', stopDrawing);
    signatureCanvas.addEventListener('mouseout', stopDrawing);
    
    // Touch events for mobile
    signatureCanvas.addEventListener('touchstart', handleTouchStart);
    signatureCanvas.addEventListener('touchmove', handleTouchMove);
    signatureCanvas.addEventListener('touchend', stopDrawing);
    
    // Buttons
    document.getElementById('clear-signature')?.addEventListener('click', clearSignature);
    document.getElementById('save-drawn-signature')?.addEventListener('click', saveDrawnSignature);
    document.getElementById('signature-upload')?.addEventListener('click', handleSignatureUpload);
    document.getElementById('save-uploaded-signature')?.addEventListener('click', saveUploadedSignature);
    document.getElementById('change-signature')?.addEventListener('click', changeSignature);
}

function setupCanvas() {
    if (!signatureCanvas) return;
    
    const rect = signatureCanvas.getBoundingClientRect();
    signatureCanvas.width = rect.width;
    signatureCanvas.height = rect.height;
    
    signatureCtx.strokeStyle = '#000';
    signatureCtx.lineWidth = 2;
    signatureCtx.lineCap = 'round';
    signatureCtx.lineJoin = 'round';
}

function switchTab(tab) {
    const drawTab = document.getElementById('draw-tab');
    const uploadTab = document.getElementById('upload-tab');
    const drawPanel = document.getElementById('draw-signature-panel');
    const uploadPanel = document.getElementById('upload-signature-panel');
    
    if (tab === 'draw') {
        drawTab.classList.add('active');
        uploadTab.classList.remove('active');
        drawPanel.classList.add('active');
        uploadPanel.classList.remove('active');
    } else {
        drawTab.classList.remove('active');
        uploadTab.classList.add('active');
        drawPanel.classList.remove('active');
        uploadPanel.classList.add('active');
    }
}

function startDrawing(e) {
    isDrawing = true;
    const rect = signatureCanvas.getBoundingClientRect();
    signatureCtx.beginPath();
    signatureCtx.moveTo(e.clientX - rect.left, e.clientY - rect.top);
}

function draw(e) {
    if (!isDrawing) return;
    const rect = signatureCanvas.getBoundingClientRect();
    signatureCtx.lineTo(e.clientX - rect.left, e.clientY - rect.top);
    signatureCtx.stroke();
}

function stopDrawing() {
    isDrawing = false;
}

function handleTouchStart(e) {
    e.preventDefault();
    const touch = e.touches[0];
    const mouseEvent = new MouseEvent('mousedown', {
        clientX: touch.clientX,
        clientY: touch.clientY
    });
    signatureCanvas.dispatchEvent(mouseEvent);
}

function handleTouchMove(e) {
    e.preventDefault();
    const touch = e.touches[0];
    const mouseEvent = new MouseEvent('mousemove', {
        clientX: touch.clientX,
        clientY: touch.clientY
    });
    signatureCanvas.dispatchEvent(mouseEvent);
}

function clearSignature() {
    if (!signatureCtx) return;
    signatureCtx.clearRect(0, 0, signatureCanvas.width, signatureCanvas.height);
}

function saveDrawnSignature() {
    if (!signatureCanvas) return;
    
    // Check if canvas is empty
    const imageData = signatureCtx.getImageData(0, 0, signatureCanvas.width, signatureCanvas.height);
    const isEmpty = !imageData.data.some(channel => channel !== 0);
    
    if (isEmpty) {
        showNotification('Please draw your signature first', 'warning');
        return;
    }
    
    userSignature = signatureCanvas.toDataURL('image/png');
    localStorage.setItem('userSignature', userSignature);
    displayCurrentSignature();
    showNotification('Signature saved successfully', 'success');
}

function handleSignatureUpload(e) {
    const file = e.target.files[0];
    if (!file) return;
    
    const reader = new FileReader();
    reader.onload = function(event) {
        const img = document.getElementById('signature-preview');
        img.src = event.target.result;
        document.getElementById('signature-preview-container').style.display = 'block';
        userSignature = event.target.result;
    };
    reader.readAsDataURL(file);
}

function saveUploadedSignature() {
    if (!userSignature) {
        showNotification('Please upload a signature image first', 'warning');
        return;
    }
    
    localStorage.setItem('userSignature', userSignature);
    displayCurrentSignature();
    showNotification('Signature saved successfully', 'success');
}

function displayCurrentSignature() {
    const container = document.getElementById('current-signature-container');
    const img = document.getElementById('current-signature');
    
    if (userSignature && container && img) {
        img.src = userSignature;
        container.style.display = 'block';
    }
}

function changeSignature() {
    const container = document.getElementById('current-signature-container');
    if (container) {
        container.style.display = 'none';
    }
    clearSignature();
}

// Reset all data
function resetAllData() {
    if (!confirm('Are you sure you want to reset all data? This will clear all expenses and your signature.')) {
        return;
    }
    
    // Clear expenses
    expenses = [];
    filteredExpenses = [];
    renderExpenses();
    updateTotalAmount();
    
    // Clear signature
    if (signatureCanvas && signatureCtx) {
        clearSignature();
        localStorage.removeItem('userSignature');
        userSignature = null;
        
        const currentContainer = document.getElementById('current-signature-container');
        if (currentContainer) {
            currentContainer.style.display = 'none';
        }
    }
    
    // Reset form inputs
    const nameInput = document.getElementById('user-name');
    const departmentInput = document.getElementById('user-department');
    const pdfFileInput = document.getElementById('pdf-file');
    
    if (nameInput) nameInput.value = '';
    if (departmentInput) departmentInput.value = '';
    if (pdfFileInput) pdfFileInput.value = '';
    
    // Clear results
    const resultsContainer = document.getElementById('pdf-results-container');
    if (resultsContainer) resultsContainer.innerHTML = '';
    
    const uploadStatus = document.getElementById('pdf-upload-status');
    if (uploadStatus) uploadStatus.innerHTML = '';
    
    currentSessionId = null;
    
    showNotification('All data has been reset', 'success');
}

// Handle Excel export
async function handleExcelExport() {
    if (expenses.length === 0) {
        showNotification('No expenses to export', 'warning');
        return;
    }
    
    try {
        loadingOverlay.classList.add('active');
        document.querySelector('.loading-overlay p').textContent = 'Generating Excel report...';
        
        const response = await fetch('/api/export-excel', {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json'
            },
            body: JSON.stringify({ 
                expenses: expenses,
                signature: userSignature
            })
        });
        
        if (!response.ok) {
            throw new Error('Failed to generate Excel report');
        }
        
        // Download the file
        const blob = await response.blob();
        const url = window.URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = `Expense_Report_${new Date().toISOString().split('T')[0]}.xlsx`;
        document.body.appendChild(a);
        a.click();
        window.URL.revokeObjectURL(url);
        document.body.removeChild(a);
        
        loadingOverlay.classList.remove('active');
        showNotification('Excel report generated successfully', 'success');
    } catch (error) {
        console.error('Error exporting to Excel:', error);
        loadingOverlay.classList.remove('active');
        showNotification('Failed to generate Excel report: ' + error.message, 'error');
    }
}

// Handle PDF export (session-aware)
async function handlePdfExport() {
    if (!currentSessionId) {
        showNotification('No receipts to export. Please upload PDF receipts first.', 'warning');
        return;
    }

    try {
        loadingOverlay.classList.add('active');
        document.querySelector('.loading-overlay p').textContent = 'Merging PDF receipts...';

        // Use session-aware endpoint so only current upload’s PDFs are merged
        const urlWithSession = `/api/export-pdf?sessionId=${encodeURIComponent(currentSessionId)}`;
        const response = await fetch(urlWithSession, { method: 'GET' });

        if (!response.ok) {
            // Try to extract server-provided error if JSON
            try {
                const err = await response.json();
                throw new Error(err.error || 'Failed to merge PDFs');
            } catch (_) {
                throw new Error('Failed to merge PDFs');
            }
        }

        // Download the merged file
        const blob = await response.blob();
        const url = window.URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = `Greenwin_Merged_PDF_${new Date().toISOString().split('T')[0]}.pdf`;
        document.body.appendChild(a);
        a.click();
        window.URL.revokeObjectURL(url);
        document.body.removeChild(a);

        // Server cleans up the session; clear local reference
        currentSessionId = null;

        loadingOverlay.classList.remove('active');
        showNotification('PDF merged successfully', 'success');
    } catch (error) {
        console.error('Error exporting PDF:', error);
        loadingOverlay.classList.remove('active');
        showNotification('Failed to merge PDFs: ' + error.message, 'error');
    }
}
