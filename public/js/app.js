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
        
        // Store original values in case we need to show edit form
        const originalExpense = {...expense};
        
        // Check if switching to mileage code
        const isMileageCode = newGLCode === '6026-000';
        
        // If switching to mileage, open edit modal to get mileage details
        if (isMileageCode) {
            // Open the edit modal first to get mileage details
            openEditModal(expense);
            
            // Pre-select the mileage code
            const glCodeSelect = document.getElementById('edit-gl-code');
            if (glCodeSelect) {
                glCodeSelect.value = '6026-000';
                // Trigger the change event manually to show mileage fields
                glCodeSelect.dispatchEvent(new Event('change'));
            }
            
            // Show a notification about entering mileage details
            showNotification('Please enter mileage details', 'info');
            return;
        }
        
        // Update local state first for immediate feedback
        expense.glCode = newGLCode;
        
        // If changing from mileage to normal expense, reset any mileage-specific fields
        if (expense.kilometers) {
            delete expense.kilometers;
            delete expense.fromLocation;
            delete expense.toLocation;
            delete expense.tripPurpose;
        }
        
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

// Update the total amount display
function updateTotalAmount() {
    if (!expensesTotal) return;
    
    // Calculate regular total
    const total = filteredExpenses.reduce((sum, expense) => sum + (parseFloat(expense.amount) || 0), 0);
    
    // Calculate mileage totals
    const mileageExpenses = filteredExpenses.filter(e => e.glCode === '6026-000');
    const totalKilometers = mileageExpenses.reduce((sum, expense) => sum + (parseFloat(expense.kilometers) || 0), 0);
    const totalMileageAmount = mileageExpenses.reduce((sum, expense) => sum + (parseFloat(expense.amount) || 0), 0);
    
    // Update the display
    expensesTotal.innerHTML = `
        <div>Total: <span class="total-amount">$${total.toFixed(2)}</span></div>
        ${totalKilometers > 0 ? `
        <div class="mileage-summary">
            <div>Total Kilometers: <span class="km-amount">${totalKilometers.toFixed(1)} km</span></div>
            <div>Total Mileage Reimbursement: <span class="mileage-amount">$${totalMileageAmount.toFixed(2)}</span></div>
        </div>
        ` : ''}
    `;
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
        
        // Include signature data if available
        const exportData = {
            expenses: dataToExport,
            signature: userSignature || null
        };
        
        // Debug logging for signature
        console.log('Signature data present:', !!userSignature);
        if (userSignature) {
            console.log('Signature data length:', userSignature.length);
            console.log('Signature data starts with:', userSignature.substring(0, 50) + '...');
        }
        
        // Send request to server
        const response = await fetch('/api/export-excel', {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json'
            },
            body: JSON.stringify(exportData)
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

// Update the editExpense function to include mileage handling
function editExpense(event) {
    // Get the row element
    const row = event.target.closest('tr');
    if (!row) return;
    
    // Get the row data
    const id = row.dataset.id;
    const date = row.querySelector('td:nth-child(1)').textContent.trim();
    const title = row.querySelector('td:nth-child(2) .title-text')?.textContent || '';
    const description = row.querySelector('td:nth-child(2) .description-text')?.textContent || '';
    const amount = row.querySelector('td:nth-child(3)').textContent.trim().replace('$', '');
    const tax = row.querySelector('td:nth-child(4)').textContent.trim().replace('$', '');
    const glCode = row.querySelector('td:nth-child(5) select')?.value || '';
    
    // Create modal with form
    const modal = document.createElement('div');
    modal.className = 'modal';
    modal.innerHTML = `
        <div class="modal-content">
            <span class="close">&times;</span>
            <h2>Edit Expense</h2>
            <form id="edit-expense-form">
                <input type="hidden" name="id" value="${id}">
                <div class="form-group">
                    <label for="date">Date</label>
                    <input type="date" id="date" name="date" value="${date}" required>
                </div>
                <div class="form-group">
                    <label for="title">Merchant/Title</label>
                    <input type="text" id="title" name="title" value="${title}" required>
                </div>
                <div class="form-group">
                    <label for="description">Description</label>
                    <textarea id="description" name="description">${description}</textarea>
                </div>
                <div class="form-group">
                    <label for="amount" id="amount-label">Amount ($)</label>
                    <input type="number" id="amount" name="amount" step="0.01" value="${amount}" required>
                    <div id="calculated-amount" style="display: none;"></div>
                </div>
                <div class="form-group" id="tax-group">
                    <label for="tax">Tax (HST)</label>
                    <input type="number" id="tax" name="tax" step="0.01" value="${tax}">
                </div>
                <div class="form-group">
                    <label for="glCode">G/L Code</label>
                    <select id="glCode" name="glCode" required>
                        <option value="6026-000" ${glCode === '6026-000' ? 'selected' : ''}>6026-000 (Mileage/ETR)</option>
                        <option value="6210-000" ${glCode === '6210-000' ? 'selected' : ''}>6210-000 (Computer Supplies)</option>
                        <option value="6400-000" ${glCode === '6400-000' ? 'selected' : ''}>6400-000 (Small Tools)</option>
                        <option value="6500-000" ${glCode === '6500-000' ? 'selected' : ''}>6500-000 (Office Supplies)</option>
                        <option value="8200-000" ${glCode === '8200-000' ? 'selected' : ''}>8200-000 (Training & Development)</option>
                    </select>
                </div>
                <div class="form-actions">
                    <button type="submit" class="btn btn-primary">Save Changes</button>
                    <button type="button" class="btn btn-secondary modal-cancel">Cancel</button>
                </div>
            </form>
        </div>
    `;
    
    // Add modal to page
    document.body.appendChild(modal);
    
    // Setup mileage mode
    const glCodeSelect = document.getElementById('glCode');
    const amountLabel = document.getElementById('amount-label');
    const amountInput = document.getElementById('amount');
    const taxGroup = document.getElementById('tax-group');
    const calculatedAmount = document.getElementById('calculated-amount');
    
    function updateMileageMode() {
        if (glCodeSelect.value === '6026-000') {
            // Mileage mode
            amountLabel.textContent = 'Kilometers';
            amountInput.placeholder = 'Enter kilometers';
            taxGroup.style.display = 'none';
            calculatedAmount.style.display = 'block';
            
            // Calculate amount based on kilometers
            const km = parseFloat(amountInput.value) || 0;
            const amount = (km * 0.72).toFixed(2);
            calculatedAmount.textContent = `Amount: $${amount}`;
            calculatedAmount.style.color = '#2e7d32';
            calculatedAmount.style.fontWeight = 'bold';
            calculatedAmount.style.marginTop = '5px';
        } else {
            // Regular expense mode
            amountLabel.textContent = 'Amount ($)';
            amountInput.placeholder = 'Enter amount';
            taxGroup.style.display = 'block';
            calculatedAmount.style.display = 'none';
        }
    }
    
    // Initial update
    updateMileageMode();
    
    // Add event listener for G/L code changes
    glCodeSelect.addEventListener('change', updateMileageMode);
    
    // Add event listener for kilometer input
    amountInput.addEventListener('input', function() {
        if (glCodeSelect.value === '6026-000') {
            const km = parseFloat(this.value) || 0;
            const amount = (km * 0.72).toFixed(2);
            calculatedAmount.textContent = `Amount: $${amount}`;
        }
    });
    
    // Close modal
    const closeBtn = modal.querySelector('.close');
    closeBtn.addEventListener('click', () => {
        document.body.removeChild(modal);
    });
    
    // Cancel button
    const cancelBtn = modal.querySelector('.modal-cancel');
    cancelBtn.addEventListener('click', () => {
        document.body.removeChild(modal);
    });
    
    // Form submit
    const form = modal.querySelector('form');
    form.addEventListener('submit', (e) => {
        e.preventDefault();
        
        const formData = new FormData(form);
        const data = {};
        formData.forEach((value, key) => {
            data[key] = value;
        });
        
        // Handle mileage calculation
        if (data.glCode === '6026-000') {
            data.kilometers = parseFloat(data.amount) || 0;
            data.amount = (data.kilometers * 0.72).toFixed(2);
            data.tax = '0.00';
        }
        
        // Show loading indicator
        const submitBtn = form.querySelector('button[type="submit"]');
        const originalText = submitBtn.textContent;
        submitBtn.textContent = 'Saving...';
        submitBtn.disabled = true;
        
        // Make API request to update expense
        fetch(`/api/expenses/${id}`, {
            method: 'PUT',
            headers: {
                'Content-Type': 'application/json'
            },
            body: JSON.stringify(data)
        })
        .then(response => {
            if (!response.ok) {
                return response.json().then(errData => {
                    throw new Error(errData.error || 'Failed to update expense');
                });
            }
            return response.json();
        })
        .then(updatedExpense => {
            console.log('Expense updated:', updatedExpense);
            
            // Update row in table
            row.querySelector('td:nth-child(1)').textContent = data.date;
            
            const titleElement = row.querySelector('td:nth-child(2)');
            titleElement.innerHTML = `
                <div class="title-text">${data.title}</div>
                <div class="description-text">${data.description}</div>
            `;
            
            row.querySelector('td:nth-child(3)').textContent = `$${data.amount}`;
            row.querySelector('td:nth-child(4)').textContent = `$${data.tax || '0.00'}`;
            
            const selectElement = row.querySelector('td:nth-child(5) select');
            if (selectElement) {
                selectElement.value = data.glCode;
            }
            
            // Close modal
            document.body.removeChild(modal);
        })
        .catch(error => {
            console.error('Error updating expense:', error);
            alert(`Error updating expense: ${error.message || 'Please try again'}`);
            
            // Reset button
            submitBtn.textContent = originalText;
            submitBtn.disabled = false;
        });
    });
}

// Handle Email export (combines Excel and PDF)
async function handleEmailExport() {
    if (expenses.length === 0) {
        showNotification('No expenses to export', 'error');
        return;
    }
    
    try {
        // Show loading indicator
        loadingOverlay.classList.add('active');
        document.querySelector('.loading-overlay p').textContent = 'Preparing files for email...';
        
        // Get the currently filtered expenses or all expenses
        const dataToExport = filteredExpenses.length > 0 ? filteredExpenses : expenses;
        
        // Send request to server to prepare files
        const response = await fetch('/api/export-email', {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json'
            },
            body: JSON.stringify({ 
                expenses: dataToExport,
                sessionId: currentSessionId 
            })
        });
        
        if (!response.ok) {
            throw new Error('Failed to prepare files for email');
        }
        
        const result = await response.json();
        
        if (!result.success) {
            throw new Error(result.error || 'Failed to prepare files for email');
        }
        
        // Get file paths from the response
        const { excelPath, pdfPath } = result;
        
        // Get full URLs for the files
        const fileUrls = [];
        if (excelPath) {
            fileUrls.push(`${window.location.origin}${excelPath}`);
        }
        if (pdfPath) {
            fileUrls.push(`${window.location.origin}${pdfPath}`);
        }
        
        // Hide loading indicator
        loadingOverlay.classList.remove('active');
        
        // Compose email subject and body
        const subject = 'Greenwin Expense Report';
        const body = 'Please find attached the expense report files.';
        
        // Attempt to open Outlook directly
        try {
            // Use the 'ms-outlook:' protocol to open Outlook with attachments
            let outlookUrl = `ms-outlook:compose?subject=${encodeURIComponent(subject)}&body=${encodeURIComponent(body)}`;
            
            // Add attachments if we have files
            if (fileUrls.length > 0) {
                outlookUrl += `&attachment=${encodeURIComponent(fileUrls.join(','))}`;
            }
            
            // Attempt to open Outlook
            window.location.href = outlookUrl;
            
            // Show a notification about opening Outlook
            showNotification('Opening Outlook with expense report files...', 'success');
            
            // After a short delay, check if Outlook opened successfully
            setTimeout(() => {
                // Fallback if Outlook didn't open
                if (document.hasFocus()) {
                    console.log('Outlook may not have opened, offering fallback method');
                    offerFallbackMethod(fileUrls, subject, body);
                }
            }, 2000);
        } catch (outlookError) {
            console.error('Error opening Outlook:', outlookError);
            // Fallback to alternative method
            offerFallbackMethod(fileUrls, subject, body);
        }
    } catch (error) {
        console.error('Error preparing email:', error);
        loadingOverlay.classList.remove('active');
        showNotification('Failed to prepare email: ' + error.message, 'error');
    }
}

// Function to provide a fallback method when Outlook can't be opened
function offerFallbackMethod(fileUrls, subject, body) {
    // Create a modal or notification to inform the user
    const fallbackModal = document.createElement('div');
    fallbackModal.className = 'modal fade show';
    fallbackModal.style.display = 'block';
    fallbackModal.style.backgroundColor = 'rgba(0,0,0,0.5)';
    
    fallbackModal.innerHTML = `
        <div class="modal-dialog" style="margin-top: 100px;">
            <div class="modal-content">
                <div class="modal-header">
                    <h5 class="modal-title">Outlook Not Detected</h5>
                    <button type="button" class="close" data-dismiss="modal" aria-label="Close">
                        <span aria-hidden="true">&times;</span>
                    </button>
                </div>
                <div class="modal-body">
                    <p>We couldn't open Outlook automatically. Would you like to:</p>
                    <div class="mt-3">
                        <button id="download-files-btn" class="btn btn-primary mr-2">
                            Download Files
                        </button>
                        <button id="use-default-email-btn" class="btn btn-secondary">
                            Use Default Email App
                        </button>
                    </div>
                </div>
            </div>
        </div>
    `;
    
    document.body.appendChild(fallbackModal);
    
    // Add event listeners to buttons
    document.getElementById('download-files-btn').addEventListener('click', async () => {
        // Download the files
        for (const url of fileUrls) {
            const filename = url.split('/').pop();
            await downloadFile(url, filename);
        }
        
        fallbackModal.remove();
        showNotification('Files downloaded successfully. Please attach them to your email manually.', 'success', 10000);
    });
    
    document.getElementById('use-default-email-btn').addEventListener('click', () => {
        // Use the default mailto link
        const mailtoLink = `mailto:?subject=${encodeURIComponent(subject)}&body=${encodeURIComponent(body)}`;
        window.location.href = mailtoLink;
        
        fallbackModal.remove();
        showNotification('Default email client opened. Please attach the files manually.', 'info', 10000);
    });
    
    // Close button event
    fallbackModal.querySelector('.close').addEventListener('click', () => {
        fallbackModal.remove();
    });
}

// Function to download a file
async function downloadFile(url, filename) {
    // Create a temporary link element
    const link = document.createElement('a');
    link.href = url;
    link.download = filename;
    link.style.display = 'none';
    
    // Add the link to the document and trigger the download
    document.body.appendChild(link);
    link.click();
    
    // Remove the link from the DOM
    setTimeout(() => {
        document.body.removeChild(link);
    }, 100);
    
    // Return a promise that resolves after a delay to ensure the download starts
    return new Promise(resolve => setTimeout(resolve, 500));
}

// Initialize signature functionality
function initSignature() {
    // Canvas drawing signature
    signatureCanvas = document.getElementById('signature-canvas');
    if (!signatureCanvas) return;
    
    // Fix canvas resolution to match display size
    const dpr = window.devicePixelRatio || 1;
    const rect = signatureCanvas.getBoundingClientRect();
    
    // Set the canvas dimensions to match its CSS dimensions
    signatureCanvas.width = rect.width * dpr;
    signatureCanvas.height = rect.height * dpr;
    
    signatureCtx = signatureCanvas.getContext('2d');
    
    // Scale the context to account for the device pixel ratio
    signatureCtx.scale(dpr, dpr);
    
    // Set up canvas for drawing
    signatureCtx.lineWidth = 2;
    signatureCtx.lineCap = 'round';
    signatureCtx.strokeStyle = '#000000';
    
    // Event listeners for canvas drawing
    signatureCanvas.addEventListener('mousedown', startDrawing);
    signatureCanvas.addEventListener('touchstart', handleTouchStart);
    signatureCanvas.addEventListener('mousemove', draw);
    signatureCanvas.addEventListener('touchmove', handleTouchMove);
    signatureCanvas.addEventListener('mouseup', stopDrawing);
    signatureCanvas.addEventListener('touchend', stopDrawing);
    signatureCanvas.addEventListener('mouseout', stopDrawing);
    
    // Clear button
    const clearBtn = document.getElementById('clear-signature');
    if (clearBtn) {
        clearBtn.addEventListener('click', clearSignature);
    }
    
    // Save drawn signature button
    const saveDrawnBtn = document.getElementById('save-drawn-signature');
    if (saveDrawnBtn) {
        saveDrawnBtn.addEventListener('click', saveDrawnSignature);
    }
    
    // Signature upload
    const signatureUpload = document.getElementById('signature-upload');
    if (signatureUpload) {
        signatureUpload.addEventListener('change', handleSignatureUpload);
    }
    
    // Save uploaded signature button
    const saveUploadedBtn = document.getElementById('save-uploaded-signature');
    if (saveUploadedBtn) {
        saveUploadedBtn.addEventListener('click', saveUploadedSignature);
    }
    
    // Tab switching
    const drawTab = document.getElementById('draw-tab');
    const uploadTab = document.getElementById('upload-tab');
    
    if (drawTab && uploadTab) {
        drawTab.addEventListener('click', () => switchSignatureTab('draw'));
        uploadTab.addEventListener('click', () => switchSignatureTab('upload'));
    }
    
    // Change signature button
    const changeSignatureBtn = document.getElementById('change-signature');
    if (changeSignatureBtn) {
        changeSignatureBtn.addEventListener('click', changeSignature);
    }
    
    // Check for saved signature in localStorage
    loadSavedSignature();
    
    // Add window resize handler to update canvas dimensions
    window.addEventListener('resize', debounce(function() {
        // Re-initialize the canvas when the window is resized
        updateCanvasDimensions();
    }, 250));
}

// Update canvas dimensions on resize
function updateCanvasDimensions() {
    if (!signatureCanvas) return;
    
    const dpr = window.devicePixelRatio || 1;
    const rect = signatureCanvas.getBoundingClientRect();
    
    // Save current drawing
    const tempCanvas = document.createElement('canvas');
    tempCanvas.width = signatureCanvas.width;
    tempCanvas.height = signatureCanvas.height;
    const tempCtx = tempCanvas.getContext('2d');
    tempCtx.drawImage(signatureCanvas, 0, 0);
    
    // Resize canvas
    signatureCanvas.width = rect.width * dpr;
    signatureCanvas.height = rect.height * dpr;
    
    // Restore context settings
    signatureCtx = signatureCanvas.getContext('2d');
    signatureCtx.scale(dpr, dpr);
    signatureCtx.lineWidth = 2;
    signatureCtx.lineCap = 'round';
    signatureCtx.strokeStyle = '#000000';
    
    // Restore drawing
    signatureCtx.drawImage(tempCanvas, 0, 0, signatureCanvas.width, signatureCanvas.height);
}

// Handle touch events for mobile drawing
function handleTouchStart(e) {
    e.preventDefault();
    const touch = e.touches[0];
    const rect = signatureCanvas.getBoundingClientRect();
    const x = touch.clientX - rect.left;
    const y = touch.clientY - rect.top;
    
    isDrawing = true;
    signatureCtx.beginPath();
    signatureCtx.moveTo(x, y);
}

function handleTouchMove(e) {
    if (!isDrawing) return;
    e.preventDefault();
    
    const touch = e.touches[0];
    const rect = signatureCanvas.getBoundingClientRect();
    const x = touch.clientX - rect.left;
    const y = touch.clientY - rect.top;
    
    signatureCtx.lineTo(x, y);
    signatureCtx.stroke();
}

// Start drawing on canvas
function startDrawing(e) {
    isDrawing = true;
    
    const rect = signatureCanvas.getBoundingClientRect();
    const x = e.clientX - rect.left;
    const y = e.clientY - rect.top;
    
    signatureCtx.beginPath();
    signatureCtx.moveTo(x, y);
}

// Draw on canvas as mouse/touch moves
function draw(e) {
    if (!isDrawing) return;
    
    const rect = signatureCanvas.getBoundingClientRect();
    const x = e.clientX - rect.left;
    const y = e.clientY - rect.top;
    
    signatureCtx.lineTo(x, y);
    signatureCtx.stroke();
    
    // Hide the "Sign here" instructions once drawing starts
    const instructions = document.querySelector('.signature-instructions');
    if (instructions) {
        instructions.style.display = 'none';
    }
}

// Stop drawing
function stopDrawing() {
    isDrawing = false;
}

// Clear the signature canvas
function clearSignature() {
    signatureCtx.clearRect(0, 0, signatureCanvas.width, signatureCanvas.height);
    
    // Show the "Sign here" instructions again
    const instructions = document.querySelector('.signature-instructions');
    if (instructions) {
        instructions.style.display = 'flex';
    }
}

// Save the drawn signature
function saveDrawnSignature() {
    // Check if canvas is empty
    const isCanvasEmpty = isSignatureCanvasEmpty();
    if (isCanvasEmpty) {
        showNotification('Please draw your signature first', 'error');
        return;
    }
    
    // Create a temporary canvas with white background for better visibility in Excel
    const tempCanvas = document.createElement('canvas');
    tempCanvas.width = signatureCanvas.width;
    tempCanvas.height = signatureCanvas.height;
    const tempCtx = tempCanvas.getContext('2d');
    
    // Fill with white background
    tempCtx.fillStyle = 'white';
    tempCtx.fillRect(0, 0, tempCanvas.width, tempCanvas.height);
    
    // Draw the signature on top
    tempCtx.drawImage(signatureCanvas, 0, 0);
    
    // Convert enhanced canvas to base64 image with white background
    const signatureData = tempCanvas.toDataURL('image/png');
    
    // Save signature to localStorage
    saveSignature(signatureData);
    
    console.log('Signature saved with white background');
}

// Check if the signature canvas is empty
function isSignatureCanvasEmpty() {
    const pixelData = signatureCtx.getImageData(0, 0, signatureCanvas.width, signatureCanvas.height).data;
    
    // Check if all pixels are transparent (alpha = 0)
    for (let i = 3; i < pixelData.length; i += 4) {
        if (pixelData[i] > 0) {
            return false; // Found a non-transparent pixel
        }
    }
    
    return true; // Canvas is empty
}

// Handle signature image upload
function handleSignatureUpload(e) {
    const file = e.target.files[0];
    if (!file) return;
    
    // Validate file type
    if (!['image/jpeg', 'image/png'].includes(file.type)) {
        showNotification('Please upload a PNG or JPG image', 'error');
        return;
    }
    
    // Validate file size (max 2MB)
    if (file.size > 2 * 1024 * 1024) {
        showNotification('Signature image must be less than 2MB', 'error');
        return;
    }
    
    // Show preview
    const reader = new FileReader();
    reader.onload = function(event) {
        const previewContainer = document.getElementById('signature-preview-container');
        const preview = document.getElementById('signature-preview');
        
        if (previewContainer && preview) {
            preview.src = event.target.result;
            previewContainer.style.display = 'block';
        }
    };
    reader.readAsDataURL(file);
}

// Save the uploaded signature image
function saveUploadedSignature() {
    const preview = document.getElementById('signature-preview');
    if (!preview || !preview.src || preview.src === '') {
        showNotification('Please upload a signature image first', 'error');
        return;
    }
    
    // Save signature to localStorage
    saveSignature(preview.src);
}

// Save signature to localStorage and update UI
function saveSignature(signatureData) {
    // Save to localStorage
    localStorage.setItem('userSignature', signatureData);
    userSignature = signatureData;
    
    // Update UI to show current signature
    showCurrentSignature();
    
    showNotification('Signature saved successfully', 'success');
}

// Load saved signature from localStorage
function loadSavedSignature() {
    const savedSignature = localStorage.getItem('userSignature');
    if (savedSignature) {
        userSignature = savedSignature;
        showCurrentSignature();
    }
}

// Show the current saved signature
function showCurrentSignature() {
    const drawPanel = document.getElementById('draw-signature-panel');
    const uploadPanel = document.getElementById('upload-signature-panel');
    const currentContainer = document.getElementById('current-signature-container');
    const currentSignatureImg = document.getElementById('current-signature');
    
    if (drawPanel && uploadPanel && currentContainer && currentSignatureImg) {
        // Hide signature panels
        drawPanel.classList.remove('active');
        uploadPanel.classList.remove('active');
        
        // Show current signature container
        currentContainer.style.display = 'block';
        currentSignatureImg.src = userSignature;
    }
}

// Change signature (remove current and show options again)
function changeSignature() {
    const drawPanel = document.getElementById('draw-signature-panel');
    const drawTab = document.getElementById('draw-tab');
    const uploadPanel = document.getElementById('upload-signature-panel');
    const uploadTab = document.getElementById('upload-tab');
    const currentContainer = document.getElementById('current-signature-container');
    
    if (drawPanel && uploadPanel && currentContainer) {
        // Show drawing panel and tab
        drawPanel.classList.add('active');
        drawTab.classList.add('active');
        
        // Hide upload panel and tab
        uploadPanel.classList.remove('active');
        uploadTab.classList.remove('active');
        
        // Hide current signature container
        currentContainer.style.display = 'none';
        
        // Clear the canvas
        clearSignature();
    }
}

// Switch between draw and upload tabs
function switchSignatureTab(tab) {
    const drawPanel = document.getElementById('draw-signature-panel');
    const drawTab = document.getElementById('draw-tab');
    const uploadPanel = document.getElementById('upload-signature-panel');
    const uploadTab = document.getElementById('upload-tab');
    
    if (drawPanel && uploadPanel && drawTab && uploadTab) {
        if (tab === 'draw') {
            drawPanel.classList.add('active');
            drawTab.classList.add('active');
            uploadPanel.classList.remove('active');
            uploadTab.classList.remove('active');
        } else {
            uploadPanel.classList.add('active');
            uploadTab.classList.add('active');
            drawPanel.classList.remove('active');
            drawTab.classList.remove('active');
        }
    }
}

// Reset all data
function resetAllData() {
    // Show confirmation dialog
    if (!confirm("Are you sure you want to reset all data? This will clear all expenses, signatures, and form inputs.")) {
        return; // User cancelled
    }
    
    // Show loading indicator
    loadingOverlay.classList.add('active');
    document.querySelector('.loading-overlay p').textContent = 'Resetting application...';
    
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
        
        // Hide loading indicator
        loadingOverlay.classList.remove('active');
        
        // Show success notification
        showNotification('Application reset successfully', 'success');
        
    } catch (error) {
        console.error('Error resetting application:', error);
        loadingOverlay.classList.remove('active');
        showNotification('Failed to reset application: ' + error.message, 'error');
    }
}

// ... existing code ... 