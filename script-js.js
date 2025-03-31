// Global variables
let xerData = {
    tables: {},
    metadata: {}
};
let dataDate = null;
let kpiData = [
    {
        id: 1,
        name: "Logic",
        description: "% of incomplete tasks without predecessors and/or successors",
        currentValue: 0,
        threshold: 5,
        status: false,
        calculate: calculateLogic,
        getNonCompliantTasks: getNonCompliantLogicTasks
    },
    {
        id: 2,
        name: "Leads",
        description: "% of relationships with negative lag",
        currentValue: 0,
        threshold: 0,
        status: false,
        calculate: calculateLeads,
        getNonCompliantTasks: getNonCompliantLeadsTasks
    },
    {
        id: 3,
        name: "Lags",
        description: "% of relationships with positive lag",
        currentValue: 0,
        threshold: 5,
        status: false,
        calculate: calculateLags,
        getNonCompliantTasks: getNonCompliantLagsTasks
    },
    {
        id: 4,
        name: "Relationship Types",
        description: "% of FS type relationships",
        currentValue: 0,
        threshold: 90,
        status: false,
        calculate: calculateRelationshipTypes,
        getNonCompliantTasks: getNonCompliantRelationshipTypesTasks,
        reverseComparison: true
    },
    {
        id: 5,
        name: "Hard Constraints",
        description: "% of activities with hard constraints",
        currentValue: 0,
        threshold: 5,
        status: false,
        calculate: calculateHardConstraints,
        getNonCompliantTasks: getNonCompliantHardConstraintsTasks
    },
    {
        id: 6,
        name: "High Float",
        description: "% of activities with float > 44 days",
        currentValue: 0,
        threshold: 5,
        status: false,
        calculate: calculateHighFloat,
        getNonCompliantTasks: getNonCompliantHighFloatTasks
    },
    {
        id: 7,
        name: "Negative Float",
        description: "% of activities with float < 0 days",
        currentValue: 0,
        threshold: 0,
        status: false,
        calculate: calculateNegativeFloat,
        getNonCompliantTasks: getNonCompliantNegativeFloatTasks
    },
    {
        id: 8,
        name: "High Duration",
        description: "% of activities with duration > 44 days",
        currentValue: 0,
        threshold: 5,
        status: false,
        calculate: calculateHighDuration,
        getNonCompliantTasks: getNonCompliantHighDurationTasks
    },
    {
        id: 9,
        name: "Invalid Dates",
        description: "Number of activities with invalid dates",
        currentValue: 0,
        threshold: 0,
        status: false,
        calculate: calculateInvalidDates,
        getNonCompliantTasks: getNonCompliantInvalidDatesTasks,
        isCount: true
    },
    {
        id: 10,
        name: "Soft Constraints",
        description: "% of activities with soft constraints",
        currentValue: 0,
        threshold: 5,
        status: false,
        calculate: calculateSoftConstraints,
        getNonCompliantTasks: getNonCompliantSoftConstraintsTasks
    }
];

// Define required tables for XER analysis
const requiredTables = ['TASK', 'TASKPRED'];

// DOM Elements
const fileInput = document.getElementById('fileInput');
const uploadButton = document.getElementById('uploadButton');
const fileName = document.getElementById('fileName');
const dataDateInput = document.getElementById('dataDate');
const processButton = document.getElementById('processButton');
const loadingIndicator = document.getElementById('loadingIndicator');
const dashboardContent = document.getElementById('dashboardContent');
const errorMessage = document.getElementById('errorMessage');
const successMessage = document.getElementById('successMessage');
const downloadExcelButton = document.getElementById('downloadExcelButton');

// Event Listeners
uploadButton.addEventListener('click', function() {
    fileInput.click();
});
fileInput.addEventListener('change', handleFileSelect);
processButton.addEventListener('click', processData);
dataDateInput.addEventListener('input', validateInputs);
downloadExcelButton.addEventListener('click', downloadNonCompliantTasksExcel);

// Event handlers
function handleFileSelect(event) {
    const file = event.target.files[0];
    if (file) {
        fileName.textContent = file.name;
        validateInputs();
        console.log("File selected:", file.name);
        
        // Clear previous error messages when selecting a new file
        errorMessage.style.display = 'none';
    } else {
        fileName.textContent = "No file selected";
    }
}

function validateInputs() {
    const isDateValid = dataDateInput.value !== '';
    const isFileSelected = fileInput.files.length > 0;
    
    processButton.disabled = !(isDateValid && isFileSelected);
}

// Helper function to get relevant tasks (not completed and not WBS Summary)
function getRelevantTasks() {
    const tasks = xerData.tables['TASK'].data;
    
    // Filter tasks that are not completed and not WBS Summary
    return tasks.filter(task => 
        task.status_code !== 'TK_Complete' && 
        task.task_type !== 'TT_WBS'
    );
}

// Helper functions
function parseDate(dateString) {
    if (!dateString) return null;
    
    // Try to parse date string
    const date = new Date(dateString);
    return isNaN(date.getTime()) ? null : date;
}

function formatPercentage(value) {
    return value.toFixed(2) + '%';
}

function showMessage(type, message) {
    if (type === 'error') {
        errorMessage.textContent = message;
        errorMessage.style.display = 'block';
        successMessage.style.display = 'none';
        // No timer for errors - they will remain visible
    } else {
        successMessage.textContent = message;
        successMessage.style.display = 'block';
        errorMessage.style.display = 'none';
        
        // Only success messages disappear automatically
        setTimeout(() => {
            successMessage.style.display = 'none';
        }, 5000);
    }
}

// Function to parse XER file
function parseXERFile(content) {
    const lines = content.split('\n');
    let currentTableName = "";
    let currentColumns = [];
    let lineCount = 0;

    xerData.tables = {};
    xerData.metadata = {};

    for (const line of lines) {
        lineCount++;
        const trimmedLine = line.trim();
        
        // Skip empty lines
        if (trimmedLine === '') continue;
        
        // Analyse the line based on its type
        if (trimmedLine.startsWith('%T')) {
            // Table definition
            currentTableName = trimmedLine.substring(2).trim();
            xerData.tables[currentTableName] = {
                columns: [],
                data: [],
                rowCount: 0
            };
        } else if (trimmedLine.startsWith('%F')) {
            // Column definitions
            if (currentTableName) {
                currentColumns = trimmedLine.substring(2).split('\t');
                xerData.tables[currentTableName].columns = currentColumns;
            }
        } else if (trimmedLine.startsWith('%R')) {
            // Row data
            if (currentTableName && currentColumns.length > 0) {
                const rowData = trimmedLine.substring(2).split('\t');
                const rowObj = {};
                
                // Create an object with row values
                for (let i = 0; i < currentColumns.length && i < rowData.length; i++) {
                    rowObj[currentColumns[i]] = rowData[i];
                }
                
                xerData.tables[currentTableName].data.push(rowObj);
                xerData.tables[currentTableName].rowCount++;
            }
        } else if (trimmedLine.startsWith('%E')) {
            // End of table, no special action needed
        } else if (trimmedLine.startsWith('%')) {
            // Other metadata
            const metaMatch = trimmedLine.match(/%(\w+)(?:\s+(.*))?/);
            if (metaMatch) {
                const metaKey = metaMatch[1];
                const metaValue = metaMatch[2] || '';
                xerData.metadata[metaKey] = metaValue;
            }
        }
    }
    
    return {
        tableCount: Object.keys(xerData.tables).length,
        lineCount: lineCount,
        metadata: xerData.metadata
    };
}

// Function to validate required tables
function validateRequiredTables() {
    const missingTables = [];
    
    for (const tableName of requiredTables) {
        if (!xerData.tables[tableName]) {
            missingTables.push(tableName);
        }
    }
    
    return missingTables;
}

// Process data from XER file
async function processData() {
    try {
        // Show loading indicator
        loadingIndicator.style.display = 'block';
        dashboardContent.classList.add('hidden');
        downloadExcelButton.disabled = true;
        
        // Parse data date from the date input (which returns yyyy-mm-dd format)
        const dateDateValue = dataDateInput.value;
        dataDate = new Date(dateDateValue + 'T00:00:00'); // Add time component for consistent parsing
        
        // Read the file
        const file = fileInput.files[0];
        const text = await file.text();
        
        // Parse the XER file
        const fileStats = parseXERFile(text);
        
        // Validate required tables
        const missingTables = validateRequiredTables();
        if (missingTables.length > 0) {
            throw new Error(`The following required tables do not exist in the file: ${missingTables.join(', ')}`);
        }
        
        // Calculate all KPIs
        for (const kpi of kpiData) {
            kpi.currentValue = kpi.calculate();
            kpi.status = kpi.reverseComparison ? 
                kpi.currentValue >= kpi.threshold : 
                kpi.currentValue <= kpi.threshold;
        }
        
        // Update the dashboard
        updateDashboard();
        
        // Enable the download button if there are non-compliant KPIs
        const hasNonCompliantKpis = kpiData.some(kpi => !kpi.status);
        downloadExcelButton.disabled = !hasNonCompliantKpis;
        
        // Hide loading indicator and show dashboard
        loadingIndicator.style.display = 'none';
        dashboardContent.classList.remove('hidden');
        
        showMessage('success', 'Data processed successfully.');
    } catch (error) {
        console.error('Error processing data:', error);
        loadingIndicator.style.display = 'none';
        downloadExcelButton.disabled = true;
        showMessage('error', 'Error processing data: ' + error.message);
    }
}

// Update all dashboard elements
function updateDashboard() {
    updateTaskCards();
    updateKpiTable();
    updateExecutiveSummary();
}

// Update the task cards in Section 1
function updateTaskCards() {
    const tasks = xerData.tables['TASK'].data;
    
    // Filter out WBS Summary tasks for all counts
    const nonWbsTasks = tasks.filter(task => task.task_type !== 'TT_WBS');
    
    const totalTasks = nonWbsTasks.length;
    const completedTasks = nonWbsTasks.filter(task => task.status_code === 'TK_Complete').length;
    const inProgressTasks = nonWbsTasks.filter(task => task.status_code === 'TK_Active').length;
    const notStartedTasks = nonWbsTasks.filter(task => task.status_code === 'TK_NotStart').length;
    
    const kpisFulfilled = kpiData.filter(kpi => kpi.status).length;
    const totalKpis = kpiData.length;
    
    // Update the cards
    document.getElementById('kpisFulfilled').textContent = `${kpisFulfilled}/${totalKpis}`;
    document.getElementById('kpisFulfilledPercentage').textContent = 
        formatPercentage((kpisFulfilled / totalKpis) * 100);
    
    document.getElementById('tasksCompleted').textContent = completedTasks;
    document.getElementById('tasksCompletedPercentage').textContent = 
        formatPercentage((completedTasks / totalTasks) * 100);
    
    document.getElementById('tasksInProgress').textContent = inProgressTasks;
    document.getElementById('tasksInProgressPercentage').textContent = 
        formatPercentage((inProgressTasks / totalTasks) * 100);
    
    document.getElementById('tasksNotStarted').textContent = notStartedTasks;
    document.getElementById('tasksNotStartedPercentage').textContent = 
        formatPercentage((notStartedTasks / totalTasks) * 100);
}

// Update the KPI table in Section 2
function updateKpiTable() {
    const tableBody = document.querySelector('#kpiTable tbody');
    tableBody.innerHTML = '';
    
    for (const kpi of kpiData) {
        const row = document.createElement('tr');
        
        // Format current value based on whether it's a count or percentage
        const formattedValue = kpi.isCount ? 
            kpi.currentValue : 
            formatPercentage(kpi.currentValue);
        
        row.innerHTML = `
            <td>${kpi.id}</td>
            <td>${kpi.name}</td>
            <td>${kpi.description}</td>
            <td>${formattedValue}</td>
            <td>${kpi.isCount ? kpi.threshold : formatPercentage(kpi.threshold)}</td>
            <td>
                <input type="number" 
                       class="threshold-input" 
                       value="${kpi.threshold}" 
                       min="0" 
                       max="${kpi.reverseComparison ? 100 : 100}" 
                       step="${kpi.isCount ? 1 : 0.1}" 
                       data-kpi-id="${kpi.id}">
            </td>
            <td class="${kpi.status ? 'status-compliant' : 'status-non-compliant'}">
                ${kpi.status ? 'Compliant' : 'Non-Compliant'}
            </td>
        `;
        
        tableBody.appendChild(row);
    }
    
    // Add event listeners to threshold inputs
    document.querySelectorAll('.threshold-input').forEach(input => {
        input.addEventListener('change', updateThreshold);
    });
}

// Update KPI threshold and recalculate status
function updateThreshold(event) {
    const kpiId = parseInt(event.target.dataset.kpiId);
    const newThreshold = parseFloat(event.target.value);
    
    const kpi = kpiData.find(k => k.id === kpiId);
    if (kpi) {
        kpi.threshold = newThreshold;
        kpi.status = kpi.reverseComparison ? 
            kpi.currentValue >= kpi.threshold : 
            kpi.currentValue <= kpi.threshold;
        
        // Update just the status cell
        const row = event.target.closest('tr');
        const statusCell = row.querySelector('td:last-child');
        statusCell.textContent = kpi.status ? 'Compliant' : 'Non-Compliant';
        statusCell.className = kpi.status ? 'status-compliant' : 'status-non-compliant';
        
        // Update the KPIs fulfilled card
        const kpisFulfilled = kpiData.filter(kpi => kpi.status).length;
        const totalKpis = kpiData.length;
        document.getElementById('kpisFulfilled').textContent = `${kpisFulfilled}/${totalKpis}`;
        document.getElementById('kpisFulfilledPercentage').textContent = 
            formatPercentage((kpisFulfilled / totalKpis) * 100);
        
        // Update download button status
        const hasNonCompliantKpis = kpiData.some(kpi => !kpi.status);
        downloadExcelButton.disabled = !hasNonCompliantKpis;
        
        // Update executive summary as well
        updateExecutiveSummary();
    }
}

// Update the executive summary in Section 3
function updateExecutiveSummary() {
    const kpisFulfilled = kpiData.filter(kpi => kpi.status).length;
    const totalKpis = kpiData.length;
    const fulfillmentPercentage = (kpisFulfilled / totalKpis) * 100;
    
    // Current status summary
    let statusText;
    if (fulfillmentPercentage >= 90) {
        statusText = "The project schedule is in an optimal state, meeting " + 
            formatPercentage(fulfillmentPercentage) + " of the evaluated KPIs.";
    } else if (fulfillmentPercentage >= 70) {
        statusText = "The project schedule is in good condition, meeting " + 
            formatPercentage(fulfillmentPercentage) + " of the evaluated KPIs, but there are areas for improvement.";
    } else if (fulfillmentPercentage >= 50) {
        statusText = "The project schedule requires attention, meeting only " + 
            formatPercentage(fulfillmentPercentage) + " of the evaluated KPIs.";
    } else {
        statusText = "The project schedule is in a critical state, meeting only " + 
            formatPercentage(fulfillmentPercentage) + " of the evaluated KPIs. Immediate corrective actions are required.";
    }
    
    document.getElementById('currentStatus').textContent = statusText;
    
    // Improvement proposals
    const proposalsDiv = document.getElementById('improvementProposals');
    proposalsDiv.innerHTML = '';
    
    const nonCompliantKpis = kpiData.filter(kpi => !kpi.status);
    
    if (nonCompliantKpis.length === 0) {
        proposalsDiv.innerHTML = '<p>The schedule meets all KPIs. Excellent work!</p>';
        return;
    }
    
    const proposalsHtml = [];
    
    for (const kpi of nonCompliantKpis) {
        let proposal;
        
        switch (kpi.id) {
            case 1: // Logic
                proposal = `<p><strong>KPI ${kpi.id}: ${kpi.name}</strong> - Review activities without predecessors or successors. Each activity should have at least one Finish-to-Start or Start-to-Start relationship as a predecessor and at least one Finish-to-Start or Finish-to-Finish relationship as a successor.</p>`;
                break;
            case 2: // Leads
                proposal = `<p><strong>KPI ${kpi.id}: ${kpi.name}</strong> - Eliminate negative lags (leads) between activities. Instead, break down activities into smaller tasks that can start independently.</p>`;
                break;
            case 3: // Lags
                proposal = `<p><strong>KPI ${kpi.id}: ${kpi.name}</strong> - Reduce the use of positive lags. Consider creating specific activities that represent the waiting time instead of using lags.</p>`;
                break;
            case 4: // Relationship Types
                proposal = `<p><strong>KPI ${kpi.id}: ${kpi.name}</strong> - Increase the use of Finish-to-Start (FS) relationships that are clearer and easier to monitor. Avoid excessive use of Start-to-Start (SS), Finish-to-Finish (FF), or Start-to-Finish (SF) relationships.</p>`;
                break;
            case 5: // Hard Constraints
                proposal = `<p><strong>KPI ${kpi.id}: ${kpi.name}</strong> - Reduce the use of hard constraints ("Mandatory Start" or "Mandatory Finish"). These constraints can prevent the schedule from being driven by relationship logic and can create negative lags.</p>`;
                break;
            case 6: // High Float
                proposal = `<p><strong>KPI ${kpi.id}: ${kpi.name}</strong> - Review activities with float greater than 44 days. This could indicate activities without adequate predecessors or successors. Add the necessary logical relationships.</p>`;
                break;
            case 7: // Negative Float
                proposal = `<p><strong>KPI ${kpi.id}: ${kpi.name}</strong> - Eliminate all activities with negative float. This is usually related to hard constraints or imposed dates. Review relationship logic and adjust durations or constraints.</p>`;
                break;
            case 8: // High Duration
                proposal = `<p><strong>KPI ${kpi.id}: ${kpi.name}</strong> - Break down activities with duration greater than 44 days into smaller, measurable tasks. This will facilitate monitoring and controlling progress.</p>`;
                break;
            case 9: // Invalid Dates
                proposal = `<p><strong>KPI ${kpi.id}: ${kpi.name}</strong> - Correct all activities with invalid dates: actual start/finish dates in the future or planned start/finish dates in the past without actual dates.</p>`;
                break;
            case 10: // Soft Constraints
                proposal = `<p><strong>KPI ${kpi.id}: ${kpi.name}</strong> - Reduce the use of soft constraints such as "As Late As Possible", "Start On", "Start On or After", "Start On or Before", "Finish On", "Finish On or After", or "Finish On or Before". These constraints can interfere with schedule logic. Consider replacing them with appropriate logical relationships.</p>`;
                break;
            default:
                proposal = `<p><strong>KPI ${kpi.id}: ${kpi.name}</strong> - Review this KPI to improve the overall schedule health.</p>`;
        }
        
        proposalsHtml.push(proposal);
    }
    
    proposalsDiv.innerHTML = proposalsHtml.join('');
}

// ====== KPI Calculation Functions ====== //

function calculateLogic() {
    const relevantTasks = getRelevantTasks();
    const taskpred = xerData.tables['TASKPRED'].data;
    
    const totalRelevantTasks = relevantTasks.length;
    
    if (totalRelevantTasks === 0) return 0;
    
    // Create a set of task IDs for faster lookup
    const relevantTaskIds = new Set();
    relevantTasks.forEach(task => {
        relevantTaskIds.add(task.task_id);
    });
    
    // Create sets for tasks with predecessors and successors
    const tasksWithPred = new Set();
    const tasksWithSucc = new Set();
    
    // Check all relationships in TASKPRED
    taskpred.forEach(relation => {
        // If the task is in our relevant tasks, mark it as having a predecessor
        if (relevantTaskIds.has(relation.task_id)) {
            tasksWithPred.add(relation.task_id);
        }
        
        // If the predecessor is in our relevant tasks, mark it as having a successor
        if (relevantTaskIds.has(relation.pred_task_id)) {
            tasksWithSucc.add(relation.pred_task_id);
        }
    });
    
    // Count tasks without predecessors or successors
    let tasksWithoutPredCount = 0;
    let tasksWithoutSuccCount = 0;
    
    relevantTaskIds.forEach(taskId => {
        if (!tasksWithPred.has(taskId)) tasksWithoutPredCount++;
        if (!tasksWithSucc.has(taskId)) tasksWithoutSuccCount++;
    });
    
    // Calculate percentages
    const percentWithoutPred = (tasksWithoutPredCount / totalRelevantTasks) * 100;
    const percentWithoutSucc = (tasksWithoutSuccCount / totalRelevantTasks) * 100;
    
    return percentWithoutPred + percentWithoutSucc;
}

function calculateLeads() {
    const relevantTasks = getRelevantTasks();
    const taskpred = xerData.tables['TASKPRED'].data;
    
    // Get IDs of all relevant tasks
    const relevantTaskIds = new Set();
    relevantTasks.forEach(task => {
        relevantTaskIds.add(task.task_id);
    });
    
    // Filter relationships where either the successor or predecessor is a relevant task
    const relevantRelations = taskpred.filter(relation =>
        relevantTaskIds.has(relation.task_id) || 
        relevantTaskIds.has(relation.pred_task_id)
    );
    
    const totalRelevantRelations = relevantRelations.length;
    if (totalRelevantRelations === 0) return 0;
    
    // Count relations with negative lag
    const relationsWithNegativeLag = relevantRelations.filter(rel => {
        const lagValue = parseInt(rel.lag_hr_cnt);
        return !isNaN(lagValue) && lagValue < 0;
    }).length;
    
    return (relationsWithNegativeLag / totalRelevantRelations) * 100;
}

function calculateLags() {
    const relevantTasks = getRelevantTasks();
    const taskpred = xerData.tables['TASKPRED'].data;
    
    // Get IDs of all relevant tasks
    const relevantTaskIds = new Set();
    relevantTasks.forEach(task => {
        relevantTaskIds.add(task.task_id);
    });
    
    // Filter relationships where either the successor or predecessor is a relevant task
    const relevantRelations = taskpred.filter(relation =>
        relevantTaskIds.has(relation.task_id) || 
        relevantTaskIds.has(relation.pred_task_id)
    );
    
    const totalRelevantRelations = relevantRelations.length;
    if (totalRelevantRelations === 0) return 0;
    
    // Count relations with positive lag
    const relationsWithPositiveLag = relevantRelations.filter(rel => {
        const lagValue = parseInt(rel.lag_hr_cnt);
        return !isNaN(lagValue) && lagValue > 0;
    }).length;
    
    return (relationsWithPositiveLag / totalRelevantRelations) * 100;
}

function calculateRelationshipTypes() {
    const tasks = xerData.tables['TASK'].data;
    const taskpred = xerData.tables['TASKPRED'].data;
    
    // Create a map of task ID to task info for efficient lookup
    const taskInfoMap = {};
    tasks.forEach(task => {
        taskInfoMap[task.task_id] = {
            status: task.status_code,
            type: task.task_type
        };
    });
    
    // Filter relationships where BOTH the predecessor AND successor are:
    // 1. Of type TT_Task (Task Dependent)
    // 2. Not completed (status_code !== 'TK_Complete')
    const relevantRelations = taskpred.filter(relation => {
        const predTaskInfo = taskInfoMap[relation.pred_task_id];
        const succTaskInfo = taskInfoMap[relation.task_id];
        
        return predTaskInfo && succTaskInfo && 
               predTaskInfo.status !== 'TK_Complete' && 
               succTaskInfo.status !== 'TK_Complete' &&
               predTaskInfo.type === 'TT_Task' &&
               succTaskInfo.type === 'TT_Task';
    });
    
    const totalRelevantRelations = relevantRelations.length;
    if (totalRelevantRelations === 0) return 100;
    
    // Count FS relations
    const fsRelations = relevantRelations.filter(rel => rel.pred_type === 'PR_FS').length;
    
    return (fsRelations / totalRelevantRelations) * 100;
}

function calculateHardConstraints() {
    const relevantTasks = getRelevantTasks();
    const totalRelevantTasks = relevantTasks.length;
    
    if (totalRelevantTasks === 0) return 0;
    
    // Count tasks with hard constraints (CS_MANDSTART or CS_MANDFIN)
    const tasksWithHardConstraints = relevantTasks.filter(task => 
        task.cstr_type === 'CS_MANDSTART' || task.cstr_type === 'CS_MANDFIN'
    ).length;
    
    return (tasksWithHardConstraints / totalRelevantTasks) * 100;
}

function calculateHighFloat() {
    const relevantTasks = getRelevantTasks();
    const totalRelevantTasks = relevantTasks.length;
    
    if (totalRelevantTasks === 0) return 0;
    
    // Count tasks with high float (> 44 days, or > 352 hours assuming 8-hour workday)
    const tasksWithHighFloat = relevantTasks.filter(task => {
        const floatValue = parseInt(task.total_float_hr_cnt);
        return !isNaN(floatValue) && floatValue > 352; // 44 days * 8 hours
    }).length;
    
    return (tasksWithHighFloat / totalRelevantTasks) * 100;
}

function calculateNegativeFloat() {
    const relevantTasks = getRelevantTasks();
    const totalRelevantTasks = relevantTasks.length;
    
    if (totalRelevantTasks === 0) return 0;
    
    // Count tasks with negative float
    const tasksWithNegativeFloat = relevantTasks.filter(task => {
        const floatValue = parseInt(task.total_float_hr_cnt);
        return !isNaN(floatValue) && floatValue < 0;
    }).length;
    
    return (tasksWithNegativeFloat / totalRelevantTasks) * 100;
}

function calculateHighDuration() {
    const relevantTasks = getRelevantTasks();
    const totalRelevantTasks = relevantTasks.length;
    
    if (totalRelevantTasks === 0) return 0;
    
    // Count tasks with high duration (> 44 days, or > 352 hours assuming 8-hour workday)
    const tasksWithHighDuration = relevantTasks.filter(task => {
        const durationValue = parseInt(task.target_drtn_hr_cnt);
        return !isNaN(durationValue) && durationValue > 352; // 44 days * 8 hours
    }).length;
    
    return (tasksWithHighDuration / totalRelevantTasks) * 100;
}

function calculateInvalidDates() {
    const relevantTasks = getRelevantTasks();
    
    let invalidDatesCount = 0;
    
    for (const task of relevantTasks) {
        const actStartDate = task.act_start_date ? parseDate(task.act_start_date) : null;
        const actEndDate = task.act_end_date ? parseDate(task.act_end_date) : null;
        
        if ((actStartDate && actStartDate > dataDate) || 
            (actEndDate && actEndDate > dataDate)) {
            invalidDatesCount++;
        }
    }
    
    return invalidDatesCount;
}

function calculateSoftConstraints() {
    const relevantTasks = getRelevantTasks();
    const totalRelevantTasks = relevantTasks.length;
    
    if (totalRelevantTasks === 0) return 0;
    
    // List of soft constraint types in P6/XER
    const softConstraintTypes = [
        'CS_MEO',   // Must End On
        'CS_MEOA',  // Must End On or After
        'CS_MSO',   // Must Start On
        'CS_MSOA',  // Must Start On or After
        'CS_ALAP',  // As Late As Possible
        'CS_MEOB',  // Must End On or Before
        'CS_MSOB'   // Must Start On or Before
    ];
    
    // Count tasks with soft constraints
    const tasksWithSoftConstraints = relevantTasks.filter(task => 
        softConstraintTypes.includes(task.cstr_type)
    ).length;
    
    return (tasksWithSoftConstraints / totalRelevantTasks) * 100;
}

// ====== Non-Compliant Tasks Functions ====== //

function getNonCompliantLogicTasks() {
    const relevantTasks = getRelevantTasks();
    const taskpred = xerData.tables['TASKPRED'].data;
    
    // Create task ID to name mapping for reporting
    const taskIdToNameMap = {};
    relevantTasks.forEach(task => {
        taskIdToNameMap[task.task_id] = {
            name: task.task_name,
            code: task.task_code || task.task_id
        };
    });
    
    // Create sets for task IDs with predecessors and successors
    const tasksWithPred = new Set();
    const tasksWithSucc = new Set();
    
    // Check all relationships in TASKPRED
    taskpred.forEach(relation => {
        tasksWithPred.add(relation.task_id);
        tasksWithSucc.add(relation.pred_task_id);
    });
    
    // Find tasks without predecessors or successors
    const result = [];
    
    relevantTasks.forEach(task => {
        const hasPred = tasksWithPred.has(task.task_id);
        const hasSucc = tasksWithSucc.has(task.task_id);
        
        if (!hasPred || !hasSucc) {
            result.push({
                task_id: task.task_id,
                task_code: task.task_code || task.task_id,
                task_name: task.task_name,
                issue: !hasPred && !hasSucc ? 'No predecessors and no successors' :
                       !hasPred ? 'No predecessors' : 'No successors'
            });
        }
    });
    
    return result;
}

function getNonCompliantLeadsTasks() {
    const relevantTasks = getRelevantTasks();
    const taskpred = xerData.tables['TASKPRED'].data;
    
    // Create task ID to name/code mapping for reporting
    const taskMap = {};
    xerData.tables['TASK'].data.forEach(task => {
        taskMap[task.task_id] = {
            name: task.task_name,
            code: task.task_code || task.task_id
        };
    });
    
    // Get IDs of all relevant tasks
    const relevantTaskIds = new Set();
    relevantTasks.forEach(task => {
        relevantTaskIds.add(task.task_id);
    });
    
    // Filter and find relationships with negative lag
    const relationsWithNegativeLag = taskpred.filter(relation => {
        const lagValue = parseInt(relation.lag_hr_cnt);
        return (relevantTaskIds.has(relation.task_id) || 
               relevantTaskIds.has(relation.pred_task_id)) &&
               !isNaN(lagValue) && lagValue < 0;
    });
    
    // Create result with details for reporting
    return relationsWithNegativeLag.map(rel => {
        const predInfo = taskMap[rel.pred_task_id] || { name: 'Unknown Task', code: rel.pred_task_id };
        const succInfo = taskMap[rel.task_id] || { name: 'Unknown Task', code: rel.task_id };
        
        return {
            task_id: rel.task_id,
            task_code: succInfo.code,
            task_name: succInfo.name,
            pred_task_id: rel.pred_task_id,
            pred_task_code: predInfo.code,
            pred_task_name: predInfo.name,
            lag_hr_cnt: rel.lag_hr_cnt,
            pred_type: rel.pred_type
        };
    });
}

function getNonCompliantLagsTasks() {
    const relevantTasks = getRelevantTasks();
    const taskpred = xerData.tables['TASKPRED'].data;
    
    // Create task ID to name/code mapping for reporting
    const taskMap = {};
    xerData.tables['TASK'].data.forEach(task => {
        taskMap[task.task_id] = {
            name: task.task_name,
            code: task.task_code || task.task_id
        };
    });
    
    // Get IDs of all relevant tasks
    const relevantTaskIds = new Set();
    relevantTasks.forEach(task => {
        relevantTaskIds.add(task.task_id);
    });
    
    // Filter and find relationships with positive lag
    const relationsWithPositiveLag = taskpred.filter(relation => {
        const lagValue = parseInt(relation.lag_hr_cnt);
        return (relevantTaskIds.has(relation.task_id) || 
               relevantTaskIds.has(relation.pred_task_id)) &&
               !isNaN(lagValue) && lagValue > 0;
    });
    
    // Create result with details for reporting
    return relationsWithPositiveLag.map(rel => {
        const predInfo = taskMap[rel.pred_task_id] || { name: 'Unknown Task', code: rel.pred_task_id };
        const succInfo = taskMap[rel.task_id] || { name: 'Unknown Task', code: rel.task_id };
        
        return {
            task_id: rel.task_id,
            task_code: succInfo.code,
            task_name: succInfo.name,
            pred_task_id: rel.pred_task_id,
            pred_task_code: predInfo.code,
            pred_task_name: predInfo.name,
            lag_hr_cnt: rel.lag_hr_cnt,
            pred_type: rel.pred_type
        };
    });
}

function getNonCompliantRelationshipTypesTasks() {
    const tasks = xerData.tables['TASK'].data;
    const taskpred = xerData.tables['TASKPRED'].data;
    
    // Create task ID to task info mapping for reporting
    const taskMap = {};
    tasks.forEach(task => {
        taskMap[task.task_id] = {
            name: task.task_name,
            code: task.task_code || task.task_id,
            status: task.status_code,
            type: task.task_type
        };
    });
    
    // Filter for relationships where BOTH the predecessor AND successor are:
    // 1. Of type TT_Task (Task Dependent)
    // 2. Not completed (status_code !== 'TK_Complete')
    // 3. Relationship type is not FS
    const nonFsRelations = taskpred.filter(relation => {
        const predInfo = taskMap[relation.pred_task_id];
        const succInfo = taskMap[relation.task_id];
        
        return predInfo && succInfo && 
               predInfo.status !== 'TK_Complete' && 
               succInfo.status !== 'TK_Complete' &&
               predInfo.type === 'TT_Task' &&
               succInfo.type === 'TT_Task' &&
               relation.pred_type !== 'PR_FS';
    });
    
    // Create result with details for reporting
    return nonFsRelations.map(rel => {
        const predInfo = taskMap[rel.pred_task_id] || { name: 'Unknown Task', code: rel.pred_task_id };
        const succInfo = taskMap[rel.task_id] || { name: 'Unknown Task', code: rel.task_id };
        
        return {
            task_id: rel.task_id,
            task_code: succInfo.code,
            task_name: succInfo.name,
            pred_task_id: rel.pred_task_id,
            pred_task_code: predInfo.code,
            pred_task_name: predInfo.name,
            pred_type: rel.pred_type
        };
    });
}

function getNonCompliantHardConstraintsTasks() {
    const relevantTasks = getRelevantTasks();
    
    // Get relevant tasks with hard constraints (CS_MANDSTART or CS_MANDFIN)
    const tasksWithHardConstraints = relevantTasks.filter(task => 
        task.cstr_type === 'CS_MANDSTART' || task.cstr_type === 'CS_MANDFIN'
    );
    
    return tasksWithHardConstraints.map(task => {
        return {
            task_id: task.task_id,
            task_code: task.task_code || task.task_id,
            task_name: task.task_name,
            constraint_type: task.cstr_type,
            constraint_date: task.cstr_date ? parseDate(task.cstr_date).toLocaleDateString() : 'N/A'
        };
    });
}

function getNonCompliantHighFloatTasks() {
    const relevantTasks = getRelevantTasks();
    
    // Get relevant tasks with high float
    const tasksWithHighFloat = relevantTasks.filter(task => {
        const floatValue = parseInt(task.total_float_hr_cnt);
        return !isNaN(floatValue) && floatValue > 352; // 44 days * 8 hours
    });
    
    return tasksWithHighFloat.map(task => {
        return {
            task_id: task.task_id,
            task_code: task.task_code || task.task_id,
            task_name: task.task_name,
            total_float: task.total_float_hr_cnt
        };
    });
}

function getNonCompliantNegativeFloatTasks() {
    const relevantTasks = getRelevantTasks();
    
    // Get relevant tasks with negative float
    const tasksWithNegativeFloat = relevantTasks.filter(task => {
        const floatValue = parseInt(task.total_float_hr_cnt);
        return !isNaN(floatValue) && floatValue < 0;
    });
    
    return tasksWithNegativeFloat.map(task => {
        return {
            task_id: task.task_id,
            task_code: task.task_code || task.task_id,
            task_name: task.task_name,
            total_float: task.total_float_hr_cnt
        };
    });
}

function getNonCompliantHighDurationTasks() {
    const relevantTasks = getRelevantTasks();
    
    // Get relevant tasks with high duration
    const tasksWithHighDuration = relevantTasks.filter(task => {
        const durationValue = parseInt(task.target_drtn_hr_cnt);
        return !isNaN(durationValue) && durationValue > 352; // 44 days * 8 hours
    });
    
    return tasksWithHighDuration.map(task => {
        return {
            task_id: task.task_id,
            task_code: task.task_code || task.task_id,
            task_name: task.task_name,
            duration: task.target_drtn_hr_cnt
        };
    });
}

function getNonCompliantInvalidDatesTasks() {
    const relevantTasks = getRelevantTasks();
    
    // Get tasks with invalid dates
    const tasksWithInvalidDates = [];
    
    for (const task of relevantTasks) {
        const actStartDate = task.act_start_date ? parseDate(task.act_start_date) : null;
        const actEndDate = task.act_end_date ? parseDate(task.act_end_date) : null;
        
        if ((actStartDate && actStartDate > dataDate) || 
            (actEndDate && actEndDate > dataDate)) {
            tasksWithInvalidDates.push({
                task_id: task.task_id,
                task_code: task.task_code || task.task_id,
                task_name: task.task_name,
                act_start_date: actStartDate ? actStartDate.toLocaleDateString() : 'N/A',
                act_end_date: actEndDate ? actEndDate.toLocaleDateString() : 'N/A',
                issue: (actStartDate && actStartDate > dataDate) ? 
                       'Actual start date in the future' : 
                       'Actual finish date in the future'
            });
        }
    }
    
    return tasksWithInvalidDates;
}

function getNonCompliantSoftConstraintsTasks() {
    const relevantTasks = getRelevantTasks();
    
    // List of soft constraint types in P6/XER
    const softConstraintTypes = [
        'CS_MEO',   // Must End On
        'CS_MEOA',  // Must End On or After
        'CS_MSO',   // Must Start On
        'CS_MSOA',  // Must Start On or After
        'CS_ALAP',  // As Late As Possible
        'CS_MEOB',  // Must End On or Before
        'CS_MSOB'   // Must Start On or Before
    ];
    
    // Get relevant tasks with soft constraints
    const tasksWithSoftConstraints = relevantTasks.filter(task => 
        softConstraintTypes.includes(task.cstr_type)
    );
    
    return tasksWithSoftConstraints.map(task => {
        return {
            task_id: task.task_id,
            task_code: task.task_code || task.task_id,
            task_name: task.task_name,
            constraint_type: task.cstr_type,
            constraint_date: task.cstr_date ? parseDate(task.cstr_date).toLocaleDateString() : 'N/A'
        };
    });
}

// Excel generation and download
function downloadNonCompliantTasksExcel() {
    try {
        // Show loading indicator
        loadingIndicator.style.display = 'block';
        
        // Get non-compliant KPIs
        const nonCompliantKpis = kpiData.filter(kpi => !kpi.status);
        
        if (nonCompliantKpis.length === 0) {
            showMessage('success', 'There are no non-compliant KPIs to report.');
            loadingIndicator.style.display = 'none';
            return;
        }
        
        // Create a new workbook
        const workbook = XLSX.utils.book_new();
        
        // Process each non-compliant KPI
        for (const kpi of nonCompliantKpis) {
            // Get the non-compliant tasks for this KPI
            let nonCompliantTasks = kpi.getNonCompliantTasks();
            
            if (nonCompliantTasks.length === 0) continue;
            
            // Ensure tasks have at least the required columns
            nonCompliantTasks = nonCompliantTasks.map(task => {
                // Create a new object with at least task_code and task_name
                const newTask = {
                    task_code: task.task_code || '',
                    task_name: task.task_name || ''
                };
                
                // Copy all other properties
                for (const key in task) {
                    if (key !== 'task_code' && key !== 'task_name') {
                        newTask[key] = task[key];
                    }
                }
                
                return newTask;
            });
            
            // Create a worksheet from the array of objects
            const worksheet = XLSX.utils.json_to_sheet(nonCompliantTasks);
            
            // Add the worksheet to the workbook with KPI name as sheet name
            const sheetName = `KPI ${kpi.id} - ${kpi.name}`;
            XLSX.utils.book_append_sheet(workbook, worksheet, sheetName);
        }
        
        // Generate Excel file
        const excelOutput = XLSX.write(workbook, { 
            bookType: 'xlsx', 
            type: 'array' 
        });
        
        // Create a Blob from the output
        const blob = new Blob([excelOutput], { 
            type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' 
        });
        
        // Create download link
        const url = URL.createObjectURL(blob);
        const fileName = `Non_Compliant_KPIs_${new Date().toISOString().slice(0, 10)}.xlsx`;
        
        const link = document.createElement('a');
        link.href = url;
        link.download = fileName;
        document.body.appendChild(link);
        link.click();
        
        // Clean up
        setTimeout(() => {
            document.body.removeChild(link);
            URL.revokeObjectURL(url);
        }, 100);
        
        // Hide loading indicator
        loadingIndicator.style.display = 'none';
        
        showMessage('success', `Excel file "${fileName}" downloaded successfully.`);
    } catch (error) {
        console.error('Error generating Excel file:', error);
        loadingIndicator.style.display = 'none';
        showMessage('error', 'Error generating Excel file: ' + error.message);
    }
}