// Global variables for data and charts
let allData = [];
let filteredData = [];
let charts = {};
let currentDetailData = null;

// DOM Elements
const showInfoBtn = document.getElementById('showInfoBtn');
const mainButtons = document.getElementById('mainButtons');
const infoSystem = document.getElementById('infoSystem');
const searchBtn = document.getElementById('searchBtn');
const clearBtn = document.getElementById('clearBtn');
const backBtn = document.getElementById('backBtn');
const resultsSection = document.getElementById('resultsSection');
const resultsTable = document.getElementById('resultsTable');
const totalCount = document.getElementById('totalCount');
const totalHours = document.getElementById('totalHours');
const detailModal = document.getElementById('detailModal');
const closeModal = document.getElementById('closeModal');
const closeModalBtn = document.getElementById('closeModalBtn');
const modalContent = document.getElementById('modalContent');
const downloadCsvBtn = document.getElementById('downloadCsvBtn');
const downloadDetailPdfBtn = document.getElementById('downloadDetailPdfBtn');

// Initialize charts with common options and colors
function initializeCharts() {
    const chartOptions = {
        responsive: true,
        maintainAspectRatio: false,
        plugins: {
            legend: {
                position: 'right',
                labels: {
                    font: {
                        family: 'Sarabun' // Set font for chart legends
                    },
                    generateLabels: function(chart) {
                        const datasets = chart.data.datasets;
                        const labels = chart.data.labels;
                        const total = datasets[0].data.reduce((a, b) => a + b, 0);
                        
                        return labels.map((label, i) => {
                            const value = datasets[0].data[i];
                            const percentage = ((value / total) * 100).toFixed(1);
                            
                            return {
                                text: `${label}: ${value} (${percentage}%)`,
                                fillStyle: datasets[0].backgroundColor[i],
                                hidden: false,
                                index: i
                            };
                        });
                    }
                }
            },
            tooltip: {
                callbacks: {
                    label: function(context) {
                        const label = context.label || '';
                        const value = context.raw || 0;
                        const total = context.dataset.data.reduce((a, b) => a + b, 0);
                        const percentage = ((value / total) * 100).toFixed(1);
                        return `${label}: ${value} (${percentage}%)`;
                    }
                }
            }
        }
    };

    const chartColors = [
        '#FF6384', '#36A2EB', '#FFCE56', '#4BC0C0', '#9966FF',
        '#FF9F40', '#8AC249', '#EA80FC', '#00E5FF', '#FF5252'
    ];

    // Create pie charts for different data categories
    charts.status = new Chart(document.getElementById('statusChart'), {
        type: 'pie',
        data: {
            labels: [],
            datasets: [{
                data: [],
                backgroundColor: chartColors
            }]
        },
        options: chartOptions
    });

    charts.department = new Chart(document.getElementById('departmentChart'), {
        type: 'pie',
        data: {
            labels: [],
            datasets: [{
                data: [],
                backgroundColor: chartColors
            }]
        },
        options: chartOptions
    });

    charts.activityType = new Chart(document.getElementById('activityTypeChart'), {
        type: 'pie',
        data: {
            labels: [],
            datasets: [{
                data: [],
                backgroundColor: chartColors
            }]
        },
        options: chartOptions
    });

    charts.level = new Chart(document.getElementById('levelChart'), {
        type: 'pie',
        data: {
            labels: [],
            datasets: [{
                data: [],
                backgroundColor: chartColors
            }]
        },
        options: chartOptions
    });

    charts.format = new Chart(document.getElementById('formatChart'), {
        type: 'pie',
        data: {
            labels: [],
            datasets: [{
                data: [],
                backgroundColor: chartColors
            }]
        },
        options: chartOptions
    });
}

// Fetch data from Google Apps Script endpoint
async function fetchData() {
    try {
        // Replace with your actual Google Apps Script deployed URL
        const response = await fetch('https://script.google.com/macros/s/AKfycby3b89RQLEK3Ps2AGaDwnhcMLEY66MaCUVCl7mQnDlugVei7MGQt2Qgm0UHeWU4mhzT/exec');
        const data = await response.json();
        
        if (data && data.length > 0) {
            // Remove header row (assuming the first row is headers)
            allData = data.slice(1);
            return allData;
        } else {
            console.error('No data received or empty data');
            return [];
        }
    } catch (error) {
        console.error('Error fetching data:', error);
        return [];
    }
}

// Filter data based on user search criteria and then sort it
function filterData() {
    const activityType = document.getElementById('activityType').value;
    const name = document.getElementById('name').value.toLowerCase();
    const department = document.getElementById('department').value.toLowerCase();
    const level = document.getElementById('level').value.toLowerCase();
    const activityFormat = document.getElementById('activityFormat').value.toLowerCase();
    const startDate = document.getElementById('startDate').value;
    const endDate = document.getElementById('endDate').value;

    filteredData = allData.filter(row => {
        // Ensure row has enough elements to avoid errors
        if (!row || row.length < 15) return false;
        
        // Get values from specific columns (adjust indices if your spreadsheet structure changes)
        const rowActivityType = (row[4] || '').toString().toLowerCase(); // Column E: ประเภทกิจกรรม
        const rowName = (row[1] || '').toString().toLowerCase();         // Column B: ชื่อ
        const rowDepartment = (row[3] || '').toString().toLowerCase();   // Column D: สังกัด
        const rowLevel = (row[8] || '').toString().toLowerCase();        // Column I: ระดับ
        const rowFormat = (row[11] || '').toString().toLowerCase();      // Column L: รูปแบบกิจกรรม
        const rowStartDate = row[9] || '';                               // Column J: ตั้งแต่วันที่
        const rowEndDate = row[10] || '';                                // Column K: ถึงวันที่
        
        // Apply filters based on input values
        if (activityType && !rowActivityType.includes(activityType.toLowerCase())) return false;
        if (name && !rowName.includes(name)) return false;
        if (department && !rowDepartment.includes(department)) return false;
        if (level && !rowLevel.includes(level)) return false;
        if (activityFormat && !rowFormat.includes(activityFormat)) return false;
        
        // Date filtering
        if (startDate) {
            const filterStartDate = new Date(startDate);
            const rowStart = new Date(rowStartDate);
            if (isNaN(rowStart.getTime()) || rowStart < filterStartDate) return false;
        }
        
        if (endDate) {
            const filterEndDate = new Date(endDate);
            const rowEnd = new Date(rowEndDate);
            if (isNaN(rowEnd.getTime()) || rowEnd > filterEndDate) return false;
        }
        
        return true;
    });
    
    // Sort filtered data by 'ตั้งแต่วันที่' (row[9]) in ascending order
    filteredData.sort((a, b) => {
        const dateA = a[9] ? new Date(a[9]) : new Date(0); // Use epoch for empty/invalid dates
        const dateB = b[9] ? new Date(b[9]) : new Date(0);
        return dateA.getTime() - dateB.getTime();
    });

    return filteredData;
}

// Display filtered results in the HTML table
function displayResults() {
    resultsTable.innerHTML = ''; // Clear previous results
    
    if (filteredData.length === 0) {
        resultsTable.innerHTML = `
            <tr>
                <td colspan="8" class="py-4 px-4 text-center text-gray-500">ไม่พบข้อมูลที่ตรงกับเงื่อนไขการค้นหา</td>
            </tr>
        `;
        return;
    }
    
    filteredData.forEach((row, index) => {
        const tr = document.createElement('tr');
        tr.className = index % 2 === 0 ? 'bg-white' : 'bg-gray-50';
        
        // Format dates for display
        const startDate = row[9] ? formatDate(row[9]) : '-';
        const endDate = row[10] ? formatDate(row[10]) : '-';
        
        // Get training hours (column U = index 20)
        const trainingHours = row.length > 20 && row[20] ? row[20] : '-';
        
        // Populate table row with data
        tr.innerHTML = `
            <td class="py-3 px-4 border-b">${startDate}</td>
            <td class="py-3 px-4 border-b">${endDate}</td>
            <td class="py-3 px-4 border-b">${row[1] || '-'}</td>
            <td class="py-3 px-4 border-b">${row[4] || '-'}</td>
            <td class="py-3 px-4 border-b">${row[5] || '-'}</td>
            <td class="py-3 px-4 border-b">${row[6] || '-'}</td>
            <td class="py-3 px-4 border-b">${trainingHours}</td>
            <td class="py-3 px-4 border-b text-center">
                <button class="detail-btn bg-red-500 hover:bg-red-600 text-white py-1 px-3 rounded-md text-sm transition duration-300" 
                        data-index="${index}">
                    ดูรายละเอียด
                </button>
            </td>
        `;
        
        resultsTable.appendChild(tr);
    });
    
    // Add event listeners to "ดูรายละเอียด" buttons
    document.querySelectorAll('.detail-btn').forEach(btn => {
        btn.addEventListener('click', () => {
            const index = parseInt(btn.getAttribute('data-index'));
            showDetailModal(filteredData[index]);
        });
    });
}

// Helper function to format date strings
function formatDate(dateString) {
    if (!dateString) return '-';
    
    try {
        const date = new Date(dateString);
        // Check if date is valid
        if (isNaN(date.getTime())) return dateString; 
        
        return date.toLocaleDateString('th-TH', {
            year: 'numeric',
            month: 'short',
            day: 'numeric'
        });
    } catch (e) {
        return dateString; // Return original string if parsing fails
    }
}

// Helper function to extract Google Drive File ID from URL
function getGoogleDriveFileId(url) {
    if (!url) return null;
    let match;

    // Common Google Drive share link: /d/FILE_ID/view
    match = url.match(/\/d\/([a-zA-Z0-9_-]+)(?:\/view)?/);
    if (match && match[1]) return match[1];
    
    // Google Drive file ID parameter: ?id=FILE_ID
    match = url.match(/id=([a-zA-Z0-9_-]+)/);
    if (match && match[1]) return match[1];

    // Google Drive open link: /open?id=FILE_ID
    match = url.match(/\/open\?id=([a-zA-Z0-9_-]+)/);
    if (match && match[1]) return match[1];

    // Direct file link (e.g., from Google Sheets IMAGE() function or direct download links)
    // This pattern covers links like https://drive.google.com/uc?id=FILE_ID&export=download
    match = url.match(/uc\?id=([a-zA-Z0-9_-]+)/);
    if (match && match[1]) return match[1];

    // Embed links
    match = url.match(/\/file\/d\/([a-zA-Z0-9_-]+)\/preview/);
    if (match && match[1]) return match[1];

    return null;
}

// Show detailed information in a modal
function showDetailModal(rowData) {
    modalContent.innerHTML = ''; // Clear previous modal content
    currentDetailData = rowData; // Store current data for PDF download
    
    if (!rowData || rowData.length < 15) {
        modalContent.innerHTML = '<p class="text-center text-gray-500">ไม่พบข้อมูล</p>';
        detailModal.classList.remove('hidden');
        return;
    }
    
    // Get training hours (column U = index 20)
    const trainingHours = rowData.length > 20 ? rowData[20] : null;
    
    // Define fields to display in the modal (excluding 'ประทับเวลา')
    const fields = [
        // { label: 'ประทับเวลา', value: rowData[0] }, // Removed as per user request
        { label: 'ชื่อ', value: rowData[1] },
        { label: 'สถานะ', value: rowData[2] },
        { label: 'สังกัด', value: rowData[3] },
        { label: 'ประเภทกิจกรรม', value: rowData[4] },
        { label: 'ชื่อกิจกรรม', value: rowData[5] },
        { label: 'สิ่งที่ได้รับ', value: rowData[6] },
        { label: 'รายละเอียด', value: rowData[7] },
        { label: 'ระดับ', value: rowData[8] },
        { label: 'ตั้งแต่วันที่', value: formatDate(rowData[9]) },
        { label: 'ถึงวันที่', value: formatDate(rowData[10]) },
        { label: 'รูปแบบกิจกรรม', value: rowData[11] },
        { label: 'หน่วยงานที่จัด', value: rowData[12] },
        { label: 'สถานที่', value: rowData[13] },
        { label: 'จำนวนชั่วโมงอบรม', value: trainingHours || '-' }
    ];
    
    // Create and append detail fields to the modal
    const detailsDiv = document.createElement('div');
    detailsDiv.className = 'grid grid-cols-1 md:grid-cols-2 gap-4';
    
    fields.forEach(field => {
        const fieldDiv = document.createElement('div');
        fieldDiv.className = 'border-b pb-2';
        fieldDiv.innerHTML = `
            <p class="font-bold text-red-700">${field.label}</p>
            <p>${field.value || '-'}</p>
        `;
        detailsDiv.appendChild(fieldDiv);
    });
    
    modalContent.appendChild(detailsDiv);
    
    // Add attachments section with thumbnails
    const attachments = rowData.slice(14, 20).filter(Boolean); // Filter out empty strings/nulls
    
    if (attachments.length > 0) {
        const attachmentsDiv = document.createElement('div');
        attachmentsDiv.className = 'mt-4';
        attachmentsDiv.innerHTML = '<h4 class="font-bold text-lg text-red-700 mb-2">เอกสารแนบ</h4>';
        
        const attachmentsGrid = document.createElement('div');
        // Adjusted grid for better thumbnail display, ensuring responsiveness
        attachmentsGrid.className = 'grid grid-cols-2 sm:grid-cols-3 lg:grid-cols-4 gap-3'; 
        
        attachments.forEach((url, idx) => {
            if (url && url.trim() !== '') { // Ensure URL is not empty or just whitespace
                const fileId = getGoogleDriveFileId(url);
                const thumbnailUrl = fileId ? `https://drive.google.com/thumbnail?id=${fileId}&sz=w300` : null;
                
                const attachmentItem = document.createElement('div');
                attachmentItem.className = 'flex flex-col items-center p-2 bg-red-50 hover:bg-red-100 rounded-md transition duration-300';
                
                if (thumbnailUrl) {
                    attachmentItem.innerHTML = `
                        <a href="${url}" target="_blank" class="block w-full h-48 overflow-hidden rounded-md mb-2">
                            <img src="${thumbnailUrl}" alt="เอกสารแนบ ${idx + 1}" class="w-full h-full object-contain">
                        </a>
                        <a href="${url}" target="_blank" class="text-sm text-red-600 hover:underline">เอกสารแนบ ${idx + 1}</a>
                    `;
                } else {
                    attachmentItem.innerHTML = `
                        <a href="${url}" target="_blank" class="flex items-center justify-center w-full h-48 bg-gray-200 rounded-md mb-2 text-gray-500">
                            <span class="text-xl">📎</span>
                        </a>
                        <a href="${url}" target="_blank" class="text-sm text-red-600 hover:underline">เอกสารแนบ ${idx + 1}</a>
                    `;
                }
                attachmentsGrid.appendChild(attachmentItem);
            }
        });
        
        attachmentsDiv.appendChild(attachmentsGrid);
        modalContent.appendChild(attachmentsDiv);
    }
    
    detailModal.classList.remove('hidden'); // Show the modal
}

// Update chart data based on filtered results
function updateCharts() {
    if (filteredData.length === 0) return;
    
    // Helper function to count occurrences of values in a specific column
    function countOccurrences(data, columnIndex) {
        const counts = {};
        data.forEach(row => {
            const value = (row[columnIndex] || 'ไม่ระบุ').toString(); // Use 'ไม่ระบุ' for empty values
            counts[value] = (counts[value] || 0) + 1;
        });
        return counts;
    }
    
    // Update each chart with new data
    updateChart(charts.status, countOccurrences(filteredData, 2));        // Column C: สถานะ
    updateChart(charts.department, countOccurrences(filteredData, 3));     // Column D: สังกัด
    updateChart(charts.activityType, countOccurrences(filteredData, 4));   // Column E: ประเภทกิจกรรม
    updateChart(charts.level, countOccurrences(filteredData, 8));          // Column I: ระดับ
    updateChart(charts.format, countOccurrences(filteredData, 11));        // Column L: รูปแบบกิจกรรม
}

// Helper function to update a Chart.js instance
function updateChart(chart, data) {
    const labels = Object.keys(data);
    const values = Object.values(data);
    
    chart.data.labels = labels;
    chart.data.datasets[0].data = values;
    chart.update(); // Redraw the chart
}

// Update summary statistics (total count and total hours)
function updateSummary() {
    totalCount.textContent = filteredData.length;
    
    // Calculate total hours from column U (index 20)
    let totalHoursValue = 0;
    filteredData.forEach(row => {
        if (row.length > 20 && row[20]) {
            const hours = parseFloat(row[20]);
            if (!isNaN(hours)) {
                totalHoursValue += hours;
            }
        }
    });
    
    totalHours.textContent = totalHoursValue.toFixed(0); // Display as whole number
}

// Generate PDF for the currently displayed search results table (this function is not used in the main results section anymore)
function generatePDF() {
    // This function is no longer called by the main results section.
    // It remains here for completeness if needed elsewhere, but its button was removed.
    if (filteredData.length === 0) {
        const messageBox = document.createElement('div');
        messageBox.className = 'fixed inset-0 bg-black bg-opacity-50 flex items-center justify-center z-[999]';
        messageBox.innerHTML = `
            <div class="bg-white p-6 rounded-lg shadow-xl text-center">
                <p class="text-lg font-bold mb-4">ไม่พบข้อมูลที่จะดาวน์โหลด</p>
                <button class="bg-red-500 hover:bg-red-600 text-white font-bold py-2 px-4 rounded-md" onclick="this.parentNode.parentNode.remove()">ตกลง</button>
            </div>
        `;
        document.body.appendChild(messageBox);
        return;
    }

    const { jsPDF } = window.jspdf;
    const pdf = new jsPDF('l', 'mm', 'a4'); 
    pdf.setFont('Sarabun');
    pdf.setFontSize(18);
    pdf.text('ระบบสารสนเทศผลงานครูและนักเรียน', pdf.internal.pageSize.getWidth() / 2, 20, { align: 'center' });
    pdf.setFontSize(14);
    pdf.text('โรงเรียนหางดงรัฐราษฎร์อุปถัมภ์', pdf.internal.pageSize.getWidth() / 2, 30, { align: 'center' });
    
    const today = new Date();
    pdf.setFontSize(10);
    pdf.text(`วันที่พิมพ์: ${today.toLocaleDateString('th-TH')}`, 10, 40);
    
    pdf.setFontSize(12);
    pdf.text(`จำนวนรายการทั้งหมด: ${totalCount.textContent} รายการ`, 10, 50);
    pdf.text(`จำนวนชั่วโมงอบรมรวม: ${totalHours.textContent} ชั่วโมง`, 10, 60);
    
    const tableHeaders = [
        { name: 'ตั้งแต่วันที่', width: 28 },
        { name: 'ถึงวันที่', width: 28 },
        { name: 'ชื่อ', width: 45 },
        { name: 'ประเภทกิจกรรม', width: 45 },
        { name: 'ชื่อกิจกรรม', width: 55 },
        { name: 'สิ่งที่ได้รับ', width: 55 },
        { name: 'ชั่วโมงอบรม', width: 25 }
    ];
    
    const totalTableWidth = tableHeaders.reduce((sum, h) => sum + h.width, 0);
    const startX = (pdf.internal.pageSize.getWidth() - totalTableWidth) / 2;
    let currentY = 70; 
    const cellPadding = 2; 

    const addTableHeader = (yPos) => {
        pdf.setFillColor(255, 200, 200); 
        pdf.rect(startX, yPos, totalTableWidth, 10, 'F'); 
        pdf.setTextColor(0, 0, 0); 
        pdf.setFontSize(10);
        
        let currentX = startX;
        tableHeaders.forEach(header => {
            pdf.text(header.name, currentX + cellPadding, yPos + 7); 
            currentX += header.width;
        });
    };

    addTableHeader(currentY);
    currentY += 10; 

    filteredData.forEach((row, index) => {
        const rowDataForPdf = [
            formatDate(row[9]),  
            formatDate(row[10]), 
            row[1] || '-',       
            row[4] || '-',       
            row[5] || '-',       
            row[6] || '-',       
            (row.length > 20 && row[20] ? row[20].toString() : '-') 
        ];

        let maxRowContentHeight = 0;
        rowDataForPdf.forEach((data, colIndex) => {
            const columnWidth = tableHeaders[colIndex].width;
            const splitText = pdf.splitTextToSize(data.toString(), columnWidth - (cellPadding * 2));
            const textHeight = splitText.length * (pdf.internal.getLineHeight() / pdf.internal.scaleFactor);
            maxRowContentHeight = Math.max(maxRowContentHeight, textHeight);
        });
        const actualRowHeight = maxRowContentHeight + (cellPadding * 2); 

        if (currentY + actualRowHeight > pdf.internal.pageSize.getHeight() - 20) { 
            pdf.addPage();
            currentY = 20; 
            addTableHeader(currentY); 
            currentY += 10; 
        }
        
        if (index % 2 === 0) {
            pdf.setFillColor(245, 245, 245); 
        } else {
            pdf.setFillColor(255, 255, 255); 
        }
        pdf.rect(startX, currentY, totalTableWidth, actualRowHeight, 'F'); 
        
        pdf.setTextColor(0, 0, 0); 
        pdf.setFontSize(9); 
        
        let currentX = startX;
        rowDataForPdf.forEach((data, colIndex) => {
            const columnWidth = tableHeaders[colIndex].width;
            const splitText = pdf.splitTextToSize(data.toString(), columnWidth - (cellPadding * 2));
            pdf.text(splitText, currentX + cellPadding, currentY + cellPadding + (pdf.internal.getLineHeight() / pdf.internal.scaleFactor)); 
            currentX += columnWidth;
        });
        
        currentY += actualRowHeight; 
    });
    
    pdf.setFontSize(10);
    pdf.text('จัดทำโดย กลุ่มงานวิชาการ โรงเรียนหางดงรัฐราษฎร์อุปถัมภ์', pdf.internal.pageSize.getWidth() / 2, pdf.internal.pageSize.getHeight() - 10, { align: 'center' });
    
    pdf.save('ข้อมูลผลงานครูและนักเรียน.pdf');
}

// Function to wait for all images inside an element to load
function waitForImagesToLoad(element) {
    return new Promise(resolve => {
        const images = element.querySelectorAll('img');
        let imagesToLoadPromises = [];

        images.forEach(img => {
            // If the image is already loaded (or has failed but completed its attempt)
            if (img.complete) {
                imagesToLoadPromises.push(Promise.resolve()); // Resolve immediately
                return;
            }

            // Otherwise, create a promise to wait for it
            imagesToLoadPromises.push(new Promise(imgResolve => {
                const onLoad = () => {
                    img.removeEventListener('load', onLoad);
                    img.removeEventListener('error', onError);
                    imgResolve();
                };
                const onError = () => {
                    console.warn('Failed to load image for html2canvas:', img.src);
                    img.removeEventListener('load', onLoad);
                    img.removeEventListener('error', onError);
                    imgResolve(); // Resolve even on error to prevent blocking
                };

                img.addEventListener('load', onLoad);
                img.addEventListener('error', onError);

                // Add a timeout fallback in case events don't fire
                setTimeout(() => {
                    if (!img.complete) { // If still not complete after timeout
                        console.warn('Image load timed out (fallback) for:', img.src);
                        imgResolve();
                    }
                }, 15000); // Increased timeout to 15 seconds
            }));
        });

        if (imagesToLoadPromises.length === 0) {
            resolve(); // No images to load
        } else {
            Promise.all(imagesToLoadPromises).then(resolve);
        }
    });
}

// Generate PDF for detail view (modal content) by converting HTML to image
async function generateDetailPDF() {
    if (!currentDetailData) {
        const messageBox = document.createElement('div');
        messageBox.className = 'fixed inset-0 bg-black bg-opacity-50 flex items-center justify-center z-[999]';
        messageBox.innerHTML = `
            <div class="bg-white p-6 rounded-lg shadow-xl text-center">
                <p class="text-lg font-bold mb-4">ไม่พบข้อมูลที่จะดาวน์โหลด</p>
                <button class="bg-red-500 hover:bg-red-600 text-white font-bold py-2 px-4 rounded-md" onclick="this.parentNode.parentNode.remove()">ตกลง</button>
            </div>
        `;
        document.body.appendChild(messageBox);
        return;
    }

    // Show a loading indicator
    const loadingIndicator = document.createElement('div');
    loadingIndicator.className = 'fixed inset-0 bg-black bg-opacity-75 flex items-center justify-center z-[1000] text-white text-xl flex-col';
    loadingIndicator.innerHTML = `
        <div class="animate-spin rounded-full h-16 w-16 border-t-4 border-b-4 border-red-500 mb-4"></div>
        กำลังสร้าง PDF...
    `;
    document.body.appendChild(loadingIndicator);

    try {
        const { jsPDF } = window.jspdf;
        
        // Temporarily hide modal controls to prevent them from being part of the screenshot
        const modalControls = detailModal.querySelector('.mt-6.flex.justify-center.gap-4');
        if (modalControls) modalControls.style.display = 'none';

        // Wait for all images in the modal content to load
        await waitForImagesToLoad(modalContent);

        // Capture the modal content as a canvas (image)
        const canvas = await html2canvas(modalContent, {
            scale: 2, // Increase scale for better resolution
            useCORS: true, // Important for images loaded from external sources (like Google Drive thumbnails)
            logging: false, // Disable logging for cleaner console
            backgroundColor: '#ffffff' // Ensure white background
        });

        // Restore modal controls display
        if (modalControls) modalControls.style.display = 'flex';

        const imgData = canvas.toDataURL('image/png');
        const pdf = new jsPDF('p', 'mm', 'a4'); // Portrait A4
        const pdfWidth = pdf.internal.pageSize.getWidth(); // A4 width in mm (210)
        const pdfHeight = pdf.internal.pageSize.getHeight(); // A4 height in mm (297)

        const imgRatio = canvas.width / canvas.height;
        
        const topMargin = 25.4; // 1 inch in mm
        const sideMargin = 10; // A reasonable side margin for the image content
        const footerSpace = 20; // Space reserved for the footer at the bottom

        const availableWidth = pdfWidth - (sideMargin * 2);
        const availableHeightForImage = pdfHeight - topMargin - footerSpace; 

        let finalImgWidth = availableWidth;
        let finalImgHeight = finalImgWidth / imgRatio;

        // If the calculated height exceeds the available height, scale it down to fit, maintaining aspect ratio
        if (finalImgHeight > availableHeightForImage) {
            finalImgHeight = availableHeightForImage;
            finalImgWidth = finalImgHeight * imgRatio;
        }

        const xOffset = (pdfWidth - finalImgWidth) / 2; // Center horizontally
        const yOffset = topMargin; // Fixed 1-inch top margin

        pdf.addImage(imgData, 'PNG', xOffset, yOffset, finalImgWidth, finalImgHeight);
        
        // Add footer (ensure it's below the image and not overlapping, and no strange chars)
        pdf.setFont('Sarabun', 'normal'); // Ensure normal font for footer
        pdf.setFontSize(10);
        const footerText = 'จัดทำโดย กลุ่มงานวิชาการ โรงเรียนหางดงรัฐราษฎร์อุปถัมภ์';
        const footerY = pdfHeight - 15; // Position 15mm from the bottom
        pdf.text(footerText, pdfWidth / 2, footerY, { align: 'center' });

        pdf.save(`รายละเอียด-${currentDetailData[1] || 'ข้อมูล'}.pdf`);

    } catch (error) {
        console.error('Error generating detail PDF:', error);
        const messageBox = document.createElement('div');
        messageBox.className = 'fixed inset-0 bg-black bg-opacity-50 flex items-center justify-center z-[999]';
        messageBox.innerHTML = `
            <div class="bg-white p-6 rounded-lg shadow-xl text-center">
                <p class="text-lg font-bold mb-4">เกิดข้อผิดพลาดในการสร้าง PDF: ${error.message}</p>
                <button class="bg-red-500 hover:bg-red-600 text-white font-bold py-2 px-4 rounded-md" onclick="this.parentNode.parentNode.remove()">ตกลง</button>
            </div>
        `;
        document.body.appendChild(messageBox);
    } finally {
        // Remove loading indicator
        if (loadingIndicator.parentNode) {
            loadingIndicator.parentNode.removeChild(loadingIndicator);
        }
    }
}

// Generate CSV file from filtered data
function generateCSV() {
    if (filteredData.length === 0) {
        // Use a custom message box instead of alert()
        const messageBox = document.createElement('div');
        messageBox.className = 'fixed inset-0 bg-black bg-opacity-50 flex items-center justify-center z-[999]';
        messageBox.innerHTML = `
            <div class="bg-white p-6 rounded-lg shadow-xl text-center">
                <p class="text-lg font-bold mb-4">ไม่พบข้อมูลที่จะดาวน์โหลด</p>
                <button class="bg-red-500 hover:bg-red-600 text-white font-bold py-2 px-4 rounded-md" onclick="this.parentNode.parentNode.remove()">ตกลง</button>
            </div>
        `;
        document.body.appendChild(messageBox);
        return;
    }
    
    // Define CSV headers, aligning with the comprehensive data available
    const headers = [
        'ประทับเวลา', 'ชื่อ', 'สถานะ', 'สังกัด', 'ประเภทกิจกรรม', 
        'ชื่อกิจกรรม', 'สิ่งที่ได้รับ', 'รายละเอียด', 'ระดับ', 
        'ตั้งแต่วันที่', 'ถึงวันที่', 'รูปแบบกิจกรรม', 'หน่วยงานที่จัด', 
        'สถานที่', 'เอกสารแนบ 1', 'เอกสารแนบ 2', 'เอกสารแนบ 3', 
        'เอกสารแนบ 4', 'เอกสารแนบ 5', 'เอกสารแนบ 6', 'จำนวนชั่วโมงอบรม'
    ];
    
    // Create CSV content string
    let csvContent = headers.join(',') + '\n';
    
    filteredData.forEach(row => {
        // Map row data to CSV columns, handling potential missing data and escaping quotes
        const csvRow = [
            row[0] || '',                      // ประทับเวลา (Timestamp)
            `"${(row[1] || '').replace(/"/g, '""')}"`, // ชื่อ (Name)
            `"${(row[2] || '').replace(/"/g, '""')}"`, // สถานะ (Status)
            `"${(row[3] || '').replace(/"/g, '""')}"`, // สังกัด (Department)
            `"${(row[4] || '').replace(/"/g, '""')}"`, // ประเภทกิจกรรม (Activity Type)
            `"${(row[5] || '').replace(/"/g, '""')}"`, // ชื่อกิจกรรม (Activity Name)
            `"${(row[6] || '').replace(/"/g, '""')}"`, // สิ่งที่ได้รับ (What was received)
            `"${(row[7] || '').replace(/"/g, '""')}"`, // รายละเอียด (Details)
            `"${(row[8] || '').replace(/"/g, '""')}"`, // ระดับ (Level)
            row[9] || '',                      // ตั้งแต่วันที่ (Start Date)
            row[10] || '',                     // ถึงวันที่ (End Date)
            `"${(row[11] || '').replace(/"/g, '""')}"`, // รูปแบบกิจกรรม (Activity Format)
            `"${(row[12] || '').replace(/"/g, '""')}"`, // หน่วยงานที่จัด (Organizing Unit)
            `"${(row[13] || '').replace(/"/g, '""')}"`, // สถานที่ (Location)
            `"${(row[14] || '').replace(/"/g, '""')}"`, // เอกสารแนบ 1 (Attachment 1)
            `"${(row[15] || '').replace(/"/g, '""')}"`, // เอกสารแนบ 2 (Attachment 2)
            `"${(row[16] || '').replace(/"/g, '""')}"`, // เอกสารแนบ 3 (Attachment 3)
            `"${(row[17] || '').replace(/"/g, '""')}"`, // เอกสารแนบ 4 (Attachment 4)
            `"${(row[18] || '').replace(/"/g, '""')}"`, // เอกสารแนบ 5 (Attachment 5)
            `"${(row[19] || '').replace(/"/g, '""')}"`, // เอกสารแนบ 6 (Attachment 6)
            row.length > 20 ? row[20] || '' : '' // จำนวนชั่วโมงอบรม (Training Hours) - check if column exists
        ];
        
        csvContent += csvRow.join(',') + '\n';
    });
    
    // Create a Blob object with UTF-8 BOM for proper Excel display
    const blob = new Blob(["\uFEFF" + csvContent], { type: 'text/csv;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a'); // Create a temporary anchor element
    link.setAttribute('href', url);
    link.setAttribute('download', 'ข้อมูลผลงานครูและนักเรียน.csv'); // Set download file name
    link.style.visibility = 'hidden'; // Hide the link
    document.body.appendChild(link); // Append to body
    link.click(); // Programmatically click the link to trigger download
    document.body.removeChild(link); // Remove the link after download
}

// Event Listeners for various interactions
document.addEventListener('DOMContentLoaded', async () => {
    // Initialize all charts on page load
    initializeCharts();
    
    // Fetch initial data from Google Apps Script
    await fetchData();
    
    // Event listener for "ข้อมูลสารสนเทศ" button
    showInfoBtn.addEventListener('click', () => {
        mainButtons.classList.add('hidden'); // Hide main buttons
        infoSystem.classList.remove('hidden'); // Show info system section
    });
    
    // Event listener for "กลับไปหน้าหลัก" button
    backBtn.addEventListener('click', () => {
        infoSystem.classList.add('hidden'); // Hide info system section
        resultsSection.classList.add('hidden'); // Hide results section
        mainButtons.classList.remove('hidden'); // Show main buttons
    });
    
    // Event listener for "ค้นหา" button
    searchBtn.addEventListener('click', async () => {
        // Re-fetch data if it's empty (e.g., first load or error)
        if (allData.length === 0) {
            await fetchData();
        }
        
        filterData();      // Filter data based on current criteria and sort it
        displayResults();  // Display filtered data in the table
        updateSummary();   // Update summary statistics
        updateCharts();    // Update charts
        resultsSection.classList.remove('hidden'); // Show results section
    });
    
    // Event listener for "ล้างข้อมูลการค้นหา" button
    clearBtn.addEventListener('click', () => {
        // Clear all search input fields
        document.getElementById('activityType').value = '';
        document.getElementById('name').value = '';
        document.getElementById('department').value = '';
        document.getElementById('level').value = '';
        document.getElementById('activityFormat').value = '';
        document.getElementById('startDate').value = '';
        document.getElementById('endDate').value = '';
    });
    
    // Event listeners for closing the detail modal
    closeModal.addEventListener('click', () => {
        detailModal.classList.add('hidden');
    });
    
    closeModalBtn.addEventListener('click', () => {
        detailModal.classList.add('hidden');
    });
    
    // Event listeners for download buttons
    downloadCsvBtn.addEventListener('click', generateCSV);
    downloadDetailPdfBtn.addEventListener('click', generateDetailPDF);
    
    // Close modal when clicking outside of its content
    detailModal.addEventListener('click', (e) => {
        if (e.target === detailModal) {
            detailModal.classList.add('hidden');
        }
    });
});
