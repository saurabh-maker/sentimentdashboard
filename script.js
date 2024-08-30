document.getElementById('fileUpload').addEventListener('change', handleFileUpload);

let jsonData = [];
let filteredData = [];

function handleFileUpload(event) {
    const file = event.target.files[0];
    const reader = new FileReader();

    reader.onload = function(e) {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: 'array' });
        const worksheet = workbook.Sheets[workbook.SheetNames[0]];
        jsonData = XLSX.utils.sheet_to_json(worksheet, { raw: false });
        filteredData = jsonData; // Initialize filtered data with all data

        console.log('Loaded JSON Data:', jsonData); // Log to verify data

        // Populate region filter
        populateRegionFilter(jsonData);

        // Update insights and display data
        updateInsights(filteredData);
        displayData(filteredData);
        generateRegionChart(filteredData);
        generateLocationChart(filteredData);
    };

    reader.readAsArrayBuffer(file);
}

function updateInsights(data) {
    const totalReviews = data.length;
    let totalPros = 0;
    let totalCons = 0;
    let totalResponses = 0;

    data.forEach(row => {
        if (row.Pros) totalPros++;
        if (row.Cons) totalCons++;
        if (row['Ciena Response']) totalResponses++;
    });

    document.getElementById('totalReviews').textContent = totalReviews;
    document.getElementById('totalPros').textContent = totalPros;
    document.getElementById('totalCons').textContent = totalCons;
    document.getElementById('totalResponses').textContent = totalResponses;
}

function displayData(data) {
    const tbody = document.querySelector('#dataTable tbody');
    tbody.innerHTML = ''; // Clear existing data

    // Extract headers dynamically
    if (data.length > 0) {
        const headers = Object.keys(data[0]);
        const thead = document.querySelector('#dataTable thead');
        
        // Generate table headers
        thead.innerHTML = '';
        const headerRow = document.createElement('tr');
        headers.forEach(header => {
            const th = document.createElement('th');
            th.textContent = header;
            headerRow.appendChild(th);
        });
        thead.appendChild(headerRow);

        // Populate table rows
        data.forEach(row => {
            const tr = document.createElement('tr');
            headers.forEach(header => {
                const td = document.createElement('td');
                // Format dates properly
                if (header.toLowerCase().includes('date') && typeof row[header] === 'number') {
                    td.textContent = formatExcelDate(row[header]);
                } else {
                    td.textContent = row[header] !== undefined ? row[header] : '';
                }
                tr.appendChild(td);
            });
            tbody.appendChild(tr);
        });
    } else {
        tbody.innerHTML = '<tr><td colspan="100%">No data available</td></tr>';
    }
}

function formatExcelDate(serial) {
    const date = new Date(Math.round((serial - 25569) * 86400 * 1000));
    return date.toLocaleDateString();
}

function generateRegionChart(data) {
    const ctx = document.getElementById('regionChart').getContext('2d');

    // Aggregate data by region
    const regionCounts = data.reduce((acc, row) => {
        const region = row.Region || 'Unknown';
        acc[region] = (acc[region] || 0) + 1;
        return acc;
    }, {});

    const labels = Object.keys(regionCounts);
    const values = Object.values(regionCounts);

    new Chart(ctx, {
        type: 'pie',
        data: {
            labels: labels,
            datasets: [{
                label: 'Number of Reviews by Region',
                data: values,
                backgroundColor: ['#FF6384', '#36A2EB', '#FFCE56', '#4BC0C0', '#F7464A'],
                borderColor: '#fff',
                borderWidth: 1
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                title: {
                    display: true,
                    text: 'Reviews by Region',
                    font: {
                        size: 16
                    }
                }
            }
        }
    });
}

function generateLocationChart(data) {
    const ctx = document.getElementById('locationChart').getContext('2d');

    // Aggregate data by location
    const locationCounts = data.reduce((acc, row) => {
        const location = row.location || 'Unknown';
        acc[location] = (acc[location] || 0) + 1;
        return acc;
    }, {});

    const labels = Object.keys(locationCounts);
    const values = Object.values(locationCounts);

    new Chart(ctx, {
        type: 'bar',
        data: {
            labels: labels,
            datasets: [{
                label: 'Number of Reviews by Location',
                data: values,
                backgroundColor: '#36A2EB',
                borderColor: '#0078D7',
                borderWidth: 1
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                title: {
                    display: true,
                    text: 'Reviews by Location',
                    font: {
                        size: 16
                    }
                }
            },
            scales: {
                x: {
                    beginAtZero: true
                },
                y: {
                    beginAtZero: true
                }
            }
        }
    });
}

function populateRegionFilter(data) {
    const regionFilter = document.getElementById('regionFilter');
    const regions = Array.from(new Set(data.map(row => row.Region).filter(Boolean)));

    regions.forEach(region => {
        const option = document.createElement('option');
        option.value = region;
        option.textContent = region;
        regionFilter.appendChild(option);
    });
}

function filterData() {
    const selectedRegion = document.getElementById('regionFilter').value;
    filteredData = selectedRegion ? jsonData.filter(row => row.Region === selectedRegion) : jsonData;

    updateInsights(filteredData);
    displayData(filteredData);
    generateRegionChart(filteredData);
    generateLocationChart(filteredData);
}

function clearFilters() {
    document.getElementById('regionFilter').value = '';
    filterData();
}
