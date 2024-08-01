const XLSX = require('xlsx');
const XLSXStyle = require('xlsx-style');

// Sample data
const data = [
     { Name: 'John Doe', Email: 'john@example.com', Phone: '555-555-5555' },
     { Name: 'Jane Smith', Email: 'jane@example.com', Phone: '555-555-5556' },
];

// Convert JSON data to worksheet using xlsx
const worksheet = XLSX.utils.json_to_sheet(data);

// Set header style using xlsx-style
const headerStyle = { fill: { fgColor: { rgb: 'FFFF0000' } } }; // Red background color

// Apply style to header cells
worksheet['A1'].s = headerStyle;
worksheet['B1'].s = headerStyle;
worksheet['C1'].s = headerStyle;

// Save workbook using xlsx-style
const workbook = XLSX.utils.book_new();
XLSX.utils.book_append_sheet(workbook, worksheet, 'Sheet1');
XLSXStyle.writeFile(workbook, 'output.xlsx');
