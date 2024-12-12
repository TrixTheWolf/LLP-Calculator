// Author: Ben Renner
// Title: LLP Calculator

// Function to process the Excel file and extract data
async function processExcelFile(file) {
  const data = await file.arrayBuffer();
  const workbook = XLSX.read(data, { type: 'array' });
  const sheet = workbook.Sheets[workbook.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(sheet, { header: 1 });
  return rows;
}

// Function to find and calculate the results
function calculateResults(partNb, cycleRm, rows) {
  let matchedRow = null;

  for (let i = 1; i < rows.length; i++) { // Start from row index 1 to skip headers
    if (rows[i][0]?.toUpperCase() === partNb) {
      matchedRow = rows[i];
      break;
    }
  }

  if (matchedRow) {
    const partNm = matchedRow[1];
    const clp = parseFloat(matchedRow[2]);
    const cycleLm = parseInt(matchedRow[3], 10);

    if (isNaN(clp) || isNaN(cycleLm)) {
      return { error: 'Data error in the spreadsheet.' };
    }

    const proRt = (clp / cycleLm) * cycleRm;

    const results = {
      partName: partNm,
      clp: clp,
      proRates: []
    };

    for (let percent = 100; percent > 5; percent -= 5) {
      results.proRates.push({ percent, value: (proRt * (percent / 100)).toFixed(2) });
    }

    return results;
  } else {
    return { error: 'The inputted Part Number does not match any parts in the current database.' };
  }
}

// Set up the UI interactions
document.getElementById('calculator-form').addEventListener('submit', async (event) => {
  event.preventDefault();

  const fileInput = document.getElementById('file-input');
  const partNumber = document.getElementById('part-number').value.toUpperCase();
  const cyclesRemaining = parseInt(document.getElementById('cycles-remaining').value, 10);

  if (!fileInput.files.length || isNaN(cyclesRemaining)) {
    alert('Please provide a valid file and input.');
    return;
  }

  const file = fileInput.files[0];
  const rows = await processExcelFile(file);

  const results = calculateResults(partNumber, cyclesRemaining, rows);
  const resultsDiv = document.getElementById('results');
  resultsDiv.innerHTML = ''; // Clear previous results

  if (results.error) {
    resultsDiv.innerHTML = `<p>${results.error}</p>`;
  } else {
    resultsDiv.innerHTML = `<h2>Results</h2>
      <p>Description: ${results.partName}</p>
      <p>CLP: $${results.clp}</p>`;

    results.proRates.forEach(rate => {
      resultsDiv.innerHTML += `<p>${rate.percent}% Pro Rate: $${rate.value}</p>`;
    });
  }
});
