document.getElementById('excelFile').addEventListener('change', handleFile);
let tagData = [];

function handleFile(event) {
  const file = event.target.files[0];
  const reader = new FileReader();
  reader.onload = (e) => {
    const data = new Uint8Array(e.target.result);
    const workbook = XLSX.read(data, { type: 'array' });
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    tagData = XLSX.utils.sheet_to_json(sheet);
    displayTags(tagData);
  };
  reader.readAsArrayBuffer(file);
}

function displayTags(data) {
  const container = document.getElementById('tagContainer');
  container.innerHTML = '';

  data.forEach((item, index) => {
    const tag = document.createElement('div');
    tag.className = 'tag';
    const barcodeId = `barcode-${index}`;

    tag.innerHTML = `
      <div class="top-line"></div>
      <center><strong>MRP (Incl. of all taxes)</strong><br>
      <strong style="font-size: 14px;">₹ ${parseFloat(item['Price']).toLocaleString('en-IN', { minimumFractionDigits: 2 })}</strong><br></center>
      <div class="middle-line"></div>
      <div>Size :&emsp;<strong  style="font-size: 14px;">${item['Size']}</strong></div>
      <div class="middle-line"></div>
      <div class="product-info">
      <div class="product-left">
      <strong>${item['Product Name']}</strong> <br> <strong>${item['Code']}<br></strong></div>
      <div class="product-right">Art No: ${item['Art No']}<br>
      Net Qty: 1</div>
      

      <div class="barcode-container">
        <svg id="${barcodeId}"></svg>
      </div>

<div class="middle-line"></div>
      <center><div style="font-size: 8px;">
        <strong>Marketed by:</strong><br>
        JC Fashion for Women<br>
        #34-2-2, Ground Floor, Vani Mahal Center,<br>
        Rythu Bazar Road, Mandapeta - 533308, AP<br>
        Made in India
        <div class="middle-line"></div>
        For Complaints/ Feedback, pls write to our customer care executive<br>
        Email: jcfashionforwomen@gmail.com<br>
        Contact: +91 9391718898
        <center><div class="recycle">♻</div><center>
      </div></center>
    `;

    container.appendChild(tag);

    JsBarcode(`#${barcodeId}`, item['Barcode'] || item['Code'] || '0000000000', {
  format: 'CODE128',
  lineColor: "#000",
  width: 1,
  height: 30,
  fontSize: 10,
  displayValue: true
});


    // Add page break after every 16 tags (4x4 grid)
    if ((index + 1) % 12 === 0) {
      const breakDiv = document.createElement('div');
      breakDiv.className = 'page-break';
      container.appendChild(breakDiv);
    }
  });
}

function generatePDF() {
  const element = document.getElementById('tagContainer');
  const opt = {
    margin: 0,
    filename: 'Garment_Tags.pdf',
    image: { type: 'jpeg', quality: 0.98 },
    html2canvas: { scale: 2 },
    jsPDF: { unit: 'mm', format: 'a4', orientation: 'portrait' },
    pagebreak: { mode: ['avoid-all', 'css', 'legacy'] }
  };
  html2pdf().from(element).set(opt).save();
}

