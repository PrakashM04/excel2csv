let fileData = null;
    let fileName = null;

    function handleFileChange(event) {
      const file = event.target.files[0];
      if (file) {
        readFile(file);
        fileName = file.name;
        updateFileName();
      } else {
        fileData = null;
        fileName = null;
        updateFileName();
      }
    }

    function handleDrop(event) {
      event.preventDefault();
      const file = event.dataTransfer.files[0];
      if (file) {
        readFile(file);
        fileName = file.name;
        updateFileName();
      } else {
        fileData = null;
        fileName = null;
        updateFileName();
      }
    }

    function openFileBrowser() {
      document.getElementById('file-input').click();
    }

    function readFile(file) {
      const reader = new FileReader();
      reader.onload = function (e) {
        fileData = e.target.result;
        document.getElementById('save-btn').style.display = 'block';
      };
      reader.readAsBinaryString(file);
    }

    function saveSheets() {
      const workbook = XLSX.read(fileData, { type: 'binary' });
      const sheetNames = workbook.SheetNames;

      const outputDiv = document.getElementById('output');
      outputDiv.innerHTML = '';

      sheetNames.forEach(function (sheetName) {
        const sheetData = XLSX.utils.sheet_to_csv(workbook.Sheets[sheetName]);
        const blob = new Blob([sheetData], { type: 'text/csv;charset=utf-8' });
        // const csvFileName = sheetName + '.csv';

            // Modify the csvFileName to remove "_Data" suffix
    let csvFileName = sheetName.replace('_Data', '') + '.csv';

        const downloadLink = document.createElement('a');
        downloadLink.href = URL.createObjectURL(blob);
        downloadLink.download = csvFileName;
        downloadLink.innerText = csvFileName;

        const p = document.createElement('p');
        p.appendChild(downloadLink);
        outputDiv.appendChild(p);
      });

      const reloadBtn = document.getElementById('reload-btn');
  reloadBtn.style.display = 'block';

  // Disable the upload area after converting sheets
  const dropArea = document.getElementById('drop-area');
  dropArea.classList.add('muted');

    }
    
    
function reloadPage() {
  // Reload the current page
  window.location.reload();
}


    function updateFileName() {
      const fileNameDiv = document.getElementById('file-name');
      const saveBtn = document.getElementById('save-btn');

      if (fileName) {
        fileNameDiv.innerText = 'File Name: ' + fileName;
        fileNameDiv.style.display = 'block';
        saveBtn.disabled = false;
        saveBtn.style.display = 'block';
      } else {
        fileNameDiv.innerText = '';
        fileNameDiv.style.display = 'none';
        saveBtn.disabled = true;
        saveBtn.style.display = 'none';
      }
    }