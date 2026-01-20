/**
 * THE MULTI-TAB DOCUMENT ENGINE v5.7
 * Features: Automatic 2-Decimal Formatting & Smart Currency Detection
 */

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('🚀 DOCUMENT ENGINE')
      .addItem('1. Generate Customer Quote', 'runForQuote')
      .addItem('2. Generate Contractor Offer', 'runForContractor')
      .addSeparator()
      .addItem('Initial Setup Check', 'checkConnections')
      .addToUi();
}

function runForQuote() { runEngine("quote"); }
function runForContractor() { runEngine("contractor"); }

function runEngine(mode) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  
  const response = ui.prompt('Generate Document', 'Enter the exact Site Name:', ui.ButtonSet.OK_CANCEL);
  if (response.getSelectedButton() !== ui.Button.OK) return;
  const searchTerm = response.getResponseText().trim();

  const settingsSheet = ss.getSheetByName('⚙️ Settings');
  const settingsValues = settingsSheet.getDataRange().getValues();
  let templateId, folderId, bizName, adminEmail;

  settingsValues.forEach(row => {
    let label = row[0].toString().toLowerCase().trim();
    let val = row[1] ? row[1].toString().trim() : "";
    
    if (mode === "quote" && label.includes("costing template id")) templateId = val;
    if (mode === "contractor" && label.includes("contractor offer template id")) templateId = val;
    if (label.includes("folder id")) folderId = val;
    if (label.includes("business name")) bizName = val;
    if (label.includes("admin email")) adminEmail = val;
  });

  if (!templateId || !folderId) {
    ui.alert('❌ Setup Error: Check Template/Folder IDs in Settings.');
    return;
  }

  let masterData = {}; 
  let boqRows = [];    
  let gwRows = [];    

  const tabs = [
    { name: 'Quote gen', siteCol: 'Site' },
    { name: 'Bill of quantities', siteCol: 'Project Name' },
    { name: 'Ground works B&Q', siteCol: 'Site Name' }
  ];

  tabs.forEach(tabInfo => {
    let sheet = ss.getSheetByName(tabInfo.name);
    if (!sheet) return;
    let data = sheet.getDataRange().getValues();
    let headers = data[0];
    let siteColIdx = headers.indexOf(tabInfo.siteCol);

    if (siteColIdx === -1) return;

    for (let r = 1; r < data.length; r++) {
      if (data[r][siteColIdx].toString().toLowerCase() === searchTerm.toLowerCase()) {
        headers.forEach((h, c) => { 
          if (h) masterData[h] = data[r][c]; 
        });

        if (tabInfo.name === 'Bill of quantities') {
          let item = data[r][headers.indexOf('Install items')] || "Item";
          let qty = data[r][headers.indexOf('quantity')] || "0";
          let total = data[r][headers.indexOf('Item total cost')] || 0;
          // Format BOQ row amount to 2 decimal places
          boqRows.push(`• ${item} (Qty: ${qty}) - £${Number(total).toFixed(2)}`);
        }
        if (tabInfo.name === 'Ground works B&Q') {
          let gwItem = data[r][headers.indexOf('Ground works item')] || "Work";
          let gwQty = data[r][headers.indexOf('Ground works quantity')] || "0";
          gwRows.push(`• ${gwItem} (Qty: ${gwQty})`);
        }
      }
    }
  });

  masterData['Full_BOQ_Table'] = boqRows.join('\n');
  masterData['Full_Groundworks_Table'] = gwRows.join('\n');

  try {
    const folder = DriveApp.getFolderById(folderId);
    const docTitle = mode === "quote" ? `Quote - ${searchTerm}` : `Contractor Offer - ${searchTerm}`;
    const copy = DriveApp.getFileById(templateId).makeCopy(docTitle, folder);
    const doc = DocumentApp.openById(copy.getId());
    const body = doc.getBody();

    for (let key in masterData) {
      let val = masterData[key];
      let formattedVal = "";

      // Logic to handle numbers and decimal places
      if (typeof val === 'number' && !key.includes('Table')) {
        formattedVal = val.toFixed(2);
        
        // Add currency symbol if header relates to money
        let kLower = key.toLowerCase();
        if (kLower.includes('total') || kLower.includes('cost') || kLower.includes('price')) {
          formattedVal = "£" + formattedVal;
        }
      } else {
        formattedVal = (val || "").toString();
      }

      body.replaceText(`{{${key}}}`, formattedVal);
    }

    body.replaceText('{{Date}}', Utilities.formatDate(new Date(), "GMT", "dd/MM/yyyy"));
    doc.saveAndClose();

    const pdfBlob = copy.getAs(MimeType.PDF);
    const finalPdf = folder.createFile(pdfBlob);
    finalPdf.setName(`${docTitle}.pdf`);

    if (adminEmail) {
      MailApp.sendEmail({
        to: adminEmail,
        subject: `New Document Generated: ${docTitle}`,
        body: `Hello,\n\nPlease find the attached ${mode} for ${searchTerm}.\n\nBest regards,\n${bizName} Engine`,
        attachments: [pdfBlob]
      });
    }

    copy.setTrashed(true);
    ss.toast(`Successfully created and emailed ${mode} for ${searchTerm}!`);

  } catch (err) {
    ui.alert('Error: ' + err.message);
  }
}
