import CRUD from './sheet/CRUD';

const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

const confirmationSave = () => {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    'สร้างเอกสาร',
    'คุณต้องการจะสร้างเอกสารใหม่ใช่หรือไม่',
    ui.ButtonSet.YES_NO
  );

  if (response === ui.Button.YES) {
    CRUD.saveNewDocument();
  }
};

const onEdit = e => {
  const { range } = e;
  const workingPaperSheet = spreadsheet.getSheetByName('สร้างเอกสาร');
  if (range.getSheet().getSheetId() === workingPaperSheet.getSheetId()) {
    if (range.getA1Notation() === 'B1') {
      CRUD.updateNewDocumentId();
    }
  }
};

const openPrintSheet = () => {
  const docId = spreadsheet.getActiveRange().getDisplayValue();
  const docType = CRUD.getDocumentType(docId);
  const docSheet = spreadsheet.getSheetByName(docType);
  docSheet.getRange('E2').setValue(docId);
  spreadsheet.setActiveSheet(docSheet);
};

global.confirmationSave = confirmationSave;
global.saveNewDocument = CRUD.saveNewDocument;
global.updateNewDocumentId = CRUD.updateNewDocumentId;
global.prefillDO = CRUD.prefillDeliveryOrder;
global.prefillIN = CRUD.prefillInvoice;
global.prefillRT = CRUD.prefillReciept;
global.prefillTV = CRUD.prefillTaxInvoice;

global.openPrintSheet = openPrintSheet;
global.onEdit = onEdit;
