import dayjs from 'dayjs';
import NameRage from './NameRanges';

const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

const zeroPadding = (number, length) => {
  return (Array(length).join('0') + number).slice(-length);
};

const updateNewDocumentId = () => {
  const type = spreadsheet
    .getRangeByName(NameRage.DOC_TYPE)
    .getDisplayValue()
    .replace('สร้าง', '');

  const typeConfig = spreadsheet
    .getRangeByName(NameRage.TYPE_CONFIG)
    .getValues()
    .map(val => {
      return {
        title: val[0],
        code: val[1]
      };
    });

  const allHeaders = spreadsheet
    .getSheetByName('ฐานข้อมูลเอกสาร')
    .getRange('A2:C')
    .getValues()
    .map(val => {
      return {
        id: val[0],
        type: val[1],
        createdDate: val[2]
      };
    })
    .filter(header => {
      const creadtedAt = dayjs(header.createdDate);
      return header.type === type && dayjs().isSame(creadtedAt, 'month');
    });

  const amount = allHeaders.length;
  const docCode = typeConfig.filter(config => {
    return config.title === type;
  })[0].code;
  const year = dayjs().format('YY');
  const month = dayjs().format('MM');
  const nextId = zeroPadding(amount + 1, 4);
  const newDocumentId = `${docCode}${year}${month}${nextId}`;
  spreadsheet.getRangeByName(NameRage.DOC_NEW_ID).setValue(newDocumentId);
};

const getDocumentType = docId => {
  const allHeaders = spreadsheet
    .getSheetByName('ฐานข้อมูลเอกสาร')
    .getRange('A2:C')
    .getValues()
    .map(val => {
      return {
        id: val[0],
        type: val[1]
      };
    })
    .filter(header => {
      return header.id === docId;
    });
  return allHeaders[0].type;
};

const getDocument = docId => {
  const allHeaders = spreadsheet
    .getSheetByName('ฐานข้อมูลเอกสาร')
    .getRange('A2:O')
    .getValues()
    .map(val => {
      return {
        id: val[0],
        type: val[1],
        createdDate: val[2],
        availableUntilDate: val[3],
        refId: val[4],
        customer: {
          name: val[5],
          address1: val[6],
          address2: val[7],
          taxId: val[8],
          phone: val[9],
          email: val[10]
        },
        note: val[11],
        keywords: val[12],
        discount: val[13],
        isVatIncluded: val[14]
      };
    })
    .filter(doc => {
      return doc.id === docId;
    });
  return allHeaders[0];
};

const getNewDocumentHeader = () => {
  const type = spreadsheet
    .getRangeByName(NameRage.DOC_TYPE)
    .getDisplayValue()
    .replace('สร้าง', '');

  const customer = {
    name: spreadsheet.getRangeByName(NameRage.DOC_CUSTOMER_NAME).getDisplayValue(),
    address1: spreadsheet.getRangeByName(NameRage.DOC_CUSTOMER_ADDRESS1).getDisplayValue(),
    address2: spreadsheet.getRangeByName(NameRage.DOC_CUSTOMER_ADDRESS2).getDisplayValue(),
    taxId: spreadsheet.getRangeByName(NameRage.DOC_CUSTOMER_TAX_ID).getDisplayValue(),
    phone: spreadsheet.getRangeByName(NameRage.DOC_CUSTOMER_PHONE).getDisplayValue(),
    email: spreadsheet.getRangeByName(NameRage.DOC_CUSTOMER_EMAIL).getDisplayValue()
  };
  const documentHeader = {
    id: spreadsheet.getRangeByName(NameRage.DOC_NEW_ID).getDisplayValue(),
    type,
    createdDate: spreadsheet.getRangeByName(NameRage.DOC_CREATED_DATE).getValue(),
    availableUntilDate:
      type === 'ใบเสนอราคา'
        ? spreadsheet.getRangeByName(NameRage.DOC_AVAILABLE_UNTIL_DATE).getValue()
        : '',
    refId: spreadsheet.getRangeByName(NameRage.DOC_REF_ID).getDisplayValue(),
    discount: spreadsheet.getRangeByName(NameRage.DOC_DISCOUNT).getValue(),
    isVatIncluded: spreadsheet.getRangeByName(NameRage.DOC_IS_VAT_INCLUDED).getValue(),
    note: spreadsheet.getRangeByName(NameRage.DOC_NOTE).getDisplayValue(),
    keywords: spreadsheet.getRangeByName(NameRage.DOC_KEYWORDS).getDisplayValue(),
    customer
  };

  return documentHeader;
};

const getNewDocumentItems = () => {
  const values = spreadsheet.getRangeByName(NameRage.DOC_ITEMS).getValues();
  return values
    .filter(value => value[3] > 0)
    .map(value => {
      return {
        title: value[1],
        amount: value[2],
        pricePerUnit: value[3]
      };
    });
};

const getDocumentItems = docId => {
  const items = spreadsheet
    .getSheetByName('ฐานข้อมูลรายการ')
    .getRange('A2:D')
    .getValues()
    .map(val => {
      return {
        docId: val[0],
        title: val[1],
        amount: val[2],
        pricePerUnit: val[3]
      };
    })
    .filter(item => {
      return item.docId === docId;
    });
  return items;
};

const saveNewHeader = () => {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('ฐานข้อมูลเอกสาร');
  const {
    id,
    type,
    createdDate,
    availableUntilDate,
    refId,
    discount,
    isVatIncluded,
    note,
    keywords,
    customer
  } = getNewDocumentHeader();
  const rowData = [
    id,
    type,
    createdDate,
    availableUntilDate,
    refId,
    customer.name,
    customer.address1,
    customer.address2,
    `'${customer.taxId}`,
    `'${customer.phone}`,
    customer.email,
    note,
    keywords,
    discount,
    isVatIncluded
  ];
  sheet.appendRow(rowData);
};

const saveNewItems = () => {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('ฐานข้อมูลรายการ');
  const newDocumentId = getNewDocumentHeader().id;
  const newItems = getNewDocumentItems();
  newItems.forEach(item => {
    sheet.appendRow([newDocumentId, item.title, item.amount, item.pricePerUnit]);
  });
};

const preSearch = () => {
  const newKeyword = spreadsheet.getRangeByName(NameRage.DOC_KEYWORDS).getDisplayValue();
  const sheet = spreadsheet.getSheetByName('ค้นหาเอกสาร');
  const textfield = sheet.getRange('B1');
  textfield.setValue(newKeyword);
  spreadsheet.setActiveSheet(sheet);
};

const saveNewDocument = () => {
  try {
    saveNewHeader();
    saveNewItems();
    updateNewDocumentId();
    preSearch();
  } catch (error) {
    Logger.log(JSON.stringify(error));
  }
};

const setNewCustomer = customer => {
  spreadsheet.getRangeByName(NameRage.DOC_CUSTOMER_NAME).setValue(customer.name);
  spreadsheet.getRangeByName(NameRage.DOC_CUSTOMER_ADDRESS1).setValue(customer.address1);
  spreadsheet.getRangeByName(NameRage.DOC_CUSTOMER_ADDRESS2).setValue(customer.address2);
  spreadsheet.getRangeByName(NameRage.DOC_CUSTOMER_TAX_ID).setValue(`'${customer.taxId}`);
  spreadsheet.getRangeByName(NameRage.DOC_CUSTOMER_PHONE).setValue(`'${customer.phone}`);
  spreadsheet.getRangeByName(NameRage.DOC_CUSTOMER_EMAIL).setValue(customer.email);
};

const setNewHeader = (refDoc, type) => {
  spreadsheet.getRangeByName(NameRage.DOC_TYPE).setValue(type);
  spreadsheet.getRangeByName(NameRage.DOC_CREATED_DATE).setValue(dayjs().format('MM/DD/YYYY'));
  spreadsheet.getRangeByName(NameRage.DOC_REF_ID).setValue(refDoc.id);
  spreadsheet.getRangeByName(NameRage.DOC_DISCOUNT).setValue(refDoc.discount);
  spreadsheet.getRangeByName(NameRage.DOC_IS_VAT_INCLUDED).setValue(refDoc.isVatIncluded);
  spreadsheet.getRangeByName(NameRage.DOC_NOTE).setValue(refDoc.note);
  spreadsheet.getRangeByName(NameRage.DOC_KEYWORDS).setValue(refDoc.keywords);
};

const setNewItems = refDoc => {
  const items = getDocumentItems(refDoc.id);
  const rowDatas = items.map(item => {
    return [item.title, item.amount, item.pricePerUnit];
  });
  const startRow = 17;
  const range = spreadsheet
    .getSheetByName('สร้างเอกสาร')
    .getRange(`B${startRow}:D${startRow + items.length - 1}`);
  spreadsheet
    .getSheetByName('สร้างเอกสาร')
    .getRange('B17:D36')
    .clearContent();
  range.setValues(rowDatas);
};

const prefillNewDocument = type => {
  const refId = spreadsheet.getActiveRange().getDisplayValue();
  const refDoc = getDocument(refId);
  setNewHeader(refDoc, type);
  setNewCustomer(refDoc.customer);
  setNewItems(refDoc);
  updateNewDocumentId();
  const newDocSheet = spreadsheet.getSheetByName('สร้างเอกสาร');
  spreadsheet.setActiveSheet(newDocSheet);
};

const prefillDeliveryOrder = () => {
  prefillNewDocument('สร้างใบส่งของ');
};

const prefillInvoice = () => {
  prefillNewDocument('สร้างใบแจ้งหนี้');
};

const prefillReciept = () => {
  prefillNewDocument('สร้างใบเสร็จรับเงิน');
};

const prefillTaxInvoice = () => {
  prefillNewDocument('สร้างใบกำกับภาษี');
};

export default {
  getDocumentType,
  getNewDocumentItems,
  getNewDocumentHeader,
  updateNewDocumentId,
  saveNewItems,
  saveNewHeader,
  saveNewDocument,
  prefillDeliveryOrder,
  prefillInvoice,
  prefillReciept,
  prefillTaxInvoice
};
