import listFilesExcel from "./excel/list-files.js";
import listFoldersExcel from "./excel/list-folders.js";
import listWorksheetsExcel from "./excel/list-worksheets.js";
import listFolders from "./mail/list-folders.js";
import listMailboxes from "./mail/list-mailboxes.js";
import listSendAsAddresses from "./mail/list-send-as-addresses.js";
import listExcelColumns from "./excel/list-excel-columns.js";

export default [
    listFolders,
    listMailboxes,
    listSendAsAddresses,
    listFilesExcel,
    listFoldersExcel,
    listWorksheetsExcel,
    listExcelColumns
];