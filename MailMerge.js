var MAIL_LIST_FILE_NAME = "mail-list";
var TEMPLATE_FILE_NAME = "templatedoc";
var SHEET_FILE_NAME = "Automatizacion";
var SENDER_MAIL = "gustavo.camargo@pwc.com";
var SENDER_NAME = "Gustavo Camargo";
var SENDER_TITLE = "PM - PSM | Contractor";
var STAKEHOLDERS_NAMES = "Gustavo Camargo";
var RESULT_TEMPLATE = "mailmergeresulttemplate";
var OUTPUT_FOLDER = "MailMergeOutput";
var subject = "AÇÃO NECESSÁRIA - DESATIVAÇÃO DE BASES NOTES (PDF Migration)";
var SETTINGS_FILE_NAME = "MailMergeSettings.json";
var EXCEL_FILE_NAME = "InputFile.xlsx";
var PROCESS_FOLDER = "MailMergeProcess"
var tt = "";
var errors = "";
var mails = [];
var names = [];
var notFound = [];
var notFoundString = "";
var notFoundEntries = "";
var invalidMailEntries = "";
var lines = 0;
var totalLines = 0;
var totalMails = 0;
var sentMails = 0;
var notFoundMails = 0;
var invalidMails = 0;
var overAllResult = "";
var emailsList = "";
var notFound;
var invalidMails;
var title = "";
var server = "";
var fileName = "";
var excelFileId = "";
var inputFolder;
var p;
var notFoundFileName = "";
var ssNotFound;
var notFoundRow = [];

var invalidMailsFileName = "";
var ssInvalidMails;
var invalidMailRow = [];
var imRow = 0;
var inRow = 0;


var notFoundFileObject;
var invalidMailsFileObject;

var mailListText = "";
var templateText = "";
var sheetsFileId = "";
var processFolder;


/**
 * Convert Excel file to Sheets
 * @param {Blob} excelFile The Excel file blob data; Required
 * @param {String} filename File name on uploading drive; Required
 * @param {Array} arrParents Array of folder ids to put converted file in; Optional, will default to Drive root folder
 * @return {Spreadsheet} Converted Google Spreadsheet instance
 **/
function convertExcel2Sheets(excelFile, filename, arrParents) {

    var parents = arrParents || []; // check if optional arrParents argument was provided, default to empty array if not
    if (!parents.isArray) parents = []; // make sure parents is an array, reset to empty array if not

    // Parameters for Drive API Simple Upload request (see https://developers.google.com/drive/web/manage-uploads#simple)
    var uploadParams = {
        method: 'post',
        contentType: 'application/vnd.ms-excel', // works for both .xls and .xlsx files
        contentLength: excelFile.getBytes().length,
        headers: { 'Authorization': 'Bearer ' + ScriptApp.getOAuthToken() },
        payload: excelFile.getBytes()
    };

    // Upload file to Drive root folder and convert to Sheets
    var uploadResponse = UrlFetchApp.fetch('https://www.googleapis.com/upload/drive/v2/files/?uploadType=media&convert=true', uploadParams);

    // Parse upload&convert response data (need this to be able to get id of converted sheet)
    var fileDataResponse = JSON.parse(uploadResponse.getContentText());

    // Create payload (body) data for updating converted file's name and parent folder(s)
    var payloadData = {
        title: filename,
        parents: []
    };
    if (parents.length) { // Add provided parent folder(s) id(s) to payloadData, if any
        for (var i = 0; i < parents.length; i++) {
            try {
                var folder = DriveApp.getFolderById(parents[i]); // check that this folder id exists in drive and user can write to it
                payloadData.parents.push({ id: parents[i] });
            }
            catch (e) { } // fail silently if no such folder id exists in Drive
        }
    }
    // Parameters for Drive API File Update request (see https://developers.google.com/drive/v2/reference/files/update)
    var updateParams = {
        method: 'put',
        headers: { 'Authorization': 'Bearer ' + ScriptApp.getOAuthToken() },
        contentType: 'application/json',
        payload: JSON.stringify(payloadData)
    };

    // Update metadata (filename and parent folder(s)) of converted sheet
    UrlFetchApp.fetch('https://www.googleapis.com/drive/v2/files/' + fileDataResponse.id, updateParams);

    return SpreadsheetApp.openById(fileDataResponse.id);
}


function defaultParameters() {
    var pars = {
        SHEET_FILE_NAME: SHEET_FILE_NAME,
        MAIL_LIST_FILE_NAME: MAIL_LIST_FILE_NAME,
        TEMPLATE_FILE_NAME: TEMPLATE_FILE_NAME,
        RESULT_TEMPLATE: RESULT_TEMPLATE,
        OUTPUT_FOLDER: OUTPUT_FOLDER,
        SUBJECT: subject,
        SENDER_MAIL: SENDER_MAIL,
        SENDER_NAME: SENDER_NAME,
        SENDER_TITLE: SENDER_TITLE,
        STAKEHOLDERS_NAMES: STAKEHOLDERS_NAMES,
        EXCEL_FILE_NAME: EXCEL_FILE_NAME
    }
    return pars;
}

function loadSettingsJSON() {
    var jsonText = "";
    var file = getFileByName(SETTINGS_FILE_NAME);
    if (file != undefined) {
        jsonText = file.getAs('application/json')
        p = JSON.parse(jsonText);
    }
    else {
        p = defaultParameters();
        jsonText = JSON.stringify(p);
    }

    return jsonText;
}


function saveAsJSON() {
    var blob, file, fileSets, obj;


    fileSets = {
        title: SETTINGS_FILE_NAME,
        mimeType: 'application/json'
    };

    blob = Utilities.newBlob(JSON.stringify(p), "application/vnd.google-apps.script+json");
    file = DriveApp.Files.insert(fileSets, blob);
    Logger.log('ID: %s, File size (bytes): %s, type: %s', file.id, file.fileSize, file.mimeType);
}

function saveSettings() {
    var file = getFileByName(SETTINGS_FILE_NAME);
    if (file != undefined)
        file.setTrashed();

    var doc = DocumentApp.create(SETTINGS_FILE_NAME);


    var json = JSON.stringify(p);

}
/* @Include JavaScript and CSS Files */
function include(filename) {
    return HtmlService.createHtmlOutputFromFile(filename)
        .getContent(); 303
}


function convertToSheets(xlsBlob) {
    //var xlsId = "0B9**************OFE"; // ID of Excel file to convert
    //var xlsFile = DriveApp.getFileById(xlsId); // File instance of Excel file
    //var xlsBlob = xlsFile.getBlob(); // Blob source of Excel file for conversion
    Logger.log("converting excel file");
    //var xlsFilename = xlsFile.getName(); // File name to give to converted file; defaults to same as source file
    var destFolders = ["MailMergeInput"]; // array of IDs of Drive folders to put converted file in; empty array = root folder
    var ss = convertExcel2Sheets(xlsBlob, p.SHEET_FILE_NAME, destFolders);
    sheetsFileId = ss.getId();
    Logger.log(ss.getId());
}

/* @Process Form */
function processForm(formObject) {
    Logger.log("processForm()");
    loadSettingsJSON();
    try {
        p.SHEET_FILE_NAME = formObject.SHEET_FILE_NAME
        p.MAIL_LIST_FILE_NAME = formObject.MAIL_LIST_FILE_NAME
        p.TEMPLATE_FILE_NAME = formObject.TEMPLATE_FILE_NAME
        p.RESULT_TEMPLATE = formObject.RESULT_TEMPLATE
        p.OUTPUT_FOLDER = formObject.OUTPUT_FOLDER
        p.SUBJECT = formObject.SUBJECT
        p.SENDER_MAIL = formObject.SENDER_MAIL
        p.SENDER_TITLE = formObject.SENDER_TITLE
        p.STAKEHOLDERS_NAMES = formObject.STAKEHOLDERS_NAMES

        //https://stackoverflow.com/questions/56063156/script-to-convert-xlsx-to-google-sheet-and-move-converted-file
        /*
        Logger.log("Receiving excel file");
        var fileBlob = formObject.EXCEL_FILE_NAME;  //.upload;
        inputFolder = DriveApp.createFolder("MailMergeInput");
        if (fileBlob != undefined) {
            Logger.log("excel file uploaded");
            //var excel = DriveApp.createFile(fileBlob);
            var excel = inputFolder.createFile(fileBlob);
            Logger.log("excel file created");
            excelFileId = excel.getId();

            Logger.log("converting to sheets");
            var newFile = {
                title: p.SHEET_FILE_NAME,
                key: excel.getId()
            }
            var file = DriveApp.Files.insert(newFile, fileBlob, {
                convert: true
            });
            Logger.log("excel file converted");
        }
        */
    }
    catch (Exception) {
        Logger.log("Exception. parameters received:");
        Logger.log(p);
        Logger.log(Exception);
    }
    var result = mailMerge();
    return result;
    //return "excel loaded";
    //saveSettings();
}


function doGet(e) {
    //return HtmlService.createTemplateFromFile('Parameters_AppKit2').evaluate();
    return HtmlService.createTemplateFromFile('Parameters_AppKit').evaluate();
}



function moveFiles(sourceFileId, targetFolderId) {
    try
    {
        var file = DriveApp.getFileById(sourceFileId);
        var folder = DriveApp.getFolderById(targetFolderId);
        file.moveTo(folder);
    }
    catch(ex)
    {
        Logger.log("Exception moving file.");
    }
}
function reCreateFolder(folderName) {
    var folders = DriveApp.getFoldersByName(folderName);
    if (folders.hasNext()) {
        while (folders.hasNext()) {
            var folder = folders.next();
            folder.setTrashed(true);
            break;
        }
    }
    folder = DriveApp.createFolder(folderName);
    return folder;
}


function selectFolder(folderName) {
    var folders = DriveApp.getFoldersByName(folderName);
    var folder = null;
    if (folders.hasNext()) {
        while (folders.hasNext()) {
            folder = folders.next();
            break;
        }
    }
    else
        folder = DriveApp.createFolder(folderName);
    return folder;
}

function getFileByName(fileName) {
    var files = DriveApp.getFilesByName(fileName);
    while (files.hasNext()) {
        var file = files.next();
        return file;
        break;
    }
    return null;
}


//Reads text from a file given its fielId
function getText(fileId) {
    var doc = DocumentApp.openById(fileId);
    return doc.getBody().getText();
}

function getTextByName(fileName) {
    var text = "";
    var file = getFileByName(fileName);
    if (file != undefined) {
        //var doc = DocumentApp.open(file);
        var doc = DocumentApp.openById(file.getId());
        text = doc.getBody().getText();
    }
    return text;
}

function getBlobByName(fileName) {
    var text = "";
    var file = getFileByName(fileName);
    if (file != undefined) {
        text = file.getAs(MimeType.PLAIN_TEXT)
    }
    return text;
}


//Gets a file given its folder id and fileName
function getFile(folder_id, name) {
    var folder = DriveApp.getFolderById(folder_id);
    var files = folder.getFiles();
    while (files.hasNext()) {
        file = files.next();
        if (file.getName().toLowerCase() == name.toLowerCase()) {
            return file.getId();
            break;
        }
        return "";
    };
}



function getTemplate(fileName, subject, namesList, sendername, sendermail, sendertitle, html) {
    html = html.replace("${names}", namesList);
    html = html.replace("${files}", `${fileName}`);
    html = html.replace("${titleDBs}", `${fileName}`);
    html = html.replace("${sendername}", sendername);
    html = html.replace("${sendername}", sendername);
    html = html.replace("${sendername}", sendername);
    index = html.indexOf("${sendermail}");
    html = html.replace("${sendermail}", sendermail);
    html = html.replace("${sendermail}", sendermail);
    html = html.replace("${sendertitle}", sendertitle);
    html = html.replace("${subject}", subject);
    return html;
}

function AppendRow(sheet,data)
{
    if ( sheet != undefined )
    {
        var cols = data.split('\t');
        sheet.appendRow(cols);
    }
}

function ValidateEmail(mail) {
    if (/^[a-zA-Z0-9.!#$%&'*+/=?^_`{|}~-]+@[a-zA-Z0-9-]+(?:\.[a-zA-Z0-9-]+)*$/.test(mail))
        return true;
    return false;
}

function findMail(name) {
    name = name.toLowerCase().trim();
    var index = mailListText.indexOf(name);
    var email = "";
    var index2 = -1;
    if (index >= 0) {
        index += name.length;
        index2 = mailListText.indexOf(",", index);
        if (index2 >= index) {
            index2++;
            var index3 = mailListText.indexOf("\n", index2);
            if (index3 > index2) {
                var len = index3 - index2;
                email = mailListText.substr(index2, len);
            }
        }
    }
    Logger.log("findMail " + name + " " + email);
    return email;
}


function processNames(namesList) {
    emailsList = "";
    if (namesList.indexOf('\n') >= 0)
        names = namesList.split('\n');
    else
        names = namesList.split(',');


    mails = [];
    notFound = [];
    notFoundString = "";
    var rowText = "";

    for (k = 0; k < names.length; k++) {
        totalLines++;
        var email = findMail(names[k]);
        rowText = `${server}\t\t${title}\t${fileName}\t${names[k]}\t${email}`;
        if (email != "") {
            if (ValidateEmail(email)) {
                mails.push(email);
                emailsList += emailsList + email + ",";
            }
            else {
                invalidMailEntries += `${rowText}\n`;
                AppendRow(ssInvalidMails,rowText);
                imRow++;
                invalidMails++;
            }
        }
        else {
            notFound.push(names[k]);
            notFoundEntries += `${rowText}\n`;
            inRow++;
            AppendRow(ssNotFound,rowText);
        }
    }
    return emailsList;
}

function RenameFile(file, newName, sDate) {

    var fileId = file.getId();
    var SourceFolder = file.getParents();
    while (SourceFolder.hasNext()) {
        var folder = SourceFolder.next();
        if (folder != undefined) {
            newName = `${newName}.${sDate}`;
            var file2 = file.makeCopy(newName);
            SourceFolder.removeFile(file);
        }
        break;
    }
    return newName;
}


function mailMerge() {

    try {
        //loadSettingsJSON();
        Logger.log("mailMerge() p:" + JSON.stringify(p));
        var ssheet = getFileByName(p.SHEET_FILE_NAME);
        var todayStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH-mm-ss');
        processFolder = selectFolder(PROCESS_FOLDER);


        if (ssheet != undefined) {

            //var dbTab = getSheetByName(name)
            var sheet = SpreadsheetApp.open(ssheet);
            var rangeData = sheet.getDataRange();
            var lastColumn = rangeData.getLastColumn();
            var lastRow = rangeData.getLastRow();

            mailListText = getTextByName(p.MAIL_LIST_FILE_NAME).toLowerCase();
            templateText = getTextByName(p.TEMPLATE_FILE_NAME);

            if (mailListText != "" && templateText != "") {
                notFoundFileName = `NotFoundNames.${todayStr}`;
                ssNotFound = SpreadsheetApp.create(notFoundFileName);
                AppendRow(ssNotFound,"Server\tADO\tTitle\tFile Name\tOwner\tStatus");
                notFoundRow = [];
        
                invalidMailsFileName = `InvalidMails.${todayStr}`;
                ssInvalidMails = SpreadsheetApp.create(invalidMailsFileName);
                AppendRow(ssInvalidMails,"Server\tADO\tTitle\tFile Name\tOwner\tStatus");
                invalidMailRow = [];
        

                for (var i = 2; i <= lastRow; i++) {
                    server = rangeData.getCell(i, 1).getValue(); 
                    title = rangeData.getCell(i, 3).getValue();
                    fileName = rangeData.getCell(i, 4).getValue();
                    var namesList = rangeData.getCell(i, 5).getValue();

                    processNames(namesList);
                    totalMails += names.length;
                    notFoundMails += notFound.length;
                    if (mails.length > 0) {

                        var merge = getTemplate(`${title} (${server}\\${fileName})`, p.SUBJECT, namesList, p.SENDER_NAME, p.SENDER_MAIL, p.SENDER_TITLE, templateText);
                        MailApp.sendEmail({
                            to: emailsList,
                            subject: p.SUBJECT,
                            htmlBody: merge
                        });
                        sentMails++;
                        lines++;
                    }

                }
                if (totalLines > 0) {
                    var folder = reCreateFolder(p.OUTPUT_FOLDER);
                    moveFiles(ssNotFound.getId(), processFolder.getId());
                    moveFiles(ssInvalidMails.getId(), processFolder.getId());
                    overAllResult = "";
                    if (folder != undefined) {
                        overAllResult = "OK;"
                        if (notFoundEntries.length > 0) {
                            if (overAllResult == "OK")
                                overAllResult = "SOME ENTRIES NOT FOUND. "
                            else
                                overAllResult += "SOME ENTRIES NOT FOUND. "

                            notFoundFileObject = folder.createFile(`${p.SHEET_FILE_NAME}.NotFound.${todayStr}.txt`, notFoundEntries, MimeType.PLAIN_TEXT);
                        }
                        if (invalidMailEntries.length > 0) {
                            if (overAllResult == "OK")
                                overAllResult = "SOME MAIL ADDRESSES INVALID."
                            else
                                overAllResult += "SOME MAILS ARE INVALID."

                            invalidMailsFileObject = folder.createFile(`${p.SHEET_FILE_NAME}.InvalidEmails.${todayStr}.txt`, invalidMailEntries, MimeType.PLAIN_TEXT);
                        }
                    }
                    else {
                        errors += `<p>Can't find output folder ${p.OUTPUT_FOLDER}</p>`;
                    }
                }
                else {
                    errors += `<p>Sheet File empty ${p.SHEET_FILE_NAME} </p>`;
                }
            }
            else {
                if (mailListText == "")
                    errors += `<p>Mail List not found. ${p.MAIL_LIST_FILE_NAME}</p>`;
                if (templateText == "")
                    errors += `<p>Template file not found. ${p.TEMPLATE_FILE_NAME}</p>`
            }

        }
        else
            errors += `"<p>Can't find sheet file ${p.SHEET_FILE_NAME} </p>"`;
    }
    catch (exception) {
        tt = `${tt}</br><p>Errors</p></br>Exception:${exception}`;
    }

    //final results

    tt = getTextByName(p.RESULT_TEMPLATE);
    tt = tt.replace("${executiontime}", todayStr);
    tt = tt.replace("${stakeholders}", p.STAKEHOLDERS_NAMES);
    tt = tt.replace("${overallresult}", overAllResult);
    tt = tt.replace("${linesprocessed}", `${lines}/${totalLines}`);
    tt = tt.replace("${sendername}", p.SENDER_NAME);
    tt = tt.replace("${sendertitle}", p.SENDER_TITLE);
    tt = tt.replace("${invalidEmailCount}", `${invalidMails}`);
    tt = tt.replace("${notfoundcount}", `${notFoundMails}`);
    tt = tt.replace("${mailsSentCount}", `${sentMails}/${totalMails}`);

    //rename sheet
    var newName = `${p.SHEET_FILE_NAME}.Processed.${todayStr}`;
    var newFile = null;
    try {
        var SourceFolder = ssheet.getParents();
        while (SourceFolder.hasNext()) {
            var folder = SourceFolder.next();
            if (folder != undefined) {
                newFile = ssheet.makeCopy(newName);
                ssheet.setTrashed(true);
            }
            break;
        }
    }
    catch (exception) {
        newFile = null;

    }

    try {
        if (newFile != undefined)
            tt = tt.replace("${inputfile}", `<a href='${newFile.getUrl()}'>${newName}</a>`);
        else
            tt = tt.replace("${inputfile}", p.SHEET_FILE_NAME);

        if ( imRow > 0 )
            tt = tt.replace("${notfoundlink}", `<a href='${ssNotFound.getUrl()}'>${notFoundFileName}</a>`)            
        else 
            tt = tt.replace("${notfoundlink}", "");

        if ( inRow > 0 )
            tt = tt.replace("${invalidmailslink}", `<a href='${ssInvalidMails.getUrl()}'>${invalidMailsFileName}</a>`);
        else 
            tt = tt.replace("${invalidmailslink}", "");
        /*
        if (notFoundFileObject != undefined)
            tt = tt.replace("${notfoundlink}", `<a href='${notFoundFileObject.getUrl()}'>Not Found Entries</a>`)
        else
            tt = tt.replace("${notfoundlink}", "")

        if (invalidMailsFileObject != undefined)
            tt = tt.replace("${invalidmailslink}", `<a href='${invalidMailsFileObject.getUrl()}'>Invalid Email Entries</a>`);
        else
            tt = tt.replace("${invalidmailslink}", "");
            */
    }
    catch (exception) {
        tt = tt.replace("${notfoundlink}", "")
        tt = tt.replace("${invalidmailslink}", "");
    }

    if (errors != "")
        tt = `${tt}</br><p>Errors</p></br>${errors}`;

    title = "Stakeholders";
    processNames(p.STAKEHOLDERS_NAMES);
    if (emailsList != "") {
        MailApp.sendEmail({
            to: emailsList,
            subject: "Mail Merge process report.",
            htmlBody: tt
        });
    }
    else {
        MailApp.sendEmail({
            to: p.SENDER_MAIL,
            subject: "Mail Merge process report. Stake Holder mails not found",
            htmlBody: tt
        });

    }
    return tt;
}
