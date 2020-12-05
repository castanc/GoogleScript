import { Utils } from "./utils";

export class GSLog {
    static log = "";
    static logFileName = "GSLogs";
    static folder;
    static folderID = "";
    static ss;
    static Level = 0;
    static initialized = false;


    static _file = "";
    static _method = "";

    static logException(file, method, text) {
        if (!this.initialized)
            this.initialize();
        this._file = file;
        this._method = method;

        this.log = `${this.log}\nEXCEPTION\t${file}\t${method}\t${Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH-mm-ss')}\t\t${text}`;
        Logger.log(this.log);
        this.flushLog(text);
        return this.log;
    }

    static logText(text, flush: boolean = false) {
        this.log = `${this.log}\nEXCEPTION\t${this._file}\t${this._method}\t${Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH-mm-ss')}\t\t${text}`;
        Logger.log(this.log);
        if (flush)
            this.flushLog(text);
        return this.log;
    }


    static async logMessage(level, file, method, text, additional: string = "", flush: boolean = false) {
        this._file = file;
        this._method = method;
        if (level >= this.Level) {
            this.log = `${this.log}\nMSG\t${file}\t${method}\t${Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH-mm-ss')}\t${additional}\t${text}`;
            Logger.log(this.log);
            if (flush)
                this.flushLog();
            return text;
        }

    }

    static async flushLog(text = "") {
        if (text.length == 0)
            text = this.log;
        if (this.initialized) {
            if (this.log.length > 0) {
                let lines = text.split('\n');
                let sheet = this.ss.getActiveSheet();
                for (let i = 0; i < lines.length; i++) {
                    let cols = lines[i].split('\t');
                    this.ss.appendRow(cols);
                }
            }
        }
        else
            DriveApp.createFile(this.logFileName, text, MimeType.PLAIN_TEXT);


    }
    static async initialize(_logFileName: string = "", _folderId: string = "") {
        if (_logFileName.length > 0)
            this.logFileName = _logFileName;

        if (_folderId.length > 0)
            this.folderID = _folderId;

        //Utils.deleteFiles(this.logFileName);

        if (this.folderID.length > 0) {
            this.folder = DriveApp.getFolderById(this.folderID);
            Utils.deleteFiles(this.logFileName, this.folder);

            this.ss = SpreadsheetApp.create(this.logFileName);
            Logger.log("GSLog.tx", "initialize()", this.ss.getId(), "ID Original");
            Utils.moveFiles(this.ss.getId(), this.folder.getId());
            let file = Utils.getFileFromFolder(this.folder, this.logFileName);
            this.ss = SpreadsheetApp.openById(file.getId());
            Logger.log("GSLog.tx", "initialize()", this.ss.getId(), "ID New");
        }
        else {
            Logger.log("Initializing logs to root folder");
            await Utils.removeFileByName(this.logFileName);
            this.ss = SpreadsheetApp.create(this.logFileName);
        }
        this.initialized = true;
    }

}
