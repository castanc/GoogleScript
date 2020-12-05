import { Application } from "../../Models/Application";
import { Authorization } from "../../Models/authorization";
import { G } from "../Globals/G";
import { BaseBusiness } from "./baseBusiness";
import { Role } from '../../Models/role'
import { AuthBase } from "../../Models/AuthBase";
import { IResult } from "../../Models/IResult";
import { IDataAccess } from "../DataAccess/IDataAccess";
import { DALFactory } from "../DataAccess/DALFactory";
import { Utils } from "../utils";
import { GSLog } from "../GSLog";
import { GSObject } from "../../Models/GSObject";
import { UserProfile } from "../../Models/UserProfile";

export class boAuthorization {

    result: IResult<Authorization>;
    SH_APP_ROLE_USERS = "AppRoleUsers";
    DAL: IDataAccess;
    fileName: string = "";
    folder;
    ssFile;
    applicationName: string = "";

    constructor(_appName: string) {
        this.applicationName = _appName;
        this.folder = DriveApp.getFolderById(G.folderId);
        let file = Utils.getFileFromFolder(this.folder, this.applicationName);
        this.DAL = DALFactory.GetDAL(G.DAL_SHEETS, G.folderId, G.authorizationSpreadSheetID);
    }

    getUserName(email:string):string{
        let index = email.indexOf("@");
        return email.substring(0,index);
    }
    getUserProfile(app:Application, email: string = ""):UserProfile
    {
        if ( email.length == 0 )
            email = Session.getActiveUser().getEmail();

        let roles = app.Roles.filter(x=>x.Users.toLowerCase().indexOf(`${email.toLowerCase()}`)>=0);

        let index = email.indexOf("@");
        let name = email.substring(0,index);
        let fileName = `UserProfile.${G.ApplicationName}.${name}.json`;
        let json = Utils.getTextFileFromFolder(this.folder,fileName);
        if ( json.length > 0 )
            return JSON.parse(json);

        if ( roles.length > 0 )
        {
            let up = new UserProfile();
            up.Role = roles[0];
            up.Name = name;
            up.Email = email;
            up.Objects = app.Objects.filter(x=>x.roles.toLowerCase().indexOf(`,${roles[0].Name.toLowerCase()},`)>=0);


            up.Role.Users = "";
            for(let i=0; i< up.Objects.length; i++)
            {
                up.Objects[i].roles = "";
            }

            json = JSON.stringify(up);
            Utils.writeTextFile(fileName,json, this.folder);
            return up;
        }
        return null;
    }


    getApplicationSecurity():Application {
        let SH_APPLICATION = 0;
        let SH_ROLE = 1;
        let SH_OBJECT = 2;
        let app = new Application();

        this.folder = DriveApp.getFolderById(G.folderId);
        let jsonText = Utils.getTextFileFromFolder(this.folder,`${this.applicationName}.json`);
        if ( jsonText.length>0)
            return JSON.parse(jsonText)

        let file = Utils.getFileFromFolder(this.folder, this.applicationName);
        if (file == null) return null;

        this.ssFile = SpreadsheetApp.openById(file.getId());
        let sheets = this.ssFile.getSheets()
        for (let i = 0; i < sheets.length; i++) {
            var rangeData = sheets[i].getDataRange();
            var lastColumn = rangeData.getLastColumn();
            var lastRow = rangeData.getLastRow();

            let grid = rangeData.getValues();
            try {
                if (i == SH_APPLICATION) {
                    app.appId = grid[0][1].trim();
                    app.name = grid[1][1].trim();
                    app.description = grid[2][1];
                }
                else if (i == SH_ROLE) {
                    app.Roles = new Array<Role>();
                    for (let j = 1; j < lastRow; j++) {
                        let r = new Role();
                        r.Name = grid[j][0].trim();
                        r.level = Number(grid[j][1]);
                        r.Permissions = 0;
                        for (let k = 3; k < 8; k++) {
                            let p = Number(grid[j][k]);
                            r.Permissions += p;
                        }
                        
                        r.Read = (r.Permissions & 1) > 0;
                        r.Write = (r.Permissions & 2) > 0;
                        r.Add = (r.Permissions & 4) > 0;
                        r.Delete = (r.Permissions & 8) > 0;
                        r.Drop = (r.Permissions & 16) > 0;
                        
                        r.Users = `,${grid[j][8].trim()},`;
                        app.Roles.push(r);
                    }
                }
                else if (i == SH_OBJECT) {

                    app.Objects = new Array<GSObject>();

                    for (let j = 1; j < lastRow; j++) {
                        let o = new GSObject();
                        o.type = grid[j][0];
                        o.name = grid[j][1].trim();
                        o.text = grid[j][2];
                        o.roles =`,${grid[j][3].trim()},`;
                        o.options = 0;
                        for(let k=5; k<8; k++)
                        {
                            o.options += Number(grid[j][k]);
                        }
                        o.notLoad = (o.options & 1 )>0;
                        o.hide = (o.options & 2 )>0;
                        o.disable = (o.options & 4 )>0;
                        o.protect = (o.options & 8 )>0;

                        app.Objects.push(o);
                    }

                }
            }
            catch (ex) {
                Logger.log("exception exportTOJSON()", ex);
                GSLog.logException("boAuthorization", "exportToJSON()", ex);
            }
            finally {
                Logger.log("Exported App", JSON.stringify(app));
                let json =JSON.stringify(grid);
            }
        }
        let json = JSON.stringify(app);
        this.folder.createFile(`${app.name}.json`, json, MimeType.PLAIN_TEXT);
        return app;
    }


    getApplication(appName: string): Application {
        return null;
    }

    getRole(id: number): Role {
        return null;
    }





    isAuthorized(appName: string, userEmail: string, actionCode: string): Authorization {
        /*
        let app = this.getApplication(appName);
        let appRoleUsers = this.DAL.getAll<any>(this.SH_APP_ROLE_USERS);
        const EMAIL_COL = 3;
        const APP_COL = 1;
        const ROLE_COL = 2;
 
        let filtered = appRoleUsers?.filter(innerArray => innerArray[EMAIL_COL] == userEmail &&
            innerArray[APP_COL] == app.id );
            
        let maxRole: number = -1;
        let maxValue = 0;
        if ( filtered?.length > 0 )
        {
            let roles = new Array<Role>();
            for(let i=0; i< filtered.length; i++)
            {
                let role = this.getRole(filtered[i][ROLE_COL]);
                if ( role != null)
                {
                    if ( role.level >= maxValue )
                    {
                        maxValue = role.level;
                        maxRole = i;
                    }
 
                    roles.push(role);
                }
            }
            let highestRole = roles.filter(x=>x.level == maxValue);
            let finalRole = null;
            if ( highestRole.length > 0 )
                finalRole = roles[maxRole];
 
            
            
        }
        */
        return null;
    }

}
