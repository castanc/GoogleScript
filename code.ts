import { boAuthorization } from "./Business/boAuthorization";
import { Business } from "./Business/business";
import { G } from "./Globals/G";
import { GSLog } from "./GSLog";
import { Utils } from "./utils";
import { Application } from "../Models/Application";
import { KeyValuePair } from "../Models/KeyValuePair";
import { ServerResponse } from "../Models/ServerResponse";
import { UserProfile } from "../Models/UserProfile";


let jsonUserProfile = "";
let userProfile: UserProfile;

/* @Include JavaScript and CSS Files */
function include(filename) {
    return HtmlService.createHtmlOutputFromFile(filename)
        .getContent(); 303
}


function buildMenu() {
    GSLog.logMessage(0, "code.ts", "buildMenu()", "", "userProfile");
    if (userProfile != null) {
      let profileName = userProfile.Role.Name.toLowerCase();
      return HtmlService.createHtmlOutputFromFile(`frontend/menu${profileName}`).getContent();
    }
    else {
        GSLog.flushLog();
        return HtmlService.createHtmlOutputFromFile('frontend/forbidden').getContent();
    }
}

function doGet(e) {
    //GSLog.initialize(`GSLog.${G.ApplicationName}`, G.folderId);
    GSLog.logMessage(GSLog.Level, "code.ts", "doGet(e)", "e", JSON.stringify(e), true);
    let userEmail = "";

    try {
        userEmail = e.pathInfo;
    }
    catch (ex) {
        userEmail = Session.getActiveUser().getEmail();
        GSLog.logException("code.ts", "doGet()", ex);
    }

    GSLog.logMessage(GSLog.Level, "code.ts", "doGet(e)", "userEmail", userEmail, true);

    try {
        let b = new boAuthorization(G.ApplicationName);
        GSLog.logText("bo built")
        let app = getApplicationSecurity(G.ApplicationName);
        GSLog.logText("app obj obtained");

        if (app != null) {
            GSLog.logMessage(0, "code.ts", "doGet()", JSON.stringify(app), "app");
            userProfile = b.getUserProfile(app, userEmail);
        }
        else
            GSLog.logText("app objecy is null ");

        if (userProfile != null) {
            GSLog.logMessage(0, "code.ts", "doGet()", JSON.stringify(userProfile), "Returned user profile");
            jsonUserProfile = `var userProfile =${JSON.stringify(userProfile)};console.log("userProfile:",userProfile);`;

            //let html = HtmlService.createTemplateFromFile('frontend/pwcindex').evaluate().getContent();

            //if (userProfile.Role.Name.toLowerCase() == "admin")
            GSLog.flushLog();
            return HtmlService.createTemplateFromFile('frontend/pwcindex').evaluate();
        }
        else
            GSLog.logText("USER PROFIULE IS NULL")
    }
    catch (ex) {
        GSLog.logException("code.ts", "doGet()", "exception");
    }
    GSLog.flushLog();
    return HtmlService.createHtmlOutputFromFile('frontend/forbidden');

}


function renderUserProfile() {
    return jsonUserProfile;
}

function getForm(role, email) {
    return `    <p class="h4 mb-4 text-center">Current User Info</p>

    <div class="form-row">
        <div class="form-group col-md-6">
            <label for="USER_ROLE">User Role: ${role}</label>
        </div>
        <div class="form-group col-md-6">
            <label for="USER_EMAIL">User Email: ${email}</label>
        </div>
    </div>
`
}

function SelectUser(user: string): ServerResponse {
    let sr = new ServerResponse();
    sr.message = "User selected";
    sr.result = 200;
    let html = "";
    let resultValue = 200;
    GSLog.initialize(`GSLog.${G.ApplicationName}`, G.folderId);
    let userProfile = getUserProfile(G.ApplicationName, user);

    jsonUserProfile = `var userProfile =${JSON.stringify(userProfile)};console.log("userProfile:",userProfile);`;


    if (userProfile != null) {
        //html = HtmlService.createTemplateFromFile('frontend/form').evaluate().getContent();
        html = getForm(userProfile.Role.Name, userProfile.Email);
    }
    else {
        html = getForm("", "User not found in saved configuration. Refresh configuration");
        sr.result = 404;
        sr.message = "User not found";
    }
    GSLog.logMessage(0, "code.ts", "SelectUser", "user Profile:", JSON.stringify(userProfile), true);

    sr.html.push(new KeyValuePair<string, string>("content", html));
    sr.html.push(new KeyValuePair<string, string>("dynamicScript", jsonUserProfile));
    sr.json.push(new KeyValuePair<string, string>("userProfile", JSON.stringify(userProfile)));
    return sr;
}

function getUserProfile(appName, email: string = ""): UserProfile {
    let b = new boAuthorization(G.ApplicationName);
    let app = getApplicationSecurity(appName);
    if (app != null) {
        GSLog.logMessage(0, "code.ts", "getUserProfile()", JSON.stringify(app), "app");
        let up = b.getUserProfile(app, email);
        GSLog.logMessage(0, "code.ts", "getUserProfile()", JSON.stringify(up), "Returned user profile");
        return up;
    }
    return null;
}

function getApplicationSecurity(appName): Application {
    let b = new boAuthorization(appName);
    return b.getApplicationSecurity();
}

function doPost(e) {

}

function profileInfo(up:UserProfile):string{
  let html = 
  `<div class="form-row">
      <div class="form-group col-md-6">
          <label for="ROLE_NAME" class="question-title" >Role Name:${up.Role.Name}</label>
      </div>
      <div class="form-group col-md-6">
          <label for="ROLE_NAME" class="question-title" >User Email:${up.Email}</label>
      </div>
  </div>
`;
  return html;
}

function LoadForm(formName:string, email: string):string{
  /*
  //this is required if we are going to generate the form server side
  //where I will need the profile,
  //or if the forms are coded to a dynamic form builder
  //
  //also we can have a single form which is manipulated client side with the
  let b = new boAuthorization(G.ApplicationName);
  let name = b.getUserName(email);
  let fileName = `UserProfile.${G.ApplicationName}.${name}.json`;
  let json = Utils.getTextFileFromFolder(this.folder,fileName);
  let userProfile: UserProfile;

  if ( json.length > 0 )
      userProfile = JSON.parse(json);
  else {
    let app = b.getApplicationSecurity();
    userProfile = b.getUserProfile(app,email);
  }
  */

  return HtmlService.createTemplateFromFile(`frontend/${formName}`).evaluate().getContent();
}

function logException(file, method, text) {
    GSLog.logException(file, method, text);
}

function logMessage(level, file, method, text, additional: string = "", flush: boolean = false) {
    GSLog.logMessage(level, file, method, text, additional, flush);
}

