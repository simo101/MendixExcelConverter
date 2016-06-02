/// <reference path='./typings/tsd.d.ts' />

import { MendixSdkClient, OnlineWorkingCopy, Project, Revision, Branch, loadAsPromise } from "mendixplatformsdk";
import { ModelSdkClient, IModel, projects, domainmodels } from "mendixmodelsdk";
import when = require('when');

const XLSX = require('xslx');

const username = "simon.black@mendix.com";
const apikey = "c665aa21-d314-4ef8-8280-822b055a3d9d";
const projectId = "799c70bc-1a75-4d40-80a6-ef8f4303f6e5";
const projectName = "ExcelImport";
const revNo = -1; // -1 for latest
const branchName = null; // null for mainline
const wc = null;
const client = new MendixSdkClient(username, apikey);
const fileName = "excelimport.xlsx";


/*
 * PROJECT TO ANALYZE
 */
const project = new Project(client, projectId, projectName);

client.platform().createOnlineWorkingCopy(project, new Revision(revNo, new Branch(project, branchName))).then(wc=>{return parseExcelDocument(wc,fileName)}).done(()=>{
    console.log("Excel Document Parsed");
});

function parseExcelDocument(workingCopy: OnlineWorkingCopy, nameOfDocument: String): when.Promise<void> {
    var workbook = XLSX.readFile(nameOfDocument);
     var worksheet = workbook.Sheets[0];
     var z = null;
    for (z in worksheet) {
        if(z[0] === '!') continue;
        console.log(worksheet + "!" + z + "=" + JSON.stringify(worksheet[z].v));
    }
  
    return;
}