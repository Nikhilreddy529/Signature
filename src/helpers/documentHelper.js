/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */
import globalEvent from "../commands/commands";
/* global Excel, Office, Word, Promise */

export function writeDataToOfficeDocument(result) {
  return new Promise(function (resolve, reject) {
    try {
      switch (Office.context.host) {
        case Office.HostType.Excel:
          writeDataToExcel(result);
          break;
        case Office.HostType.Outlook:
          writeDataToOutlook(result);
          break;
        case Office.HostType.PowerPoint:
          writeDataToPowerPoint(result);
          break;
        case Office.HostType.Word:
          writeDataToWord(result);
          break;
        default:
          throw "Unsupported Office host application: This add-in only runs on Excel, Outlook, PowerPoint, or Word.";
      }
      resolve();
    } catch (error) {
      reject(Error("Unable to write data to document. " + error.toString()));
    }
  });
}

function filterUserProfileInfo(result) {
  let userProfileInfo = [];
  userProfileInfo.push(result["displayName"]);
  userProfileInfo.push(result["jobTitle"]);
  userProfileInfo.push(result["mail"]);
  userProfileInfo.push(result["mobilePhone"]);
  userProfileInfo.push(result["officeLocation"]);
  return userProfileInfo;
}

function writeDataToExcel(result) {
  return Excel.run(function (context) {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    let data = [];
    let userProfileInfo = filterUserProfileInfo(result);

    for (let i = 0; i < userProfileInfo.length; i++) {
      if (userProfileInfo[i] !== null) {
        let innerArray = [];
        innerArray.push(userProfileInfo[i]);
        data.push(innerArray);
      }
    }
    const rangeAddress = `B5:B${5 + (data.length - 1)}`;
    const range = sheet.getRange(rangeAddress);
    range.values = data;
    range.format.autofitColumns();

    return context.sync();
  });
}

function writeDataToOutlook(result) {
  let data = [];
  let userProfileInfo = filterUserProfileInfo(result);

  for (let i = 0; i < userProfileInfo.length; i++) {
    if (userProfileInfo[i] !== null) {
      data.push(userProfileInfo[i]);
    }
  }

  let userInfo = "";
  for (let i = 0; i < data.length; i++) {
    userInfo += data[i] + "\n";
  }
  var formattedtext = ReplaceSignature(result);
  
  Office.context.mailbox.item.body.setSelectedDataAsync(
    formattedtext,
    {
      "asyncContext" : globalEvent,
      "coercionType" : Office.CoercionType.Html
    },
    function (asyncResult) {
      // Handle success or error.
      if (asyncResult.status !== Office.AsyncResultStatus.Succeeded) {
        console.error("Failed to set signature: " + JSON.stringify(asyncResult.error));
      }

      // Call event.completed() after all work is done.
      asyncResult.asyncContext.completed();
    });
}

function writeDataToPowerPoint(result) {
  let data = [];
  let userProfileInfo = filterUserProfileInfo(result);

  for (let i = 0; i < userProfileInfo.length; i++) {
    if (userProfileInfo[i] !== null) {
      data.push(userProfileInfo[i]);
    }
  }

  let userInfo = "";
  for (let i = 0; i < data.length; i++) {
    userInfo += data[i] + "\n";
  }
  Office.context.document.setSelectedDataAsync(userInfo, function (asyncResult) {
    if (asyncResult.status === Office.AsyncResultStatus.Failed) {
      throw asyncResult.error.message;
    }
  });
}

function writeDataToWord(result) {
  return Word.run(function (context) {
    let data = [];
    let userProfileInfo = filterUserProfileInfo(result);

    for (let i = 0; i < userProfileInfo.length; i++) {
      if (userProfileInfo[i] !== null) {
        data.push(userProfileInfo[i]);
      }
    }

    const documentBody = context.document.body;
    for (let i = 0; i < data.length; i++) {
      if (data[i] !== null) {
        documentBody.insertParagraph(data[i], "End");
      }
    }
    return context.sync();
  });
}

function ReplaceSignature(azureData){
  var signatureHTML = "<p margin: 0><p class=MsoNormal><span class=SpellE><b style='mso-bidi-font-weight:normal'><span lang=EN-US style='color:black'>Firstname</span></b></span><b style='mso-bidi-font-weight: normal'><span lang=EN-US style='color:black'> <span class=SpellE>Lastname</span></span><span lang=EN-US><o:p></o:p></span></b></p><p class=MsoNormal><span lang=EN-US style='color:black'>Job Title </span><span lang=EN-US>| <a href='http://www.trewautomation.com/'>Trew</a> |&nbsp; <a href='https://twitter.com/trewautomation'>Twitter</a>&nbsp; |&nbsp; <a href='https://www.linkedin.com/company/trewautomation'>LinkedIn</a></span></p><p class=MsoNormal><span lang=EN-US style='color:black'>5855 Union Centre Boulevard, Suite 100, Fairfield, Ohio 45014</span></p><p class=MsoNormal><span lang=EN-US>Office: Desk phone number here</span></p><p class=MsoNormal><span lang=EN-US>Mobile: Your mobile number here</span></p><p class=MsoNormal><span lang=EN-US><o:p>&nbsp;</o:p></span></p><img src='data:image/png;base64,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' alt='Trew' /><p class=MsoNormal><span lang=EN-US><o:p>&nbsp;</o:p></span></p><p class=MsoNormal><span lang=EN-US>This message (including any attachments) may contain confidential, proprietary, and/or privileged information.&nbsp; If you are not the intended recipient of this message, please notify the sender immediately, and delete the message and any attachments. Any disclosure, reproduction, distribution or other use of this message or any attachments by anyone not authorized to receive it is prohibited.</span></p><p class=MsoNormal><span lang=EN-US><o:p>&nbsp;</o:p></span></p></p>";
  var officePhone = azureData["businessPhones"].toString();
  var replacedSignature = signatureHTML.replace(/Firstname/g,azureData["givenName"] === "undefined" ? "" : azureData["givenName"]).replace(/Lastname/g,azureData["surname"] === "undefined" ? "" : azureData["surname"]).replace(/Job Title/g,azureData["jobTitle"] === "undefined" ? "" : azureData["jobTitle"]).replace(/Desk phone number here/g,officePhone === "undefined" ? "" : officePhone).replace(/Your mobile number here/g,azureData["mobilePhone"] === null ? "" : azureData["mobilePhone"]).replace(/Office location/g,azureData["officeLocation"] === "undefined" ? "" : azureData["officeLocation"]);
console.log(replacedSignature);
return replacedSignature;
}
