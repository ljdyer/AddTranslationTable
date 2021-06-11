/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

// images references in the manifest
// import "../../assets/icon-16.png";
// import "../../assets/icon-32.png";
// import "../../assets/icon-80.png";

/* global document, Office */

Office.onReady((info) => {
    if (info.host === Office.HostType.Outlook) {
        document.getElementById("sideload-msg").style.display = "none";
        document.getElementById("app-body").style.display = "flex";
        document.getElementById("click-here").onclick = run;
    }
});

export async function run() {
    // Get a reference to the current message
    var item = Office.context.mailbox.item;
    var showHtml = document.getElementById("showHtml");

    item.getSelectedDataAsync(
        Office.CoercionType.Html,
        {},
        function (asyncResult) {
            if (asyncResult.status == Office.AsyncResultStatus.Failed) {
                showHtml.innerHTML = "Failed.";
            }
            else {
                let htmlDataRaw = escapeHTML(processHtmlData(asyncResult.value.data));
                showHtml.innerHTML = htmlDataRaw;
            }
        }
    );
}


function tabulateSrcLeft(content) {
    let result =
        `<table style="width:100%; border-width: 1px; border-color: black; border-collapse: collapse">
    <tr style="border-style: solid; border-width: 1px; border-color: black;">
      <td style="width:50%; border-width: 1px; border-style: solid; border-color: black;">${content}</td>
      <td style="width:50%; border-width: 1px; border-style: solid; border-color: black;"></td>
    </tr>
  </table>`;
    return result;
}


function tabulateSrcRight(content) {
    let result =
        `<table style="width:100%; border-width: 1px; border-color: black; border-collapse: collapse">
    <tr style="border-style: solid; border-width: 1px; border-color: black;">
      <td style="width:50%; border-width: 1px; border-style: solid; border-color: black;"></td>
      <td style="width:50%; border-width: 1px; border-style: solid; border-color: black;">${content}</td>
    </tr>
  </table>`;
    return result;
}


function processHtmlData(htmlDataRaw) {
    let result = removeExtraLBs(removeOuterDiv(getBody(htmlDataRaw)));
    return result;
}


function processHtmlDataLeavingLBs(htmlDataRaw) {
    let result = removeOuterDiv(getBody(htmlDataRaw));
    return result;
}


function getBody(content) {
    // getSelectedDataAsync wraps data in a body tag which we do not need, so remove that tag using RegEx
    var regexp = /<body.*>([^]*)<\/body>/g;
    let match = regexp.exec(content);
    let result = match[1];
    return result;
}


function removeOuterDiv(content) {
    // getSelectedDataAsync wraps data in a div tag which we do not need, so remove that tag using RegEx
    var regexp = /<div.*>([^]*)<\/div>/g;
    let match = regexp.exec(content);
    let result = match[1];
    return result;
}


function removeExtraLBs(content) {
    // Remove any extra <p ***>&nbsp;</p> added by getSelectedDataAsync
    var regexp = /<p((?! class=MsoPlainText).)[^>]*>&nbsp;<\/p[^>]*>/g;
    let result = content.replace(regexp, "");
    return result;
}


function escapeHTML(content) {
    // For debugging
    let result = content.replace(/</g, "&lt;").replace(/&/g, "&amp;");
    return result;
}