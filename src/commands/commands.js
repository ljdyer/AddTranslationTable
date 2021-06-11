var btnEvent;
var item;
var srcSide;

// Strings
const initSuccessMsg = "Add Translation Table add-in initialized.";
const addingLeftMsg = "Attempting to add a translation table (source on left).";
const addingRightMsg = "Attempting to add a translation table (source on right).";
const gettingMsg = "Getting selection.";
const getFailedMsg = "getSelectedDataAsync failed. You may be composing in plain text format. Select 'Format Text' > 'HTML' and try again."
const setFailedMsg = "setSelectedDataAsync failed. You may be composing in plain text format. Select 'Format Text' > 'HTML' and try again."
const successMsg = "Table added successfully."


Office.initialize = function (reason) {
  // The initialize function must be run each time a new page is loaded.
  showError(initSuccessMsg);
};


function showError(error) {
  Office.context.mailbox.item.notificationMessages.replaceAsync('addtable-msg', {
    type: 'errorMessage',
    message: error
  }, function (result) {
  });
}


function addTableLeft(event) {
  btnEvent = event;
  showError(addingLeftMsg);
  srcSide = "left";
  addTable();
}


function addTableRight(event) {
  btnEvent = event;
  showError(addingRightMsg);
  srcSide = "right";
  addTable();
}


function addTable() {
  item = Office.context.mailbox.item;
  showError(gettingMsg);
  item.getSelectedDataAsync(
    Office.CoercionType.Html,
    {},
    function (asyncResult) {
      if (asyncResult.status == Office.AsyncResultStatus.Failed) {
        showError(getFailedMsg);
        btnEvent.completed();
      }
      else {
        let htmlDataRaw = asyncResult.value.data;
        let srcContent = prepareSrcContent(htmlDataRaw);
        setSelection(srcContent, Office.CoercionType.Html);
      }
    }
  );
}


function prepareSrcContent(htmlDataRaw){
  let htmlDataProcessed = processHtmlData(htmlDataRaw);
  let tabulated = "";
  if (srcSide == "left") {
    tabulated = tabulateSrcLeft(htmlDataProcessed);
  }
  else {
    tabulated = tabulateSrcRight(htmlDataProcessed);
  }
  return tabulated;
}


function setSelection(content, cType) {
  item.body.setSelectedDataAsync(
    content,
    {
      coercionType: cType,
      asyncContext: { var3: 1, var4: 2 }
    },
    function (asyncResult) {
      if (asyncResult.status == Office.AsyncResultStatus.Failed) {
        showError(setFailedMsg);
        btnEvent.completed();
      }
      else {
        showError(successMsg);
        btnEvent.completed();
      }
    }
  );
}


function getGlobal() {
  return (typeof self !== "undefined") ? self :
    (typeof window !== "undefined") ? window :
      (typeof global !== "undefined") ? global :
        undefined;
}


function tabulateSrcLeft(content) {
  let result =
    `<table style="width:100%; border-width: 1px; border-color: black; border-collapse: collapse">
    <tr style="border-style: solid; border-width: 1px; border-color: black;">
      <td style="width:50%; border-width: 1px; border-style: solid; border-color: black; vertical-align: top;">${content}</td>
      <td style="width:50%; border-width: 1px; border-style: solid; border-color: black; vertical-align: top;"></td>
    </tr>
  </table>`;
  return result;
}


function tabulateSrcRight(content) {
  let result =
    `<table style="width:100%; border-width: 1px; border-color: black; border-collapse: collapse">
    <tr style="border-style: solid; border-width: 1px; border-color: black;">
      <td style="width:50%; border-width: 1px; border-style: solid; border-color: black; vertical-align: top;"></td>
      <td style="width:50%; border-width: 1px; border-style: solid; border-color: black; vertical-align: top;">${content}</td>
    </tr>
  </table>`;
  return result;
}


function processHtmlData(htmlDataRaw) {
  let result = removeExtraLBs(removeOuterDiv(getBody(htmlDataRaw)));
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
  // Remove any extra <p ***>&nbsp;</p> added by getSelectedDataAsync, except in the case of plain text
  var regexp = /<p((?! class=MsoPlainText).)[^>]*>&nbsp;<\/p[^>]*>/g;
  let result = content.replace(regexp, "");
  return result;
}


function escapeHTML(content) {
  // For debugging
  let result = content.replace(/</g, "&lt;").replace(/&/g, "&amp;");
  return result;
}


var g = getGlobal();

g.addTableLeft = addTableLeft;
g.addTableRight = addTableRight;