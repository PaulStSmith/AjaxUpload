/*
 '**************************************************
 ' FILE      : AjaxUpload.js
 ' AUTHOR    : Paulo.Santos
 ' CREATION  : 10/25/2007 8:43:29 PM
 ' COPYRIGHT : Copyright © 2007
 '             PJ on Development
 '             All Rights Reserved.
 '
 ' Description:
 '       Functions for asynchronous file upload
 '
 ' Change log:
 ' 0.1   10/25/2007 8:43:29 PM
 '       Paulo.Santos
 '       Created.
 '***************************************************
 */

/* Function  : __AjaxUploadInit()
 * Objective : Initializes the client DOM.
 */
function __AjaxUploadInit(ID)
{
    var obj = document.getElementById(ID);
    if (!obj)
        return;

    /*
     * Set the properties
     */
    if (!obj.name             && obj.attributes["name"])             obj.name             = obj.attributes["name"].nodeValue;
    if (!obj.displayIcon      && obj.attributes["displayIcon"])      obj.displayIcon      = obj.attributes["displayIcon"].nodeValue;
    if (!obj.controlLocation  && obj.attributes["controlLocation"])  obj.controlLocation  = obj.attributes["controlLocation"].nodeValue;
    if (!obj.removeMessage    && obj.attributes["removeMessage"])    obj.removeMessage    = obj.attributes["removeMessage"].nodeValue;
    if (!obj.errorMessage     && obj.attributes["errorMessage"])     obj.errorMessage     = obj.attributes["errorMessage"].nodeValue;
    if (!obj.confirmMessage   && obj.attributes["confirmMessage"])   obj.confirmMessage   = obj.attributes["confirmMessage"].nodeValue;
    if (!obj.postbackFunction && obj.attributes["postbackFunction"]) obj.postbackFunction = obj.attributes["postbackFunction"].nodeValue;

    if (typeof(obj.displayIcon) == "string") obj.displayIcon = parseInt(obj.displayIcon);

    /*
     * Get the form
     */
    var oForm = obj;
    while (oForm && oForm.tagName != "FORM")
        oForm = oForm.parentNode;

    /*
     * No form was found.
     */
    if (!oForm)
        return;

    obj.ownerForm = oForm;
    obj.lblFileName = document.getElementById(ID + "_lblFileName");
    obj.uploadingPanel = document.getElementById(ID + "_uploading");
    obj.tableFiles = document.getElementById(obj.id + "_files");
    obj.tableFiles.ownerControl = obj;

    /*
     * Create the iFrame
     */
    var oDivFrame = document.createElement("DIV");
    var iFrameID = "uploadIFrame_" + Math.floor(Math.random() * 9999);
    oDivFrame.innerHTML = "<iframe style='display:none' onload='UploadFileComplete(event)' src='about:blank' id='" + iFrameID + "' name='" + iFrameID + "'></iframe>";
    oForm.appendChild(oDivFrame);
    var iFrame = oDivFrame.childNodes[0];
    iFrame.associatedUploadControl = obj;
    obj.iFrame = iFrame;

    /*
     * Store the new values for the form
     */
    oForm.__newValues = new Object();
    oForm.__newValues.enctype = "multipart/form-data";
    oForm.__newValues.method  = "POST";
    oForm.__newValues.action  = obj.uploadPage ? obj.uploadPage : document.location.href;
    oForm.__newValues.target  = iFrameID;

    /*
     * Add the input type=file
     */
    var oSpan = document.createElement("SPAN");
    oSpan.id = obj.id + "_span";
    obj.appendChild(oSpan);
    obj.inputFileContainer = oSpan;
    AddInputFile(ID);

    /*
     * Control initialized
     */
    obj.isAjaxUpload = true;
}

/* Function  : UploadFileInternal()
 * Objective : Uploads a form
 */
function UploadFileInternal(controlID, postArg)
{
    if (!controlID) return;
    if (!postArg)   postArg = "";
    var oControl =  document.getElementById(controlID);
    if (!oControl)  return;

    var oInputFile = oControl.inputFile;
    var oForm      = oControl.ownerForm;

    if (oForm.__newValues)
    {
        try
        {
            /*
             * Store the current Values
             */
            if (!oForm.__oldValues)
            {
                oForm.__oldValues = new Object();
                oForm.__oldValues.enctype = oForm.enctype;
                oForm.__oldValues.encoding = oForm.encoding;
                oForm.__oldValues.method  = oForm.method ;
                oForm.__oldValues.action  = oForm.action ;
                oForm.__oldValues.target  = oForm.target ;
            }

            /*
             * Update to the new Values
             */
            oForm.setAttribute("enctype",oForm.__newValues.enctype);
            oForm.setAttribute("encoding",oForm.__newValues.enctype);
            oForm.setAttribute("method", oForm.__newValues.method);
            oForm.setAttribute("action", oForm.__newValues.action);
            oForm.setAttribute("target", oForm.__newValues.target);

            if (document.getElementById("__EVENTTARGET"))
                oForm.__currentEventTarget = document.getElementById("__EVENTTARGET").value
            else
                oForm.__currentEventTarget = null;
            if (document.getElementById("__EVENTARGUMENT"))
                oForm.__currentEventArgument = document.getElementById("__EVENTARGUMENT").value
            else
                oForm.__currentEventArgument = null;

            /*
             * Submit the form
             */
            eval(oInputFile.uploadControl.postbackFunction + "('" + oControl.name + "', '" + postArg + "')");
        }
        catch(e)
        {
        }
        finally
        {
            /*
             * Return the old values
             */
            oForm.setAttribute("enctype",oForm.__oldValues.enctype);
            oForm.setAttribute("encoding",oForm.__oldValues.encoding);
            oForm.setAttribute("method", oForm.__oldValues.method);
            oForm.setAttribute("action", oForm.__oldValues.action);
            oForm.setAttribute("target", oForm.__oldValues.target);

            document.getElementById("__EVENTTARGET").value = oForm.__currentEventTarget;
            document.getElementById("__EVENTARGUMENT").value = oForm.__currentEventArgument;

            oForm.__currentEventTarget = null;
            oForm.__currentEventArgument = null;
        }
    }
}

/* Function  : AddInputFile
 * Objective : Add the input type file in the document.
 */
function AddInputFile(ID)
{
    var obj = document.getElementById(ID);
    var oSpan = obj.inputFileContainer;
    var inputFileID = (obj.id + "_file");
    oSpan.innerHTML = "<input type='file' id='" + inputFileID + "' name='" + inputFileID + "' />";

    var inputFile = document.getElementById(inputFileID);
    inputFile.uploadControl = obj;
    inputFile.ownerForm = obj.ownerForm;
    inputFile.onchange = function()
    {
        if (this.value && this.value != "")
        {
            inputFile.uploadControl.lblFileName.innerHTML = this.value.substr(this.value.lastIndexOf("\\") + 1);
            inputFile.uploadControl.uploadingPanel.style.display = "";
            window.setTimeout("UploadFileInternal('" + this.uploadControl.id + "')", 2000);
            this.style.display = "none";
        }
    };
    obj.inputFile = inputFile;
}
/* Function  : RemoveInputFile
 * Objective : Removes the input type=file when submiting the form.
 */
function RemoveInputFile(ID, message)
{
    var obj = document.getElementById(ID);
    if (!obj)
        return;
    var oSpan = obj.inputFileContainer;
    var oForm = obj.ownerForm;
    if (oForm.target != obj.iFrame.id)
    {
        oSpan.removeChild(oSpan.childNodes[0]);
        if (message)
            oSpan.innerHTML = "<a href='javascript:AddInputFile(\"" + ID + "\")'>" + message + "</a>";
    }
}

/* Function  : UploadFileComplete
 * Objective : Called when the upload is complete
 */
function UploadFileComplete(e)
{
    e = e ? e : window.event;
    var obj = e.srcElement ? e.srcElement : e.target ? e.target : e.currentTarget;

    var doc = null;
    try
    {
        doc = obj.contentDocument ? obj.contentDocument : (obj.contentWindow.document ? obj.contentWindow.document : window.frames[obj.id].document);
    }
    catch(e)
    {
        doc = null;
    }

    if (!doc || (doc && doc.location.href == "about:blank"))
        return;

    /*
     * Call the onUploadComplete
     */
    var uploadControl = obj.associatedUploadControl;
    if (uploadControl)
    {
        uploadControl.uploadingPanel.style.display = "none";

        var fileInfo = new Object();
        fileInfo.id          = !doc ? null : doc.getElementById("ID")          ? doc.getElementById("ID").value          : null;
        fileInfo.fileName    = !doc ? ''   : doc.getElementById("fileName")    ? doc.getElementById("fileName").value    : '';
        fileInfo.fileSize    = !doc ? 0    : doc.getElementById("fileSize")    ? doc.getElementById("fileSize").value    : 0;
        fileInfo.type        = !doc ? ''   : doc.getElementById("ContentType") ? doc.getElementById("ContentType").value : '';
        fileInfo.url         = !doc ? ''   : doc.getElementById("url")         ? doc.getElementById("url").value         : '';
        fileInfo.removed     = !doc ? false: doc.getElementById("removed")     ? doc.getElementById("removed").value     : false;
        fileInfo.donotremove = !doc ? false: doc.getElementById("donotremove") ? doc.getElementById("donotremove").value : false;
        fileInfo.message     = !doc ? null : doc.getElementById("message")     ? doc.getElementById("message").value     : null;
        fileInfo.fileName    = fileInfo.fileName ? fileInfo.fileName : uploadControl.inputFile.value.substr(uploadControl.inputFile.value.lastIndexOf("\\") + 1);

        /*
         * The file was removed from the list
         */
        if (fileInfo.removed)
        {
            if (!fileInfo.donotremove)
                uploadControl.tableFiles.deleteRow(document.getElementById(fileInfo.id).rowIndex)
            else
                if (fileInfo.message && fileInfo.message != "") alert(fileInfo.message);

            return;
        }

        /*
         * Create a new line on the file list
         */
        var oTR = uploadControl.tableFiles.insertRow(uploadControl.tableFiles.rows.length);
        var oTD = oTR.insertCell(oTR.cells.length);

        /*
         * Check if the file was accepted by the server
         */
        if (fileInfo.id)
        {
            /*
             * The file was accepted
             */
            oTR.id = fileInfo.id;
            if (uploadControl.displayIcon)
            {
                var img = document.createElement("IMG");
                img.src = uploadControl.controlLocation + "/icon.aspx?extension=" + fileInfo.fileName.substr(fileInfo.fileName.lastIndexOf(".") + 1);
                oTD.appendChild(img);
                oTD = oTR.insertCell(oTR.cells.length);
            }
            if (fileInfo.url)
            {
                var a = document.createElement("A");
                a.href = fileInfo.url;
                a.innerHTML = fileInfo.fileName;
                a.target="_blank";
                oTD.appendChild(a);
            }
            else
                oTD.innerHTML = fileInfo.fileName;
            oTD = oTR.insertCell(oTR.cells.length);
            var img = document.createElement("IMG");
            img.width = 16;
            img.height = 16;
            img.border = 0;
            img.style.cursor = "pointer";
            img.src = uploadControl.controlLocation + "/remove.png";
            img.title = uploadControl.removeMessage;
            img.onclick = function (e)
            {
                RemoveFile(e, this, fileInfo.id, fileInfo.fileName, fileInfo.fileSize, fileInfo.type);
            }
            oTD.appendChild(img);
            /*
             * Add the file properties for PostBack
             */
            var inputHidden = null;
            /*
             * FileID
             */
            inputHidden = document.createElement("INPUT");
            inputHidden.type = "hidden";
            inputHidden.id = uploadControl.id + "_fileID";
            inputHidden.name = uploadControl.name + "$fileID";
            inputHidden.value = fileInfo.id + "|" + fileInfo.fileName + "|" + fileInfo.type + "|" + fileInfo.url + "|" + fileInfo.fileSize;
            oTD.appendChild(inputHidden);
        }
        else
        {
            if (uploadControl.displayIcon)
                oTD = oTR.insertCell(oTR.cells.length);
            oTD.innerHTML = fileInfo.fileName;
            oTD = oTR.insertCell(oTR.cells.length);
            var img = document.createElement("IMG");
            img.width = 16;
            img.height = 16;
            img.uploadControl = uploadControl;
            img.border = 0;
            img.style.cursor = "pointer";
            img.src = uploadControl.controlLocation + "/warning.png";
            img.title = fileInfo.message ? fileInfo.message : uploadControl.errorMessage;
            img.onclick = function ()
            {
                if (confirm(this.title + "\n\n" + this.uploadControl.removeMessage))
                {
                    var obj = this;
                    while (obj.tagName != "TR")
                        obj = obj.parentNode;

                    this.uploadControl.tableFiles.deleteRow(obj.rowIndex);
                }
            };
            oTD.appendChild(img);
        }

        /*
         * Reload input type=file
         */
        RemoveInputFile(uploadControl.id, "");
        AddInputFile(uploadControl.id);
    }
}

/*
 *
 */
function RemoveFile(e, obj, fileID, fileName, fileSize, type)
{
    e = e ? e : window.event;
    var table = obj;
    while (table.tagName != "TABLE")
        table = table.parentNode;

    var uploadControl = table.ownerControl;
    if (confirm(uploadControl.confirmMessage))
        UploadFileInternal(uploadControl.id,
                   fileID + "|" +
                   fileName + "|" +
                   fileSize + "|" +
                   type);
}