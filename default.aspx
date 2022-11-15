<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="Default.aspx.vb" Inherits="AjaxUpload._Default" %>
<%@ Register Src="ajaxUpload.ascx" TagName="ajaxUpload" TagPrefix="uc1" %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>Async Upload</title>
</head>
<body>
    <form id="form1" runat="server">
        <h1>Ajax Upload Sample</h1>
            <br />
            <uc1:ajaxUpload id="AjaxUpload1" runat="server" />
    </form>
</body>
</html>
