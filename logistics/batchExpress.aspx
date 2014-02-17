<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="batchExpress.aspx.cs" Inherits="batchExpress.batchExpress" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8"/>
    <title></title>
    <style>
        table {
        width:auto;height:1px;table-layout:auto;border-collapse:collapse;margin-left:20px;border:1px solid black;
        }
        td, th {
        width:auto;height:1px;overflow:hidden;visibility:visible;border:1px solid black;padding:5px;background:#E1F5A9;text-align:center;vertical-align:middle;text-indent:5px;
        }
    </style>
</head>
<body>
    <form id="form1" runat="server">
    <div>
        <asp:FileUpload ID="fulExcel" runat="server" Width="400" />
        <asp:Button ID="Button_Import" runat="server" Text="导入" Width="100" Height="30" OnClick="Button_Import_Click" />
        <br />
        <div id="divExcelData" runat="server"></div>
    </div>
    </form>
</body>
</html>
