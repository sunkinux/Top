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
        <p>
            session: <input id="access_token" type="text" runat="server" style="width:500px"/><span style="color:red">*</span>
            登陆<a href="http://api.taobao.com/apitools/sessionPage.htm" target="_blank">授权页面</a>，填写 appKey：21723219，获得 session
        </p>
        <p>            
            <asp:FileUpload ID="fulExcel" runat="server" Width="400" />
        </p>
        <p>
            <asp:Button ID="Button_Import" runat="server" Text="提交批量处理" Width="100" Height="30" OnClick="Button_Import_Click" />
        </p>
        
        <div id="divExcelData" runat="server"></div>
    </div>
    </form>
</body>
</html>
