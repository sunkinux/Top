<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="batchExpress.aspx.cs" Inherits="batchExpress.batchExpress" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8"/>
    <title></title>
<%--    <script>
        function sessionValidate() {
            if（form1.access_token.value.length === 0){
                alert("请填入 session！");
                form1.access_token.focus();
                return false;
            }
            return true;
        }
    </script>--%>
    <style>
        body {
        text-align:center;
        }
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
            session: <input id="sessionKey" type="text" runat="server" style="width:500px"/><span style="color:red">*</span>
            <br />（登陆<a href="http://api.taobao.com/apitools/sessionPage.htm" target="_blank">授权页面</a>，填写 appKey：21723219，获得 session）
        </p>
        <p>            
            <asp:FileUpload ID="fulExcel" runat="server" Width="400" />
        </p>
        <p>
            <asp:Button ID="Button_Import" runat="server" Text="提交批量处理" Width="100" Height="30" OnClick="Button_Import_Click" />
            <br />支持 .xls, .xlsx 两种格式；请将数据放置在默认的 Sheet1 中，第一行的前三列请分别填入列名：“订单号”、“运单号”、“快递公司 code”<br />（注意：淘宝的订单号是一串很长的数字，需要在数字前加上单引号，将其转为文本，否则程序无法正确识别）
        </p>
        
        <div id="divExcelData" runat="server"></div>
    </div>
    </form>
</body>
</html>
