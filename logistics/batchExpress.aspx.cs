using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using Top.Api;
using Top.Api.Domain;
using Top.Api.Response;
using Top.Api.Request;
using System.Web.Script.Serialization;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Text;

namespace batchExpress
{
    public partial class batchExpress : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {

        }


        //上传文件
        protected void Button_Import_Click(object sender, EventArgs e)
        {
            string fileName = DateTime.Now.ToString("yyMMddhhmmssfffff") + Path.GetExtension(fulExcel.FileName); //excel 第一行是字段名，另外数字前需加'转换成字符串
            fileName = Path.Combine(Server.MapPath("~/tmpexcel/"), fileName); 
            fulExcel.SaveAs(fileName);

            DataSet excelData = GetExcelData(fileName);
            DataTable dt = excelData.Tables[0];

            StringBuilder excelTable = new StringBuilder();
            excelTable.Append("<div>没有数据</div>");

            if (dt.Rows.Count != 0)
            {
                string result = string.Empty;
                List<string> dataHouse = new List<string>(); //临时的数据仓库
                excelTable.Remove(0, excelTable.Length);
                excelTable.Append("<div><table><th>订单号</th><th>运单号</th><th>快递公司</th><th>结果</th>");
                for (int i = 0; i < dt.Rows.Count; i++) //按行读取数据
                {
                    excelTable.Append("<tr>");
                    for (int j = 0; j < dt.Columns.Count; j++) //按列读取数据
                    {
                        excelTable.Append("<td>" + dt.Rows[i][j] + "</td>");
                        dataHouse.Add(dt.Rows[i][j].ToString()); //将数据放入仓库中
                    }

                    result = responseFromAPI(Convert.ToInt64(dataHouse[0]), dataHouse[1], dataHouse[2]);//调用API方法，返回结果
                    excelTable.Append("<td>" + result + "</td></tr>"); //将结果加在表的最后一行

                    dataHouse.RemoveRange(0, dataHouse.Count); //清空List
                }
                excelTable.Append("</table></div>");
            }

            divExcelData.InnerHtml = excelTable.ToString();
        }

        //读取excel文件，取得数据
        protected DataSet GetExcelData(string str)
        {
            string strCon = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + str + ";Extended Properties='Excel 12.0;HDR=YES'";
            OleDbConnection myConn = new OleDbConnection(strCon);
            string strCom = " SELECT * FROM [Sheet1$]";
            myConn.Open();
            OleDbDataAdapter myCommand = new OleDbDataAdapter(strCom, myConn);
            DataSet myDataSet = new DataSet();
            myCommand.Fill(myDataSet, "[Sheet1$]");
            myConn.Close();
            return myDataSet;
        }

        //调用API
        protected string responseFromAPI(long? dataTid, string dataOutSid, string dataCompanyCode)
        {
            string url = "http://gw.api.tbsandbox.com/router/rest";//沙箱环境调用地址,
            string appkey = "1021723219";
            string appsecret = "sandboxa85cf30610ff927555aa6e825";
            string sessionkey = "6100e236b608af9b78530e8fd2b236b1b63be8dab8ffff22074082786";   //如沙箱测试帐号sandbox_c_1授权后得到的sessionkey

            /*
            //正式环境需要设置为:
            string url = "http://gw.api.taobao.com/router/rest";
            string appkey = "21723219";
            string appsecret = "789a305a85cf30610ff927555aa6e825";
            string sessionkey = "test";   //如沙箱测试帐号sandbox_c_1授权后得到的sessionkey
            */

            ITopClient client = new DefaultTopClient(url, appkey, appsecret,"json");
            LogisticsOfflineSendRequest req = new LogisticsOfflineSendRequest();
            req.Tid = dataTid;
            req.OutSid = dataOutSid;
            req.CompanyCode = dataCompanyCode;
            LogisticsOfflineSendResponse response = client.Execute(req, sessionkey);

            return response.Body;
        }

    }
}