using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data.Sql;
using System.Data.SqlClient;
using Microsoft.Office.Interop.Owc11;
using System.Configuration;
using System.Data;

public partial class excel : System.Web.UI.Page
{
    protected void Page_Load(object sender, EventArgs e)
    {
        //创建图表控件
        ChartSpace laySpace = new ChartSpace();
        //添加一个表容器
        SpreadsheetClass myexcel = new SpreadsheetClass();
        Worksheet mysheet = myexcel.ActiveSheet;
        //添加表标题
        myexcel.Cells[1, 1] = "籍贯";
        myexcel.Cells[1, 2] = "人数";
        //连接数据库
        String str = ConfigurationManager.ConnectionStrings["connection"].ConnectionString.ToString();
        SqlConnection con = new SqlConnection(str);
        con.Open();
        String sel = "select jiguan,count(jiguan) as number from xx group by jiguan";
        SqlDataAdapter adsa = new SqlDataAdapter(sel, con);
        DataSet adds = new DataSet();
        adsa.Fill(adds);
        if (adds.Tables[0].Rows.Count > 0)
        {
            for (int i = 0; i < adds.Tables[0].Rows.Count; i++)
            {
                mysheet.Cells[i + 2, 1] = adds.Tables[0].Rows[i][0].ToString();
                mysheet.Cells[i + 2, 2] = adds.Tables[0].Rows[i][1].ToString();
            }
            //导出表格
            myexcel.Export(Server.MapPath(".")+@"\test.xls", SheetExportActionEnum.ssExportActionOpenInExcel, SheetExportFormat.ssExportXMLSpreadsheet);
        }
        con.Close();     
    }
}