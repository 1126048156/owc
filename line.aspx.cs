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

public partial class line: System.Web.UI.Page
{
    protected void Page_Load(object sender, EventArgs e)
    {
        //创建ChartSpace对象来放置图表
        ChartSpace laySpace=new ChartSpace();
        //在ChartSpace对象中添加图表
        ChChart InsertChart=laySpace.Charts.Add(0);
        //指定绘制图表类型
        InsertChart.Type=ChartChartTypeEnum.chChartTypeLine;//折线图
        //指定图表是否需要图例标注
        InsertChart.HasLegend=false;
        InsertChart.HasTitle=true;//为图表添加标题
        InsertChart.Title.Caption = "员工信息表";
        //为x,y轴添加图示说明
        InsertChart.Axes[0].HasTitle=true;
        InsertChart.Axes[0].Title.Caption="籍贯";
        InsertChart.Axes[1].HasTitle=true;
        InsertChart.Axes[1].Title.Caption="人数";
        //连接数据库
        String str = ConfigurationManager.ConnectionStrings["connection"].ConnectionString.ToString();
        SqlConnection con = new SqlConnection(str);
        con.Open();
        String sel = "select jiguan,count(jiguan) as number from xx group by jiguan";
        SqlDataAdapter adsa = new SqlDataAdapter(sel, con);
        DataSet adds = new DataSet();
        adsa.Fill(adds);
        //为x，y轴指定特定字符串，以便显示数据
        string strXdata = String.Empty;
        string strYdata = String.Empty;
        if (adds.Tables[0].Rows.Count > 0)
        {
            for (int i = 0; i < adds.Tables[0].Rows.Count;i++)
            {
                strXdata = strXdata + adds.Tables[0].Rows[i][0].ToString() + "\t";
                strYdata = strYdata + adds.Tables[0].Rows[i][1].ToString() + "\t";
            }
        }
     //添加图表块
        InsertChart.SeriesCollection.Add(0);
     //设置图表的属性
        InsertChart.SeriesCollection[0].SetData(ChartDimensionsEnum.chDimCategories, (int)ChartSpecialDataSourcesEnum.chDataLiteral, strXdata);
        InsertChart.SeriesCollection[0].SetData(ChartDimensionsEnum.chDimValues, (int)ChartSpecialDataSourcesEnum.chDataLiteral,strYdata);        
        con.Close();
        laySpace.ExportPicture(Server.MapPath(".") + @"\temp.jpg", "jpg", 600, 450);
    }
}