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

public partial class pie : System.Web.UI.Page
{
    protected void Page_Load(object sender, EventArgs e)
    {
        //创建图表控件
        ChartSpace myspace = new ChartSpace();
        //添加一个图表对象
        ChChart mychart = myspace.Charts.Add(0);
        //设置图表类型为柱形
        mychart.Type = ChartChartTypeEnum.chChartTypePie;
        //设置图表的相关属性
        mychart.HasLegend = true;//添加图列
        mychart.HasTitle = true;//添加主题
        mychart.Title.Caption = "员工信息图表";//设置主题内容
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
            string strDataName = "";
            string strData = "";
            //添加图表块
            mychart.SeriesCollection.Add(0);
            //添加图表数据
            for (int j = 0; j < adds.Tables[0].Rows.Count; j++)
            {
                if (j == adds.Tables[0].Rows.Count - 1)
                {
                    strDataName += adds.Tables[0].Rows[j][0].ToString();
                }
                else
                {
                    strDataName += adds.Tables[0].Rows[j][0].ToString() + "\t";
                }
                strData += adds.Tables[0].Rows[j][1].ToString() + "\t";
            }
                //设置图表的属性
                mychart.SeriesCollection[0].SetData(ChartDimensionsEnum.chDimCategories, (int)ChartSpecialDataSourcesEnum.chDataLiteral,strDataName);
                mychart.SeriesCollection[0].SetData(ChartDimensionsEnum.chDimValues, (int)ChartSpecialDataSourcesEnum.chDataLiteral, strData);
            //设置百分比
            ChDataLabels labels=mychart.SeriesCollection[0].DataLabelsCollection.Add();
            labels.HasPercentage=true;
            
        }
        con.Close();
        myspace.ExportPicture(Server.MapPath(".") + @"\temp.jpg", "jpg", 600, 450);
    }
}