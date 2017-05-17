using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

using ExportToExcel.Models;
using System.Data;
using System.Reflection;
using ClosedXML.Excel;
using System.IO;

namespace ExportToExcel
{
    public partial class WebForm1 : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {
            string a;
            string b; 
        }

        protected void Button1_Click(object sender, EventArgs e)
        {
            List<people> peopleList = new List<people>();
            peopleList = GetReports(1);
            DataTable dt = new DataTable();
            PropertyInfo[] Props = typeof(people).GetProperties();
          //  Props = Props.Where(x => (x.Name != "Seconds" && x.Name != "ExecutionDate" && x.Name != "ClientId" && x.Name != "ClientJobId")).ToArray();

            foreach (PropertyInfo prop in Props)
            {
                //Setting column names as Property names
                dt.Columns.Add(prop.Name);
            }
            foreach (people item in peopleList)
            {
                var values = new object[dt.Columns.Count];
                for (int i = 0; i < dt.Columns.Count; i++)
                {
                    values[i] = Props[i].GetValue(item, null);
                }
                dt.Rows.Add(values);
            }


            #region closedxmltype
            using (XLWorkbook xlwb = new XLWorkbook())
            {
                string stringvalue = "attachment; filename=test.xlsx";
                IXLWorksheet ws = xlwb.AddWorksheet(dt, "ws_1");
                ws.ShowGridLines = true;
                ws.TabColor = XLColor.BabyBlue;
                Response.Clear();
                Response.Buffer = true;
                Response.Charset = "";
                Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                Response.AddHeader("content-disposition", stringvalue);
                using (MemoryStream ms = new MemoryStream())
                {
                    xlwb.SaveAs(ms);
                    ms.WriteTo(Response.OutputStream);
                    Response.Flush();
                    Response.End();
                }
            }
            #endregion

            //#region without using Interop
            //string stringvalue = "attachment; filename=test.xls";
            //var cd = new System.Net.Mime.ContentDisposition
            //{
            //    FileName = stringvalue,
            //    Inline = false,
            //};
            //HttpContext.Current.Response.Clear();
            //HttpContext.Current.Response.ClearContent();
            //HttpContext.Current.Response.ClearHeaders();
            //HttpContext.Current.Response.Buffer = true;
            //HttpContext.Current.Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
            //HttpContext.Current.Response.Write(@"<!DOCTYPE HTML PUBLIC ""-//W3C//DTD HTML 4.0 Transitional//EN"">");
            //HttpContext.Current.Response.AddHeader("Content-Disposition", cd.ToString());

            //HttpContext.Current.Response.Charset = "utf-8";
            //HttpContext.Current.Response.ContentEncoding = System.Text.Encoding.GetEncoding("windows-1250");
            ////sets font
            //HttpContext.Current.Response.Write("<font style='font-size:10.0pt; font-family:Calibri;'>");
            //HttpContext.Current.Response.Write("<BR><BR><BR>");
            ////sets the table border, cell spacing, border color, font of the text, background, foreground, font height
            //HttpContext.Current.Response.Write("<Table border='1' borderColor='#000000'> <TR>");
            //string tab = "";
            //foreach (DataColumn dc in dt.Columns)
            //{      //write in new column
            //    HttpContext.Current.Response.Write("<Td>");
            //    //Get column headers  and make it as bold in excel columns
            //    HttpContext.Current.Response.Write("<B>");
            //    HttpContext.Current.Response.Write(tab + dc.ColumnName);
            //    tab = "\t";
            //    HttpContext.Current.Response.Write("</B>");
            //    HttpContext.Current.Response.Write("</Td>");
            //}
            //HttpContext.Current.Response.Write("</TR>");
            //foreach (DataRow row in dt.Rows)
            //{//write in new row
            //    tab = "";
            //    HttpContext.Current.Response.Write("<TR>");
            //    for (int i = 0; i < dt.Columns.Count; i++)
            //    {
            //        HttpContext.Current.Response.Write("<Td>");
            //        HttpContext.Current.Response.Write(tab + row[i].ToString());
            //        tab = "\t";
            //        HttpContext.Current.Response.Write("</Td>");
            //    }

            //    HttpContext.Current.Response.Write("</TR>");
            //}
            //HttpContext.Current.Response.Write("</Table>");
            //HttpContext.Current.Response.Write("</font>");
            //HttpContext.Current.Response.Flush();
            //HttpContext.Current.Response.End();
            //#endregion

        }
        public List<people> GetReports(int duid)
        {
            List<people> peopleList = new List<people>();
            people p = new people { UserId=1, CareerLevel="ASE",Du="HRPortal ", EnterpriseId="k.basavaraj.mudhol" , Project="HRL" , Supervisor="kiranaa"  };
            people p1 = new people { UserId = 2, CareerLevel = "SE", Du = "Portal ", EnterpriseId = "k.basavaraj.mudhol", Project = "HL", Supervisor = "kiransaa" };

            peopleList.Add(p);
            peopleList.Add(p1);
            return peopleList;  
        }
        //
        //private ActionResult ExportToExcel(List<People> peopleList)
        //{
        //    DataTable dt = new DataTable();
        //    PropertyInfo[] Props = typeof(People).GetProperties();
        //    //  Props = Props.Where(x => (x.Name != "Seconds" && x.Name != "ExecutionDate" && x.Name != "ClientId" && x.Name != "ClientJobId")).ToArray();
        //    foreach (PropertyInfo prop in Props)
        //    {
        //        //Setting column names as Property names
        //        dt.Columns.Add(prop.Name);
        //    }
        //    foreach (People item in peopleList)
        //    {
        //        var values = new object[dt.Columns.Count];
        //        for (int i = 0; i < dt.Columns.Count; i++)
        //        {
        //            values[i] = Props[i].GetValue(item, null);
        //        }
        //        dt.Rows.Add(values);
        //    }
        //    try
        //    {
        //        #region closedxmltype
        //        using (XLWorkbook xlwb = new XLWorkbook())
        //        {
        //            string stringvalue = "attachment; filename=test.xlsx";
        //            IXLWorksheet ws = xlwb.AddWorksheet(dt, "ws_1");
        //            ws.ShowGridLines = true;
        //            ws.TabColor = XLColor.Black;
        //            Response.Clear();
        //            Response.Buffer = true;
        //            Response.Charset = "";
        //            Response.ContentType = "application/vnd.ms-excel";
        //            Response.AddHeader("content-disposition", stringvalue);
        //            using (MemoryStream ms = new MemoryStream())
        //            {
        //                xlwb.SaveAs(ms);
        //                ms.WriteTo(Response.OutputStream);
        //                Response.Flush();
        //                Response.End();
        //            }
        //        }
        //        #endregion
        //    }
        //    catch (Exception ex)
        //    {
        //        throw (ex);
        //    }

        //    return View();


        //}
    }

}