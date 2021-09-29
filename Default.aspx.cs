using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Text;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

namespace ExportToExcel
{
    public partial class _Default : Page
    {
        string DB = ConfigurationManager.ConnectionStrings["ExportExcelDB"].ConnectionString;
        protected void Page_Load(object sender, EventArgs e)
        {
            if (!IsPostBack)
            {
                LoadResultTable();
            }
        }
        void LoadResultTable()
        {
            try
            {
                using (SqlConnection conn = new SqlConnection(DB))
                {
                    string com = "Select * from ExamResult";
                    SqlDataAdapter adpt = new SqlDataAdapter(com, conn);
                    DataTable dt = new DataTable();
                    adpt.Fill(dt);



                    StringBuilder sb = new StringBuilder();

                    sb.Append("<div>");

                    sb.Append("<div style='float:left;width:5%;text-align:center;'>");
                    sb.Append("<img src='images/logo.png' alt='Logo' style='width:100%;'/>");
                    sb.Append("</div>");

                    sb.Append("<div style='float:right;width:95%;text-align:center;'>");

                    sb.Append("<div style='font-weight:bold; font-size:xx-large; text-align:center;'>");
                    sb.Append("VIKASH GROUP OF INSTITUTIONS");
                    sb.Append("</div>");

                    sb.Append("<div style='font-weight:bold; font-size:medium; text-align:center;'>");
                    //sb.Append("ANSWER KEYS OF " + Session["_examname"].ToString().ToUpper() + ", " + Session["_papername"].ToString().ToUpper());
                    sb.Append("</div>");

                    sb.Append("</div>");


                    sb.Append("<table class='table table-sm table-bordered'>");
                    sb.Append("<thead>");
                    sb.Append("<tr>");
                    sb.Append("<th rowspan='2'>ROLL NO</th>");
                    sb.Append("<th rowspan='2'>CAMPUS</th>");
                    sb.Append("<th rowspan='2'>NAME</th>");
                    sb.Append("<th rowspan='2'>SECTION</th>");
                    sb.Append("<th colspan='2'>ENGLISH</th>");
                    sb.Append("<th colspan='2'>ACCOUNTANCY</th>");
                    sb.Append("<th colspan='2'>BST</th>");
                    sb.Append("<th rowspan='2'>TOTAL MARKS</th>");
                    sb.Append("<th rowspan='2'>%AGE</th>");
                    sb.Append("<th rowspan='2'>CAMPUS RANK</th>");
                    sb.Append("<th rowspan='2'>VIKASH RANK</th>");
                    sb.Append("<th rowspan='2'>TOTAL CORRECT</th>");
                    sb.Append("<th rowspan='2'>TOTAL INCORRECT</th>");
                    sb.Append("<th rowspan='2'>TIME TAKEN</th>");
                    sb.Append("</tr>");
                    sb.Append("<tr>");
                    sb.Append("<th>Mark</th>");
                    sb.Append("<th>Rank</th>");
                    sb.Append("<th>Mark</th>");
                    sb.Append("<th>Rank</th>");
                    sb.Append("<th>Mark</th>");
                    sb.Append("<th>Rank</th>");
                    sb.Append("</tr>");
                    sb.Append("</thead>");

                    sb.Append("<tbody>");
                    foreach (DataRow row in dt.Rows)
                    {
                        sb.Append("<tr>");
                        sb.Append("<td>"+row["RollNo"]+"</td>");
                        sb.Append("<td>"+row["Campus"] +"</td>");
                        sb.Append("<td>"+row["Name"] +"</td>");
                        sb.Append("<td>"+row["Section"] +"</td>");
                        sb.Append("<td>"+row["Eng_Mark"] +"</td>");
                        sb.Append("<td>"+row["Eng_Rank"] +"</td>");
                        sb.Append("<td>"+row["Acc_Mark"] +"</td>");
                        sb.Append("<td>"+row["Acc_Rank"] +"</td>");
                        sb.Append("<td>"+row["Bst_Mark"] +"</td>");
                        sb.Append("<td>"+row["Bst_Rank"] +"</td>");
                        sb.Append("<td>"+row["TotalMarks"] +"</td>");
                        sb.Append("<td>"+row["Percentage"] +"</td>");
                        sb.Append("<td>"+row["CampusRank"] +"</td>");
                        sb.Append("<td>"+row["VikashRank"] +"</td>");
                        sb.Append("<td>"+row["TotalCorrect"] +"</td>");
                        sb.Append("<td>"+row["TotalIncorrect"] +"</td>");
                        sb.Append("<td>"+row["TimeTaken"] +"</td>");
                        sb.Append("</tr>");
                    }
                    sb.Append("</tbody>");

                    sb.Append("</table>");


                    Results.Text = sb.ToString();

                }
            }
            catch { }
        }

        protected void btn_ExportExcel_Click(object sender, EventArgs e)
        {
            using (XLWorkbook wb = new XLWorkbook())
            {
                DataTable dt = new DataTable();
                using (SqlConnection conn = new SqlConnection(DB))
                {
                    string com = "Select * from ExamResult";
                    SqlDataAdapter adpt = new SqlDataAdapter(com, conn);
                    adpt.Fill(dt);
                }
                var count = dt.Columns.Count;
                var ws = wb.Worksheets.Add("Results");
                ws.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                var cr = 1;
                ws.Cell(cr, 1).Value = "VIKASH GROUP OF INSTITUTIONS";
                ws.Range(cr, 1, cr, count).Merge();
                ws.Row(cr).Style.Font.Bold = true;

                cr++;
                ws.Cell(cr, 1).Value = "FINAL RESULT OF XII COMMERCE, SPT-01 (ACC, BST, ENG) [19-09-2021]";
                ws.Range(cr, 1, cr, count).Merge();
                ws.Row(cr).Style.Font.Bold = true;

                cr++;
                ws.Cell(cr, 1).Value = "VRS, BRGH";
                ws.Cell(cr, 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                ws.Range(cr, 1, cr, count).Merge();
                ws.Row(cr).Style.Font.Bold = true;


                cr++;
                ws.Cell(cr, 1).Value = "ROLL NO";
                ws.Range(cr, 1, cr+1, 1).Merge();

                ws.Cell(cr, 2).Value = "CAMPUS";
                ws.Range(cr, 2, cr+1, 2).Merge();

                ws.Cell(cr, 3).Value = "NAME";
                ws.Range(cr, 3, cr+1, 3).Merge();
                ws.Cell(cr, 4).Value = "SECTION";
                ws.Range(cr, 4, cr + 1, 4).Merge();


                ws.Cell(cr, 5).Value = "ENGLISH";
                ws.Range(cr, 5, cr, 6).Merge();
                ws.Cell(cr, 7).Value = "ACCOUNTANCY";
                ws.Range(cr, 7, cr, 8).Merge();
                ws.Cell(cr, 9).Value = "BST";
                ws.Range(cr, 9, cr, 10).Merge();




                ws.Cell(cr, 11).Value = "TOTAL MARKS";
                ws.Cell(cr, 12).Value = "%AGE";
                ws.Cell(cr, 13).Value = "CAMPUS RANK";
                ws.Cell(cr, 14).Value = "VIKASH RANK";
                ws.Cell(cr, 15).Value = "TOTAL CORRECT";
                ws.Cell(cr, 16).Value = "TOTAL INCORRECT";
                ws.Cell(cr, 17).Value = "TOTAL TIME TAKEN";
                ws.Row(cr).Style.Font.Bold = true;

                cr++;
                ws.Cell(cr, 5).Value = "Mark";
                ws.Cell(cr, 6).Value = "Rank";
                ws.Cell(cr, 7).Value = "Mark";
                ws.Cell(cr, 8).Value = "Rank";
                ws.Cell(cr, 9).Value = "Mark";
                ws.Cell(cr, 10).Value = "Rank";

                foreach (DataRow row in dt.Rows)
                {
                    cr++;
                    ws.Cell(cr, 1).Value = row["RollNo"].ToString();
                    ws.Cell(cr, 2).Value = row["Campus"].ToString();
                    ws.Cell(cr, 3).Value = row["Name"].ToString();
                    ws.Cell(cr, 4).Value = row["Section"].ToString();
                    ws.Cell(cr, 5).Value = row["Eng_Mark"].ToString();
                    ws.Cell(cr, 6).Value = row["Eng_Rank"].ToString();
                    ws.Cell(cr, 7).Value = row["Acc_Mark"].ToString();
                    ws.Cell(cr, 8).Value = row["Acc_Rank"].ToString();
                    ws.Cell(cr, 9).Value = row["Bst_Mark"].ToString();
                    ws.Cell(cr, 10).Value = row["Bst_Rank"].ToString();
                    ws.Cell(cr, 11).Value = row["TotalMarks"].ToString();
                    ws.Cell(cr, 12).Value = row["Percentage"].ToString();
                    ws.Cell(cr, 13).Value = row["CampusRank"].ToString();
                    ws.Cell(cr, 14).Value = row["VikashRank"].ToString();
                    ws.Cell(cr, 15).Value = row["TotalCorrect"].ToString();
                    ws.Cell(cr, 16).Value = row["TotalIncorrect"].ToString();
                    ws.Cell(cr, 17).Value = row["TimeTaken"].ToString();
                }
                cr++;
                ws.Cell(cr, 1).Value = "Vikash | Vidya & copy; " + DateTime.Now.Year;
                ws.Cell(cr, 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;
                ws.Range(cr, 1, cr, count / 2).Merge();
                ws.Cell(cr, 5).Value = "Generated On :" + DateTime.Now.ToString("dddd, dd-MMMM-yyyy HH:mm");
                ws.Cell(cr, 5).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                ws.Range(cr, 5, cr, count).Merge();
                ws.Row(cr).Style.Font.Bold = true;


                Response.Clear();
                Response.Buffer = true;
                Response.Charset = "US-ASCII";
                Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                Response.AddHeader("content-disposition", "attachment;filename=ResultExport.xlsx");
                using (MemoryStream MyMemoryStream = new MemoryStream())
                {
                    wb.SaveAs(MyMemoryStream);
                    MyMemoryStream.WriteTo(Response.OutputStream);
                    Response.Flush();
                    Response.End();
                }
            }
        }
    }
}