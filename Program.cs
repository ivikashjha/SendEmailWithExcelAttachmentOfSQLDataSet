using ClosedXML.Excel;
using Microsoft.Practices.EnterpriseLibrary.Data;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.Common;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net.Mail;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace SendEmailWithExcelAttachmentOfSQLDataSet
{
    public class Program
    {
        static void Main(string[] args)
        {
            Program pObj=new Program();
            string report = pObj.FetchSummary();
            pObj.SendEmail();
            Console.WriteLine(report);
        }

        public void SendEmail()
        {
            try
            {
                using (SmtpClient smtpClient = new SmtpClient())
                {
                    using (MailMessage message = new MailMessage())
                    {
                        string AppLocation = "";
                        AppLocation = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().CodeBase);
                        AppLocation = AppLocation.Replace("file:\\", "");
                        var mailTo = ConfigurationManager.AppSettings["mailto"].ToString();
                        var reportedById = ConfigurationManager.AppSettings["mailfrom"].ToString();
                        var smtpHost = ConfigurationManager.AppSettings["smtphost"].ToString();

                        var emailto = ConfigurationManager.AppSettings["mailto"].ToString();
                        message.From = new MailAddress(ConfigurationManager.AppSettings["mailfrom"].ToString());
                        message.To.Add(emailto);
                        message.Subject = DateTime.UtcNow.ToString("yyyy-MM-dd HH:mm") + ":- " + ConfigurationManager.AppSettings["subject"].ToString();
                        message.Body = ConfigurationManager.AppSettings["body"].ToString();

                        StringBuilder strBuilder = new StringBuilder();
                        strBuilder.Append("Hello Team, <br/><br/>");
                        strBuilder.Append(message.Body + "<br/><br/>");
                        strBuilder.Append("<i>*Time is in Greenwich Mean Time (GMT). Use the <a href='http://worldclock.accenture.com/'/>World Clock</a> to convert to your local time zone.</i><br/><br/>");
                        strBuilder.Append("Thanks\n");

                        message.Body = strBuilder.ToString();
                        message.IsBodyHtml = true;
                        message.BodyEncoding = Encoding.UTF8;

                        string path = AppLocation + "\\ExcelFiles\\";
                        string[] files = Directory.GetFiles(path);

                        foreach (string f in files)
                        {
                            try
                            {
                                System.Net.Mail.Attachment attachment = new System.Net.Mail.Attachment(f);
                                message.Attachments.Add(attachment);
                            }
                            catch (Exception ex)
                            {
                                throw ex;
                            }
                        }

                        if (message.Attachments.Count != 0)
                        {
                            smtpClient.Send(message);
                        }
                        else
                        {
                            //log the failure.
                        }
                    }
                }
                this.Delete();
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public void Delete()
        {
            string fileDirectory = string.Empty;
            string sourceDir = "";
            sourceDir = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().CodeBase);
            sourceDir = sourceDir.Replace("file:\\", "");
            fileDirectory = sourceDir + "\\ExcelFiles\\";
            string[] txtList = Directory.GetFiles(fileDirectory, "*.xlsx");
            foreach (string f in txtList)
            {
                File.Delete(f);
            }
        }

        private string FetchSummary()
        {
            string fileName = "Report";
            string message = string.Empty;
            var excelData = FetchDomainDataSummary();
            if (excelData != null)
            {
                var dsReport = excelData.ConvertToDataSet("RPT");
                message = ExportDataSetToExcel(dsReport, fileName);
            }
            return fileName;
        }

        private string ExportDataSetToExcel(DataSet reportingDS, string fileName)
        {
            string file = string.Empty;
            try
            {
                if (reportingDS.Tables[0].Rows.Count != 0)
                {
                    string AppLocation = "";
                    AppLocation = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().CodeBase);
                    AppLocation = AppLocation.Replace("file:\\", "");
                    file = AppLocation + "\\ExcelFiles\\" + fileName + DateTime.UtcNow.ToString("yyyyMMddHHmm") + ".xlsx";
                    IXLWorkbook wb = new XLWorkbook();
                    IXLWorksheet ws = wb.Worksheets.Add(reportingDS.Tables[0]);
                    ws.Style.Fill.BackgroundColor = XLColor.White;
                    ws.Row(1).Style.Fill.BackgroundColor = XLColor.Gray;
                    ws.RowHeight = 25;
                    ws.ColumnWidth = 50;
                    ws.Style.Border.LeftBorder = XLBorderStyleValues.Thin;
                    ws.Columns().AdjustToContents();
                    ws.Style.Alignment.WrapText = true;
                    //ws.Style.Alignment.Indent = 5;
                    wb.SaveAs(file);
                }
            }
            catch (Exception exp)
            {

                throw exp;
            }
            return file;
        }

        private IList<MyDomainData> FetchDomainDataSummary()
        {
            MyDomainData mydomObj = new MyDomainData();                       
            List<MyDomainData> list = null;
            MyDomainData excel = null;
            var ConnectionStringName = "";
            Stopwatch sw = Stopwatch.StartNew();
            DatabaseProviderFactory factory = new DatabaseProviderFactory();
            Database database = factory.Create(ConnectionStringName);
            using (DbCommand dbCommand = database.GetSqlStringCommand(ExportQuery()))
            {
                using (IDataReader reader = database.ExecuteReader(dbCommand))
                {
                    while (reader.Read())
                    {
                        if (list == null) list = new List<MyDomainData>();
                        excel = new MyDomainData();
                        excel.ID = (int)reader[0];
                        excel.Name = (string)reader[1];
                        excel.Status = (string)reader[2];
                        list.Add(excel);
                    }
                }
            }

            return list;
        }

        private string ExportQuery()
        {
            return "select * from myTable";
        }
    }
}
