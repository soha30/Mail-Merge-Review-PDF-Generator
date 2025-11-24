using ali1982ReviewMailmerge.App_Code;
using Syncfusion.DocIO.DLS;
using Syncfusion.DocToPDFConverter;
using Syncfusion.Pdf;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

namespace ali1982ReviewMailmerge.demo
{

    public partial class _default : System.Web.UI.Page
    {
        //  public static string pathExcel = @"C:\ds\ds_excelStudent";
        //   public static string conStr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source;=" + pathExcel + ";Extended Properties=Excel 12.0;";
        protected void Page_Load(object sender, EventArgs e)
        {
            populateGvStudent();
        }
        protected void populateGvStudent()
        {
            CRUD myCrud = new CRUD();
            string mySql = @"select studentId,studentName,nId,gpa,major,certificateDate
                        from student s inner  join major ma on s.majorId = ma.majorId";
            SqlDataReader dr = myCrud.getDrPassSql(mySql);
            gvStudent.DataSource = dr;
            gvStudent.DataBind();
        }
        protected DataTable getDt()
        {
            string mySql = @"select studentId,studentName,gpa,major,certificateDate
                        from student s inner
                        join major ma on s.majorId = ma.majorId";

            CRUD myCrud = new CRUD();
            // mySql = @"select * from v_dsTemplate order by refno";// sepcify any data in db  created in case above 
            DataTable dt = myCrud.getDT(mySql);
            return dt;
        }
        protected Array getDtColNames(DataTable myDt)
        {
            // capture gv column header 
            int intColCount = myDt.Columns.Count;
            string[] colNames = new string[intColCount];
            for (int i = 0; i <= myDt.Columns.Count - 1; i++)
            {
                colNames[i] = myDt.Columns[i].ToString();
            }
            return colNames;
        }
        protected Array getDtColValues(DataTable myDt, DataRow myDataRow)
        {
            int myindex = 0;
            int columnCount = myDt.Columns.Count;
            string[] results = new string[columnCount];
            foreach (DataColumn dc in myDt.Columns)
            {
                if (myindex <= columnCount)
                {
                    results[myindex] = myDataRow[dc].ToString();
                }
                myindex += 1;
            }
            return results;
        }


        protected DataTable GetRecipients()
        {
            CRUD myCrud = new CRUD();
            string mySql = @"select studentId,studentName,gpa,major,certificateDate
                    from student s inner  join major ma on s.majorId = ma.majorId ";// sepcify any data in db  where id in (95,96)
            DataTable dt = myCrud.getDT(mySql);
            return dt;
        }

        protected void btnIssueCertificateViaDT_Click1(object sender, EventArgs e)
        {
            //Creates new Word document instance for Word processing.
            using (WordDocument template = new WordDocument())
            {
                //Opens the template Word document.
                //  template.Open(Path.GetFullPath(@"../../LetterTemplate.docx"), FormatType.Docx);  // error = Could not find file 'C:\LetterTemplate.Docx'.
                // template.Open(Path.GetFullPath(@"~/Data/LetterTemplate.docx"), FormatType.Docx);  // error = Could not find a part of the path 'C:\Program Files (x86)\IIS Express\~\Data\LetterTemplate.Docx'.
                template.Open(Server.MapPath("~/myTemplate/sample_Template.dotx"));
                //Gets the recipient details as a DataTable.
                DataTable recipients = GetRecipients();
                //Creates folder for saving generated documents.
                if (!Directory.Exists(Path.GetFullPath(@"../../Result/")))
                    Directory.CreateDirectory(Path.GetFullPath(@"../../Result/"));
                foreach (DataRow dataRow in recipients.Rows)
                {
                    //Clones the template document for creating new document for each record in the data source.
                    WordDocument document = template.Clone();
                    //Performs the mail merge.
                    document.MailMerge.Execute(dataRow);

                    #region to save as word
                    ////..Save the file in the given path.
                    ////document.Save(Path.GetFullPath(@"../../Result/Letter_" + dataRow.ItemArray[2].ToString() + ".docx"), FormatType.Docx);
                    ////document.Save(Path.GetFullPath(@"/Result/Letter_ali.docx"), FormatType.Docx);
                    //document.Save(Server.MapPath(@"~/Result/Letter_"+dataRow.ItemArray[2].ToString()+".docx"));
                    ////Releases the resources occupied by WordDocument instance.
                    //document.Dispose();
                    #endregion

                    #region To save as pdf 
                    // if I want to save in pdf
                    // Creates an instance of the DocToPDFConverter
                    DocToPDFConverter converter = new DocToPDFConverter();
                    // Converts Word document to PDF document
                    PdfDocument pdfDocument = converter.ConvertToPDF(document);
                    // Closes the instance of Word document object
                    document.Close();
                    //Releases all resources used by DocToPDFConverter object
                    converter.Dispose();
                    pdfDocument.Save(Server.MapPath(@"~/Result/Letter_" + dataRow.ItemArray[1].ToString() + ".pdf"));
                    pdfDocument.Close(true);
                    #endregion
                }
            }
            lblOutput.Text = "Issuce Certificate Completed!";
        }

        protected void btnIssueCertificateViaArray_Click(object sender, EventArgs e)
        {
            // put the code here 
            string sname = "";
            // get dt
            DataTable myDt = getDt();
            if (myDt.Rows.Count >= 1)
            {
                // if DT has records
                string[] fieldNames = (string[])getDtColNames(myDt); // note gv must have explicit columns , myGv 
                string[] fieldValues;
                List<int> intRefnoList = new List<int>();
                foreach (DataRow row in myDt.Rows)  // to get DataTable column names
                {
                    fieldValues = (string[])getDtColValues(myDt, row);
                    //  int intRefno = int.Parse(fieldValues[4]);// capture intern Refno to make update in table 
                    Syncfusion.DocIO.DLS.WordDocument document = new Syncfusion.DocIO.DLS.WordDocument(Server.MapPath("~/myTemplate/updatedDocumentWithWatermark.dotx"));
                    //يحذف الحقول الفارغه 
                    document.MailMerge.RemoveEmptyParagraphs = true;
                    // mail merge
                    document.MailMerge.Execute(fieldNames, fieldValues);
                    // Creates an instance of the DocToPDFConverter
                    DocToPDFConverter converter = new DocToPDFConverter();
                    // Converts Word document to PDF document
                    PdfDocument pdfDocument = converter.ConvertToPDF(document);
                    // Closes the instance of Word document object
                    document.Close();
                    //Releases all resources used by DocToPDFConverter object
                    converter.Dispose();
                    pdfDocument.Save(Server.MapPath("~/myPdf/" + fieldValues[3] + "_sample_Template.pdf"));
                    pdfDocument.Close(true);
                }
                lblOutput.Text = "Documents generated Successfully!";
            }
        }

        protected void btnIssueCertificateViaExce_Click(object sender, EventArgs e)
        {
            //CRUD myCrud = new CRUD();
            //string mySql = @"select studentId,studentName,gpa,majorId,certificateDate
            //            from [ds_excelStudent] ";
            //DataSet ds = myCrud.getDS(mySql); // put data from excel 

            //SqlDataReader dr = myCrud.getDrPassSql(mySql);

            //SqlBulkCopy bulkInsert = new SqlBulkCopy(CRUD.conStr);
            //bulkInsert.DestinationTableName = "student";
            //bulkInsert.WriteToServer(dr);

        }

    }     //cls
}// ns

