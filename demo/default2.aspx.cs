using ali1982ReviewMailmerge.App_Code;
using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.DocToPDFConverter;
using Syncfusion.Pdf;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using ali1982ReviewMailmerge;
using System.Configuration;
using System.Net.Mail;

namespace ali1982ReviewMailmerge.demo
{
    public partial class default2 : System.Web.UI.Page
    {

        protected void Page_Load(object sender, EventArgs e)
        {
           
            if (!IsPostBack)
            {
                populateGvStudent();

                // Load certificate data in view certificate link
                List<string> pdfFilePaths = new List<string>();
                // Add the paths of the PDF files here
                ViewState["AllPdfFilePaths"] = pdfFilePaths;

                // Ensure the button is visible only if there are PDF files
                //btnViewCertificate.Visible = pdfFilePaths.Any();
            }
        }
        protected void populateGvStudent()
        {
            CRUD myCrud = new CRUD();
            string mySql = @"select studentId,studentName,gpa,major,certificateDate,email
                        from student s inner
                        join major ma on s.majorId = ma.majorId";
            SqlDataReader dr = myCrud.getDrPassSql(mySql);
            gvStudent.DataSource = dr;
            gvStudent.DataBind();
        }
        protected DataTable getDt()
        {
            string mySql = @"select studentId,studentName,gpa,major,certificateDate,email
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
        protected void btnIssueCertificateViaExce_Click(object sender, EventArgs e)
        {
            CRUD myCrud = new CRUD();
            string mySql = @"select studentId,studentName,gpa,majorId,certificateDate
                        from [ds_excelStudent] ";
            DataSet ds = myCrud.getDS(mySql); // put data from excel 

            SqlDataReader dr = myCrud.getDrPassSql(mySql);

            SqlBulkCopy bulkInsert = new SqlBulkCopy(CRUD.conStr);
            bulkInsert.DestinationTableName = "student";
            bulkInsert.WriteToServer(dr);

        }
        protected void btnAddWatermark_Click(object sender, EventArgs e)
        {
            try
            {
                // Retrieve inputs from user interface
                string watermarkTextInput = lblWatermarkText.Text; // Text for the watermark
                float fontSizeInput = float.Parse(lblFontSize.Text); // Font size for the watermark text
                string fontColorName = lblFontColor.Text; // Font color name from a label or input
                Color fontColorInput = Color.FromName(fontColorName); // Convert color name to Color object

                // Create a WordDocument instance and load an existing document template
                WordDocument document = new WordDocument();
                document.Open(Server.MapPath("~/Template/sample_Template.dotx")); // Opening the Word document template

                // Add a text watermark to the document
                AddTextWatermark(document, watermarkTextInput, fontSizeInput, fontColorInput);

                // Save the document after adding the watermark to a specified path
                document.Save(Server.MapPath("~/Template/UpdatedDocumentWithWatermark.docx"), FormatType.Docx);

                // Display a success message to the user
                lblOutput.Text = "Watermark added successfully!";
            }
            catch (Exception ex)
            {
                // Display an error message if an exception occurs
                lblOutput.Text = "An error occurred: " + ex.Message;
            }
        }

        /// <summary>
        /// Adds a text watermark to the specified Word document.
        /// This function demonstrates how to insert a watermark into the header of a Word document
        /// using Syncfusion's DocIO library. The watermark is customized with text, font size, and color.
        /// </summary>
       
        private void AddTextWatermark(WordDocument document, string watermarkText, float fontSize, Color fontColor)
        {
            // Access the first section (or all sections if desired)
            foreach (WSection section in document.Sections)
            {
                // Create a new paragraph for adding the watermark in the header
                IWParagraph watermarkParagraph = section.HeadersFooters.Header.AddParagraph();

                // Add the watermark text to the paragraph
                IWTextRange textRange = watermarkParagraph.AppendText(watermarkText);
                textRange.CharacterFormat.FontSize = fontSize;
                textRange.CharacterFormat.TextColor = fontColor;

                // Additional settings for text alignment or rotation can be configured here
            }
        }
        /// <summary>
        /// Handles the click event of the button that initiates certificate generation for each student.
        /// This method fetches student data, generates certificates using a Word document template, and saves them as PDFs.
        /// It stores all generated field values and PDF paths for further use, possibly for emailing or other processing.
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="e">An EventArgs object that contains the event data.</param>
        protected void btnIssueCertificateViaArray_Click(object sender, EventArgs e)
        {
            // Temporary variable for student name initialization
            string sname = "";
            // Fetch the data table
            DataTable myDt = getDt();

            // Check if there are records in the data table
            if (myDt.Rows.Count >= 1)
            {
                // Retrieve column names to use in the mail merge process
                string[] fieldNames = (string[])getDtColNames(myDt);
                // Lists to store field values and PDF file paths
                List<string[]> allFieldValues = new List<string[]>();
                List<string> allPdfFilePaths = new List<string>();

                // Iterate over each record in the data table
                foreach (DataRow row in myDt.Rows)
                {
                    string[] fieldValues = (string[])getDtColValues(myDt, row);

                    // Load the template and perform mail merge
                    Syncfusion.DocIO.DLS.WordDocument document = new Syncfusion.DocIO.DLS.WordDocument(Server.MapPath("~/Template/updatedDocumentWithWatermark.dotx"));
                    document.MailMerge.RemoveEmptyParagraphs = true;
                    document.MailMerge.Execute(fieldNames, fieldValues);

                    // Convert the document to PDF
                    DocToPDFConverter converter = new DocToPDFConverter();
                    PdfDocument pdfDocument = converter.ConvertToPDF(document);
                    document.Close();
                    converter.Dispose();

                    // Save the PDF file using Student ID and Certificate ID as the file name
                    string pdfFilePath = Server.MapPath("~/myPdf/" + fieldValues[0] + fieldValues[3] + "_certificate.pdf");
                    pdfDocument.Save(pdfFilePath);
                    pdfDocument.Close(true);

                    // Add field values and PDF file path to the lists
                    allFieldValues.Add(fieldValues);
                    allPdfFilePaths.Add(pdfFilePath);
                }

                // Store the lists in ViewState for later use
                ViewState["AllFieldValues"] = allFieldValues;
                ViewState["AllPdfFilePaths"] = allPdfFilePaths;

                // Display a success message to the user
                lblOutput.Text = "Documents generated Successfully!";
            }
        }



        /// <summary>
        /// Sends the graduation certificate to the specified student via email.
        /// </summary>
        /// <param name="studentEmail">The student's email address where the certificate will be sent.</param>
        /// <param name="pdfFilePath">The file path of the PDF certificate to be attached to the email.</param>
        private void SendCertificateByEmail(string studentEmail, string pdfFilePath)
        {
            // Create a new MailMessage object
            MailMessage mail = new MailMessage();

            // Set the sender email address from the app settings
            mail.From = new MailAddress(ConfigurationManager.AppSettings["emailFrom"]);

            // Add the student's email as the recipient
            mail.To.Add(studentEmail);

            // Set the subject of the email
            mail.Subject = "Your Graduation Certificate";

            // Set the body of the email
            mail.Body = "Dear student,\n\nPlease find your graduation certificate attached.\n\nBest regards,\nYour University";
            mail.IsBodyHtml = false; // Set the email body format as plain text

            // Attach the PDF file to the email
            Attachment attachment = new Attachment(pdfFilePath);
            mail.Attachments.Add(attachment);

            // Set up the SMTP client using settings from the configuration file
            SmtpClient smtpClient = new SmtpClient
            {
                Host = ConfigurationManager.AppSettings["HostsmtpAddress"], // SMTP host
                Port = int.Parse(ConfigurationManager.AppSettings["PortNumber"]), // SMTP port
                Credentials = new System.Net.NetworkCredential(
                    ConfigurationManager.AppSettings["emailUserName"], // SMTP username
                    ConfigurationManager.AppSettings["emailPassword"] // SMTP password
                ),
                EnableSsl = bool.Parse(ConfigurationManager.AppSettings["EnableSSL"]) // Enable SSL if specified
            };

            try
            {
                // Send the email
                smtpClient.Send(mail);
            }
            catch (Exception ex)
            {
                // Handle any exceptions that occur during the email sending process
                lblOutput.Text = "An error occurred while sending the email: " + ex.Message;
            }
            finally
            {
                // Release resources used by the attachment and mail objects
                attachment.Dispose();
                mail.Dispose();
            }
        }

        /// <summary>
        /// Handles the event when the 'Send Certificate By Email' button is clicked.
        /// This method retrieves the necessary data from ViewState and sends each certificate to the corresponding student via email.
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="e">The event data.</param>
        protected void btnSendCertificateByEmail_Click(object sender, EventArgs e)
        {
            // Retrieve the lists from ViewState
            List<string[]> allFieldValues = (List<string[]>)ViewState["AllFieldValues"];
            List<string> allPdfFilePaths = (List<string>)ViewState["AllPdfFilePaths"];

            // Send each certificate via email
            for (int i = 0; i < allFieldValues.Count; i++)
            {
                // Send the certificate to the student's email address
                // 'allFieldValues[i][5]' contains the student's email, and 'allPdfFilePaths[i]' contains the file path to the PDF certificate
                SendCertificateByEmail(allFieldValues[i][5], allPdfFilePaths[i]);
            }
            lblOutput.Text = "sent successfully";
        }


        /// <summary>
        /// Handles the event when the 'Issue Certificate Via Email' button is clicked.
        /// This method retrieves student data, generates certificates using a Word template, 
        /// converts them to PDF, and sends them via email to each student.
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="e">The event data.</param>
        protected void btnIssueCertificateViaEmail_Click(object sender, EventArgs e)
        {
            // Fetch data
            string sname = "";
            DataTable myDt = getDt();

            if (myDt.Rows.Count >= 1)
            {
                // Check if there are any records
                string[] fieldNames = (string[])getDtColNames(myDt);
                string[] fieldValues;
                List<int> intRefnoList = new List<int>();

                foreach (DataRow row in myDt.Rows)
                {
                    // Get column values for each student
                    fieldValues = (string[])getDtColValues(myDt, row);

                    // Create a document using the Word template
                    Syncfusion.DocIO.DLS.WordDocument document = new Syncfusion.DocIO.DLS.WordDocument(Server.MapPath("~/Template/updatedDocumentWithWatermark.dotx"));

                    // Remove empty fields in the template
                    document.MailMerge.RemoveEmptyParagraphs = true;

                    // Perform Mail Merge to fill in the student data into the template
                    document.MailMerge.Execute(fieldNames, fieldValues);

                    // Create a converter to convert the document to PDF
                    DocToPDFConverter converter = new DocToPDFConverter();

                    // Convert the document to PDF
                    PdfDocument pdfDocument = converter.ConvertToPDF(document);

                    // Close the Word document to release resources
                    document.Close();
                    converter.Dispose();

                    // Save the PDF file with a custom name including the student's name
                    string pdfFilePath = Server.MapPath("~/myPdf/" + fieldValues[3] + "_sample_Template.pdf");
                    pdfDocument.Save(pdfFilePath);
                    pdfDocument.Close(true);

                    // Send the certificate via email
                    SendCertificateByEmail(fieldValues[5], pdfFilePath); // Assuming the email is in the sixth field
                }

                lblOutput.Text = "Documents generated and sent successfully!";
            }
            else
            {
                lblOutput.Text = "No records found!";
            }
        }
        /// <summary>
        /// This method is designed to display the certificates generated in a previous step.
        /// It retrieves file paths stored in ViewState and binds them to a repeater control to show links to the PDF certificates.
        /// This method could be triggered by a user action like a button click.
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="e">An EventArgs object that contains the event data.</param>
        protected void ShowCertificates(object sender, EventArgs e)
        {
            // Check if the ViewState contains any stored PDF file paths
            if (ViewState["AllPdfFilePaths"] != null)
            {
                // Retrieve the list of PDF file paths from ViewState
                List<string> AllpdfFilePaths = (List<string>)ViewState["AllPdfFilePaths"];

                // Create a list of anonymous objects containing file names and URLs for each PDF
                var certificates = AllpdfFilePaths.Select(path => new
                {
                    FileName = Path.GetFileName(path), // Extracts the file name from the path
                    FileUrl = ResolveUrl("~/myPdf/" + Path.GetFileName(path)) // Creates a URL that can be used in a web environment
                }).ToList();

                // Set the data source of a repeater control to display the certificates
                rptCertificates.DataSource = certificates;
                rptCertificates.DataBind(); // Bind the data to the repeater control to update the UI
            }
        }

    }     //cls
}// ns