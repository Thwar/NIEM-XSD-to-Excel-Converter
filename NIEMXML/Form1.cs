using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Deployment.Application;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml.Linq;

namespace NIEMXML
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            this.Text = "NIEM XSD to Excel Converter " + GetRunningVersion();
        }

        string elementName = "";
        string extensionClass = "";
        XNamespace xs = XNamespace.Get("http://www.w3.org/2001/XMLSchema");
        XDocument doc;
        XLWorkbook wb = new XLWorkbook();
        bool jobCanceled = false;
        string errorMsg;
        string outputPath;
        string filename;

        private string GetRunningVersion()
        {
            //if (System.Deployment.Application.ApplicationDeployment.IsNetworkDeployed)
            //{
            //    System.Deployment.Application.ApplicationDeployment ad = System.Deployment.Application.ApplicationDeployment.CurrentDeployment;
            //    return "Version: " + ad.CurrentVersion.ToString();
            //}

            return "v2.3";
        }

        private void button1_Click(object sender, EventArgs e)
        {
        }

        //Get Element Types
        public static string searchForElementTypes(string elementName, XDocument doc, XNamespace xs, BackgroundWorker backgroundWorker1, int total)
        {
            foreach (var el in doc.Root.Elements(xs + "element"))
            {
                if (el.Attribute("name").Value == elementName)
                {
                    backgroundWorker1.ReportProgress(total, total);
                    return el.Attribute("type") != null ? el.Attribute("type").Value : "";
                }
            }
            return "";
        }


        //Get Element Documentation
        public static Tuple<string, string> searchForElementDocumentation(string elementName, XDocument doc, XNamespace xs)
        {
            string description = "";
            string source = "";

            foreach (var el in doc.Root.Elements(xs + "element"))
            {
                if (el.Attribute("name").Value == elementName)
                {
                    //Check if Documentation exists for Element
                    var documentation = el.Elements(xs + "annotation").Elements(xs + "documentation").ToList();

                    //Check if more than one Documentation entry exists and throw error
                    if (documentation.Count() > 2)
                        throw new documentationEntryException(elementName);

                    //Get Documentation
                    if (documentation.Count() == 1)
                    {
                        source = documentation[0].Value.StartsWith("Source:") ? documentation[0].Value : "";
                        description = documentation[0].Value.StartsWith("Source:") ? "" : documentation[0].Value;
                    }

                    //Get Source
                    if (documentation.Count() == 2)
                    {
                        source = documentation[1].Value.StartsWith("Source:") ? documentation[1].Value : documentation[0].Value;
                        description = documentation[1].Value.StartsWith("Source:") ? documentation[0].Value : documentation[1].Value;
                    }

                }
            }
            return new Tuple<string, string>(description, source);
        }


        private void writeClassDocumentation(XElement el, IXLWorksheet ws, int row)
        {
            //Check if Documentation exists for Class
            var classDocumentation = el.Elements(xs + "annotation").Elements(xs + "documentation").ToList();

            //Check if more than one Documentation entry exists and throw error
            if (classDocumentation.Count() > 2)
                throw new documentationEntryException(el.Attribute("name").Value);

            if (classDocumentation.Count() == 1)
            {
                ws.Cell("E" + row).Value = classDocumentation[0].Value.StartsWith("Source:") ? classDocumentation[0].Value : "";
                ws.Cell("D" + row).Value = classDocumentation[0].Value.StartsWith("Source:") ? "" : classDocumentation[0].Value;
            }

            if (classDocumentation.Count() == 2)
            {
                ws.Cell("E" + row).Value = classDocumentation[0].Value.StartsWith("Source:") ? classDocumentation[0].Value : classDocumentation[1].Value;
                ws.Cell("D" + row).Value = classDocumentation[0].Value.StartsWith("Source:") ? classDocumentation[1].Value : classDocumentation[0].Value;
            }

        }

        private void openFileDialog1_FileOk(object sender, CancelEventArgs e)
        {

        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        public class documentationEntryException : Exception
        {
            public documentationEntryException(string message) : base(message) { }
        }

        private void selectXSD_Click(object sender, EventArgs e)
        {
            openFileDialog1.Filter = "XSD NIEM (.xsd)|*.xsd";
            openFileDialog1.FilterIndex = 1;

            DialogResult result = openFileDialog1.ShowDialog(); // Show the dialog.
            if (result == DialogResult.OK) // Test result.
            {
                string file = openFileDialog1.FileName;
                outputPath = Path.GetDirectoryName(file);
                filename = Path.GetFileNameWithoutExtension(file);

                textBox1.Text = file;

                try
                {
                    string text = File.ReadAllText(file);
                    createExcel.Enabled = true;
                    label1.Text = "Step 2: Click Convert To Excel button to begin.";

                }
                catch (IOException)
                {
                }
            }
        }

        private void aboutToolStripMenuItem_Click(object sender, EventArgs e)
        {
            MessageBox.Show("This tool is used to convert NIEM XSD schemas into Excel Spreadsheets. \n\nAuthor: Ruben T. Rosales\nRuben.Rosales@acf.hhs.gov\n\nClick Help button for GitHub repository. \nVersion: " + GetRunningVersion(), "About NIEM XSD to Excel Converter ", MessageBoxButtons.OK, MessageBoxIcon.Asterisk, MessageBoxDefaultButton.Button1, 0, "https://github.com/Thwar/NIEM-XSD-to-Excel-Converter", "");
        }

        private void createExcel_Click(object sender, EventArgs e)
        {
            progressBar1.Value = 0;
            progressBar1.Step = 1;
            label3.Text = "0%";

            createExcel.Enabled = false;
            selectXSD.Enabled = false;
            label1.Text = "Creating Excel Spreadsheet. Please wait...";
            label2.Text = "Do not close or click on Excel. \nDoing so will terminate the operation.";

            backgroundWorker1.RunWorkerAsync();
        }

        private void backgroundWorker1_DoWork(object sender, DoWorkEventArgs e)
        {
            try
            {
                jobCanceled = false;
                errorMsg = "";
                wb = new XLWorkbook();
                var ws = wb.Worksheets.Add("NIEM");
                doc = XDocument.Load(openFileDialog1.FileName);

                //creates the main header
                ws.Cell("A2").Value = openFileDialog1.FileName;
                ws.Range("A2:E2").Merge().Style.Fill.BackgroundColor = XLColor.Khaki;
                if (doc.Root.Elements(xs + "annotation") != null)
                    if (doc.Root.Elements(xs + "annotation").Elements(xs + "documentation").Count() > 0)
                        ws.Cell("A3").Value = doc.Root.Elements(xs + "annotation").Elements(xs + "documentation").Single().Value;

                ws.Range("A3:E3").Merge().Style.Fill.BackgroundColor = XLColor.Khaki;

                //creates subheaders
                ws.Cell("A1").Value = "Class Name (Extension Class)";
                ws.Cell("B1").Value = "(namespace:) Element Name";
                ws.Cell("C1").Value = "(namespace:) Element Type";
                ws.Cell("D1").Value = "Documentation ";
                ws.Cell("E1").Value = "Source ";
                ws.Range("A1:E1").Style.Font.Bold = true;
                ws.Range("A1:E1").Style.Fill.BackgroundColor = XLColor.LightBlue;
                ws.Range("A1:E1").Style.Font.FontSize = 16;
                ws.SheetView.FreezeRows(1);

                //Row Start
                int row = 4;
                int complex = doc.Descendants(xs + "complexType").Count();
                int simple = doc.Descendants(xs + "simpleType").Count();
                int elements = doc.Root.Elements(xs + "element").Count();
                int total = complex + simple + elements;
                decimal percentage = 100 / total;

                //complexType Process
                foreach (var el in doc.Descendants(xs + "complexType"))
                {
                    ws.Range("A" + row + ":E" + row).Style.Fill.BackgroundColor = XLColor.LightBlue;
                    backgroundWorker1.ReportProgress(total, total);
                    if (cancelJob(e)) break;

                    //Get extension class
                  //  if ((el.Elements(xs + "complexContent")).Count() != 0) not needed maybe
                    if (el.Elements(xs + "complexContent").Elements(xs + "extension").Count() != 0)
                            extensionClass = el.Elements(xs + "complexContent").Elements(xs + "extension").Single().Attribute("base").Value;

                    if (el.Elements(xs + "complexContent").Elements(xs + "restriction").Count() != 0)
                        extensionClass = el.Elements(xs + "complexContent").Elements(xs + "restriction").Single().Attribute("base").Value;

                    if (el.Elements(xs + "simpleContent").Elements(xs + "extension").Count() != 0)
                        extensionClass = el.Elements(xs + "simpleContent").Elements(xs + "extension").Single().Attribute("base").Value;


                    ws.Cell("A" + row).Value = el.Attribute("name").Value + " (" + extensionClass + ")"; // write class name and extension class
                    ws.Cell("A" + row).Style.Font.Bold = true;
                    extensionClass = "";

                    writeClassDocumentation(el, ws, row);


                    row++;

                    //write attributes from extension tag
                    foreach (var attr in el.Elements(xs + "complexContent").Elements(xs + "extension").Elements(xs + "sequence").Elements(xs + "element"))
                    {
                        if (cancelJob(e)) break;

                        writeElement(attr, row, ws, backgroundWorker1, total);
                        row++;
                    }

                    //write attributes from restriction tag
                    foreach (var attr in el.Elements(xs + "complexContent").Elements(xs + "restriction").Elements(xs + "sequence").Elements(xs + "element"))
                    {
                        if (cancelJob(e)) break;

                        writeElement(attr, row, ws, backgroundWorker1, total);
                        row++;
                    }
                    row++;
                }

                //simpleType Process
                foreach (var el in doc.Descendants(xs + "simpleType"))
                {
                    ws.Range("A" + row + ":E" + row).Style.Fill.BackgroundColor = XLColor.LightBlue;
                    backgroundWorker1.ReportProgress(total, total);
                    if (cancelJob(e)) break;

                    //Write simpleType name
                    ws.Cell("A" + row).Value = el.Attribute("name").Value + " (Enumerable)";
                    ws.Cell("A" + row).Style.Font.Bold = true;

                    writeClassDocumentation(el, ws, row);

                    row++;

                    foreach (var attr in el.Elements(xs + "restriction").Elements(xs + "enumeration"))
                    {
                        var edocu = "";

                        writeClassDocumentation(attr, ws, row);

                        //Documentation Enumerable
                        if (attr.Elements(xs + "annotation").Elements(xs + "documentation").Count() != 0)
                            edocu = attr.Elements(xs + "annotation").Elements(xs + "documentation").Single().Value;

                        var ename = attr.Attribute("value").Value;
                        ws.Cell("B" + row).Value = ename;
                        ws.Cell("C" + row).Value = "enumeration";
                        ws.Cell("D" + row).Value = edocu;
                        row++;
                    }
                    row++;
                }
                ws.Columns().AdjustToContents();
            }

            catch (documentationEntryException ex)
            {
                errorMsg = "Remove multiple documentation entry. You have three or more entries for documentation. At element: " + ex.Message;
            }
            catch (System.Runtime.InteropServices.COMException ex)
            {
                errorMsg = "You probably clicked on Excel or Closed it. Therefore the operation will now terminate.\nPlease restart the conversion.";
            }
            catch (Exception ex)
            {
                errorMsg = "The schema MAY be invalid. Please check the schema for conformance. If nothing is wrong, this may be a bug, please contact the developer." + ex.ToString();
            }

        }

        public bool cancelJob(DoWorkEventArgs e)
        {
            if ((backgroundWorker1.CancellationPending == true))
            {
                e.Cancel = true;
                return true;
            }
            return false;
        }

        public void writeElement(XElement attr, int row, IXLWorksheet ws, BackgroundWorker backgroundWorker1, int total)
        {
            //Element Name
            if (attr.Attribute("ref") != null)
            {
                elementName = attr.Attribute("ref").Value;
            }
            else if (attr.Attribute("name") != null)
            {
                elementName = attr.Attribute("name").Value;
            }
            else
            {
                elementName = "";
            }

            // elementName = attr.Attribute("ref") != null ? attr.Attribute("ref").Value : "";
            // elementName = attr.Attribute("name") != null ? attr.Attribute("name").Value : ;

            ws.Cell("B" + row).Value = elementName;

            //Element Type
            elementName = elementName.Substring(elementName.IndexOf(":") + 1);
            string elementType = searchForElementTypes(elementName, doc, xs, backgroundWorker1, total);
            ws.Cell("C" + row).Value = elementType;

            //Element Documentation/Source
            var tuple = searchForElementDocumentation(elementName, doc, xs);
            ws.Cell("D" + row).Value = tuple.Item1;
            ws.Cell("E" + row).Value = tuple.Item2;
        }

        private void backgroundWorker1_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            // The progress percentage is a property of e
            progressBar1.Maximum = Convert.ToInt32(e.UserState);
            progressBar1.Increment(1);
            label3.Text = (progressBar1.Value * 100) / progressBar1.Maximum + "%";
        }

        private void backgroundWorker1_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            progressBar1.Maximum = 1;
            var outputFilePath = "";
            try
            {
                outputFilePath = outputPath + "\\" + filename + ".xlsx";
                wb.SaveAs(outputFilePath);

             
            }
            catch (Exception ex)
            {
                errorMsg = "Unable to save. Please close Excel Sheet. \n " + ex.Message;
            }

            label3.Text = "100%";
            createExcel.Enabled = true;
            selectXSD.Enabled = true;
            label2.Text = "";
            label1.Text = "";

            if (errorMsg != "")
            {
                MessageBox.Show(new Form() { TopMost = true }, errorMsg, "Fatal Error!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                label2.Text = "An Error Ocurred. Please try again.";
            }
            else
            {
                if (!jobCanceled)
                {
                    MessageBox.Show(new Form() { TopMost = true }, "Done! \nThe spreadsheet is located at: \n" + outputFilePath, "Operation Succesfull", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                    label1.Text = "Done";
                    Process.Start(@"" + outputPath);
                }
                else
                {
                    MessageBox.Show(new Form() { TopMost = true }, "Operation was Canceled!", "Operation Canceled", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                    label1.Text = "Canceled!";
                }
            }
        }

        private void button1_Click_1(object sender, EventArgs e)
        {
            backgroundWorker1.WorkerSupportsCancellation = true;
            jobCanceled = true;
            backgroundWorker1.CancelAsync();
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }
    }
}
