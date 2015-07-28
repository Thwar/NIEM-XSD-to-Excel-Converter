using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
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
        }

        string elementName = "";
        string extensionClass = "";
        XNamespace xs = XNamespace.Get("http://www.w3.org/2001/XMLSchema");
        XDocument doc;
        bool jobCanceled = false;
        string errorMsg;

        private void button1_Click(object sender, EventArgs e)
        {

        }

        //Get Element Types
        public static string searchForElementTypes(string elementName, XDocument doc, XNamespace xs)
        {
            foreach (var el in doc.Root.Elements(xs + "element"))
            {
                if (el.Attribute("name").Value == elementName)
                {
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
                    //Check if Documentation exists for Class

                    var documentation = el.Elements(xs + "annotation").Elements(xs + "documentation").ToList();

                    //Check if more than one Documentation entry exists and throw error
                    if (documentation.Count() > 2)
                        throw new documentationEntryException(elementName);

                    //Get Documentation
                    if (documentation.Count() == 1)
                        description =  documentation[0].Value;

                    //Get Source
                    if (documentation.Count() == 2)
                        source =  documentation[1].Value;

                }
            }
            return new Tuple<string, string>(description, source);
        }

        private void openFileDialog1_FileOk(object sender, CancelEventArgs e)
        {

        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        public class documentationEntryException : Exception
        {
            public documentationEntryException(string message)  : base(message) { }
        }

        private void selectXSD_Click(object sender, EventArgs e)
        {
            openFileDialog1.Filter = "XSD NIEM (.xsd)|*.xsd";
            openFileDialog1.FilterIndex = 1;

            DialogResult result = openFileDialog1.ShowDialog(); // Show the dialog.
            if (result == DialogResult.OK) // Test result.
            {
                string file = openFileDialog1.FileName;
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
            MessageBox.Show("This tool is used to convert NIEM XSD schemas into Excel Spreadsheets. \n\nAuthor: Ruben T. Rosales\nVersion v1.0", "About NIEM XSD to Excel Converter ", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
        }

        private void createExcel_Click(object sender, EventArgs e)
        {
            progressBar1.Value = 0;
            progressBar1.Style = ProgressBarStyle.Marquee;
            progressBar1.MarqueeAnimationSpeed = 30;
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
            CreateExcelDoc excell_app = new CreateExcelDoc();
            doc = XDocument.Load(openFileDialog1.FileName);

            //creates the main header
            excell_app.createHeaders(2, 1, openFileDialog1.FileName, "A2", "D2", 2, "YELLOW", true, 10, "n");
            //creates subheaders
            excell_app.createHeaders(1, 1, "Class Name (Extension Class)", "A1", "A1", 0, "GRAY", true, 10, "");
            excell_app.createHeaders(1, 2, "Element Name", "B1", "B1", 0, "GRAY", true, 10, "");
            excell_app.createHeaders(1, 3, "Element Type", "C1", "C1", 0, "GRAY", true, 10, "");
            excell_app.createHeaders(1, 4, "Documentation ", "D1", "D1", 0, "GRAY", true, 10, "");
            excell_app.createHeaders(1, 5, "Source ", "E1", "E1", 0, "GRAY", true, 10, "");

            //Row Start
            int row = 3;
            string documentation = "";


                foreach (var el in doc.Descendants(xs + "complexType"))
                {
                    if ((backgroundWorker1.CancellationPending == true))
                    {
                        e.Cancel = true;
                        break;
                    }
                  
                    //Write Class Name
                    if ((el.Elements(xs + "complexContent")).Count() != 0)
                     extensionClass = el.Elements(xs + "complexContent").Elements(xs + "extension").Single().Attribute("base").Value;

                    excell_app.addData(row, 1, el.Attribute("name").Value + " (" + extensionClass + ")", "A" + row, "A" + row, "");

                    //Check if Documentation exists for Class
                    if (el.Elements(xs + "annotation").Elements(xs + "documentation").Count() != 0)
                    {
                        documentation = el.Elements(xs + "annotation").Elements(xs + "documentation").Single().Value;
                        //Write Documentation for Class
                        excell_app.addData(row, 4, documentation, "D" + row, "D" + row, "");
                    }

                    row++;

                    foreach (var attr in el.Elements(xs + "complexContent").Elements(xs + "extension").Elements(xs + "sequence").Elements(xs + "element"))
                    {
                        if ((backgroundWorker1.CancellationPending == true))
                        {
                            e.Cancel = true;
                            break;
                        }

                        //Element Name
                        elementName = attr.Attribute("ref") != null ? attr.Attribute("ref").Value : "";
                        excell_app.addData(row, 2, elementName, "B" + row, "B" + row, "");

                        //Element Type
                        elementName = elementName.Substring(elementName.IndexOf(":") + 1);                      
                        string elementType = searchForElementTypes(elementName, doc, xs);
                        excell_app.addData(row, 3, elementType, "C" + row, "C" + row, "");

                        //Element Documentation/Source
                        var tuple = searchForElementDocumentation(elementName, doc, xs);
                        excell_app.addData(row, 4,  tuple.Item1, "D" + row, "D" + row, "");
                        excell_app.addData(row, 5, tuple.Item2, "E" + row, "E" + row, "");                                   
                        
                        row++;
                    }
                    excell_app.createHeaders(row, 2, "", "A" + row, "D" + row, 2, "GAINSBORO", true, 10, "");
                    row++;
                }
            }

            catch (documentationEntryException ex)
            {
                errorMsg = "Remove multiple documentation entry. You have three or more entries for documentation. At element: " + ex.Message;
            }
            catch (System.Runtime.InteropServices.COMException ex)
            {
                errorMsg = "You probably clicked on Excel or Closed it. Therefore the operation will now terminate.\nPlease restart the conversion.";
            }
            catch(Exception ex) 
            {
               errorMsg = ex.ToString();
            }
          
        }

        private void backgroundWorker1_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
        
            progressBar1.Style = ProgressBarStyle.Continuous;
            createExcel.Enabled = true;
            selectXSD.Enabled = true;
            label2.Text = "";
            label1.Text = "";

            if (errorMsg != "")
            {
                MessageBox.Show(errorMsg, "Fatal Error!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                label2.Text = "An Error Ocurred. Please try again.";
            }
            else
            {
                if (!jobCanceled)
                {
                    MessageBox.Show("Done!", "Operation Succesfull", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                    label1.Text = "Done!";
                }
                else
                {
                    MessageBox.Show("Operation was Canceled!", "Operation Canceled", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
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
