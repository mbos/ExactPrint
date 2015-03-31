using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.IO;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Text.RegularExpressions;
using Microsoft.Office.Interop.Word; 

namespace ExactPrint
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                string text = File.ReadAllText(openFileDialog1.FileName);
                text = Regex.Replace(text, @"[^\u0009-\u0010|^\u0020-\u007E]", string.Empty);
  
                object missing = System.Reflection.Missing.Value;
                object Visible=true;
                object start1 = 0;
                object end1 = 0;

                Microsoft.Office.Interop.Word.Application WordApp = new Microsoft.Office.Interop.Word.Application();
                Document adoc = WordApp.Documents.Add(ref missing, ref missing, ref missing, ref missing);
                Range rng = adoc.Range(ref start1, ref missing);

 

                try
                {              
                    rng.Font.Size = 5;
                    rng.Font.Name = "Courier New";
                    // rng.PageSetup.Orientation = WdOrientation.wdOrientLandscape;
                    rng.InsertAfter(text);
                    Microsoft.Office.Interop.Word.Paragraphs paragraphs = adoc.Paragraphs;
                    foreach (Microsoft.Office.Interop.Word.Paragraph paragraph in paragraphs)
                    {
                        if (paragraph.Range.Text.Trim() == string.Empty)
                        {
                            paragraph.Range.Select();
                            WordApp.Selection.Delete();
                        }
                    } 
                    WordApp.Visible = true;
                }

                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }        
            }
            System.Windows.Forms.Application.Exit();
            
        }
    }
}
