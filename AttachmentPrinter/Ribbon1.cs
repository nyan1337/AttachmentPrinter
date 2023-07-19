using Microsoft.Office.Interop.Outlook;
using Microsoft.Office.Tools.Ribbon;
using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing.Printing;
using System.IO;
using System.Linq;
using System.Security.Policy;
using System.Text;
using System.Threading;
using System.Windows.Forms;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace AttachmentPrinter
{
    public partial class Ribbon1
    {
        const string outputFileName = "attachments.pdf";

        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            Outlook.Application outlookApp = new Outlook.Application();
            Outlook.Explorer explorer = outlookApp.ActiveExplorer();
            Outlook.Selection selectedItems = explorer.Selection;

            if (selectedItems.Count > 0)
            {
                string tempPath = System.IO.Path.GetTempPath();
                int counter = 0;
                List<string> inputFilePaths = new List<string>();

                foreach (object selectedItem in selectedItems)
                {
                    if (selectedItem is Outlook.MailItem mailItem)
                    {
                        if (mailItem.Attachments.Count > 0)
                        {
                            foreach (Outlook.Attachment attachment in mailItem.Attachments)
                            {
                                if (attachment.Type == Outlook.OlAttachmentType.olByValue && attachment.FileName.EndsWith(".pdf"))
                                {
                                    string tempFilePath = tempPath + "file" + counter + ".pdf";
                                    attachment.SaveAsFile(tempFilePath);
                                    inputFilePaths.Add(tempFilePath);
                                    counter++;
                                }
                            }
                        }
                    }
                }

                if (inputFilePaths.Count > 0)
                {
                    using (PdfDocument outPdf = new PdfDocument())
                    {
                        foreach (string path in inputFilePaths)
                        {
                            CopyPages(PdfReader.Open(path, PdfDocumentOpenMode.Import), outPdf);
                            File.Delete(path);
                        }

                        outPdf.Save(tempPath + outputFileName);
                    }

                    Process p = new Process();
                    p.StartInfo = new ProcessStartInfo()
                    {
                        FileName = tempPath + outputFileName //put the correct path here
                    };
                    p.Start();
                }
                else
                {
                    MessageBox.Show("Es wurden keine PDF-Anhänge in den ausgewählten E-Mails gefunden", "AttachmentPrinter: Fehler", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }

            }
            else
            {
                MessageBox.Show("Wählen Sie E-Mails aus, um fortzufahren", "AttachmentPrinter: Fehler", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        void CopyPages(PdfDocument from, PdfDocument to)
        {
            for (int i = 0; i < from.PageCount; i++)
            {
                to.AddPage(from.Pages[i]);
            }
        }


    }
}
