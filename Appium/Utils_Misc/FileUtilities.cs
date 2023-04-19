using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Text;
using Appium.Utilities;
using iTextSharp.text.pdf;
using NPOI.SS.Formula.Functions;
using Excel = Microsoft.Office.Interop.Excel;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace UIAutomation.Utils_Misc
{
    public class FileUtilities : Base
    {
        private List<string> excelExtensions = new List<string> {"xls", "xlsx"};
        private List<string> wordExtensions = new List<string> {"doc", "docx"};
        private List<string> powerPointExtensions = new List<string> {"ppt", "pptx"};

        private FileForCompare File1 { get; set; }
        private FileForCompare File2 { get; set; }

        //Constructor
        public FileUtilities(string filePath1, string filePath2)
        {
            File1 = new FileForCompare(filePath1);
            File2 = new FileForCompare(filePath2);
        }

        public FileUtilities(string filePath)
        {
            File1 = new FileForCompare(filePath);
        }

        public bool CheckExcelContainsText(string text)
        {
            Excel.Application excelApp = new Excel.Application();
            Excel.Workbook excelWorkBook;
            bool result = false;
            try
            {
                Log.Debug("Opening Microsoft Excel");
                excelApp.Visible = true;
                excelApp.DisplayAlerts = false;
                log.Debug("Opening document");
                excelWorkBook = excelApp.Workbooks.Open(File1.Path, ReadOnly: false);
                excelWorkBook.Activate();
                // search text in worksheets
                foreach (Excel.Worksheet excelWorkSheet in excelWorkBook.Worksheets)
                {
                    Excel.Range oRng = GetSpecifiedRange(text, excelWorkSheet);
                    if (oRng != null)
                    {
                        result = true;
                        break;
                    }
                }
                log.Debug("Saving document");
                excelWorkBook.Close(Excel.XlSaveAction.xlSaveChanges);
                log.Info(MethodBase.GetCurrentMethod().Name + " completed");
                return result;
            }
            catch (Exception e)
            {
                throw e;
            }
            finally
            {
                log.Debug("closing Microsoft Excel");
                try
                {
                    excelApp.Quit();
                }
                catch (Exception)
                {
                    // ignored
                }
            }
        }

        public string ExtractTextFromPDF(string pdfFileName)
        {
            StringBuilder result = new StringBuilder();
            // Create a reader for the given PDF file
            using (PdfReader reader = new PdfReader(pdfFileName))
            {
                // Read pages
                for (int page = 1; page <= reader.NumberOfPages; page++)
                {
                    SimpleTextExtractionStrategy strategy =
                        new SimpleTextExtractionStrategy();
                    string pageText =
                        PdfTextExtractor.GetTextFromPage(reader, page, strategy);
                    result.Append(pageText);
                }
            }
            return result.ToString();
        }
        
        // search range with needed text in excel worksheet
        private static Excel.Range GetSpecifiedRange(string matchStr, Excel.Worksheet objWs)
        {
            object missing = Missing.Value;
            Excel.Range currentFind = null;
            //Microsoft.Office.Interop.Excel.Range firstFind = null;
            currentFind = objWs.get_Range("A1", "AM100").Find(matchStr, missing,
                Excel.XlFindLookIn.xlValues,
                Excel.XlLookAt.xlPart,
                Excel.XlSearchOrder.xlByRows,
                Excel.XlSearchDirection.xlNext, false, missing, missing);
            return currentFind;
        }

        public bool CheckFilesIdentical()
        {
            if (!File1.Extension.Equals(File2.Extension))
            {
                log.Info("Files has different extensions.");
                return false;
            }

            bool result;
            if (wordExtensions.Contains(File1.Extension))
            {
                result = CheckWordFilesIdentical(File1.Path, File2.Path);
            }
            else if (powerPointExtensions.Contains(File1.Extension))
            {
                result = CheckPowerPointFilesIdentical(File1.Path, File2.Path);
            }
            //                else if (file1Extension.Equals("pdf"))
            //                {
            //                    result = CheckPDFFilesIdentical(file1, file2);
            //                }else if (excelExtensions.Contains(file1Extension))
            //                {
            //                    result = CheckExcelFilesIdentical(file1, file2);
            //                }
            else
            {
                throw new Exception("Files are not supported");
            }
            log.Info($"Files identical: [{result}]");
            return result;
        }
        
        private bool CheckWordFilesIdentical(string file1, string file2)
        {
            var result = false;
            var expected = File.ReadAllBytes(file1);
            var actual = File.ReadAllBytes(file2);
            log.Debug($"Loading file one: [{file1}]");
            var expectedResult = new WmlDocument("expected.docx", expected);
            log.Debug($"Loading file two: [{file2}]");
            var actualDocument = new WmlDocument("result.docx", actual);
            var comparisonSettings = new WmlComparerSettings();
            log.Debug("Comparing files");
            var comparisonResults = WmlComparer.Compare(expectedResult, actualDocument, comparisonSettings);
            var revisions = WmlComparer.GetRevisions(comparisonResults, comparisonSettings);
            log.Debug("Revisions count = " + revisions.Count);
            if (revisions.Count == 0)
            {
                result = true;
            }
            else
            {
                log.Error("File differs found:\n");
                foreach (var revision in revisions)
                {
                    log.Error(revision.Text);
                }
            }
            return result;
        }
        
        private bool CheckPowerPointFilesIdentical(string file1, string file2)
        {
            PowerPoint.Application PowerPoint_App;
            PowerPoint_App = new PowerPoint.Application();
            PowerPoint.Presentations multi_presentations = PowerPoint_App.Presentations;
            PowerPoint.Presentation presentation1 = multi_presentations.Open(file1);
            PowerPoint.Presentation presentation2 = multi_presentations.Open(file2);
            string presentation1_text = Helpers.GetPresentationText(presentation1);
            string presentation2_text = Helpers.GetPresentationText(presentation2);
            presentation1.Close();
            presentation2.Close();
            System.Runtime.InteropServices.Marshal.FinalReleaseComObject(presentation1);
            System.Runtime.InteropServices.Marshal.FinalReleaseComObject(presentation2);
            Helpers.ClosePowerPoint(PowerPoint_App);
            return presentation1_text.Equals(presentation2_text);
        }  

        private class FileForCompare
        {
            public string Path { get; }
            public string Name { get; }
            public string Extension { get; }

            public FileForCompare(string filePath)
            {
                if (!new FileInfo(filePath).Exists)
                {
                    throw new FileNotFoundException($"File [{filePath}] not found.");
                }

                Path = filePath;
                Name = System.IO.Path.GetFileName(Path);
                Extension = System.IO.Path.GetExtension(Path).Replace(".", "").ToLower();
            }
        }

        private class Helpers
        {
            protected internal static string GetPresentationText(PowerPoint.Presentation presentation)
            {
                string presentation_text = "";
                for (int i = 0; i < presentation.Slides.Count; i++)
                {
                    foreach (var item in presentation.Slides[i + 1].Shapes)
                    {
                        var shape = (PowerPoint.Shape)item;
                        if (shape.HasTextFrame == MsoTriState.msoTrue)
                        {
                            if (shape.TextFrame.HasText == MsoTriState.msoTrue)
                            {
                                var textRange = shape.TextFrame.TextRange;
                                var text = textRange.Text;
                                presentation_text += text + " ";
                            }
                        }
                    }
                }
                return presentation_text;
            }
            
                    
            protected internal static void ClosePowerPoint(PowerPoint._Application powerPointApp)
            {
                powerPointApp?.Quit();
                //Commented out lines could break tests run in parallel. To discuss: Do we need it. If yes - make it with some other way
//                Process[] processes = Process.GetProcessesByName("powerpnt");
//                System.Runtime.InteropServices.Marshal.FinalReleaseComObject(powerPointApp);
//                for (int i = 0; i < processes.Count(); i++)
//                {
//                    try
//                    {
//                        processes[i].Kill();
//                    }
//                    catch (Exception)
//                    {
//                        // ignored
//                    }
//                }
            }
        }
    }
}