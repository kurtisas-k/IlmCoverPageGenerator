using Newtonsoft.Json;
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
using Microsoft.Office.Interop.Word;
using Application = Microsoft.Office.Interop.Word.Application;

namespace IlmCoverPageGenerator
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            var sectionData = loadSectionData();
            // prepare cover data
            var prepareCover = true;
            prepareCoverData(prepareCover,sectionData);
            // get data
            List<ModuleInfo> data = getCoverData();
            
            // update documents
            updateDocuments(data);
        }

        private Dictionary<string,string> loadSectionData()
        {
            var outFile = new Dictionary<string, string>();
            var path = @"C:\Users\kstaples\source\repos\GenerateModuleToSectionTable\GenerateModuleToSectionTable\bin\Debug\sectionTextData.txt";
            var lines = File.ReadAllLines(path);
            for(var i = 0; i<lines.Length; i++)
            {
                var line = lines[i];
                var lineData = line.Split('\t');
                var moduleNumber = lineData[0];
                var moduleSection = lineData[1];
                try
                {
                    outFile.Add(moduleNumber, moduleSection);
                }
                catch(Exception ex)
                {
                    Console.WriteLine(ex.Message);
                }
            }
            return outFile;
        }

        private List<ModuleInfo> getCoverData()
        {
            var inFilePath = @"C:\Users\kstaples\Documents\Projects\ILM Covers\CoverData.json";
            var data = JsonConvert.DeserializeObject<List<ModuleInfo>>(File.ReadAllText(inFilePath));
            return data;
        }

        private void updateDocuments(List<ModuleInfo> moduleList)
        {
            // get list of files
            var files = Directory.GetFiles(@"E:\V21", "*.docx",SearchOption.AllDirectories);
            // for all files try and find associative module
            for(var i = 0; i < files.Length; i++)
            {
                // extract key per file and get module information
                var file = files[i];
                int start = file.LastIndexOf("\\");
                var moduleKeyWithExtension = file.Substring(start + 1);
                var moduleKey = moduleKeyWithExtension.Substring(0, moduleKeyWithExtension.LastIndexOf("p"));
                var module = getModuleByKey(moduleKey, moduleList);
                FileInfo fi = new FileInfo(file);
                var nm = fi.Name;
                if(nm[0] == '~') { continue; }
                //if (module.ModuleShortcode != "MIL_1") { continue; }
                Directory.CreateDirectory(@"C:\Users\kstaples\Documents\Projects\Update ILMS\" + module.ModuleShortcode);
                var path = @"C:\Users\kstaples\Documents\Projects\Update ILMS\"+ module.ModuleShortcode +"\\" + fi.Name.Replace(".docx", "_updated.docx");
                var fileExists = File.Exists(path);
                if (fileExists) { continue; };
                
                var frontCover = createFrontCover(module);
                var backCover = createBackCover(module);
                try
                {
                    updateDocument(file, frontCover, backCover, module);
                }
                catch(Exception ex)
                {
                    updateDocument(file, frontCover, backCover, module);
                }
            }
        }

        private string createBackCover(ModuleInfo module)
        {
            Application wrdApp = new Application();
            wrdApp.Visible = false;
            
            var root = @"C:\Users\kstaples\Documents\Projects\Update ILMS\";
            var outPath = @root+module.ModuleNumber+"_backcover.docx";
            if (File.Exists(outPath)) { return outPath; }
            // template path
            var backCoverTemplatePath = @"C:\Users\kstaples\Documents\Projects\Update ILMS\Cover Templates\ILM Example Back Cover BW-Rev3.docx";
            // open template
            var backCoverDoc = wrdApp.Documents.Open(backCoverTemplatePath, false, true);
            backCoverDoc.Activate();
            // swap values
            FindAndReplace(wrdApp, "Module Number | Version", module.ModuleNumber + " | Version 21");
            // populate fields with data
            // save as outpath
            wrdApp.ActiveDocument.SaveAs(outPath);
            wrdApp.ActiveDocument.Close();
            wrdApp.Quit(false);
            GC.Collect();
            GC.WaitForPendingFinalizers();
            // return outpath if file exists
            if (File.Exists(outPath))
            {
                return outPath;
            }
            else return "";
        }

        private string createFrontCover(ModuleInfo module)
        {
            Application wrdApp = new Application();
            wrdApp.Visible = false;

            var root = @"C:\Users\kstaples\Documents\Projects\Update ILMS\";
            var outPath = @root + module.ModuleNumber + "_frontcover.docx";
            if (File.Exists(outPath)) { return outPath; }
            // template path
            var frontCoverTemplatePath = @"C:\Users\kstaples\Documents\Projects\Update ILMS\Cover Templates\ILM Example Front Cover BW-Rev3.docx";
            // open template
            var frontCoverDoc = wrdApp.Documents.Open(frontCoverTemplatePath, false, true);
            try
            {
                frontCoverDoc.Activate();
            }
            catch(Exception ex)
            {
                createFrontCover(module);
            }
            // swap values
            FindAndReplace(wrdApp, "Module Name", module.ModuleTitle);
            FindAndReplace(wrdApp, "MODULE", module.ModuleNumber);
            FindAndReplace(wrdApp, "TRADE", module.ModuleTrade);
            FindAndReplace(wrdApp, "SECTION", module.ModuleSection);
            FindAndReplace(wrdApp, "PERIOD", module.ModulePeriod);            // populate fields with data
            // save as outpath
            frontCoverDoc.SaveAs(outPath);
            frontCoverDoc.Close(WdSaveOptions.wdDoNotSaveChanges);
            wrdApp.Quit(false);
            GC.Collect();
            GC.WaitForPendingFinalizers();
            // return outpath if file exists
            if (File.Exists(outPath))
            {
                return outPath;
            }
            else return "";
        }

        private void updateDocument(string filePath, string frontCoverPath, string backCoverPath, ModuleInfo moduleInfo)
        {
            Application wrdApp = new Application();
            wrdApp.Visible = false;
            var moduleDoc = wrdApp.Documents.Open(filePath, false, false);
            
            var frontCoverDoc = wrdApp.Documents.Open(frontCoverPath, false, false);
            var backCoverDoc = wrdApp.Documents.Open(backCoverPath, false, false);

            /* remove existing content from covers*/
            moduleDoc.Activate();
            removeFirstTwoPages(wrdApp);
            removeLastTwoPages(wrdApp);

            frontCoverDoc.Activate();
            wrdApp.ActiveDocument.Content.Copy();

            moduleDoc.Content.Characters.First.Select();
            wrdApp.Selection.Collapse();
            wrdApp.Selection.Paste();

            backCoverDoc.Activate();
            wrdApp.ActiveDocument.Content.Copy();

            moduleDoc.Content.Characters.Last.Select();
            wrdApp.Selection.Collapse();
            wrdApp.Selection.Paste();
            removeAndUnlinkHeadersAndFootersFromFinalSection(moduleDoc);
            changeMargin(moduleDoc);
            frontCoverDoc.Close(WdSaveOptions.wdDoNotSaveChanges);
            backCoverDoc.Close(WdSaveOptions.wdDoNotSaveChanges);
            var path = @"C:\Users\kstaples\Documents\Projects\Update ILMS\" + moduleInfo.ModuleShortcode + "\\" + moduleDoc.Name.Replace(".docx", "_updated.docx");
            moduleDoc.SaveAs(path);
            moduleDoc.Close(WdSaveOptions.wdDoNotSaveChanges);
            wrdApp.Quit(false);
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }

        private void changeMargin(Document module)
        {
            module.Content.Sections.Last.Range.PageSetup.HeaderDistance = 0;
            module.Content.Sections.Last.Range.PageSetup.FooterDistance = 0;
        }

        private void removeAndUnlinkHeadersAndFootersFromFinalSection(Document module)
        {
            module.Content.Characters.Last.Delete();
            HeadersFooters headers = module.Content.Sections.Last.Headers;
            HeadersFooters footers = module.Content.Sections.Last.Footers;
            headers[WdHeaderFooterIndex.wdHeaderFooterPrimary].LinkToPrevious = false;
            headers[WdHeaderFooterIndex.wdHeaderFooterEvenPages].LinkToPrevious = false;
            headers[WdHeaderFooterIndex.wdHeaderFooterFirstPage].LinkToPrevious = false;
            footers[WdHeaderFooterIndex.wdHeaderFooterPrimary].LinkToPrevious = false;
            footers[WdHeaderFooterIndex.wdHeaderFooterEvenPages].LinkToPrevious = false;
            footers[WdHeaderFooterIndex.wdHeaderFooterFirstPage].LinkToPrevious = false;
            module.Content.Sections.Last.Headers[WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.Delete();
            module.Content.Sections.Last.Headers[WdHeaderFooterIndex.wdHeaderFooterEvenPages].Range.Delete();
            module.Content.Sections.Last.Headers[WdHeaderFooterIndex.wdHeaderFooterFirstPage].Range.Delete();
            module.Content.Sections.Last.Footers[WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.Delete();
            module.Content.Sections.Last.Footers[WdHeaderFooterIndex.wdHeaderFooterEvenPages].Range.Delete();
            module.Content.Sections.Last.Footers[WdHeaderFooterIndex.wdHeaderFooterFirstPage].Range.Delete();
            module.PrintPreview();
        }

        private void removeFirstTwoPages(Application wrdApp)
        {
            selectFirstTwoPages(wrdApp);
            wrdApp.Selection.Range.Delete();
        }

        private void removeAllLocks(Application wrdApp)
        {
            //foreach(Shape shape in wrdApp.ActiveDocument.Shapes)
            //{
            //    shape.Anchor.Delete();
                
            //}

            foreach(ContentControl control in wrdApp.ActiveDocument.ContentControls)
            {
                control.LockContentControl = false;
            }
        }

        private void removeLastTwoPages(Application wrdApp)
        {
            removeAllLocks(wrdApp);
            selectLastTwoPages(wrdApp);
            wrdApp.Selection.Range.Delete();
        }

        private void selectLastTwoPages(Application wrdApp)
        {
            object what = WdGoToItem.wdGoToPage;
            object which = WdGoToDirection.wdGoToAbsolute;

            wrdApp.Visible = false;
            object readOnly = false;
            object missing = System.Reflection.Missing.Value;

            object count = wrdApp.ActiveDocument.ComputeStatistics(Microsoft.Office.Interop.Word.WdStatistic.wdStatisticPages) - 1;
            Range startRange = wrdApp.Selection.GoTo(ref what, ref which, ref count, ref missing);
            object count2 = (int)count + 2;
            wrdApp.ActiveDocument.Characters.Last.Select();
            Range endRange = wrdApp.Selection.Range;
            //if you want to select last page
            if (endRange.Start == startRange.Start)
            {
                which = WdGoToDirection.wdGoToLast;
                what = WdGoToItem.wdGoToObject;
                endRange = wrdApp.Selection.GoTo(ref what, ref which, count2, ref missing);
            }

            endRange.SetRange(startRange.Start, endRange.End);
            endRange.Select();
        }



        public void selectFirstTwoPages(Application wrdApp)
        {
            object what = WdGoToItem.wdGoToPage;
            object which = WdGoToDirection.wdGoToAbsolute;
            object count = 1;

            object readOnly = false;
            object missing = System.Reflection.Missing.Value;


            Range startRange = wrdApp.Selection.GoTo(ref what, ref which, ref count, ref missing);
            object count2 = (int)count + 2;
            Range endRange = wrdApp.Selection.GoTo(ref what, ref which, ref count2, ref missing);
            endRange.SetRange(startRange.Start, endRange.End);
            endRange.Select();
        }

        public void deleteSelection(Application wrdApp)
        {
            wrdApp.Selection.Delete();
        }

        private ModuleInfo getModuleByKey(string moduleKey, List<ModuleInfo> moduleList)
        {
            ModuleInfo module = moduleList.Where(i => i.ModuleNumber == moduleKey).FirstOrDefault();
            return module;
        }

        private void createCover(ModuleInfo module, string docPath)
        {
            Microsoft.Office.Interop.Word.Application wrdApp = new Microsoft.Office.Interop.Word.Application();
            wrdApp.Visible = false;
            // get section text from V19 cover
            // if file exists create cover
            // if file does not exist move on
            Document moduleDocument = wrdApp.Documents.Open(docPath,false,true);
            // open template document and go through each paragraph
            //var sectionText = getSectionText(wrdApp, moduleDocument);

            // create cover for module
            docPath = @"C:\Users\kstaples\Documents\Cover Template BW Revised.dotx";
            populateTemplateValues(wrdApp, docPath, module);
            wrdApp.ActiveDocument.SaveAs("C:\\Users\\kstaples\\Documents\\ilm_covers\\"+module.ModuleNumber+"_cover.docx");
        }

        private void populateTemplateValues(Microsoft.Office.Interop.Word.Application wrdApp, string docPath, ModuleInfo module)
        {
            Document activeDocument = wrdApp.Documents.Open(docPath, false, false);
            activeDocument.Content.Select();
            FindAndReplace(wrdApp, "Module Number | Version", module.ModuleNumber + " | 21");
            FindAndReplace(wrdApp, "Module Name", module.ModuleTitle);
            FindAndReplace(wrdApp, "MODULE", module.ModuleNumber);
            FindAndReplace(wrdApp, "TRADE", module.ModuleTrade);
            FindAndReplace(wrdApp, "SECTION", module.ModuleSection);
            FindAndReplace(wrdApp, "PERIOD", module.ModulePeriod);
            // populate fields with data
        }
        

        private string getSectionText(Microsoft.Office.Interop.Word.Application wrdApp, Document activeDocument)
        {
            activeDocument.Content.Select();
            for (var i = 1; i < wrdApp.Selection.ContentControls.Count; i++)
            {
                var contentControl = wrdApp.Selection.ContentControls[i];
                Console.WriteLine("CC"+contentControl.Title);
                if (contentControl.Title == "Section")
                {
                    var x = contentControl;
                    var SectionText = x.Range.Text;
                    activeDocument.Close(false);
                    return SectionText.Replace("  "," ");
                }
            }
            MessageBox.Show("Section not found.");
            return "";
        }

        private void FindAndReplace(Microsoft.Office.Interop.Word.Application doc, object findText, object replaceWithText)
        {
            //options
            object matchCase = true;
            object matchWholeWord = false;
            object matchWildCards = false;
            object matchSoundsLike = false;
            object matchAllWordForms = false;
            object forward = true;
            object format = false;
            object matchKashida = false;
            object matchDiacritics = false;
            object matchAlefHamza = false;
            object matchControl = false;
            object read_only = false;
            object visible = true;
            object replace = 2;
            object wrap = 1;
            //execute find and replace
            try
            {
                doc.ActiveDocument.Content.Select();
                doc.Selection.Find.Execute(ref findText, ref matchCase, ref matchWholeWord,
                    ref matchWildCards, ref matchSoundsLike, ref matchAllWordForms, ref forward, ref wrap, ref format, ref replaceWithText, ref replace,
                    ref matchKashida, ref matchDiacritics, ref matchAlefHamza, ref matchControl);
            }
            catch (Exception ex) {
                FindAndReplace(doc, findText, replaceWithText);
            }
        }

        private void prepareCoverData(bool v, Dictionary<string, string> sectionData)
        {
            if (v == true)
            {
                List<ModuleInfo> allModulesInfo = new List<ModuleInfo>();
                var allData = File.ReadAllLines(@"C:\Users\kstaples\Documents\Projects\ILM Script\ILM\ILM_ModuleList_ExtractTradeSecrets.csv");
                for (var i = 0; i < allData.Length; i++)
                {
                    var recordData = allData[i].Split(',');
                    var moduleNumber = recordData[0];
                    var modulePeriod = recordData[1];
                    var moduleTrade = recordData[2];
                    var moduleTitle = recordData[3].Replace("[comma]", ",");
                    var moduleSection = getSectionData(recordData[0],sectionData);

                    var module = new ModuleInfo(moduleNumber, modulePeriod, moduleTrade, moduleTitle, moduleSection);
                    allModulesInfo.Add(module);
                }
                var outData = JsonConvert.SerializeObject(allModulesInfo);

                var outFilePath = @"C:\Users\kstaples\Documents\Projects\ILM Covers\CoverData.json";
                using (System.IO.StreamWriter file = new System.IO.StreamWriter(outFilePath))
                {
                    file.Write(outData);
                }
            }
            else { };
        }

        private string getSectionData(string key, Dictionary<string, string> sectionData)
        {
            string sectionText;

            try
            {
                sectionText = sectionData[key];
            }
            catch
            {
                sectionText = "";
            }

            return sectionText;
        }

        private void copyFirstTwoPages(Microsoft.Office.Interop.Word.Application wrdApp) {
            object what = WdGoToItem.wdGoToPage;
            object which = WdGoToDirection.wdGoToAbsolute;
            object count = 2;

            const string fileName = @"C:\Users\kstaples\Documents\ilm_covers\100101c_cover.dotx";
            object fileNameAsObject = fileName;
            object missing = System.Reflection.Missing.Value;
            Range startRange = wrdApp.Selection.GoTo(ref what, ref which, ref count, ref missing);
            object count2 = (int)count + 2;
            Range endRange = wrdApp.Selection.GoTo(ref what, ref which, ref count2, ref missing);

            endRange.SetRange(startRange.Start, endRange.End);
            endRange.Select();
            wrdApp.Selection.Copy();
        }

        private void copyLastTwoPages(Microsoft.Office.Interop.Word.Application wrdApp)
        {
            FileInfo fi = new FileInfo(@"C:\Users\kstaples\Documents\ilm_covers\100101c_cover.dotx");

        }

        public class ModulesInfo
        {
            public ModuleInfo record;
        }

        [Serializable]
        public class ModuleInfo
        {
            public string ModuleTrade { get; set; }
            public string ModuleNumber { get; set; }
            public string ModuleTitle { get; set; }
            public string ModulePeriod { get; set; }
            public string ModuleSection { get; set; }
            public string ModuleShortcode { get; set; }

            public ModuleInfo(string moduleNumber, string moduleShortcode, string moduleTitle, string moduleTrade, string moduleSection)
            {
                ModuleNumber = moduleNumber;
                ModulePeriod = ToPeriodString(moduleShortcode);
                ModuleTitle = moduleTitle;
                ModuleTrade = moduleTrade;
                ModuleSection = moduleSection;
                ModuleShortcode = moduleShortcode;
            }

            private string ToPeriodString(string modulePeriod)
            {
                if(modulePeriod.Contains('_'))
                {
                    string periodInt = modulePeriod.Split('_')[1];
                    if (periodInt == "1")
                    {
                        return "FIRST PERIOD";
                    }
                    else if (periodInt == "2")
                    {
                        return "SECOND PERIOD";
                    }
                    else if (periodInt == "3")
                    {
                        return "THIRD PERIOD";
                    }
                    else if (periodInt == "4")
                    {
                        return "FOURTH PERIOD";
                    }
                    else if (periodInt == "0")

                    {
                        return "";
                    }
                    else
                    {
                        return "";
                    }
                }
                return modulePeriod;
                
            }
        }
    }
}
