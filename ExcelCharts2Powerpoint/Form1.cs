using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Serilog;
using Serilog.Configuration;
using Serilog.Core;
using Serilog.Events;

using Office = NetOffice.OfficeApi;
using Excel = NetOffice.ExcelApi;
using PowerPoint = NetOffice.PowerPointApi;

using PowerPointEnums = NetOffice.PowerPointApi.Enums;
using System.Threading;
using System.IO;
using System.Runtime.InteropServices;
using System.Diagnostics;

namespace ExcelCharts2Powerpoint
{
    public partial class Form1 : Form
    {
        public Form1 thisForm;
        public Form1()
        {
            InitializeComponent();
            thisForm = this;
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            Log.Logger = new LoggerConfiguration()
                   .MinimumLevel.Information()
                   .WriteTo.CustomSinkLogger()
                   .CreateLogger();
            CustomSink.instance.OnLogUpdate += Instance_OnLogUpdate;
        }

        private void Instance_OnLogUpdate(string logText)
        {
            this.BeginInvoke((MethodInvoker)(() => { this.logBox.AppendText(logText + Environment.NewLine); }));

        }
        static string xlsx;
        static List<string> pptx;
        static string output_folder;
        private void button3_Click(object sender, EventArgs e)
        {
            bool error = false;
            button1.Enabled = false;
            button2.Enabled = false;
            buttonFolder.Enabled = false;
            button3.Enabled = false;
            listView1.Enabled = false;
            listView2.Enabled = false;
            listView3.Enabled = false;
            logBox.Clear();
            pptx = new List<string>();
            try
            {
                xlsx = listView2.Items[0].SubItems[1].Text;
            }
            catch (Exception ee)
            {
                MessageBox.Show("Input Excel file not definded in Step 1.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                error = true;
            }
            foreach (ListViewItem listViewItem in listView1.Items)
            {
                String filename = listViewItem.SubItems[1].Text;
                pptx.Add(filename);
            }
            if (pptx.Count == 0)
            {
                MessageBox.Show("No input presentations defined in Step 2.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                error = true;
            }
            // pptx.AddRange(openFileDialogPPTX.FileNames);
            try
            {

                output_folder = listView3.Items[0].SubItems[1].Text;
            }
            catch (Exception ee)
            {
                MessageBox.Show("Output folder not definded in Step 3.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                error = true;
            }
            if (error)
            {

                button1.Enabled = true;
                button2.Enabled = true;
                buttonFolder.Enabled = true;
                button3.Enabled = true;

                listView1.Enabled = true;
                listView2.Enabled = true;
                listView3.Enabled = true;
            }
            else
            {
                Thread workThread = new Thread(doWork);
                workThread.Start();
            }
            //  button1.Enabled = false;
        }

        private void doWork()
        {
            PowerPoint.Presentation presentation = null;
            PowerPoint.Application applicationPowerPoint = null;
            Excel.Workbook book = null;
            Excel.Application applicationExcel = null;

            try
            {
                Dictionary<String, Excel.Shape> excelShapesDictionary = new Dictionary<string, Excel.Shape>();


                Log.Information("Opening Excel Applicaiton");
                applicationExcel = new Excel.Application();
                applicationExcel.DisplayAlerts = false;

                Log.Information("Opening Excel File {0}", xlsx);
                book = applicationExcel.Workbooks.Open(xlsx, true, true);

                Boolean flagDouble = false;
                Log.Information("Iterating all shapes in all sheets and filtering shapes with name staring with \"#\"");
                foreach (Excel.Worksheet sheet in book.Worksheets)
                {
                    foreach (Excel.Shape shape in sheet.Shapes)
                    {
                        if (shape.Name.Length > 1 && shape.Name.Substring(0, 1) == "#")
                        {

                            if (excelShapesDictionary.ContainsKey(shape.Name))
                            {
                                Log.Error("\t\tSheet: {0} Shape Name : {1} Size(w x h) : {2} x {3} Position(left x top) : {4} , {5} Error!!! Shape with the same name exists", sheet.Name, shape.Name, sheet.Name, shape.Name, shape.Width, shape.Height, shape.Left, shape.Top);
                                flagDouble = true;
                            }
                            else
                            {
                                Log.Information("\tSheet: {0} Shape Name : {1} Size(w x h) : {2} x {3} Position(left x top) : {4} , {5}", sheet.Name, shape.Name, shape.Width, shape.Height, shape.Left, shape.Top);

                                excelShapesDictionary.Add(shape.Name.ToLower().Trim(), shape);
                            }
                        }
                    }
                }
                if (flagDouble)
                {
                    Log.Error("ERROR -Found shapes with the same name : Duplicated shapes need to be manually renamed in the Excel file before proceeding. Terminating run.");
                    thisForm.BeginInvoke((MethodInvoker)(() =>
                    {
                        MessageBox.Show(thisForm, "Duplicated shapes need to be manually renamed in the Excel file before proceeding.\nSee logs for more information", "ERROR - Found shapes with the same name", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }));
                    goto closeWorkbook;
                }

                Log.Information("Opening PowerPoint Applicaiton");
                applicationPowerPoint = new PowerPoint.Application();
                applicationPowerPoint.DisplayAlerts = PowerPoint.Enums.PpAlertLevel.ppAlertsNone;


                foreach (string pptxSingle in pptx)
                {
                    Dictionary<int, List<PowerPoint.Shape>> powerpointShapesSheetDictionary = new Dictionary<int, List<PowerPoint.Shape>>();

                    Log.Information("Opening Presentation {0}", pptxSingle);
                    presentation = applicationPowerPoint.Presentations.Open(pptxSingle, true, true, false);


                    Log.Information("Iterating all shapes in all slides and filtering shapes with name staring with \"#\"");
                    Boolean datamissing = false;
                    foreach (PowerPoint.Slide slide in presentation.Slides)
                    {
                        foreach (PowerPoint.Shape shape in slide.Shapes)
                        {
                            if (shape.Name.Length > 1 && shape.Name.Substring(0, 1) == "#")
                            {
                                if (!excelShapesDictionary.ContainsKey(shape.Name.ToLower().Trim()))
                                {
                                    Log.Information("\t\tData Missing for Slide No : {0} Shape Name : {1} Size(w x h) : {2} x {3} Position(left x top) : {4} , {5}", slide.SlideNumber, shape.Name, shape.Width, shape.Height, shape.Left, shape.Top);
                                    datamissing = true;
                                }
                                else
                                {
                                    Log.Information("\tFound data for Slide No : {0} Shape Name : {1} Size(w x h) : {2} x {3} Position(left x top) : {4} , {5}", slide.SlideNumber, shape.Name, shape.Width, shape.Height, shape.Left, shape.Top);

                                }
                                if (!powerpointShapesSheetDictionary.ContainsKey(slide.SlideNumber))
                                {
                                    powerpointShapesSheetDictionary.Add(slide.SlideNumber, new List<PowerPoint.Shape>());

                                }
                                powerpointShapesSheetDictionary[slide.SlideNumber].Add(shape);
                            }
                            else
                            {
                                /*
                                Log.Information("\tFound data for Slide No : {0} Shape Name : {1} Size(w x h) : {2} x {3} Position(left x top) : {4} , {5}", slide.SlideNumber, shape.Name, shape.Width, shape.Height, shape.Left, shape.Top);
                                if (shape.HasTextFrame == Office.Enums.MsoTriState.msoTrue && shape.TextFrame.HasText == Office.Enums.MsoTriState.msoTrue)
                                    shape.TextFrame.TextRange.Replace("|*test*|", "123");
                                    */
                            }
                        }
                    }
                    if (datamissing)
                    {
                        Log.Error("ERROR - Match not found for shape/s in presentation : All shapes starting with \"#\"in the presentation should have matching shape in the excel.");
                        thisForm.BeginInvoke((MethodInvoker)(() =>
                        {
                            MessageBox.Show(thisForm, "All shapes starting with \"#\"in the presentation should have matching shape in the excel.\nSee logs for more information", "ERROR - Match not found for shape/s in presentation", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }));
                        goto closePresentation;
                    }


                    Log.Information("Start of Update");
                    foreach (int slideNo in powerpointShapesSheetDictionary.Keys)
                    {
                        foreach (var shape in powerpointShapesSheetDictionary[slideNo])
                        {
                            excelShapesDictionary[shape.Name.ToLower()].Copy();
                            Thread.Sleep(100);
                            PowerPoint.ShapeRange shapes = presentation.Slides[slideNo].Shapes.PasteSpecial(PowerPointEnums.PpPasteDataType.ppPasteJPG);
                            string shape_name = shape.Name;
                            float shape_top = shape.Top;
                            float shape_left = shape.Left;

                            float shape_width = shape.Width;
                            float shape_height = shape.Height;

                            shapes[1].Name = shape_name;
                            shapes[1].Top = shape_top;
                            shapes[1].Left = shape_left;

                            shapes[1].ScaleWidth(shape.Width / shapes[1].Width, Office.Enums.MsoTriState.msoFalse);
                            //shapes[1].Height = shape_height;

                            shape.Delete();

                            Log.Information(" Updated {0} on slide {1}", shape_name, slideNo);
                        }
                    }
                    Log.Information("End of Update");

                    String outputfile = Path.Combine(output_folder, new FileInfo(pptxSingle).Name);
                    Log.Information("Saving a copy of updated Presentation to {0}", outputfile);
                    presentation.SaveCopyAs(outputfile);

                    closePresentation:
                    Log.Information("Closing Presentation");

                    presentation.Close();
                }


                Log.Information("Closing PowerPoint Applicaiton");
                applicationPowerPoint.Quit();
                applicationPowerPoint.Dispose();

                closeWorkbook:
                Log.Information("Closing Excel File");
                book.Close();

                Log.Information("Closing Excel Applicaiton");
                applicationExcel.Quit();
                applicationExcel.Dispose();


                Log.Information("Done");
                thisForm.BeginInvoke((MethodInvoker)(() =>
                {
                    button1.Enabled = true;
                    button2.Enabled = true;
                    buttonFolder.Enabled = true;
                    button3.Enabled = true;

                    listView1.Enabled = true;
                    listView2.Enabled = true;
                    listView3.Enabled = true;
                    MessageBox.Show(thisForm, "Task Completed. Check logs for more info", "Completed", MessageBoxButtons.OK, MessageBoxIcon.None);
                }));
            }
            catch (Exception e)
            {

                Log.Error("Fatal Error - " + e.ToString());
                thisForm.BeginInvoke((MethodInvoker)(() =>
                {
                    MessageBox.Show(thisForm, e.ToString(), "Fatal Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    button1.Enabled = true;
                    button2.Enabled = true;
                    buttonFolder.Enabled = true;
                    button3.Enabled = true;

                    listView1.Enabled = true;
                    listView2.Enabled = true;
                    listView3.Enabled = true;
                }));
                try
                {
                    presentation.Close();
                }
                catch (Exception ee) { }
                try
                {
                    applicationPowerPoint.Quit();
                    applicationPowerPoint.Dispose();
                }
                catch (Exception ee) { }
                try
                {
                    book.Close();
                }
                catch (Exception ee) { }
                try
                {
                    applicationExcel.Quit();
                    applicationExcel.Dispose();
                }
                catch (Exception ee) { }
            }
        }

        void addList2(String filename)
        {
            listView2.Items.Clear();
            appendList(ref listView2, new string[] { filename });
        }

        void addList1(String[] filenames)
        {
            List<string> alreadyInList = new List<string>();

            List<string> toAddInList = new List<string>();
            foreach (ListViewItem listViewItem in listView1.Items)
            {
                alreadyInList.Add(listViewItem.SubItems[0].Text);
            }

            foreach (string name in filenames)
            {
                if (!alreadyInList.Contains(new FileInfo(name).Name))
                {
                    toAddInList.Add(name);
                }
            }
            appendList(ref listView1, toAddInList.ToArray());
        }

        void addList3(String filename)
        {
            listView3.Items.Clear();
            appendList(ref listView3, new string[] { filename });
        }
        private void button1_Click_1(object sender, EventArgs e)
        {
            var result = openFileDialogXLSX.ShowDialog();
            if (result == DialogResult.OK)
            {
                addList2(openFileDialogXLSX.FileName);
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            var result = openFileDialogPPTX.ShowDialog();
            if (result == DialogResult.OK)
            {
                addList1(openFileDialogPPTX.FileNames);

                //listView1.Items.AddRange();
            }
        }

        private void buttonFolder_Click(object sender, EventArgs e)
        {
            var result = folderBrowserDialog1.ShowDialog();
            if (result == DialogResult.OK)
            {
                // labelFolder.Text = new DirectoryInfo(folderBrowserDialog1.SelectedPath).Name;
                addList3(folderBrowserDialog1.SelectedPath);

            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            //temporary function
            Thread workThread = new Thread(doWork2);
            workThread.Start();
        }
        private void doWork2()
        {
            PowerPoint.Presentation presentation = null;
            PowerPoint.Application applicationPowerPoint = null;
            Excel.Workbook book = null;
            Excel.Application applicationExcel = null;

            try
            {
                Dictionary<String, Excel.Shape> excelShapesDictionary = new Dictionary<string, Excel.Shape>();

                Log.Information("Opening Excel Applicaiton");
                applicationExcel = new Excel.Application();
                applicationExcel.DisplayAlerts = false;

                String xlsx = "D:\\Input\\Input.xlsx";
                String output_folder = "D:\\Output";
                String pptxSingle = "D:\\Input\\input.pptx";

                Log.Information("Opening Excel File {0}", xlsx);
                book = applicationExcel.Workbooks.Open(xlsx, true, true);

                Excel.Worksheet worksheet = (Excel.Worksheet)book.Worksheets[1];
                int index = 3;
                List<PresentationItem> presentationItems = new List<PresentationItem>();
                do
                {

                    String s = worksheet.Range("A" + index).Value2.ToString();
                    Log.Information(worksheet.Name + " " + s);

                    PresentationItem presentationItem = new PresentationItem();
                    presentationItem.OutputFileName = s;
                    int index2 = 1;
                    do
                    {
                        String attr_type = worksheet.Range(((char)((byte)'A' + index2)).ToString() + 1).Value2.ToString();
                        String attr_name = worksheet.Range(((char)((byte)'A' + index2)).ToString() + 2).Value2.ToString();
                        String attr_value = worksheet.Range(((char)((byte)'A' + index2)).ToString() + index).Value2.ToString();

                        Log.Information("{0} {1} {2} {3}", ((char)((byte)'A' + index2)).ToString() + 1, attr_type, attr_name, attr_value);
                        presentationItem.Attributes.Add(new Attribute(attr_type, attr_name, attr_value));

                        index2++;
                        if (index2 == 20)
                        {
                            break;
                        }
                    } while (worksheet.Range(((char)((byte)'A' + index2)).ToString() + 1).Value2 != null && worksheet.Range(((char)((byte)'A' + index2)).ToString() + 1).Value2.ToString() != "");

                    index++;
                    presentationItems.Add(presentationItem);
                    if (index == 100)
                    {
                        break;
                    }
                } while (worksheet.Range("A" + index).Value2 != null && worksheet.Range("A" + index).Value2.ToString() != "");


                Log.Information("Closing Excel File");
                book.Close();

                Log.Information("Opening PowerPoint Applicaiton");
                applicationPowerPoint = new PowerPoint.Application();
                applicationPowerPoint.DisplayAlerts = PowerPoint.Enums.PpAlertLevel.ppAlertsNone;


                foreach (var presentationItem in presentationItems)
                {
                    Dictionary<int, List<PowerPoint.Shape>> powerpointShapesSheetDictionary = new Dictionary<int, List<PowerPoint.Shape>>();

                    Log.Information("Opening Presentation {0}", pptxSingle);
                    presentation = applicationPowerPoint.Presentations.Open(pptxSingle, true, true, true);

                    foreach (PowerPoint.Shape shape in presentation.SlideMaster.Shapes)
                    {
                        Log.Information("\tSlide Master:  Shape Name : {0} Size(w x h) : {1} x {2} Position(left x top) : {3} , {4}", shape.Name, shape.Width, shape.Height, shape.Left, shape.Top);
                        if (shape.HasTextFrame == Office.Enums.MsoTriState.msoTrue && shape.TextFrame.HasText == Office.Enums.MsoTriState.msoTrue)
                        {
                            foreach (var attribute in presentationItem.Attributes)
                            {
                                if (attribute.type == "Text")
                                {
                                    string before = shape.TextFrame.TextRange.Text;
                                    shape.TextFrame.TextRange.Replace("|*" + attribute.name + "*|", attribute.value);
                                    string after = shape.TextFrame.TextRange.Text;
                                    if (before != after)
                                    {
                                        Log.Information("\t\t{0} >> {1}", before, after);
                                    }
                                }
                            }
                        }
                    }
                    Log.Information("Iterating all shapes in all slides and filtering shapes with name staring with \"#\"");
                    foreach (PowerPoint.Slide slide in presentation.Slides)
                    {
                        foreach (PowerPoint.Shape shape in slide.Shapes)
                        {
                            Log.Information("\tSlide No : {0} Shape Name : {1} Size(w x h) : {2} x {3} Position(left x top) : {4} , {5}", slide.SlideNumber, shape.Name, shape.Width, shape.Height, shape.Left, shape.Top);
                            if (shape.HasTextFrame == Office.Enums.MsoTriState.msoTrue && shape.TextFrame.HasText == Office.Enums.MsoTriState.msoTrue)
                            {
                                foreach (var attribute in presentationItem.Attributes)
                                {
                                    if (attribute.type == "Text")
                                    {
                                        string before = shape.TextFrame.TextRange.Text;
                                        shape.TextFrame.TextRange.Replace("|*" + attribute.name + "*|", attribute.value);
                                        string after = shape.TextFrame.TextRange.Text;
                                        if (before != after)
                                        {
                                            Log.Information("\t\t{0} >> {1}", before, after);
                                        }
                                    }
                                }
                            }
                            foreach (var attribute in presentationItem.Attributes)
                            {
                                if (attribute.type == "Chart" && shape.Name == "|*" + attribute.name + "*|")
                                {
                                    string before = shape.Name;
                                    shape.Name = attribute.value;
                                    string after = shape.Name;
                                    if (before != after)
                                    {
                                        Log.Information("\t\t{0} >> {1}", before, after);
                                    }
                                }
                            }
                        }
                        Log.Information("");
                    }


                    String outputfile = Path.Combine(output_folder, new FileInfo(pptxSingle).Name);
                    Log.Information("Saving a copy of updated Presentation to {0}", outputfile);
                    presentation.SaveCopyAs(outputfile);

                    Log.Information("Closing Presentation");
                }

                presentation.Close();

                Log.Information("Closing PowerPoint Applicaiton");
                applicationPowerPoint.Quit();
                applicationPowerPoint.Dispose();

                Log.Information("Closing Excel Applicaiton");
                applicationExcel.Quit();
                applicationExcel.Dispose();

                Log.Information("Done");
            }
            catch (Exception e)
            {

                Log.Error("Fatal Error - " + e.ToString());
                thisForm.BeginInvoke((MethodInvoker)(() =>
                {
                    MessageBox.Show(thisForm, e.ToString(), "Fatal Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }));
                try
                {
                    presentation.Close();
                }
                catch (Exception ee) { }
                try
                {
                    applicationPowerPoint.Quit();
                    applicationPowerPoint.Dispose();
                }
                catch (Exception ee) { }
                try
                {
                    book.Close();
                }
                catch (Exception ee) { }
                try
                {
                    applicationExcel.Quit();
                    applicationExcel.Dispose();
                }
                catch (Exception ee) { }
            }
        }

        void appendList(ref ListView listView, string[] filenames)
        {
            // Obtain a handle to the system image list.
            NativeMethods.SHFILEINFO shfi = new NativeMethods.SHFILEINFO();
            IntPtr hSysImgList = NativeMethods.SHGetFileInfo("",
                                                             0,
                                                             ref shfi,
                                                             (uint)Marshal.SizeOf(shfi),
                                                             NativeMethods.SHGFI_SYSICONINDEX
                                                              | NativeMethods.SHGFI_SMALLICON);
            Debug.Assert(hSysImgList != IntPtr.Zero);  // cross our fingers and hope to succeed!

            // Set the ListView control to use that image list.
            IntPtr hOldImgList = NativeMethods.SendMessage(listView.Handle,
                                                           NativeMethods.LVM_SETIMAGELIST,
                                                           NativeMethods.LVSIL_SMALL,
                                                           hSysImgList);

            // If the ListView control already had an image list, delete the old one.
            if (hOldImgList != IntPtr.Zero)
            {
                NativeMethods.ImageList_Destroy(hOldImgList);
            }

            NativeMethods.SetWindowTheme(listView.Handle, "Explorer", null);


            foreach (string file in filenames)
            {
                IntPtr himl = NativeMethods.SHGetFileInfo(file,
                                                          0,
                                                          ref shfi,
                                                          (uint)Marshal.SizeOf(shfi),
                                                          NativeMethods.SHGFI_DISPLAYNAME
                                                            | NativeMethods.SHGFI_SYSICONINDEX
                                                            | NativeMethods.SHGFI_SMALLICON);
                Debug.Assert(himl == hSysImgList); // should be the same imagelist as the one we set
                var listViewItem = listView.Items.Add(shfi.szDisplayName, shfi.iIcon);
                listViewItem.SubItems.Add(file);
            }
        }

        internal static class NativeMethods
        {
            public const uint LVM_FIRST = 0x1000;
            public const uint LVM_GETIMAGELIST = (LVM_FIRST + 2);
            public const uint LVM_SETIMAGELIST = (LVM_FIRST + 3);

            public const uint LVSIL_NORMAL = 0;
            public const uint LVSIL_SMALL = 1;
            public const uint LVSIL_STATE = 2;
            public const uint LVSIL_GROUPHEADER = 3;

            [DllImport("user32")]
            public static extern IntPtr SendMessage(IntPtr hWnd,
                                                    uint msg,
                                                    uint wParam,
                                                    IntPtr lParam);

            [DllImport("comctl32")]
            public static extern bool ImageList_Destroy(IntPtr hImageList);

            public const uint SHGFI_DISPLAYNAME = 0x200;
            public const uint SHGFI_ICON = 0x100;
            public const uint SHGFI_LARGEICON = 0x0;
            public const uint SHGFI_SMALLICON = 0x1;
            public const uint SHGFI_SYSICONINDEX = 0x4000;

            [StructLayout(LayoutKind.Sequential)]
            public struct SHFILEINFO
            {
                public IntPtr hIcon;
                public int iIcon;
                public uint dwAttributes;
                [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 260 /* MAX_PATH */)]
                public string szDisplayName;
                [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 80)]
                public string szTypeName;
            };

            [DllImport("shell32")]
            public static extern IntPtr SHGetFileInfo(string pszPath,
                                                      uint dwFileAttributes,
                                                      ref SHFILEINFO psfi,
                                                      uint cbSizeFileInfo,
                                                      uint uFlags);

            [DllImport("uxtheme", CharSet = CharSet.Unicode)]
            public static extern int SetWindowTheme(IntPtr hWnd,
                                                    string pszSubAppName,
                                                    string pszSubIdList);
        }

        private void listView1_KeyDown(object sender, KeyEventArgs e)
        {
            if (Keys.Delete == e.KeyCode)
            {
                foreach (ListViewItem listViewItem in ((ListView)sender).SelectedItems)
                {
                    listViewItem.Remove();
                }
            }
        }

        private void listView3_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        string[] xlsxExtension = new string[] { ".xls", ".xlsx", ".xlsm" };
        string[] pptxExtension = new string[] { ".ppt", ".pptx" };

        private void listView_DragEnter(object sender, DragEventArgs e)
        {
            bool handled = false;
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                String[] files = ((string[])e.Data.GetData(DataFormats.FileDrop));
                if (files.Count() == 1 && (sender.Equals(listView2) || sender.Equals(listView3)))
                {

                    String extension = new FileInfo(files[0]).Extension.ToLower();
                    if (sender.Equals(listView2) && xlsxExtension.Contains(extension))
                    {
                        handled = true;
                        e.Effect = DragDropEffects.Copy;
                    }

                    if (sender.Equals(listView3) && new DirectoryInfo(files[0]).Exists)
                    {
                        handled = true;
                        e.Effect = DragDropEffects.Copy;
                    }

                }
                if (sender.Equals(listView1))
                {
                    bool dataOK = true;
                    foreach (String file in files)
                    {
                        String extension = new FileInfo(files[0]).Extension.ToLower();
                        if (!pptxExtension.Contains(extension))
                            dataOK = false;
                    }
                    if (dataOK)
                    {
                        handled = true;
                        e.Effect = DragDropEffects.Copy;
                    }

                }
            }
            if (!handled)
            {
                e.Effect = DragDropEffects.None;
            }
        }

        private void listView_DragDrop(object sender, System.Windows.Forms.DragEventArgs e)
        {
            if (sender.Equals(listView1))
            {
                string[] data = ((string[])e.Data.GetData(DataFormats.FileDrop));
                addList1(data);
            }
            if (sender.Equals(listView2))
            {
                string[] data = ((string[])e.Data.GetData(DataFormats.FileDrop));
                addList2(data[0]);
            }
            if (sender.Equals(listView3))
            {
                string[] data = ((string[])e.Data.GetData(DataFormats.FileDrop));
                addList3(data[0]);
            }

        }

        private void listView_DoubleClick(object sender, EventArgs e)
        {
            string pathToOpen = ((ListView)sender).SelectedItems[0].SubItems[1].Text;
            FileAttributes attr = File.GetAttributes(pathToOpen);

            if ((attr & FileAttributes.Directory) != FileAttributes.Directory)
                pathToOpen = Path.GetDirectoryName(pathToOpen);

            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo()
            {
                FileName = pathToOpen,
                UseShellExecute = true,
                Verb = "open"
            });
        }
        
    }
}
