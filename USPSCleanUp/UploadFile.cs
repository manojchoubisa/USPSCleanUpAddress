using LinqToExcel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Data.OleDb;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Web;
using System.Web.UI;
using System.Collections;

using Excel = Microsoft.Office.Interop.Excel;   //namespace
using System.Diagnostics;

namespace USPSCleanUp
{
    public partial class UploadFile : Form
    {
        BackgroundWorker defaultDataLoader;
        BackgroundWorker defaultDataLoader_LoadFiles;

        Excel.Application sExcelApp;
        Excel.Workbook sWorkbook;
        private string Excel03ConString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source={0};Extended Properties='Excel 8.0;HDR={1};IMEX=1'";
        private string Excel07ConString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties='Excel 12.0;HDR={1};IMEX=1'";
        //string  sheetName,sheet;
        string housnno, addressline1, address1;
        static public List<string> dtsheetName = new List<string>();
        public static string filePath = "";
        string s1 = " ";

        public static bool upload;
        public static DataSet dt = new DataSet();
        public static DataSet olddt = new DataSet();
        System.Data.DataColumn newColumn = new System.Data.DataColumn("Error", typeof(System.String));
        string newaddress, newzip, newstate, newcountry, newcity;
        static string oldaddress1, oldaddress2, oldzip, oldstate, oldcountry, oldHouseNo, oldcity;
        int lineno = 0;
        static string filename, fullpath, MyPath, onlyFileName, fileextension, err;


        private void cbSheetList_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                dataGridView1.DataSource = null;
                dataGridView1.Rows.Clear();
                dataGridView1.Refresh();
                string tablename = cbSheetList.SelectedValue.ToString();
                dataGridView1.DataSource = dt.Tables[tablename];


            }
            catch (Exception ex)
            {
                Logger.Log(ex, null);
                MessageBox.Show(ex.Message);
            }
        }

        List<string> objexception = new List<string>();

        private void UploadFile_Load(object sender, EventArgs e)
        {

            if (upload == true)
            {
                btnCancel.Visible = true;
                btnCleasing.Visible = true;
                pnlHide.Visible = false;
            }
        }



        private void btnCancel_Click(object sender, EventArgs e)
        {
            try
            {
                btnCleasing.Visible = false;
                btnCancel.Visible = false;
                pnlHide.Visible = true;
                dt.Clear();
                olddt.Clear();
                USPSCleanUp.UploadFile.dtsheetName.Clear();
                USPSCleanUp.SetColumns.sheetCount = 0;
                USPSCleanUp.SetColumns.FullColumnList.Clear();

                USPSCleanUp.UploadFile.upload = false;
                USPSCleanUp.UploadFile.dtsheetName.Clear();
                USPSCleanUp.SetColumns.sheetCount = 0;
                USPSCleanUp.SetColumns.FullColumnList.Clear();
                USPSCleanUp.UploadFile.dt.Clear();

                var fileupload = new UploadFile();
                fileupload.Show();
                this.Hide();
            }
            catch (Exception ex)
            {
                Logger.Log(ex, null);
                MessageBox.Show(ex.Message);
            }
        }

        string Addressline1col, AddressLine2col, Zipcol, Statecol, Citycol, Housenocol;
        public static List<string> ColumnList = new List<string>();

        public UploadFile()
        {
            defaultDataLoader = new BackgroundWorker();
            defaultDataLoader.WorkerReportsProgress = true;
            defaultDataLoader.WorkerSupportsCancellation = true;
            defaultDataLoader.DoWork += new DoWorkEventHandler(DataImport_Dowork);
            defaultDataLoader.RunWorkerCompleted += new RunWorkerCompletedEventHandler(ImportOperaions_RunWorkerCompleted);

            defaultDataLoader_LoadFiles = new BackgroundWorker();
            defaultDataLoader_LoadFiles.WorkerReportsProgress = true;
            defaultDataLoader_LoadFiles.WorkerSupportsCancellation = true;
            defaultDataLoader_LoadFiles.DoWork += new DoWorkEventHandler(DataLoad_Dowork);
            defaultDataLoader_LoadFiles.RunWorkerCompleted += new RunWorkerCompletedEventHandler(LoadOperaions_RunWorkerCompleted);

            InitializeComponent();

        }
        private void DataImport_Dowork(object sender, DoWorkEventArgs e)
        {
            Logger.Log(null, "Filling contribution PDF started");
            ProcessCleansing();
        }
        private void DataLoad_Dowork(object sender, DoWorkEventArgs e)
        {
            ExportToExcelArgs args = (ExportToExcelArgs)e.Argument;

            Logger.Log(null, "Loading input contribution file started");
            PerformExportToExcel(args.ds, args.path);

        }

        private void ImportOperaions_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            HideProgressbar();
            Logger.Log(null, "Filling contribution PDF finished");
            //if (showSuccessMessage)
            //{
            //    MessageBox.Show("PDF creation is completed.");
            //}
        }

        private void LoadOperaions_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            HideProgressbar();
            Logger.Log(null, "Loading input contribution file finished");
            //MessageBox.Show("Input file loaded.");
        }
        private void ShowProgressbar()
        {
            btnCancel.Enabled = false;
            this.tabPageProcessFiles.Visible = true;
            this.tabPageProcessFiles.Style = ProgressBarStyle.Marquee;
            this.tabPageProcessFiles.MarqueeAnimationSpeed = 1;
        }
        private void HideProgressbar()
        {
            this.tabPageProcessFiles.Visible = false;
            btnCancel.Enabled = true;
        }

        private void Clear_Rec()
        {
            lblmessage.Visible = false;
            lblResult.Visible = false;
            lblStatus.Visible = false;
            lblfilepath.Visible = false;
            dt.Clear();
            olddt.Clear();
            dataGridView1.DataSource = null;
            dataGridView1.Rows.Clear();
            dataGridView1.Refresh();
        }


        private void btnUpload_Click(object sender, EventArgs e)
        {

            ColumnList.Clear();
            Text_FilePath();
            Clear_Rec();
            try
            {
                dt = new DataSet();
                olddt = new DataSet();
                DialogResult result = openFileDialog1.ShowDialog();
                if (result == DialogResult.OK)
                {
                    onlyFileName = System.IO.Path.GetFileNameWithoutExtension(openFileDialog1.FileName);
                    fileextension = System.IO.Path.GetExtension(openFileDialog1.FileName);
                    if ((fileextension == ".xls") || (fileextension == ".xlsx"))
                    {
                        textBox1.Text = onlyFileName;
                        filename = "ErrorLog.txt";
                        MyPath = Application.StartupPath.Replace("\\bin\\Debug", "");

                        if (!Directory.Exists(MyPath + "\\ErrorLog\\"))
                        {
                            Directory.CreateDirectory(MyPath + "\\ErrorLog\\");
                        }


                        fullpath = MyPath + "\\ErrorLog\\" + filename;

                        if (File.Exists(fullpath))
                        {
                            File.Delete(fullpath);
                        }

                        filePath = openFileDialog1.FileName;
                        string[] FName;
                        foreach (string s in openFileDialog1.FileNames)
                        {
                            FName = s.Split('\\');


                            if (!Directory.Exists(MyPath + "\\UploadedExcel\\"))
                            {
                                Directory.CreateDirectory(MyPath + "\\UploadedExcel\\");
                            }


                            string Oldfullpath = MyPath + "\\UploadedExcel\\";
                            if (File.Exists(Oldfullpath + FName[FName.Length - 1]))
                            {
                                File.Delete(Oldfullpath + FName[FName.Length - 1]);
                            }
                            File.Copy(s, Oldfullpath + FName[FName.Length - 1]);

                        }
                        string conStr, sheetName;
                        string header = "YES";
                        conStr = string.Empty;
                        switch (fileextension)
                        {
                            case ".xls": //Excel 97-03
                                conStr = string.Format(Excel03ConString, filePath, header);
                                break;

                            case ".xlsx": //Excel 07
                                conStr = string.Format(Excel07ConString, filePath, header);
                                break;
                        }
                        using (OleDbConnection con = new OleDbConnection(conStr))
                        {
                            using (OleDbCommand cmd = new OleDbCommand())
                            {
                                cmd.Connection = con;
                                con.Open();
                                DataTable dtExcelSchema = con.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);


                                //                      sExcelApp = new Excel.Application();


                                //                      //sWorkbook = sExcelApp.Workbooks.Open(filePath, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);  //sFilePath  Excel File Path

                                //                      try
                                //                      {
                                //                          //sWorkbook = sExcelApp.Workbooks.Open(filePath);
                                //                          sWorkbook = sExcelApp.Workbooks.Open(filePath, 0, false, 5, "", "", false, Excel.XlPlatform.xlWindows, "",
                                //true, false, 0, true, false, false);

                                //                      }
                                //                      catch (Exception ex)
                                //                      {

                                //                      }

                                //                      foreach (Excel.Worksheet wSheet in sWorkbook.Worksheets)
                                //                      {
                                //                          var Range = wSheet.UsedRange;
                                //                          //Range excelRange = wSheet.UsedRange;
                                //                          int test1 = Range.Columns.Count;
                                //                          int test2 = Range.Rows.Count;
                                //                          int test3 = Range.Count;
                                //                          if (test1 > 1 || test2 > 1 || test3 > 1)
                                //                          {
                                //                              dtsheetName.Add(wSheet.Name);
                                //                          }


                                //                      }

                                for (int i = 0; i < dtExcelSchema.Rows.Count; i++)
                                {
                                    String sheet = dtExcelSchema.Rows[i]["TABLE_NAME"].ToString();
                                    if (sheet != null && sheet.Contains('$'))
                                    {
                                        if (sheet.Contains(" ") || sheet.Contains("'"))
                                        {
                                            sheet = sheet.Substring(1, sheet.Length - 3);
                                        }
                                        else
                                        {
                                            sheet = sheet.Substring(0, sheet.Length - 1);
                                        }
                                        dtsheetName.Add(sheet);
                                    }
                                }

                                con.Close();
                            }
                        }
                        using (OleDbConnection con = new OleDbConnection(conStr))
                        {
                            using (OleDbCommand cmd = new OleDbCommand())
                            {
                                using (OleDbDataAdapter oda = new OleDbDataAdapter())
                                {
                                    for (int i = 0; i < dtsheetName.Count; i++)
                                    {
                                        sheetName = dtsheetName[i].ToString();


                                        cmd.CommandText = "SELECT * From [" + sheetName + "$]";
                                        cmd.Connection = con;
                                        con.Open();
                                        oda.SelectCommand = cmd;
                                        var tbl = new DataTable(sheetName);
                                        dt.Tables.Add(tbl);
                                        tbl = new DataTable(sheetName);

                                        tbl.Columns.Add("Error", typeof(string));




                                        olddt.Tables.Add(tbl);
                                        oda.Fill(dt.Tables[sheetName]);
                                        oda.Fill(olddt.Tables[sheetName]);
                                        con.Close();
                                    }



                                    upload = true;
                                    if (upload == true)
                                    {
                                        pnlHide.Visible = false;
                                    }
                                    var moreForm = new SetColumns();
                                    moreForm.Show();
                                    this.Hide();


                                }
                            }

                        }
                    }

                    else
                    {
                        btnCleasing.Visible = true;
                        lblmessage.Visible = true;
                        lblmessage.Text = "Please Select only .xls or .xlsx Extension files";
                        lblmessage.ForeColor = Color.Red;
                    }
                }
                else
                {
                    lblmessage.Visible = true;
                    lblmessage.Text = "No file Selected";
                    lblmessage.ForeColor = Color.Red;
                }
            }
            catch (Exception ex)
            {
                Logger.Log(ex, null);
                MessageBox.Show(ex.Message);
            }
            finally
            {

                //sWorkbook.Close(0);
                //sExcelApp.Quit();

                //var processes = from p in Process.GetProcessesByName("EXCEL")
                //                select p;

                //foreach (var process in processes)
                //{
                //    if (process.MainWindowTitle == "Microsoft Excel - " + filePath)
                //        process.Kill();
                //}
            }
        }

        public void logerrors(Exception ex, string sheetName)
        {
            lblmessage.Visible = false;
            lblStatus.Visible = false;
            lblResult.Visible = false;
            lblfilepath.Visible = false;
            Text_FilePath();
            try
            {
                //string filename = "ErrorLog_" + DateTime.Now.ToString("dd-MM-yyyy") + ".txt";
                filename = "ErrorLog.txt";
                MyPath = Application.StartupPath.Replace("\\bin\\Debug", "");


                if (!Directory.Exists(MyPath + "\\ErrorLog\\"))
                {
                    Directory.CreateDirectory(MyPath + "\\ErrorLog\\");
                }

                fullpath = MyPath + "\\ErrorLog\\" + filename;

                if (File.Exists(fullpath))
                {
                    string message = string.Format("Time: {0}", DateTime.Now.ToString("dd/MM/yyyy hh:mm:ss tt"));

                    message += Environment.NewLine;
                    message += "-----------------------------------------------------------";
                    message += string.Format("Sheet Name: {0}", sheetName);

                    message += Environment.NewLine;

                    message += "-----------------------------------------------------------";
                    message += Environment.NewLine;
                    message += string.Format("Address1: {0}", oldaddress1);
                    message += string.Format("Address2: {0}", oldaddress2);
                    message += string.Format(", Zip: {0}", oldzip);
                    message += Environment.NewLine;
                    message += string.Format("Line No: {0}", lineno);
                    message += Environment.NewLine;
                    message += string.Format("Message: {0}", ex.Message);
                    message += Environment.NewLine;
                    message += "-----------------------------------------------------------";
                    message += Environment.NewLine;
                    message += "-----------------------------------------------------------";
                    message += Environment.NewLine;
                    message += "";
                    message += Environment.NewLine;
                    using (StreamWriter writer = new StreamWriter(fullpath, true))
                    {
                        writer.WriteLine(message);
                        objexception.Add(message);
                        writer.Close();
                    }
                }
                else
                {
                    string message = string.Format("Time: {0}", DateTime.Now.ToString("dd/MM/yyyy hh:mm:ss tt"));

                    message += Environment.NewLine;
                    message += "-----------------------------------------------------------";
                    message += string.Format("Sheet Name: {0}", sheetName);

                    message += Environment.NewLine;
                    message += "-----------------------------------------------------------";
                    message += Environment.NewLine;
                    message += string.Format("Address1: {0}", oldaddress1);
                    message += string.Format("Address2: {0}", oldaddress2);
                    message += string.Format(", Zip: {0}", oldzip);
                    message += Environment.NewLine;
                    message += string.Format("Line No: {0}", lineno);
                    message += Environment.NewLine;
                    message += string.Format("Message: {0}", ex.Message);
                    message += Environment.NewLine;
                    message += "-----------------------------------------------------------";
                    message += Environment.NewLine;
                    message += "-----------------------------------------------------------";
                    message += Environment.NewLine;
                    message += "";
                    message += Environment.NewLine;
                    using (StreamWriter writer = new StreamWriter(fullpath, true))
                    {

                        objexception.Add(message);
                        writer.WriteLine(message);
                        writer.Close();
                    }
                }

                //DataColumnCollection columns = olddt.Tables[sheetName].Columns;

                //if (columns.Contains(newColumn.ColumnName))
                //{
                //    foreach (DataTable anotherdt in olddt.Tables)
                //    {
                //        DataColumnCollection lstcolumndt = anotherdt.Columns;
                //        if (!lstcolumndt.Contains(newColumn.ColumnName))
                //        {
                //            olddt.Tables[sheetName].Columns.Remove(newColumn);
                //        }
                //    }
                //    olddt.Tables[sheetName].Columns.Add(newColumn);
                //    string s = ex.Message;
                //    newColumn.DefaultValue = s;

                //}
                //else
                //{
                //    foreach (DataTable anotherdt in olddt.Tables)
                //    {
                //        DataColumnCollection lstcolumndt = anotherdt.Columns;
                //        if (lstcolumndt.Contains(newColumn.ColumnName))
                //        {
                //            if (!columns.Contains(newColumn.ColumnName))
                //            {

                //                olddt.Tables[sheetName].Columns.Add(newColumn.ColumnName, anotherdt.Columns[newColumn.ColumnName].DataType);
                //                string s = ex.Message;
                //                newColumn.DefaultValue = s;
                //            }

                //        }
                //        else
                //        {
                //            if (columns.Contains(newColumn.ColumnName))
                //            {
                //                olddt.Tables[sheetName].Columns.Remove(newColumn);
                //                olddt.Tables[sheetName].Columns.Add(newColumn);
                //                string s = ex.Message;
                //                newColumn.DefaultValue = s;

                //            }
                //            else
                //            {
                //                olddt.Tables[sheetName].Columns.Add(newColumn);
                //                string s = ex.Message;
                //                newColumn.DefaultValue = s;
                //            }
                //        }
                //    }





                //} 

            }
            catch (Exception ex1)
            {
                Logger.Log(ex1, null);
                MessageBox.Show(ex1.Message);
            }
        }

        private void btnExportNewExcel_Click(object sender, EventArgs e)
        {
            Text_FilePath();

            lblmessage.Invoke((MethodInvoker)delegate
            {
                lblmessage.Visible = false;
                lblStatus.Visible = false;
                lblResult.Visible = false;
                lblfilepath.Visible = false;
            });

            if (dt.Tables.Count > 0)
            {
                try
                {
                    MyPath = Application.StartupPath.Replace("\\bin\\Debug", "");

                    //if (!Directory.Exists(MyPath + "\\NewExcelFile\\"))
                    //{
                    //    Directory.CreateDirectory(MyPath + "\\NewExcelFile\\");
                    //}
                    //fullpath = MyPath + "\\NewExcelFile\\" + onlyFileName + fileextension;

                    //if (File.Exists(fullpath))
                    //{

                    //    File.Delete(fullpath);
                    //}

                    ExportDataSetToExcel(dt, MyPath);



                    //txtFilePath.Visible = true;
                    //txtFilePath.Text = "Exported File Path:- " + fullpath;
                    //txtFilePath.ForeColor = Color.Green;
                }
                catch (Exception ex)
                {
                    Logger.Log(ex, null);
                    MessageBox.Show(ex.Message);
                }

            }
        }

        private void ExportDataSetToExcel(DataSet ds, string filepath)
        {
            try
            {
                SaveFileDialog openDlg = new SaveFileDialog();
                openDlg.InitialDirectory = filepath;

                openDlg.Filter = "Excel (*.xls)|*.xlsx|All files (*.*)|*.*";
                openDlg.FilterIndex = 2;

                string path = "";

                if (openDlg.ShowDialog() == DialogResult.OK)
                {
                    path = openDlg.FileName;
                }

                ShowProgressbar();
                if (!defaultDataLoader_LoadFiles.IsBusy)
                    defaultDataLoader_LoadFiles.RunWorkerAsync(new ExportToExcelArgs() { ds = ds, path = path });

            }
            catch (Exception ex)
            {
                Logger.Log(ex, null);
                MessageBox.Show(ex.Message);
            }
        }


        class ExportToExcelArgs
        {
            internal DataSet ds { get; set; }
            internal string path { get; set; }
        }
        void PerformExportToExcel(DataSet ds, string path)
        {
            sExcelApp = new Excel.Application();
            sWorkbook = sExcelApp.Workbooks.Add(Type.Missing);

            foreach (DataTable table in olddt.Tables)
            {
                //Add a new worksheet to workbook with the Datatable name
                Excel.Worksheet excelWorkSheet = sWorkbook.Sheets.Add();
                excelWorkSheet.Name = table.TableName + "old";

                for (int i = 1; i < table.Columns.Count + 1; i++)
                {
                    excelWorkSheet.Cells[1, i] = table.Columns[i - 1].ColumnName;
                }

                for (int j = 0; j < table.Rows.Count; j++)
                {
                    for (int k = 0; k < table.Columns.Count; k++)
                    {
                        DataColumnCollection columns = table.Columns;
                        if (columns.Contains("Error"))
                        {
                            if (table.Rows[j]["Error"].ToString() != "")
                            {

                                excelWorkSheet.Cells[j + 2, 1].Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);

                                excelWorkSheet.Cells[j + 2, k + 1] = table.Rows[j].ItemArray[k].ToString();
                            }
                            else
                            {

                                excelWorkSheet.Cells[j + 2, k + 1] = table.Rows[j].ItemArray[k].ToString();


                            }
                        }


                        else
                        {

                            excelWorkSheet.Cells[j + 2, k + 1] = table.Rows[j].ItemArray[k].ToString();


                        }
                    }
                }


            }
            foreach (DataTable table in ds.Tables)
            {
                //Add a new worksheet to workbook with the Datatable name
                Excel.Worksheet excelWorkSheet = sWorkbook.Sheets.Add();
                excelWorkSheet.Name = table.TableName + "New";

                for (int i = 1; i < table.Columns.Count + 1; i++)
                {
                    excelWorkSheet.Cells[1, i] = table.Columns[i - 1].ColumnName;
                }

                for (int j = 0; j < table.Rows.Count; j++)
                {
                    for (int k = 0; k < table.Columns.Count; k++)
                    {
                        excelWorkSheet.Cells[j + 2, k + 1] = table.Rows[j].ItemArray[k].ToString();
                    }
                }
            }

            string file = Path.GetFileNameWithoutExtension(path);

            path = Path.Combine(Path.GetDirectoryName(path), file + ".xlsx");

            sWorkbook.SaveAs(path);

            sWorkbook.Close(0);
            sExcelApp.Quit();


            lblStatus.Invoke((MethodInvoker)delegate
            {
                lblStatus.Visible = true;
                lblStatus.Text = "File Exported Successfully";
                lblStatus.ForeColor = Color.Green;
            });
        }
        private void Text_FilePath()
        {
            txtFilePath.Invoke((MethodInvoker)delegate
            {
                txtFilePath.ReadOnly = true;
                txtFilePath.BorderStyle = 0;
                txtFilePath.BackColor = this.BackColor;
                txtFilePath.TabStop = false;
                txtFilePath.Visible = false;
            });

        }
        //private void btnUpdateExitingExcel_Click(object sender, EventArgs e)
        //{
        //    Text_FilePath();
        //    lblmessage.Visible = false;
        //    lblStatus.Visible = false;
        //    lblResult.Visible = false;
        //    lblfilepath.Visible = false;
        //    if (dt.Tables.Count > 0)
        //    {
        //        try
        //        {
        //            MyPath = Application.StartupPath.Replace("\\bin\\Debug", "");


        //            //if (!Directory.Exists(MyPath + "\\UploadedExcel\\"))
        //            //{
        //            //    Directory.CreateDirectory(MyPath + "\\UploadedExcel\\");
        //            //}


        //            //fullpath = MyPath + "\\UploadedExcel\\" + onlyFileName + fileextension;

        //            //if (File.Exists(fullpath))
        //            //{
        //            //    File.Delete(fullpath);
        //            //}
        //            ExportDataSetToExcel(dt, MyPath);

        //            lblStatus.Visible = true;
        //            lblStatus.Text = "File Updated Successfully";
        //            lblStatus.ForeColor = Color.Green;
        //            //txtFilePath.Visible = true;
        //            //txtFilePath.Text = "Updated File Path:- " + fullpath;
        //            //txtFilePath.ForeColor = Color.Green;
        //        }
        //        catch (Exception ex)
        //        {
        //            throw ex;
        //        }

        //    }
        //}

        private void btnErrorLog_Click(object sender, EventArgs e)
        {
            Text_FilePath();
            lblmessage.Visible = false;
            lblStatus.Visible = false;
            lblResult.Visible = false;
            lblfilepath.Visible = false;
            try
            {
                filename = "ErrorLog.txt";
                MyPath = Application.StartupPath.Replace("\\bin\\Debug", "");

                if (!Directory.Exists(MyPath + "\\ErrorLog\\"))
                {
                    Directory.CreateDirectory(MyPath + "\\ErrorLog\\");
                }


                fullpath = MyPath + "\\ErrorLog\\" + filename;

                if (objexception.Count > 0)
                {
                    foreach (string color in objexception)
                    {
                        err += color;
                    }
                }
                else
                {
                    err = "No Error Log";
                }
                MessageBox.Show(err, "Error Detail");
                txtFilePath.Visible = true;
                txtFilePath.Text = "Log File Path:- " + fullpath;
                txtFilePath.ForeColor = Color.Green;
            }
            catch (Exception ex)
            {
                Logger.Log(ex, null);
                MessageBox.Show(ex.Message);
            }
        }
        private void Grid_Style1()
        {
            dataGridView1.RowsDefaultCellStyle.BackColor = Color.Bisque;
        }

        private void btnCleasing_Click(object sender, EventArgs e)
        {
            ShowProgressbar();
            if (!defaultDataLoader.IsBusy)
                defaultDataLoader.RunWorkerAsync();
        }
        void ProcessCleansing()
        {

            //btnCancel.Visible = false;
            cbSheetList.DataSource = USPSCleanUp.UploadFile.dtsheetName;


            if (dt.Tables.Count > 0)
            {
                foreach (var sheet in USPSCleanUp.UploadFile.dtsheetName)
                {
                    Text_FilePath();

                    btnUpload.Invoke((MethodInvoker)delegate
                    {
                        btnUpload.Visible = false;
                        textBox1.Visible = false;
                    });

                    var currentsheet = USPSCleanUp.SetColumns.FullColumnList.Where(x => x.SheetName == sheet).FirstOrDefault();
                    if (currentsheet != null)
                    {
                        Addressline1col = currentsheet.Addressline1col;
                        AddressLine2col = currentsheet.AddressLine2col;
                        Zipcol = currentsheet.Zipcol;
                        Statecol = currentsheet.Statecol;
                        Citycol = currentsheet.Citycol;
                        Housenocol = currentsheet.Housenocol;
                        if (dt.Tables[sheet].Rows.Count > 0)
                        {
                            btnCleasing.Invoke((MethodInvoker)delegate
                            {
                                btnCleasing.Text = "Loading.....";
                                btnCleasing.Enabled = false;
                            });
                        }

                        if (dt.Tables[sheet].Rows != null && dt.Tables[sheet].Rows.Count > 0)
                        {
                            var lcolumnlist = dt.Tables[sheet].Columns;

                            string whereclasue = "";

                            foreach (DataRow row in dt.Tables[sheet].Rows)
                            {
                                whereclasue = "";
                                for (int i = 0; i < lcolumnlist.Count; i++)
                                {
                                    string columnname = lcolumnlist[i].ColumnName.ToString();
                                    string columnValue = string.Empty;
                                    try
                                    {
                                        columnValue = row[columnname].ToString();
                                    }
                                    catch (Exception ex)
                                    {

                                    }
                                    if (!string.IsNullOrEmpty(columnValue))
                                    {
                                        if (columnValue.Contains("'"))
                                            columnValue = columnValue.Replace("'", "''");

                                        if (!string.IsNullOrEmpty(whereclasue))
                                            whereclasue += " and ";
                                        whereclasue += "[" + columnname + "] =" + "'" + columnValue + "'";

                                        //if (i < lcolumnlist.Count - 1)
                                        //{
                                        //    whereclasue += " and ";
                                        //}
                                    }
                                }

                                var CurrentRow = olddt.Tables[sheet].Select(whereclasue).FirstOrDefault();

                                foreach (DataColumn column in dt.Tables[sheet].Columns)
                                {
                                    if (Housenocol != "")
                                    {
                                        if (column.ToString() == Housenocol)
                                        {
                                            foreach (DataColumn column1 in dt.Tables[sheet].Columns)
                                            {
                                                if (column1.ToString() == Addressline1col)
                                                {
                                                    try
                                                    {
                                                        addressline1 = row.Field<string>(Addressline1col);
                                                        //housnno = row.Field<string>(Housenocol);
                                                        //address1 = housnno + s1 + addressline1;
                                                        row.SetField(Addressline1col, addressline1);
                                                    }
                                                    catch (Exception ex)
                                                    {
                                                        //CurrentRow[newColumn] = ex.Message.ToString();
                                                        this.logerrors(ex, sheet);
                                                        continue;
                                                    }
                                                }
                                            }
                                        }
                                    }
                                    else
                                    {
                                        foreach (DataColumn column2 in dt.Tables[sheet].Columns)
                                        {
                                            if (column2.ToString() == Addressline1col)
                                            {
                                                try
                                                {
                                                    addressline1 = row.Field<string>(Addressline1col);
                                                    row.SetField(Addressline1col, addressline1);
                                                }
                                                catch (Exception ex)
                                                {
                                                    //CurrentRow[newColumn] = ex.Message.ToString();
                                                    this.logerrors(ex, sheet);
                                                    continue;
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                            dt.Tables[sheet].AcceptChanges();
                        }
                        USPSManager m = new USPSManager("607FRIEN1074", true);

                        if (dt.Tables[sheet].Rows != null && dt.Tables[sheet].Rows.Count > 0)
                        {
                            var lcolumnlist = dt.Tables[sheet].Columns;
                            string whereclasue = "";

                            foreach (DataRow row in dt.Tables[sheet].Rows)
                            {
                                whereclasue = "";
                                for (int i = 0; i < lcolumnlist.Count; i++)
                                {
                                    string columnname = lcolumnlist[i].ColumnName.ToString();
                                    string columnValue = string.Empty;
                                    try
                                    {
                                        columnValue = row[columnname].ToString();
                                    }
                                    catch (Exception ex)
                                    {

                                    }
                                    if (!string.IsNullOrEmpty(columnValue))
                                    {
                                        if (columnValue.Contains("'"))
                                            columnValue = columnValue.Replace("'", "''");

                                        if (!string.IsNullOrEmpty(whereclasue))
                                            whereclasue += " and ";

                                        whereclasue += "[" + columnname + "] =" + "'" + columnValue + "'";

                                        //if (i < lcolumnlist.Count - 1)
                                        //{
                                        //    whereclasue += " and ";
                                        //}
                                    }
                                }
                                var CurrentRow = olddt.Tables[sheet].Select(whereclasue).FirstOrDefault();
                                foreach (DataColumn column in dt.Tables[sheet].Columns)
                                {
                                    if (column.ToString() == Addressline1col)
                                    {
                                        foreach (DataColumn column1 in dt.Tables[sheet].Columns)
                                        {
                                            if (column1.ToString() == Zipcol)
                                            {
                                                try
                                                {
                                                    if (!string.IsNullOrEmpty(Addressline1col))
                                                    {
                                                        var add1Value = row.Field<object>(Addressline1col);
                                                        if (add1Value != null)
                                                            oldaddress1 = Convert.ToString(add1Value);
                                                    }

                                                    if (!string.IsNullOrEmpty(AddressLine2col))
                                                    {
                                                        var value = row.Field<object>(AddressLine2col);
                                                        if (value != null)
                                                            oldaddress2 = Convert.ToString(value);
                                                    }

                                                    if (!string.IsNullOrEmpty(Zipcol))
                                                    {
                                                        var zipValue = row.Field<object>(Zipcol);
                                                        //var zipValue = row.Field<double?>(Zipcol);
                                                        if (zipValue != null)
                                                            oldzip = Convert.ToString(zipValue);
                                                    }

                                                    if (!string.IsNullOrEmpty(Statecol))
                                                    {
                                                        var value = row.Field<object>(Statecol);
                                                        if (value != null)
                                                            oldstate = Convert.ToString(value);
                                                    }

                                                    if (!string.IsNullOrEmpty(Citycol))
                                                    {
                                                        var value = row.Field<object>(Citycol);
                                                        if (value != null)
                                                            oldcity = Convert.ToString(value);
                                                    }

                                                    if (!string.IsNullOrEmpty(Housenocol))
                                                    {
                                                        var value = row.Field<object>(Housenocol);
                                                        if (value != null)
                                                            oldHouseNo = Convert.ToString(value);
                                                    }

                                                    lineno++;
                                                    Address a = new Address();
                                                    a.Address1 = GetEmptyStringIfNull(oldHouseNo) + s1 + GetEmptyStringIfNull(oldaddress1);
                                                    a.Address2 = GetEmptyStringIfNull(oldaddress2);
                                                    a.Zip = GetEmptyStringIfNull(oldzip);
                                                    a.State = GetEmptyStringIfNull(oldstate);
                                                    a.City = GetEmptyStringIfNull(oldcity);

                                                    Address validatedAddress;

                                                    validatedAddress = m.ValidateAddress(a);
                                                    newaddress = validatedAddress.Address2;
                                                    newzip = validatedAddress.Zip;
                                                    newcity = validatedAddress.City;
                                                    newstate = validatedAddress.State;

                                                }
                                                catch (Exception ex)
                                                {
                                                    //CurrentRow[newColumn.ColumnName] = ex.Message.ToString();
                                                    CurrentRow.SetField(newColumn.ColumnName, ex.Message.ToString());
                                                    this.logerrors(ex, sheet);
                                                    continue;
                                                }
                                                row.SetField(Addressline1col, newaddress);
                                                row.SetField(Zipcol, newzip);
                                                if (Statecol != "")
                                                {
                                                    row.SetField(Statecol, newstate);
                                                }

                                                if (Citycol != "")
                                                {
                                                    row.SetField(Citycol, newcity);
                                                }


                                            }

                                        }
                                    }

                                }
                            }
                            dt.Tables[sheet].AcceptChanges();

                            lblResult.Invoke((MethodInvoker)delegate
                            {
                                lblResult.Visible = true;
                                lblResult.Text = "Record Successfully Updated";
                                lblResult.ForeColor = Color.Green;
                            });
                        }



                    }

                    btnUpload.Invoke((MethodInvoker)delegate
                    {
                        btnUpload.Text = "Upload";
                        btnUpload.Enabled = true;
                        btnExportNewExcel.Visible = true;
                        // btnUpdateExitingExcel.Visible = true;
                        btnErrorLog.Visible = true;
                        btnUpload.Visible = true;
                        textBox1.Visible = true;

                        btnCleasing.Text = "Cleaning Address";
                        btnCleasing.Visible = false;
                        btnDublicate.Visible = true;
                        btnClearDublicate.Visible = true;
                        lblResult.Visible = true;
                        lblSheetName.Visible = true;
                        cbSheetList.Visible = true;
                        string tablename = cbSheetList.SelectedValue.ToString();
                        dataGridView1.Visible = true;
                        dataGridView1.DataSource = dt.Tables[tablename];
                    });
                }
            }
            else
            {
                lblmessage.Invoke((MethodInvoker)delegate
                {
                    lblmessage.Visible = true;
                    lblmessage.Text = "No Record Found";
                    lblmessage.ForeColor = Color.Red;
                });
            }
        }
        private void btnDublicate_Click(object sender, EventArgs e)
        {
            Text_FilePath();
            lblmessage.Visible = false;
            lblStatus.Visible = false;
            lblResult.Visible = false;
            dataGridView1.DataSource = null;
            DataTable filterTable = new DataTable();
            string tablename = cbSheetList.SelectedValue.ToString();
            if (dt.Tables[tablename].Rows.Count > 0)
            {
                var a = dt.Tables[tablename].Rows;
                string Addressline1 = Addressline1col.ToString();
                string Houseno = Housenocol.ToString();
                string AddressLine2 = AddressLine2col.ToString();
                string City = Citycol.ToString();
                string State = Statecol.ToString();
                string Zip = Zipcol.ToString();
                //.GroupBy(r => new { Col1 = r[Addressline1], Col2 = r[Housenocol], Col3 = r[AddressLine2col], Col4 = r[Citycol], Col5 = r[Statecol], Col6 = r[Zipcol] })

                if (Addressline1 != "" && Zip != "")
                {
                    var Result = dt.Tables[tablename].AsEnumerable()
             .GroupBy(r => new { Col1 = r[Addressline1], Col6 = r[Zip] })
            .Select(group => new
            {
                MyList = group.ToList(),
                count = group.Count()
            })

                     .ToList().Where(x => x.count > 1).ToList();

                    filterTable = dt.Tables[tablename].Clone();
                    if (Result.Count > 0)
                    {
                        for (int i = 0; i < Result.Count; i++)
                        {

                            for (int k = 0; k < Result[i].MyList.Count; k++)
                            {

                                filterTable.Rows.Add(Result[i].MyList[k].ItemArray.ToArray());
                            }
                        }
                    }
                }
            }

            dataGridView1.DataSource = filterTable;
        }

        private void btnClearDublicate_Click(object sender, EventArgs e)
        {
            Text_FilePath();
            lblmessage.Visible = false;
            lblStatus.Visible = false;
            lblResult.Visible = false;
            dataGridView1.DataSource = null;
            DataTable filterTable1 = new DataTable();

            string tablename = cbSheetList.SelectedValue.ToString();

            if (dt.Tables[tablename].Rows.Count > 0)
            {
                var a = dt.Tables[tablename].Rows;
                string Addressline1 = Addressline1col.ToString();
                string Houseno = Housenocol.ToString();
                string AddressLine2 = AddressLine2col.ToString();
                string City = Citycol.ToString();
                string State = Statecol.ToString();
                string Zip = Zipcol.ToString();
                if (Addressline1 != "" && Zip != "")
                {
                    var Result = dt.Tables[tablename].AsEnumerable()
                   .GroupBy(r => new { Col1 = r[Addressline1], Col6 = r[Zip] })
                   .Select(group => new
                   {
                       MyList = group.ToList(),
                       count = group.Count()
                   })

                                .ToList();

                    filterTable1 = dt.Tables[tablename].Clone();
                    if (Result.Count > 0)
                    {
                        for (int i = 0; i <= Result.Count - 1; i++)
                        {
                            if (Result[i].MyList.Count > 0)
                            {
                                filterTable1.Rows.Add(Result[i].MyList[0].ItemArray.ToArray());
                            }
                        }
                    }
                }
                dt.Tables[tablename].Clear();




                foreach (DataRow dr in filterTable1.Rows)
                {

                    dt.Tables[tablename].Rows.Add(dr.ItemArray);
                }

            }


            dataGridView1.DataSource = filterTable1;
        }

        string GetEmptyStringIfNull(string inputString)
        {
            string res = string.Empty;
            if (!string.IsNullOrEmpty(inputString))
                res = inputString;

            return res;
        }
    }
}
