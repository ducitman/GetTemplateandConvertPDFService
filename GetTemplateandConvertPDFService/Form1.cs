﻿using DevExpress.XtraEditors;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using DevExpress.Spreadsheet;
using GetTemplateandConvertPDFService.Models.Entities;
using GetTemplateandConvertPDFService.Models;
using System.IO;
using DevExpress.XtraPrinting;
using Microsoft.Office.Interop.Excel;
using _Excel = Microsoft.Office.Interop.Excel;
using DataTable = System.Data.DataTable;
using Workbook = Microsoft.Office.Interop.Excel.Workbook;
using Worksheet = Microsoft.Office.Interop.Excel.Worksheet;
using log4net;
using log4net.Config;
using System.Diagnostics;

namespace GetTemplateandConvertPDFService
{
    public partial class Form1 : Form
    {        
        public DB _DB = new DB();
        public static string _TemplateLink = @"\\10.118.11.28\Spec_Doc\";
        public static string _TemplateCoppy = @"\\10.118.11.42\python\TS process convert";
        public static string _StoragePdf = @"\\10.118.11.42\process_pdf\MTRL";
        public DBBusiness _DbBusiness = new DBBusiness();
        public int _maxMinute = 5;
        public int _countMinute;
        public int _countSecond = 0;

        public Form1()
        {
            InitializeComponent();
        }

        public List<string> machineGroup = new List<string>();
        public string listMachineGroups = "";
        public bool isRunning = false;
        private bool IsInputInArray(string[] array, string input)
        {
            return array.Any(item => item == input);
        }
        private void Form1_Load(object sender, EventArgs e)
        {
            DataTable tbl = listMachineGroup();

            
            DataTable tbl1 = HistoryRefresh();
            gridControl1.DataSource = tbl1;

            foreach (DataRow item in tbl.Rows)
            {
                machineGroup.Add(item["MACHINEGROUP"].ToString().Trim());
            }
            getDataField();
            //string st = "";
            timer1.Enabled = true;
            timer1.Start();
        }
                
        private DataTable getLastestVersionInf(string machineGroup)
        {
            DataTable tbl = new DataTable();
            
            string query = @"SELECT c.PARTID, c.PARTNUMBER, c.MACHINEGROUP, c.REVISIONNO, c.PREVREVISIONNO, c.TRIALFLAG, c.APPLICATIONDATE
                            FROM
                                (SELECT 
                                    A.PARTNUMBER AS PARTNUMBER, 
                                    A.MACHINEGROUP AS MACHINEGROUP, 
                                    A.PARTID AS PARTID, 
                                    A.REVISIONNO AS REVISIONNO, 
                                    LAG(A.REVISIONNO, 1, 0) OVER (PARTITION BY A.PARTNUMBER ORDER BY A.REVISIONNO) AS PREVREVISIONNO, 
                                    A.TRIALFLAG AS TRIALFLAG, 
                                    B.APPLICATIONDATE AS APPLICATIONDATE
                                FROM 
                                    Z_SPECMTRLPROCESSSTATUS A
                                JOIN 
                                    Z_SPECMTRLREVISION B ON A.PARTNUMBER = B.PARTNUMBER AND A.REVISIONNO = B.REVISIONNO AND A.MACHINEGROUP = B.MACHINEGROUP
                                WHERE 
                                    A.APPROVESTATE = '9' 
                                    AND A.MACHINEGROUP = '"+machineGroup+"' " +@" 
                                    AND TO_DATE(SUBSTR(B.APPLICATIONDATE, 1, 8), 'YYYYMMDD') < SYSDATE) C
                                    JOIN
                                    (SELECT MAX(S.REVISIONNO) AS REVISIONNO , S.PARTNUMBER from Z_SPECMTRLPROCESSSTATUS S
                                    JOIN
                                    Z_SPECMTRLREVISION V ON 
                                    S.PARTNUMBER = V.PARTNUMBER AND S.REVISIONNO = V.REVISIONNO AND S.MACHINEGROUP = V.MACHINEGROUP
                                    WHERE TO_DATE(SUBSTR(V.APPLICATIONDATE,1,8),'YYYYMMDD') < SYSDATE AND S.APPROVESTATE = '9' AND S.MACHINEGROUP = '"+machineGroup+"'" +@"
                                    GROUP BY S.PARTNUMBER )D ON d.revisionno = c.revisionno AND d.partnumber = c.partnumber";
            tbl =  _DbBusiness.GetOrclData(query);
            return tbl;
        }
        
        private DataTable listMachineGroup()
        {
            string query22 = "select distinct machinegroup from Z_SPECMTRLREVISION where MACHINEGROUP NOT IN ('CR','KZ') order by machinegroup asc";
            DataTable tbl = _DbBusiness.GetOrclData(query22);
            return tbl;
        }

        private string TemplatePathReturn(string machineGroup)
        {
            string path = "";
            switch (machineGroup)
            {
                case "51":
                case "5A":
                case "5B":
                    path = "Ext-51-Top(TT).xlsm";
                    break;
                case "53":
                case "5P":
                case "5Q":
                    path = "Ext-53-Side(TS).xlsm";
                    break;
                case "55":
                    path = "Ext-55-BF(BS).xlsm";
                    break;
                case "61":
                    path = "Cut-61-PLy(PL).xlsm";
                    break;
                case "62":
                    path = "Cut-62-SQPreset(PL).xlsm";
                    break;
                case "65":
                    path = "Cut-65-BexterCut(BR).xlsm";
                    break;
                case "6D":
                    path = "Cut-64-BexterIns(SR)04.xlsm";
                    break;
                case "6C":
                    path = "Cut-64-BexterIns(SR)03.xlsm";
                    break;
                case "64":
                    path = "Cut-64-BexterIns(SR)02.xlsm";
                    break;
                case "6B":
                    path = "Cut-64-BexterIns(SR)01.xlsm";
                    break;
                case "66":
                    path = "Cut-66-Lay1st(SP).xlsm";
                    break;
                case "68":
                    path = "Cut-68-Lay2nd(SP).xlsm";
                    break;
                case "69":
                    path = "Cut-69-CCHCut(SL).xlsm";
                    break;
                case "63":
                    path = "Cut-63-CCHSlit(SL).xlsm";
                    break;
                case "70":
                    path = "Cut-70-IL(IL).xlsm";
                    break;
                case "71":
                    path = "Cut-71-PA(PA).xlsm";
                    break;
                case "74":
                    path = "Cut-74-TexCal(TX).xlsm";
                    break;
                case "75":
                    path = "Cut-75-MiniCal(GW).xlsm";
                    break;
                case "77":
                    path = "Cut-77-MiniSlit(GS).xlsm";
                    break;
                case "81":
                    path = "BP-81-BD(BD).xlsm";
                    break;
                case "85":
                    path = "BP-85-BFPreset(BP).xlsm";
                    break;
            }
            return path;
        }
        private DataTable GetAlreadyConvertSize(string Machinegroup)
        {
            string query = "SELECT * FROM MTRL_ProcessConvertLog WHERE MachineGroup = '"+Machinegroup+"'";
            DataTable tbl = _DbBusiness.GetSQLData(query);
            return tbl;
        }
        
        private DataTable GetAllFieldData(string partNumber, string revision, string prevrevision, string machineGroup, string listFieldID)
        {
            string GetCurrentVerFieldDataQuery = "";
            if ( revision != "")
            {
                GetCurrentVerFieldDataQuery  = @"SELECT TRIM(SPECM.FIELDID) AS FIELDID,TRIM(SPECM.FIELDDATA) AS FIELDDATA,TRIM(SPECM.REVISIONNO) AS REVISIONNO,FMAS.DATAKIND AS DATAKIND FROM Z_SPECMTRLMASTER specm
                                LEFT JOIN Z_SPECMTRLFIELDIDMASTER fmas  ON SPECM.FIELDID = FMAS.FIELDID AND FMAS.MACHINEGROUP = SPECM.MACHINEGROUP 
                                WHERE SPECM.REVISIONNO IN ('"+ revision.Trim() +"','"+ prevrevision.Trim() +"') AND SPECM.PARTNUMBER ='" + partNumber.Trim() + "' AND SPECM.MACHINEGROUP='" + machineGroup.Trim()
                                           + "' AND SPECM.FIELDID IN " + listFieldID
                                       + " order by SPECM.FIELDID asc";
                return _DbBusiness.GetOrclData(GetCurrentVerFieldDataQuery);
            }
            else
            {
                return null;
            }            
        }
        private DataTable GetDataUnconvertDTB(DataTable dataCompletedConvertHistory, DataTable dataLastestApprovedProcess)
        {
            DataTable dataUnConvertedProcess = new DataTable();
            dataUnConvertedProcess.Columns.Add("PARTNUMBER");
            dataUnConvertedProcess.Columns.Add("REVISIONNO");
            dataUnConvertedProcess.Columns.Add("PREVREVISIONNO");
            dataUnConvertedProcess.Columns.Add("APPLICATIONDATE");
            dataUnConvertedProcess.Columns.Add("MACHINEGROUP");

            if (dataCompletedConvertHistory.Rows.Count > 0)
            {
                foreach (DataRow row in dataLastestApprovedProcess.Rows)
                {
                    if (!ExistsInConvertHistory(row, dataCompletedConvertHistory))
                    {
                        DataRow NewRow = dataUnConvertedProcess.NewRow();
                        NewRow["PARTNUMBER"] = row["PARTNUMBER"].ToString().Trim();
                        NewRow["REVISIONNO"] = row["REVISIONNO"].ToString().Trim();
                        NewRow["PREVREVISIONNO"] = row["PREVREVISIONNO"].ToString().Trim();
                        NewRow["APPLICATIONDATE"] = row["APPLICATIONDATE"].ToString().Trim();
                        NewRow["MACHINEGROUP"] = row["MACHINEGROUP"].ToString().Trim();
                        dataUnConvertedProcess.Rows.Add(NewRow);
                    }
                }
            }
            else
            {
                dataUnConvertedProcess = dataLastestApprovedProcess;
            }

            return dataUnConvertedProcess;
        }
        private void getDataField()
         {
            isRunning = true;
            int startFileRow = 9;
            /// Get the lastest version of all size
            /// 
            string listFieldID = "";
            DataTable dataUnConvertedProcess = new DataTable();
            DataTable dataCompletedConvertHistory = new DataTable();
            DataTable dataLastestApprovedProcess = new DataTable();
            foreach (string machinegr in machineGroup)
            {
                dataCompletedConvertHistory = GetAlreadyConvertSize(machinegr);
                 dataLastestApprovedProcess = getLastestVersionInf(machinegr);

                dataUnConvertedProcess = GetDataUnconvertDTB(dataCompletedConvertHistory, dataLastestApprovedProcess);

                if(dataUnConvertedProcess.Rows.Count > 0)
                {
                    /// Get the list of FieldID for each machinegroup
                    listFieldID = GetTemplateFieldID(TemplatePathReturn(machinegr));

                    StringBuilder stringBuilder = new StringBuilder();
                    string newlistFieldID = listFieldID.Replace("(", "").Replace(")", "").Replace("'", "");

                    List<string> ListField = new List<string>(newlistFieldID.Split(','));
                    ///

                    foreach (DataRow row in dataUnConvertedProcess.Rows)
                    {
                        _Application excelApp = new _Excel.Application();
                        try
                        {
                            Workbook wb = excelApp.Workbooks.Open(Path.Combine(_TemplateCoppy, TemplatePathReturn(machinegr)));
                            Worksheet wsValues = wb.Worksheets["DataSheetValue"];
                            Worksheet wsHeaders = wb.Worksheets["DataSheetHeader"];
                            Worksheet wsMain = wb.Worksheets[1];

                            DataTable dataFieldAllVersions = GetAllFieldData(row["PARTNUMBER"].ToString(), row["REVISIONNO"].ToString(), row["PREVREVISIONNO"].ToString(), row["MACHINEGROUP"].ToString(), listFieldID);

                            ///Fill Data
                            wsHeaders.Cells[1, 2].Value = "";
                            wsHeaders.Cells[2, 2].Value = "";
                            wsHeaders.Cells[3, 2].Value = row["PARTNUMBER"].ToString().Trim();

                            wsValues.Cells[4, 2].Value = row["REVISIONNO"].ToString().Trim();
                            wsValues.Cells[5, 2].Value = row["PREVREVISIONNO"].ToString().Trim();

                            if (dataFieldAllVersions.Rows.Count > 0)
                            {
                                for (int i = 0; i < ListField.Count; i++)
                                {
                                    string fieldName = ListField[i];
                                    int dataCurrentIndex = dataFieldAllVersions.AsEnumerable()
                                    .Select((Tblrow, TblrowIndex) => new { Row = Tblrow, Index = TblrowIndex })
                                    .FirstOrDefault(item => item.Row.Field<string>("FIELDID") == fieldName && item.Row.Field<string>("REVISIONNO") == row["REVISIONNO"].ToString().Trim())?
                                    .Index ?? -1;
                                    int dataPreviousIndex = dataFieldAllVersions.AsEnumerable()
                                    .Select((Tblrow, TblrowIndex) => new { Row = Tblrow, Index = TblrowIndex })
                                    .FirstOrDefault(item => item.Row.Field<string>("FIELDID") == fieldName && item.Row.Field<string>("REVISIONNO") == row["PREVREVISIONNO"].ToString().Trim())?
                                    .Index ?? -1;

                                    if (dataCurrentIndex >= 0)
                                    {

                                        if (dataFieldAllVersions.Rows[dataCurrentIndex]["DATAKIND"].ToString() == "1" && dataFieldAllVersions.Rows[dataCurrentIndex]["FIELDDATA"].ToString() != "")
                                        {
                                            wsValues.Cells[startFileRow + i, 2].Value = dataFieldAllVersions.Rows[dataCurrentIndex]["FIELDDATA"].ToString();
                                        }
                                        else if (dataFieldAllVersions.Rows[dataCurrentIndex]["DATAKIND"].ToString() == "1" && dataFieldAllVersions.Rows[dataCurrentIndex]["FIELDDATA"].ToString() == "")
                                        {
                                            wsValues.Cells[startFileRow + i, 2].Value = "";
                                        }
                                        else if (dataFieldAllVersions.Rows[dataCurrentIndex]["DATAKIND"].ToString() == "0" && dataFieldAllVersions.Rows[dataCurrentIndex]["FIELDDATA"].ToString() != "")
                                        {
                                            if (!IsNumber(dataFieldAllVersions.Rows[dataCurrentIndex]["FIELDDATA"].ToString()))
                                            {
                                                wsValues.Cells[startFileRow + i, 2].Value = "";
                                            }
                                            else
                                            {
                                                double data = double.Parse(dataFieldAllVersions.Rows[dataCurrentIndex]["FIELDDATA"].ToString());
                                                wsValues.Cells[startFileRow + i, 2].Value = data;
                                            }
                                        }
                                        else
                                        {
                                            wsValues.Cells[startFileRow + i, 2].Value = "";
                                        }
                                    }
                                    if (dataPreviousIndex >= 0)
                                    {
                                        if (dataFieldAllVersions.Rows[dataPreviousIndex]["DATAKIND"].ToString() == "1" && dataFieldAllVersions.Rows[dataPreviousIndex]["FIELDDATA"].ToString() != "")
                                        {
                                            wsValues.Cells[startFileRow + i, 3].Value = dataFieldAllVersions.Rows[dataPreviousIndex]["FIELDDATA"].ToString();
                                        }
                                        else if (dataFieldAllVersions.Rows[dataPreviousIndex]["DATAKIND"].ToString() == "1" && dataFieldAllVersions.Rows[dataPreviousIndex]["FIELDDATA"].ToString() == "")
                                        {
                                            wsValues.Cells[startFileRow + i, 3].Value = "";
                                        }
                                        else if (dataFieldAllVersions.Rows[dataPreviousIndex]["DATAKIND"].ToString() == "0" && dataFieldAllVersions.Rows[dataPreviousIndex]["FIELDDATA"].ToString() != "")
                                        {
                                            if (!IsNumber(dataFieldAllVersions.Rows[dataPreviousIndex]["FIELDDATA"].ToString()))
                                            {
                                                wsValues.Cells[startFileRow + i, 3].Value = "";
                                            }
                                            else
                                            {
                                                double data = double.Parse(dataFieldAllVersions.Rows[dataPreviousIndex]["FIELDDATA"].ToString());
                                                wsValues.Cells[startFileRow + i, 3].Value = data;
                                            }
                                        }
                                        else
                                        {
                                            wsValues.Cells[startFileRow + i, 3].Value = "";
                                        }
                                    }
                                }
                                wsValues.Cells[6, 2].Value = row["APPLICATIONDATE"].ToString().Trim();
                                // Edit Application date , issue date of both size
                            }


                            string fileName = row["PARTNUMBER"].ToString().Trim() + "_" + row["REVISIONNO"].ToString();
                            Range range = wsMain.Range["AL75"];

                            range.Font.Name = "Calibri";
                            range.Font.Size = 10;

                            wb.ForceFullCalculation = true;
                            string outputPath = _StoragePdf + GetLocationDestinationFolder(machinegr);
                            string despath = Path.Combine(outputPath, fileName + ".pdf");

                            wsMain.ExportAsFixedFormat(XlFixedFormatType.xlTypePDF, despath, XlFixedFormatQuality.xlQualityStandard);
                            Thread.Sleep(3000);
                            bool res = CreateConvertPdfHistory(row["PARTNUMBER"].ToString(), row["REVISIONNO"].ToString(), machinegr);

                            wb.Close(false);
                            System.Runtime.InteropServices.Marshal.FinalReleaseComObject(wb);

                            //excelApp.Quit();

                            ReleaseObject(GetProcessID(excelApp));
                            System.Runtime.InteropServices.Marshal.FinalReleaseComObject(excelApp);

                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show(ex.Message);
                        }

                    }
                }
                
            }
            isRunning = false;
        }

        private bool IsNumber(string input)
        {
            // Sử dụng phương thức TryParse của lớp số nguyên (int)
            // để kiểm tra xem chuỗi có phải là số hay không.
            // Nếu TryParse trả về true, tức là chuỗi là số.
            double result;
            bool final = double.TryParse(input, out result);
            return final;
        }

        private static int GetProcessID(_Application app)
        {
            int processID = 0;
            
            try
            {
                int hwnd = app.Hwnd;
                int threadId = NativeMethods.GetWindowThreadProcessId(hwnd, out processID);
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error: " + ex.Message);
            }

            return processID;
        }
        internal static class NativeMethods
        {
            [System.Runtime.InteropServices.DllImport("user32.dll")]
            internal static extern int GetWindowThreadProcessId(int hwnd, out int processId);
        }
      
        private static void ReleaseObject(int pid)
        {
            try
            {
                Process excelProcess = Process.GetProcessById(pid);
                excelProcess.Kill();
            }
            catch(Exception ex)
            {
                Console.WriteLine("Error: " + ex.Message);
            }
            finally
            {
                GC.Collect();
            }
        }
        private bool ExistsInConvertHistory(DataRow row, DataTable dataCompletedConvertHistory)
        {
            // Assume the primary key is PARTNUMBER + REVISIONNO
            string pk = row["PARTNUMBER"].ToString().Trim() + "|" + row["REVISIONNO"].ToString();

            foreach (DataRow historyRow in dataCompletedConvertHistory.Rows)
            {
                string historyPk = historyRow["PARTNAME"].ToString().Trim() + "|" + historyRow["REVISION"].ToString();
                if (pk == historyPk)
                {
                    return true;
                }
            }

            return false;
        }
        private bool CreateConvertPdfHistory(string MaterialSize, string Revision, string MachineGroup)
        {
            bool result = false;

            string query = "INSERT INTO MTRL_ProcessConvertLog(PartName,Revision,Register_Date,Register_By,Status,MachineGroup) " +
                " VALUES ('" + MaterialSize + "','" + Revision + "','" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "','GetProcessData&Template','Successed','" + MachineGroup + "')";
            try
            {
                result = _DbBusiness.ExcuteSQLQuery(query);
            }
            catch(Exception ex)
            {
                throw ex;
            }

            return result;
        }
        private string GetTemplateFieldID(string path) ///Lỗi: Nếu giữa 1 chu kì get data mà ở thư mục Spec gốc thay đổi template thì các thông số value sẽ sai so với thực tế
        {
            //string FullFilePath = Path.Combine(_TemplateLink, path);
            
            
            string DestinationPath = Path.Combine(_TemplateCoppy, path); // Đã lấy data từ thư mục sync spec (10p 1 lần)


            StringBuilder stringBuilder = new StringBuilder();
            try
            {
                if (!Directory.Exists(_TemplateLink))
                {
                    MessageBox.Show("Không thể truy cập vào sharing folder.");// use Log4net
                }

                //if (!File.Exists(FullFilePath))
                //{
                //    MessageBox.Show("File không tồn tại trong sharing folder.");// use Log4net
                //}

                //if (!Directory.Exists(_TemplateCoppy))
                //{
                //    Directory.CreateDirectory(_TemplateCoppy);// use Log4net
                //}
                //if (!File.Exists(Path.Combine(DestinationPath)))
                //{
                //    try
                //    {
                //        using (FileStream sourceFile = new FileStream(FullFilePath, FileMode.Open, FileAccess.Read))
                //        using (FileStream destinationFile = new FileStream(DestinationPath, FileMode.Create, FileAccess.Write))
                //        {
                //            sourceFile.CopyTo(destinationFile);
                //        }

                //        DateTime lastModified = File.GetLastWriteTime(FullFilePath);
                //        //MessageBox.Show($"File đã được sao chép thành công.\nNgày sửa đổi cuối cùng: {lastModified}");// use Log4net
                //    }
                //    catch (Exception ex)
                //    {
                //        MessageBox.Show(ex.Message);
                //    }
                //}                               
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

            #region LoadFieldID
            _Application app = new _Excel.Application();
            Workbook wbDestination = app.Workbooks.Open(DestinationPath);

            Worksheet wsValue = wbDestination.Worksheets["DataSheetValue"];

            int numData = int.Parse(wsValue.Cells[1,2].Value.ToString());
            int startFileRow = 9;

           
            try
            {
                stringBuilder.Append("(");
                for (int i = 0; i < numData; i++)
                {
                    stringBuilder.Append("'" + wsValue.Cells[startFileRow + i, 1].Value.ToString() + "'");
                    stringBuilder.Append(",");
                }
                stringBuilder.Remove(stringBuilder.Length - 1, 1);
                stringBuilder.Append(")");
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message);
            }            
            
            #endregion
            wbDestination.Close(false);
            
            System.Runtime.InteropServices.Marshal.FinalReleaseComObject(wbDestination);
            //app.Quit();
            ReleaseObject(GetProcessID(app));
            System.Runtime.InteropServices.Marshal.FinalReleaseComObject(app);
            return stringBuilder.ToString();
        }
        private string GetLocationDestinationFolder(string machineGroup)
        {
            string path = "";

            switch (machineGroup)
            {
                case "51":
                    path = @"\TOP_1\";
                    break;
                case "5A":
                    path = @"\TOP_2\";
                    break;
                case "5B":
                    path = @"\TOP_3\";
                    break;
                case "53":
                    path = @"\SIDE_1\";
                    break;
                case "5P":
                    path = @"\SIDE_2\";
                    break;
                case "5Q":
                    path = @"\SIDE_3\";
                    break;
                case "55":
                    path = @"\BF\";
                    break;
                case "61":
                    path = @"\PLY\";
                    break;
                case "62":
                    path = @"\SQPRESET\";
                    break;
                case "65":
                    path = @"\BEXTERCUTTER\";
                    break;
                case "6D":
                    path = @"\BEXTERINS_4\";
                    break;
                case "6C":
                    path = @"\BEXTERINS_3\";
                    break;
                case "64":
                    path = @"\BEXTERINS_2\";
                    break;
                case "6B":
                    path = @"\BEXTERINS_1\";
                    break;
                case "66":
                    path = @"\1STLAYER\";
                    break;
                case "68":
                    path = @"\2NDLAYER\";
                    break;
                case "69":
                    path = @"\CCHCUTTER\";
                    break;
                case "63":
                    path = @"\CCHSLITTER\";
                    break;
                case "70":
                    path = @"\2RH\";
                    break;
                case "71":
                    path = @"\3PA\";
                    break;
                case "74":
                    path = @"\TEXCAL\";
                    break;
                case "75":
                    path = @"\SMALLCAL\";
                    break;
                case "77":
                    path = @"\SMALLSLITTER\";
                    break;
                case "81":
                    path = @"\BEAD\";
                    break;
                case "85":
                    path = @"\BEADPRESET\";
                    break;
            }
            return path;
        }
        private DataTable HistoryRefresh()
        {
            string query = @"SELECT PartName AS SIZE, Revision AS PHIENBAN, Register_Date AS NGAYCAPNHAT FROM MTRL_ProcessConvertLog ORDER BY Register_Date DESC";
            DataTable tbl = _DbBusiness.GetSQLData(query);
            return tbl;
        }
        private void timer1_Tick(object sender, EventArgs e)
        {
            if (isRunning == false)
            {
                timer1.Stop();
                getDataField();
                DataTable tbl = HistoryRefresh();
                gridControl1.DataSource = tbl;
                timer1.Start();
            }
        }

        private void simpleButton1_Click(object sender, EventArgs e)
        {
            timer1.Stop();
            DataTable tbl = HistoryRefresh();
            gridControl1.DataSource = tbl;
            timer1.Start();
        }
    }
}
