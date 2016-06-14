using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.Collections;
using System.Data.OleDb;
using System.Data.Odbc;
using System.IO;
using System.Windows.Forms;
using Microsoft.Win32;
using System.Diagnostics;
using System.Reflection;
using Microsoft.Office.Interop;


namespace excelFormatter
{
    /// <summary>
    /// 构造函数
    /// </summary>
  public  class ExcelHelp
    {
      string excelversion="";
      /// <summary>
      /// excel版本
      /// </summary>
      /// <param name="_excelversion">"11.0":office 2003；"12.0":office 2007;"14.0":office 2010;默认"11.0"</param>
      public  ExcelHelp(string _excelversion)
      {
          excelversion = _excelversion;
      }
        /// <summary>
        /// 获取连接字符串
        /// </summary>
        /// <param name="excel_filepath">excel路径</param>
        /// <returns></returns>
      public string ExcelConStr(String excel_filepath)
      {
          string strConn = null;
          switch (excelversion)
          {
              case "11.0"://office 2003
                  strConn = string.Format("Provider=Microsoft.Jet.OLEDB.4.0;Data Source='" + excel_filepath + "';Extended Properties='Excel 8.0;HDR=YES;IMEX=1'");//第一行作为列
                  break;
              case "12.0"://office 2007
              case "14.0"://office 2010
                  strConn = string.Format("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" + excel_filepath + "';Extended Properties='Excel 12.0;HDR=YES;IMEX=1'");//第一行作为列
                  break;
              default://默认为office2003
                  strConn = string.Format("Provider=Microsoft.Jet.OLEDB.4.0;Data Source='" + excel_filepath + "';Extended Properties='Excel 8.0;HDR=YES;IMEX=1'");//第一行作为列
                  break;
          }
          return strConn;
      }
        
        private string GetOfficeVersion()
        {
                     Type type;
                    object excel;
                    object version=null;

                    type=Type.GetTypeFromProgID("Excel.Application");

                    if(type==null)
                    {
                        //MessageBox.Show("没有安装excel");
                        return "0";
                    }
                    else
                    {
                        excel= Activator.CreateInstance(type);
                        if(excel==null)
                        {
                            //MessageBox.Show("创建对象出错");
                            return "1";
                        }
                        else
                        {
                            version=type.GetProperty("Version").GetValue(excel,null);
                            type.GetProperty("Visible").SetValue(excel,false,null);
                            type.GetMethod("Quit").Invoke(excel,null);
                            if (version != null)
                            {
                                return version.ToString();
                                //MessageBox.Show("Excel版本号是：" + version.ToString());
                            }
                            else
                                //未知错误
                                return "2";
                        }
                    }                
            }
      
        /// <summary>
        /// 获取excel表名
        /// </summary>
        /// <param name="excel_filepath">excel路径</param>
        public ArrayList GetSheelName(String excel_filepath)
        {
            string strConn = ExcelConStr(excel_filepath);
            OleDbConnection conn = new OleDbConnection(strConn);
            try
            {
                conn.Open();
                DataTable dtSheetName = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                ArrayList strTableNames = new ArrayList();
                for (int k = 0; k < dtSheetName.Rows.Count; k++)
                {
                    string sheetname = dtSheetName.Rows[k][2].ToString();
                    if (sheetname.IndexOf('$') < 0)
                    {
                        strTableNames.Add(sheetname);
                    }
                    else
                    {
                        strTableNames.Add(sheetname.Substring(0,sheetname.Length-1));
                    }
                }
                return strTableNames;
            }
            catch (Exception error)
            {
                throw new Exception(error.Message);
            }
            finally
            {
                conn.Close();
            }
        }

        /// <summary>
        /// 获取Excel  Dataset
        /// </summary>
        /// <param name="Path">excel路径</param>
        /// <param name="sheetname">excel表名</param>
        /// <returns></returns>
        public DataSet GetExcelDs(string Path, String sheetname)
        {
            string strConn = ExcelConStr(Path);
            OleDbConnection conn = new OleDbConnection(strConn); 
            try
            {
                conn.Open();

                string strExcel = "select * from [" + sheetname + "]";
                if (sheetname.IndexOf('$') < 0)//if it have not $
                {
                    strExcel = "select * from[" + sheetname + "$]";
                }

                OleDbDataAdapter da = new OleDbDataAdapter(strExcel, strConn);

                DataSet ds = new DataSet();
                da.Fill(ds, "tablename");
                conn.Close();
                return ds;
            }
            catch (Exception error)
            {
                throw new Exception(error.Message);
            }
            finally
            {
                conn.Close();
            }

        }

        /// <summary>
        /// 使用COM读取Excel
        /// </summary>
        /// <param name="excelFilePath">路径</param>
        /// <param name="index">第index张表，从1开始</param>
        /// <returns>DataTabel</returns>
        public DataSet GetExcelDsWithCom(string excelFilePath, int index) {
            Stopwatch wath = new Stopwatch();
            Microsoft.Office.Interop.Excel.Application app = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Sheets sheets;
            Microsoft.Office.Interop.Excel.Workbook workbook = null;
            object oMissiong = System.Reflection.Missing.Value;
            System.Data.DataTable dt = new System.Data.DataTable();

            wath.Start();

            try {
                if (app == null) {
                    return null;
                }

                workbook = app.Workbooks.Open(excelFilePath, oMissiong, oMissiong, oMissiong, oMissiong, oMissiong, oMissiong, oMissiong, oMissiong, oMissiong, oMissiong, oMissiong, oMissiong, oMissiong, oMissiong);

                //将数据读入到DataTable中——Start   

                sheets = workbook.Worksheets;
                Microsoft.Office.Interop.Excel.Worksheet worksheet = (Microsoft.Office.Interop.Excel.Worksheet)sheets.get_Item(index);//读取第一张表
                if (worksheet == null)
                    return null;

                string cellContent;
                int iRowCount = worksheet.UsedRange.Rows.Count;
                int iColCount = worksheet.UsedRange.Columns.Count;
                Microsoft.Office.Interop.Excel.Range range;

                //负责列头Start
                DataColumn dc;
                int ColumnID = 1;
                range = (Microsoft.Office.Interop.Excel.Range)worksheet.Cells[1, 1];
                while (dt.Columns.Count < iColCount){   //(range.Text.ToString().Trim() != "") {
                    dc = new DataColumn();
                    dc.DataType = System.Type.GetType("System.String");
                    dc.ColumnName = range.Text.ToString().Trim();
                    dt.Columns.Add(dc);

                    range = (Microsoft.Office.Interop.Excel.Range)worksheet.Cells[1, ++ColumnID];
                }
                //End

                for (int iRow = 2; iRow <= iRowCount; iRow++) {
                    DataRow dr = dt.NewRow();

                    for (int iCol = 1; iCol <= iColCount; iCol++) {
                        range = (Microsoft.Office.Interop.Excel.Range)worksheet.Cells[iRow, iCol];

                        cellContent = (range.Value2 == null) ? "" : range.Text.ToString();
                    
                        dr[iCol - 1] = cellContent;
                    }

                    //if (iRow != 1)
                    dt.Rows.Add(dr);
                }

                wath.Stop();
                TimeSpan ts = wath.Elapsed;

                //将数据读入到DataTable中——End
                DataSet ds = new DataSet();
                ds.Tables.Add(dt);
                return ds;
            } catch (Exception ex){
                Console.WriteLine("error:" + ex.Message);
                return null;
            } finally {
                workbook.Close(false, oMissiong, oMissiong);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook);
                workbook = null;
                app.Workbooks.Close();
                app.Quit();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(app);
                app = null;
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }

      /// <summary>
        /// 导出为Text文件
      /// </summary>
      /// <param name="ds"></param>
        public bool SaveAsText(DataSet ds)
        {
            SaveFileDialog saveFileDialog = new SaveFileDialog();  
            saveFileDialog.Filter = "Text files  (*.txt)|*.txt";
            saveFileDialog.FilterIndex = 0;
            saveFileDialog.RestoreDirectory = true;
            saveFileDialog.CreatePrompt = true;
            saveFileDialog.Title = "导出Text文件";
            if (saveFileDialog.ShowDialog() == DialogResult.OK)
            {
                Stream myStream;
                myStream = saveFileDialog.OpenFile();
                StreamWriter sw = new StreamWriter(myStream, System.Text.Encoding.GetEncoding(-0));
                string str = "";
                try
                {
                    for (int i = 0; i < ds.Tables[0].Columns.Count; i++)
                    {
                        if (i > 0)
                        {
                            str += "\t";
                        }
                        str += ds.Tables[0].Columns[i].ColumnName;
                    }
                    sw.WriteLine(str);
                    for (int j = 0; j < ds.Tables[0].Rows.Count; j++)
                    {
                        string tempStr = "";

                        for (int k = 0; k < ds.Tables[0].Columns.Count; k++)
                        {
                            if (k > 0)
                            {
                                tempStr += "\t";
                            }
                            if (ds.Tables[0].Rows[j][k].ToString() == "")
                            {
                                tempStr += "";
                            }
                            tempStr += ds.Tables[0].Rows[j][k].ToString().Trim();

                        }
                        sw.WriteLine(tempStr);
                    }
                    sw.Close();
                    myStream.Close();
                    //MessageBox.Show("数据导出成功！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                    return true;
                }
                catch (Exception e)
                {
                    return false;
                    throw new Exception(e.ToString());
                }
                finally
                {
                    sw.Close();
                    myStream.Close();
                }
            }
            else
            {
                return false;
            }
        }

      /// <summary>
      /// 导出为excel文档
      /// </summary>
      /// <param name="dataSet"></param>
      /// <returns>true：导出成功； false：导出失败</returns>
        public  bool ExportToExcel(DataSet dataSet)
        {
            string fileName="";
            if (dataSet.Tables.Count == 0)
            {
                throw new Exception("没有任何可导出的数据");
            }
            /////////////////////////////////////////
            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.Filter = "Excel files  2007(*.xlsx)|*.xlsx|Excel files  97-2003(*.xls)|*.xls";
            saveFileDialog.FilterIndex = 0;
            saveFileDialog.RestoreDirectory = true;
            saveFileDialog.CreatePrompt = true;
            saveFileDialog.Title = "导出excel文件";
            saveFileDialog.FileName = dataSet.Tables[0].TableName;
            if (saveFileDialog.ShowDialog() == DialogResult.OK)
            {
                fileName = saveFileDialog.FileName;
            }
            else
            {
                return false;
            }
            ////////////////////////
            Microsoft.Office.Interop.Excel.Application excelApplication = new Microsoft.Office.Interop.Excel.Application();
            excelApplication.DisplayAlerts = false;

            Microsoft.Office.Interop.Excel.Workbook workbook = excelApplication.Workbooks.Add(Microsoft.Office.Interop.Excel.XlWBATemplate.xlWBATWorksheet);//Missing.Value

            foreach (DataTable dt in dataSet.Tables)
            {
                Microsoft.Office.Interop.Excel.Worksheet lastWorksheet = (Microsoft.Office.Interop.Excel.Worksheet)workbook.Worksheets.get_Item(workbook.Worksheets.Count);
                //Microsoft.Office.Interop.Excel.Worksheet newSheet = (Microsoft.Office.Interop.Excel.Worksheet)workbook.Worksheets.Add(Type.Missing, lastWorksheet, Type.Missing, Type.Missing);
                Microsoft.Office.Interop.Excel.Worksheet newSheet = (Microsoft.Office.Interop.Excel.Worksheet)workbook.Worksheets.Add(Type.Missing, lastWorksheet, 1, Microsoft.Office.Interop.Excel.XlSheetType.xlWorksheet);


                newSheet.Name = dt.TableName;

                for (int col = 0; col < dt.Columns.Count; col++)
                {
                    newSheet.Cells[1, col + 1] = dt.Columns[col].ColumnName.ToString().Trim();
                }

                for (int row = 0; row < dt.Rows.Count; row++)
                {
                    for (int col = 0; col < dt.Columns.Count; col++)
                    {
                        var range = (Microsoft.Office.Interop.Excel.Range)newSheet.Cells[row + 2, col + 1];
                        range.NumberFormatLocal = "@";
                        //range.Text = dt.Rows[row][col].ToString().Trim();
                        newSheet.Cells[row + 2, col + 1] = dt.Rows[row][col].ToString().Trim();//.ToString()
                    }
                }
                customFormatHeadUI(newSheet);
                customFormatContentUI(newSheet);
            }
           // ((Microsoft.Office.Interop.Excel.Worksheet)workbook.Worksheets.get_Item(1)).Delete();
           // ((Microsoft.Office.Interop.Excel.Worksheet)workbook.Worksheets.get_Item(1)).Delete();
           // ((Microsoft.Office.Interop.Excel.Worksheet)workbook.Worksheets.get_Item(1)).Delete();
           //((Microsoft.Office.Interop.Excel.Worksheet)workbook.Worksheets.get_Item(1)).Activate();
            try
            {
                workbook.Close(true, fileName, System.Reflection.Missing.Value);
            }
            catch (Exception e)
            {
                throw e;
            }
            excelApplication.Quit();
            KillExcel();
            return true;
        }
       
      private void customFormatHeadUI(Microsoft.Office.Interop.Excel.Worksheet _wsh){
           for (int i = 1; i <= _wsh.UsedRange.Columns.Count; i++) {
               var range = ((Microsoft.Office.Interop.Excel.Range)_wsh.Cells[1, i]);
               range.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
               range.Font.Bold = true;
               range.Font.ColorIndex = 2;
               range.Interior.ColorIndex = 1;
               range.RowHeight = 20;
               if (i == 3 || i == 4 || i == 5 || i == 6) {
                   range.ColumnWidth = 30;
               }
           }
           //ws.get_Range(ws.Cells[x1, y1], ws.Cells[x2, y2]).Merge(Type.Missing);
       }
       private void customFormatContentUI(Microsoft.Office.Interop.Excel.Worksheet _wsh) {
           int startRowIndex = 0;
           int endRowIndex = 0;
           for (int i = 2; i <= _wsh.UsedRange.Rows.Count; i++) {
               ((Microsoft.Office.Interop.Excel.Range)_wsh.Cells[i, 1]).RowHeight = 18;
               var text = ((Microsoft.Office.Interop.Excel.Range)_wsh.Cells[i, 1]).Text.ToString().Trim();
               if (text != "") {
                   startRowIndex = i;
                   ((Microsoft.Office.Interop.Excel.Range)(_wsh.Cells[startRowIndex, 1])).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
               } else {
                   endRowIndex = i;

                   _wsh.get_Range(_wsh.Cells[startRowIndex, 1], _wsh.Cells[endRowIndex, 1]).Merge(Type.Missing);//申请人
                   _wsh.get_Range(_wsh.Cells[startRowIndex, 5], _wsh.Cells[endRowIndex, 5]).Merge(Type.Missing);//申请补偿
                   _wsh.get_Range(_wsh.Cells[startRowIndex, 6], _wsh.Cells[endRowIndex, 6]).Merge(Type.Missing);//审批

               }
           }
           //ws.get_Range(ws.Cells[x1, y1], ws.Cells[x2, y2]).Merge(Type.Missing);
       }
        private static void KillExcel()
        {
            Process[] excelProcesses = Process.GetProcessesByName("EXCEL");
            DateTime startTime = new DateTime();

            int processId = 0;
            for (int i = 0; i < excelProcesses.Length; i++)
            {
                if (startTime < excelProcesses[i].StartTime)
                {
                    startTime = excelProcesses[i].StartTime;
                    processId = i;
                }
            }

            if (excelProcesses[processId].HasExited == false)
            {
                excelProcesses[processId].Kill();
            }
        }

     //////////////////////////////////////////
    }
}
