using System.Collections.Generic;
//using Microsoft.Office.Interop.Excel;
using System.Data;
using System.Data.OleDb;

namespace ReplaceExcel
{
    public class ExcelWorker
    {
        //protected Application eApp;

        public List<string> Headers = new List<string>();
        public List<List<string>> Cells = new List<List<string>>();

        public System.Data.DataTable OpenExcel(string filename)
        {
            //var xlApp = new Application(); ;
            //var workbook = xlApp.Workbooks.Open(filename, 0, false, 5, "", "", false, XlPlatform.xlWindows, "", true, false, 0, true, 1, 0);
            //var n = workbook.Worksheets.Count;
            //var sheetSet = new string[n];
            //var al = new ArrayList();
            //for (int i = 0; i < n; i++) {
            //    sheetSet[i] = ((Worksheet)workbook.Worksheets[i + 1]).Name;
            //}

            //xlApp.Workbooks.Close();

            //var ds = new DataSet();
            //for (int i = 0; i < n; i++) {
            //    ds.Tables.Add();
            //    var strSql = "select * from [" + sheetSet[i] + "$]";

            //    string strConn = sqlconn(filename);
            //    var myConnection = new OleDbConnection(strConn);
            //    myConnection.Open();
            //    var da = new OleDbDataAdapter(strSql, myConnection);
            //    try {
            //        da.Fill(ds.Tables[i]);
            //    }
            //    catch (Exception) {
            //        throw;
            //    }
            //    finally {
            //    }
            //    if (i == n - 1) {
            //        da.Dispose();
            //        da = null;
            //        myConnection.Close();
            //        myConnection = null;
            //    }
            //}

            var hasTitle = false;
            var filePath = filename;
            var fileType = System.IO.Path.GetExtension(filePath);
            using (var ds = new DataSet()) {
                var strCon = string.Format("Provider=Microsoft.Jet.OLEDB.{0}.0;" +
                                "Extended Properties=\"Excel {1}.0;HDR={2};IMEX=1;\";" +
                                "data source={3};",
                                (fileType == ".xls" ? 4 : 12), (fileType == ".xls" ? 8 : 12), (hasTitle ? "Yes" : "NO"), filePath);
                var strCom = " SELECT * FROM [Sheet1$]";
                using (var my_coon = new OleDbConnection(strCon)) {
                    using (var my_cmd = new OleDbDataAdapter(strCom, my_coon)) {
                        my_coon.Open();
                        my_cmd.Fill(ds);
                    }
                }

                if (ds == null || ds.Tables.Count <= 0) {
                    return null;
                }

                return ds.Tables[0];
            }

            //eApp = new Application();
            //eApp.Visible = false;
            //object miss = System.Reflection.Missing.Value;

            //var book = eApp.Workbooks.Open(filename, miss, miss, miss, miss, miss, miss, miss,
            //                      miss, miss, miss, miss, miss, miss, miss);

            //var sheets = book.Worksheets;
            //var sheet = (Worksheet)sheets.get_Item(2);

            //for (var c = 1; c <= sheet.UsedRange.Columns.Count; ++c) {
            //    Headers.Add(sheet.Cells[1, c].Value.ToString());
            //}

            //for (var r = 2; r <= sheet.UsedRange.Rows.Count; ++r) {
            //    var row = new List<string>();

            //    for (var c = 1; c <= sheet.UsedRange.Columns.Count; ++c) {
            //        var cell = sheet.Cells[r, c];
            //        if (cell != null && cell.Value != null) {
            //            row.Add(cell.Value.ToString());
            //        }
            //        else {
            //            row.Add(string.Empty);
            //        }
            //    }

            //    Cells.Add(row);
            //}

            //book.Close();
            //eApp.Quit();
        }
    }
}
