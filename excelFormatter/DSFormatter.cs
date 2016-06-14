using System;
using System.Text;
using System.Data;
using System.Collections.Generic;

namespace excelFormatter {
    class DSFormatter {
        public static bool format(DataTable srcDt, DataSet targetDS) {
            if (targetDS.Tables.Count == 0) {
                DataTable firstTable = new DataTable();
                firstTable.Columns.Add("申报人", typeof(System.String));
                firstTable.Columns.Add("日期", typeof(System.String));
                firstTable.Columns.Add("Bug", typeof(System.String));
                firstTable.Columns.Add("建议", typeof(System.String));
                firstTable.Columns.Add("补偿申请", typeof(System.String));
                firstTable.Columns.Add("审批", typeof(System.String));
                targetDS.Tables.Add(firstTable);
            }

            var dt = targetDS.Tables[0];
            Record record = createRecord(srcDt);
            if (record == null) {
                return false;
            }
            for (var i = 0; i < record.bugs.Count; i++) {
                var row = dt.NewRow();
                if (i == 0) {
                    row["申报人"] = record.declarant;
                    row["补偿申请"] = record.compensate;
                    row["审批"] = record.Approval;
                }
                row["日期"] = record.bugs[i].date;
                row["Bug"] = record.bugs[i].bugDesc;
                row["建议"] = record.bugs[i].suggest;
                dt.Rows.Add(row);
            }
            return true;
        }
        private static Record createRecord(DataTable srcDt) {
            var authorIndex = getAuthorRowIndex(srcDt);
            if (authorIndex < 0) {
                Console.WriteLine("格式不对");
                return null;
            }
            var rec = new Record();
            rec.declarant = srcDt.Rows[authorIndex][0].ToString();
            rec.compensate = srcDt.Rows[authorIndex][1].ToString();
            rec.Approval = "";
            rec.bugs = createBugList(srcDt);
            if (rec.bugs == null) {
                return null;
            }
            return rec;
        }
        private static List<Bug> createBugList(DataTable srcDt) {
            var bugIndex = getBugRowIndex(srcDt);
            if (bugIndex < 0) {
                Console.WriteLine("格式不对");
                return null;
            }
            List<Bug> buglist = new List<Bug>();
            for (var i = bugIndex; i < srcDt.Rows.Count; i++) {
                var d = srcDt.Rows[i][0].ToString();
                var b = srcDt.Rows[i][1].ToString();
                var s = srcDt.Rows[i][2].ToString();
                if (d.Trim().Length == 0 && b.Trim().Length == 0 && s.Trim().Length == 0) {
                    continue;
                }
                var bug = new Bug();
                bug.date = d;
                bug.bugDesc = b;
                bug.suggest = s;
                buglist.Add(bug);
            }
            return buglist;
        }
        private static void addColumn(DataTable dt, string headText) {
            dt.Columns.Add(headText, typeof(string));
            dt.AcceptChanges();
        }
        private static DataRow addRow(DataTable dt) {
            DataRow row = dt.NewRow();
            dt.Rows.Add(row);
            dt.AcceptChanges();
            return row;
        }
        private static int getAuthorRowIndex(DataTable srcDt) {
            if(srcDt.Columns.Count < 3){
                return -1;
            }
            
            for (var i = 0; i < srcDt.Rows.Count; i++) {
                //if (srcDt.Rows[i][0].ToString() == "申报人" && srcDt.Rows[i][0].ToString() == "补偿申请") {
                //    return i;
                //}
                if (srcDt.Rows[i][0].ToString().Length > 0 && srcDt.Rows[i][0].ToString().Length > 0) {
                    return i;
                }
            }
            return -1;
        }
        private static int getBugRowIndex(DataTable srcDt) {
            if (srcDt.Columns.Count < 3) {
                return -1;
            }
            for (var i = 0; i < srcDt.Rows.Count; i++) {
                if (srcDt.Rows[i][0].ToString() == "日期" && srcDt.Rows[i][1].ToString() == "Bug" && srcDt.Rows[i][2].ToString() == "建议") {
                    return i + 1;
                }
            }
            return -1;
        }
    }
}
