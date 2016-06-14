using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace excelFormatter {
    public partial class Form1 : Form {
        List<string> fullFileNames = new List<string>();

        public Form1() {
            InitializeComponent();
            this.FormBorderStyle = FormBorderStyle.Fixed3D;
            this.lbx.AllowDrop = true;
        }

        private void btnFormat_Click(object sender, EventArgs e) {
            if (this.fullFileNames.Count == 0) {
                return;
            }
            onStartFormat();
            DataSet targetDS = new DataSet();
            var eh = new ExcelHelp("12.0");
            for (var i = 0; i < this.fullFileNames.Count; i++) {
                string fileName = this.fullFileNames[i];
                var sheets = eh.GetSheelName(fileName);
                //var ds = eh.GetExcelDs(fileName, sheets[0].ToString());
                var ds = eh.GetExcelDsWithCom(fileName, 1);
                try {
                    if (!DSFormatter.format(ds.Tables[0], targetDS)) {
                        Console.WriteLine("文件格式不正确");
                    }
                } catch (Exception ex) {
                    MessageBox.Show("error:" + ex.Message);
                }
            }

            try {
                eh.ExportToExcel(targetDS);
                MessageBox.Show("finished!");
            } catch (Exception ex) {
                MessageBox.Show("error:" + ex.Message);
            }
            onEndFormat();
        }
        private void onStartFormat() {
            this.Cursor = Cursors.WaitCursor;
            this.btnAdd.Enabled = false;
            this.btnClear.Enabled = false;
            this.btnFormat.Enabled = false;
        }
        private void onEndFormat() {
            this.Cursor = Cursors.Default;
            this.btnAdd.Enabled = true;
            this.btnClear.Enabled = true;
            this.btnFormat.Enabled = true;
        }

        private void btnAdd_Click(object sender, EventArgs e) {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Filter = "xlsx文件|*.xlsx|xls文件|*.xls";
            ofd.Multiselect = true;
            if (ofd.ShowDialog() == System.Windows.Forms.DialogResult.OK) {
                for (var i = 0; i < ofd.FileNames.Length; i++) {
                    this.fullFileNames.Add(ofd.FileNames[i]);

                    this.lbx.Items.Add(ofd.FileNames[i]);
                }
            }
        }

        private void btnClear_Click(object sender, EventArgs e) {
            this.lbx.Items.Clear();
            this.fullFileNames.Clear();
        }

        private void lbx_DragEnter(object sender, DragEventArgs e) {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
                e.Effect = DragDropEffects.Link;
            else e.Effect = DragDropEffects.None;
        }

        private void lbx_DragDrop(object sender, DragEventArgs e) {
            Array aryFiles = ((System.Array)e.Data.GetData(DataFormats.FileDrop));
            for (int i = 0; i < aryFiles.Length; i++) {
                this.fullFileNames.Add(aryFiles.GetValue(i).ToString());
                this.lbx.Items.Add(aryFiles.GetValue(i).ToString());
            }
        }
    }
}
