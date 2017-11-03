using System;
using System.Collections.Generic;
using System.IO;
using System.Windows.Forms;

namespace ReplaceExcel
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void btnWord_Click(object sender, EventArgs e)
        {
            // Create an instance of the open file dialog box.
            var openFileDialog1 = new OpenFileDialog();

            // Set filter options and filter index.
            openFileDialog1.Filter = "Word File (.doc)|*.doc";
            openFileDialog1.FilterIndex = 0;

            openFileDialog1.Multiselect = false;

            // Call the ShowDialog method to show the dialog box.
            var result = openFileDialog1.ShowDialog();

            // Process input if the user clicked OK.
            if (result == DialogResult.OK) {
                // Open the selected file to read.
                tbWord.Text = openFileDialog1.FileName;
            }
        }

        private void btnExcel_Click(object sender, EventArgs e)
        {
            // Create an instance of the open file dialog box.
            var openFileDialog1 = new OpenFileDialog();

            // Set filter options and filter index.
            openFileDialog1.Filter = "Excel File (.xls)|*.xls";
            openFileDialog1.FilterIndex = 0;

            openFileDialog1.Multiselect = false;

            // Call the ShowDialog method to show the dialog box.
            var result = openFileDialog1.ShowDialog();

            // Process input if the user clicked OK.
            if (result == DialogResult.OK) {
                // Open the selected file to read.
                tbExcel.Text = openFileDialog1.FileName;
            }
        }

        private void btnStart_Click(object sender, EventArgs e)
        {
            if (tbWord.Text == string.Empty || tbExcel.Text == string.Empty) {
                //MessageBox.Show("请指定")
            }

            var excel = new ExcelWorker();
            var word = new WordWorker();

            try {
                var dt = excel.OpenExcel(tbExcel.Text);

                var headers = new List<string>(dt.Columns.Count);
                for (var c = 0; c < dt.Columns.Count; ++c) {
                    headers.Add(dt.Rows[0][c].ToString());
                }

                var path = System.IO.Path.GetDirectoryName(tbWord.Text);
                path = path + "/输出";
                Directory.CreateDirectory(path);

                var row_cnt = dt.Rows.Count;
                for (var r = 1; r < row_cnt; ++r) {
                    var key = dt.Rows[r][0].ToString();
                    if (string.IsNullOrEmpty(key)) {
                        break;
                    }

                    progressBar1.Minimum = 1;
                    progressBar1.Maximum = row_cnt - 1;
                    progressBar1.Value = r;
                    lbProgress.Text = string.Format("{0}/{1}", r, row_cnt - 1);

                    word.OpenWord(tbWord.Text);

                    for (var c = 0; c < dt.Columns.Count; ++c) {
                        word.ReplaceString(string.Format("${0}$", headers[c]), dt.Rows[r][c].ToString());
                    }

                    word.SaveAs(string.Format(@"{0}\{1}.doc", path, key));
                    word.Close();
                }
                progressBar1.Value = row_cnt - 1;
                lbProgress.Text = "完成";

                word.Destroy();
            }
            finally {
                if (word != null) {
                    word.Close();
                    word.Destroy();
                }
            }
        }

        private void btnTest_Click(object sender, EventArgs e)
        {
            var excel = new ExcelWorker();
            var word = new WordWorker();

            try {
                var dt = excel.OpenExcel(tbExcel.Text);

                var headers = new List<string>(dt.Columns.Count);
                for (var c = 0; c < dt.Columns.Count; ++c) {
                    headers.Add(dt.Rows[0][c].ToString());
                }

                var path = System.IO.Path.GetDirectoryName(tbWord.Text);
                path = path + "/输出";
                Directory.CreateDirectory(path);

                word.OpenWord(tbWord.Text);
                var r = 1;
                var key = dt.Rows[r][0].ToString();

                for (var c = 0; c < dt.Columns.Count; ++c) {
                    word.ReplaceString(string.Format("${0}$", headers[c]), dt.Rows[r][c].ToString());
                }

                word.SaveAs(string.Format(@"{0}\{1}.doc", path, key));
                word.Close();

                lbProgress.Text = "完成";

                word.Destroy();
            }
            finally {
                if (word != null) {
                    word.Close();
                    word.Destroy();
                }
            }
        }
    }
}
