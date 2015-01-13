using Aspose.Cells;
using Aspose.Words;
using Aspose.Words.Reporting;
using FISCA.Presentation.Controls;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace PictureToPDF
{
    public partial class MainForm : BaseForm
    {
        string _source, _target, _picFolderPath, _templatePath;
        BackgroundWorker _BW;
        Workbook _WB;
        List<string> formats = new List<string>() { ".jpg", ".png" };

        public MainForm()
        {
            InitializeComponent();
            Aspose.Cells.License License_cell = new Aspose.Cells.License();
            Aspose.Words.License License_word = new Aspose.Words.License();

            MemoryStream stream = new MemoryStream(Properties.Resources.Aspose_Total);

            License_cell.SetLicense(stream);
            stream.Seek(0, SeekOrigin.Begin);
            License_word.SetLicense(stream);

            _BW = new BackgroundWorker();
            _BW.DoWork += new DoWorkEventHandler(_BW_DoWork);
            _BW.RunWorkerCompleted += new RunWorkerCompletedEventHandler(_BW_RunWorkerCompleted);
        }

        private void _BW_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            try
            {
                if (e.Error != null)
                    MessageBox.Show(e.Error.Message);
                else
                    System.Diagnostics.Process.Start(_target);
            }
            finally
            {
                btnConfirm.Enabled = true;
            }
        }

        private void _BW_DoWork(object sender, DoWorkEventArgs e)
        {
            using (FileStream stream = new FileStream(_templatePath, FileMode.Open))
            {
                DataTable table = GetData();
                foreach (DataRow row in table.Rows)
                {
                    stream.Seek(0, SeekOrigin.Begin);

                    string fileName = _target + "\\" + row["學號"] + "_" + row["學生姓名"] + ".pdf";

                    Document doc = new Document(stream);
                    doc.MailMerge.FieldMergingCallback = new MergeCallBack();

                    doc.MailMerge.Execute(row);

                    doc.Save(fileName, Aspose.Words.SaveFormat.Pdf);
                }
            }
        }

        /// <summary>
        /// 選擇來源路徑
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnSource_Click(object sender, EventArgs e)
        {
            OpenFileDialog open = new OpenFileDialog();

            open.Filter = "Excel檔案 (*.xls)|*.xls|所有檔案 (*.*)|*.*";

            if (open.ShowDialog() == DialogResult.OK)
            {
                this.txtSourcePath.Text = open.FileName;
                this._source = open.FileName;
            }
        }

        /// <summary>
        /// 選擇目的目錄
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnTarget_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog fbd = new FolderBrowserDialog();

            if (fbd.ShowDialog() == DialogResult.OK)
            {
                this.txtTargetPath.Text = fbd.SelectedPath;
                _target = fbd.SelectedPath;
            }
        }

        /// <summary>
        /// 選擇圖片來源目錄
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnPicFolder_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog fbd = new FolderBrowserDialog();

            if (fbd.ShowDialog() == DialogResult.OK)
            {
                this.txtPicFolder.Text = fbd.SelectedPath;
                _picFolderPath = fbd.SelectedPath;
            }
        }

        private void btnExit_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void btnConfirm_Click(object sender, EventArgs e)
        {
            _source = txtSourcePath.Text;
            _target = txtTargetPath.Text;
            _picFolderPath = txtPicFolder.Text;

            //檢查路徑設定
            if (string.IsNullOrWhiteSpace(_source) || string.IsNullOrWhiteSpace(_target) || string.IsNullOrWhiteSpace(_picFolderPath))
            {
                MessageBox.Show("請確認三種路徑皆已設定");
                return;
            }

            //檢查template存在
            _templatePath = Path.GetDirectoryName(Assembly.GetEntryAssembly().Location) + "\\template.doc";
            if (!File.Exists(_templatePath))
            {
                MessageBox.Show("找不到檔案:" + _templatePath);
                return;
            }

            //嘗試讀取Excel檔
            try
            {
                _WB = new Workbook(_source);
            }
            catch (Exception error)
            {
                MessageBox.Show(error.Message);
                return;
            }

            if (_BW.IsBusy)
                MessageBox.Show("正在忙碌中,請稍後再試...");
            else
            {
                _BW.RunWorkerAsync();
                btnConfirm.Enabled = false;
            }
        }

        private DataTable GetData()
        {
            //Column對照
            Dictionary<string, int> columnMapping = new Dictionary<string, int>();
            columnMapping.Add("學號", -1);
            columnMapping.Add("學生姓名", -1);
            columnMapping.Add("學年度學期", -1);
            columnMapping.Add("指導老師", -1);
            columnMapping.Add("實習老師", -1);
            columnMapping.Add("圖片路徑", -1);

            //DataTable欄位建立
            DataTable table = new DataTable();
            table.Columns.Add("學號");
            table.Columns.Add("學生姓名");
            table.Columns.Add("學年度學期");
            table.Columns.Add("指導老師");
            table.Columns.Add("實習老師");
            table.Columns.Add("圖片路徑");

            //取得有資料的最大範圍
            int maxDataCloumn = _WB.Worksheets[0].Cells.MaxDataColumn;
            int MaxDataRow = _WB.Worksheets[0].Cells.MaxDataRow;

            //記憶支援的欄位index
            for (int i = 0; i <= maxDataCloumn; i++)
            {
                string columnName = _WB.Worksheets[0].Cells[0, i].Value + "";

                if (columnMapping.ContainsKey(columnName))
                    columnMapping[columnName] = i;
            }

            //取得每個row
            for (int i = 1; i <= MaxDataRow; i++)
            {
                DataRow row = table.NewRow();

                foreach (string columnName in columnMapping.Keys)
                {
                    int columnIndex = columnMapping[columnName];

                    if (columnIndex >= 0)
                    {
                        string value = _WB.Worksheets[0].Cells[i, columnIndex].Value + "";

                        switch (columnName)
                        {
                            case "學號":
                                row["學號"] = value;
                                break;

                            case "學生姓名":
                                row["學生姓名"] = value;
                                break;

                            case "學年度學期":
                                row["學年度學期"] = value;
                                break;

                            case "指導老師":
                                row["指導老師"] = value;
                                break;

                            case "實習老師":
                                row["實習老師"] = value;
                                break;

                            case "圖片路徑":
                                foreach (string f in formats)
                                {
                                    string path = _picFolderPath + "\\" + value + f;
                                    if (File.Exists(path))
                                    {
                                        row["圖片路徑"] = path;
                                        break;
                                    }
                                }
                                break;
                        }
                    }
                }

                table.Rows.Add(row);
            }

            return table;
        }

        private void txtSourcePath_TextChanged(object sender, EventArgs e)
        {
            _source = txtSourcePath.Text;
        }

        private void txtTargetPath_TextChanged(object sender, EventArgs e)
        {
            _target = txtTargetPath.Text;
        }

        private void txtPicFolder_TextChanged(object sender, EventArgs e)
        {
            _picFolderPath = txtPicFolder.Text;
        }

        class MergeCallBack : IFieldMergingCallback
        {
            public void FieldMerging(FieldMergingArgs args)
            {
                //do nothing...
            }

            public void ImageFieldMerging(ImageFieldMergingArgs args)
            {
                string filePath = args.FieldValue + "";
                if (!string.IsNullOrWhiteSpace(filePath))
                {
                    //args.Image = GetBitmap(filePath, 300 , 200);

                    using (Bitmap pic = new Bitmap(filePath))
                    {
                        int seed = 5;
                        Size size = GetResize(pic, 40 * seed, 30 * seed);
                        Aspose.Words.Fields.MergeFieldImageDimension h = new Aspose.Words.Fields.MergeFieldImageDimension(size.Height, Aspose.Words.Fields.MergeFieldImageDimensionUnit.Point);
                        Aspose.Words.Fields.MergeFieldImageDimension w = new Aspose.Words.Fields.MergeFieldImageDimension(size.Width, Aspose.Words.Fields.MergeFieldImageDimensionUnit.Point);

                        args.ImageHeight = h;
                        args.ImageWidth = w;
                        //args.Image = pic;
                    }

                }
            }

            /// <summary>
            /// 取得比例縮小的圖片
            /// </summary>
            /// <param name="filePath"></param>
            /// <param name="maxWidth"></param>
            /// <param name="maxHeight"></param>
            /// <returns></returns>
            public static Bitmap GetBitmap(string filePath, int maxWidth, int maxHeight)
            {
                //照理說不會爆
                Bitmap photo = new Bitmap(filePath);

                int width = photo.Width;
                int height = photo.Height;
                Size newSize;

                if (width < maxWidth && height < maxHeight)
                    return photo;

                decimal maxW = Convert.ToDecimal(maxWidth);
                decimal maxH = Convert.ToDecimal(maxHeight);

                decimal mp = decimal.Divide(maxW, maxH);
                decimal p = decimal.Divide(width, height);

                // 若長寬比預設比例較寬, 則以傳入之長為縮放基準
                if (mp > p)
                {
                    decimal hp = decimal.Divide(maxH, height);
                    decimal newWidth = decimal.Multiply(hp, width);
                    newSize = new Size(decimal.ToInt32(newWidth), maxHeight);
                }
                else
                {
                    decimal wp = decimal.Divide(maxW, width);
                    decimal newHeight = decimal.Multiply(wp, height);
                    newSize = new Size(maxWidth, decimal.ToInt32(newHeight));
                }

                return new Bitmap(photo, newSize);
            }
        }

        public static Size GetResize(Bitmap photo, int maxWidth, int maxHeight)
        {
            int width = photo.Width;
            int height = photo.Height;
            Size newSize;

            if (width < maxWidth && height < maxHeight)
                return new Size(width, height);

            decimal maxW = Convert.ToDecimal(maxWidth);
            decimal maxH = Convert.ToDecimal(maxHeight);

            decimal mp = decimal.Divide(maxW, maxH);
            decimal p = decimal.Divide(width, height);


            // 若長寬比預設比例較寬, 則以傳入之長為縮放基準
            if (mp > p)
            {
                decimal hp = decimal.Divide(maxH, height);
                decimal newWidth = decimal.Multiply(hp, width);
                newSize = new Size(decimal.ToInt32(newWidth), maxHeight);
            }
            else
            {
                decimal wp = decimal.Divide(maxW, width);
                decimal newHeight = decimal.Multiply(wp, height);
                newSize = new Size(maxWidth, decimal.ToInt32(newHeight));
            }

            return newSize;
        }

        private void MainForm_FormClosed(object sender, FormClosedEventArgs e)
        {
            Properties.Settings.Default.Save();
        }
    }
}
