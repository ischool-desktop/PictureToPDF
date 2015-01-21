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
        string _source, _target, _picFolderPath, _templatePath, _vlookupPath;
        BackgroundWorker _BW;
        Workbook _WB;
        List<string> _formats = new List<string>() { "", ".jpg", ".png" };
        string _sheetName = "學生資料";

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

                IDNumberLookup lookup = new IDNumberLookup(_vlookupPath);

                int data_count = table.Rows.Count;
                int progress = 1;
                int error_count = 0;

                foreach (DataRow row in table.Rows)
                {
                    try
                    {
                        stream.Seek(0, SeekOrigin.Begin);

                        PictureToPDF.IDNumberLookup.StudentData stu = lookup.GetStudentData(row["班級名稱"] + "", row["座號"] + "");

                        if (stu == null)
                        {
                            stu = new IDNumberLookup.StudentData();
                        }

                        //名字不對就顯示查無證號 + error_count
                        if (stu.Name != row["學生姓名"] + "")
                        {
                            stu.IDNumber = string.Format("查無證號{0}{1}", error_count, row["學生姓名"] + "");
                            error_count++;
                        }
                            
                        string fileName = string.Format("{0}\\{1}.pdf", _target, Util.GenerateFileName(stu.IDNumber));

                        Document doc = new Document(stream);
                        doc.MailMerge.FieldMergingCallback = new MergeCallBack();

                        doc.MailMerge.Execute(row);
                        doc.MailMerge.DeleteFields();

                        doc.Save(fileName, Aspose.Words.SaveFormat.Pdf);
                        ReportProgress(data_count, progress);
                        progress++;
                        //doc.Save(fileName, Aspose.Words.SaveFormat.Doc);
                    }
                    catch (Exception error)
                    {
                        MessageBox.Show("過程發生錯誤:" + error.Message);
                    }
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

            //檢查檔案存在
            _templatePath = Path.GetDirectoryName(Assembly.GetEntryAssembly().Location) + "\\template.doc";
            _vlookupPath = Path.GetDirectoryName(Assembly.GetEntryAssembly().Location) + "\\vlookup.xls";

            if (!File.Exists(_templatePath))
            {
                MessageBox.Show("找不到範本檔案:" + _templatePath);
                return;
            }
            else if (!File.Exists(_vlookupPath))
            {
                MessageBox.Show("找不到對照檔案:" + _vlookupPath);
                return;
            }

            //嘗試讀取Excel檔
            try
            {
                _WB = new Workbook(_source);

                if (_WB.Worksheets[_sheetName] == null)
                {
                    MessageBox.Show(string.Format("Excel檔中找不到 '{0}' 頁面", _sheetName));
                    return;
                }
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
            //ws在此之前已檢查存在
            Worksheet ws = _WB.Worksheets[_sheetName];

            //Sheet1 Column對照
            Dictionary<string, int> sheetMapping = GetColumnMapping(ws);

            //DataTable欄位建立
            DataTable table = new DataTable();
            table.Columns.Add("學號");
            table.Columns.Add("座號");
            table.Columns.Add("學生姓名");
            table.Columns.Add("學生圖片");
            table.Columns.Add("學年度學期");
            table.Columns.Add("指導老師");
            table.Columns.Add("實習老師");
            table.Columns.Add("指導老師圖片");
            table.Columns.Add("實習老師圖片");
            table.Columns.Add("標題");
            table.Columns.Add("班級名稱");

            //讀取_WB資料
            for (int i = 1; i <= _WB.Worksheets[_sheetName].Cells.MaxDataRow; i++)
            {
                DataRow row = table.NewRow();

                string 學號 = GetColumnValue(ws, i, sheetMapping, "學號");
                string 座號 = GetColumnValue(ws, i, sheetMapping, "座號");
                string 學生姓名 = GetColumnValue(ws, i, sheetMapping, "學生姓名");
                string 學生圖片 = GetColumnValue(ws, i, sheetMapping, "學生圖片");
                string 學年度學期 = GetColumnValue(ws, i, sheetMapping, "學年度學期");
                string 指導老師 = GetColumnValue(ws, i, sheetMapping, "指導老師");
                string 實習老師 = GetColumnValue(ws, i, sheetMapping, "實習老師");
                string 指導老師圖片 = GetColumnValue(ws, i, sheetMapping, "指導老師圖片");
                string 實習老師圖片 = GetColumnValue(ws, i, sheetMapping, "實習老師圖片");
                string 標題 = GetColumnValue(ws, i, sheetMapping, "標題");
                string 班級名稱 = GetColumnValue(ws, i, sheetMapping, "班級名稱");

                row["學號"] = 學號;
                row["座號"] = 座號;
                row["學生姓名"] = 學生姓名;
                row["學生圖片"] = GetPicPath(學生圖片);
                row["學年度學期"] = 學年度學期;
                row["指導老師"] = 指導老師;
                row["實習老師"] = 實習老師;
                row["指導老師圖片"] = GetPicPath(指導老師圖片);
                row["實習老師圖片"] = GetPicPath(實習老師圖片);
                row["標題"] = 標題;
                row["班級名稱"] = 班級名稱;

                table.Rows.Add(row);
            }

            return table;
        }

        /// <summary>
        /// 取得欄位對照
        /// </summary>
        /// <param name="ws"></param>
        /// <returns></returns>
        private Dictionary<string, int> GetColumnMapping(Worksheet ws)
        {
            Dictionary<string, int> dic = new Dictionary<string, int>();

            for(int i=0;i<=ws.Cells.MaxDataColumn;i++)
            {
                string value = ws.Cells[0,i].Value + "";

                if (!dic.ContainsKey(value))
                    dic.Add(value, i);
            }

            return dic;
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
                int seed = 5;

                //if (args.FieldName != "學生圖片")
                //    seed = 3;

                string filePath = args.FieldValue + "";
                if (!string.IsNullOrWhiteSpace(filePath))
                {
                    //args.Image = GetBitmap(filePath, 300 , 200);

                    using (Bitmap pic = new Bitmap(filePath))
                    {
                        Size size = GetResize(pic, 40 * seed, 30 * seed);
                        Aspose.Words.Fields.MergeFieldImageDimension h = new Aspose.Words.Fields.MergeFieldImageDimension(size.Height, Aspose.Words.Fields.MergeFieldImageDimensionUnit.Point);
                        Aspose.Words.Fields.MergeFieldImageDimension w = new Aspose.Words.Fields.MergeFieldImageDimension(size.Width, Aspose.Words.Fields.MergeFieldImageDimensionUnit.Point);

                        args.ImageHeight = h;
                        args.ImageWidth = w;
                        //args.Image = pic;
                    }

                }
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

        /// <summary>
        /// 取得指定的欄位資料
        /// </summary>
        /// <param name="wb"></param>
        /// <param name="sheetName"></param>
        /// <param name="row"></param>
        /// <param name="dic"></param>
        /// <param name="colName"></param>
        /// <returns></returns>
        private string GetColumnValue(Worksheet ws, int row, Dictionary<string, int> dic, string colName)
        {
            string value = string.Empty;

            if (ws != null && dic.ContainsKey(colName))
            {
                int index = dic[colName];
                value = ws.Cells[row, index].Value + "";
            }

            return value;
        }

        /// <summary>
        /// 取得圖片的完整路徑
        /// </summary>
        /// <param name="fileName"></param>
        /// <returns></returns>
        private string GetPicPath(string fileName)
        {
            string value = null;
            foreach (string format in _formats)
            {
                string path = _picFolderPath + "\\" + fileName + format;
                if (File.Exists(path))
                {
                    value = path;
                    break;
                }
            }

            return value;
        }

        private void lnkSetting_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            Workbook wb = new Workbook(new MemoryStream(Properties.Resources.封面設定檔));

            SaveFileDialog sfd = new SaveFileDialog();
            sfd.Filter = "Excel檔案 (*.xls)|*.xls|所有檔案 (*.*)|*.*";

            if (sfd.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    wb.Save(sfd.FileName, Aspose.Cells.SaveFormat.Excel97To2003);
                    System.Diagnostics.Process.Start(sfd.FileName);
                }
                catch (Exception error)
                {
                    MessageBox.Show(error.Message);
                }
            }
        }

        private void lnkMerge_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            Document doc = new Document(new MemoryStream(Properties.Resources.功能變數));

            SaveFileDialog sfd = new SaveFileDialog();
            sfd.Filter = "Word檔案 (*.doc)|*.doc|所有檔案 (*.*)|*.*";

            if (sfd.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    doc.Save(sfd.FileName, Aspose.Words.SaveFormat.Doc);
                    System.Diagnostics.Process.Start(sfd.FileName);
                }
                catch (Exception error)
                {
                    MessageBox.Show(error.Message);
                }
            }
        }

        private void ReportProgress(int total, int current)
        {
            if (InvokeRequired)
                Invoke(new Action<int, int>(ReportProgress), new object[] { total, current });
            else
            {
                if (total == current)
                    lblProgress.Text = string.Format("產生完成：{0}", total);
                else
                    lblProgress.Text = string.Format("進度：{0}/{1}", current, total);
            }
        }
    }
}
