using Microsoft.VisualBasic;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using Excel = Microsoft.Office.Interop.Excel;

namespace KeyWordFinder_Tool
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        string strFilePath = String.Empty;
        string strFolderPath = string.Empty;
        #region Prepare Report
        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                if(!String.IsNullOrEmpty(label5.Text) && !File.Exists(label5.Text))
                {
                    MessageBox.Show("File is not exists on given location. Please check again.");
                    return;
                }
                string[] keys = File.ReadAllLines(strFilePath);
                if (keys.Length == 0)
                {
                    MessageBox.Show("File is empty.");
                    return;
                }
                string fileContents = String.Empty;
                string folderPath = label4.Text;
                System.IO.DirectoryInfo dir = new System.IO.DirectoryInfo(folderPath);
                IEnumerable<System.IO.FileInfo> fileList = dir.GetFiles("*.*", System.IO.SearchOption.AllDirectories);
                string searchTearm =  String.Join(",", keys);
                searchTearm  = Regex.Replace(searchTearm, "<[^>]+>", string.Empty);
                List<string> listOfNames = new List<string>(searchTearm.Split(','));
                listOfNames.RemoveAll(s => string.IsNullOrEmpty(s));
                Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
                if (xlApp == null)
                {
                    MessageBox.Show("Excel is not properly installed!!");
                    return;
                }
                Excel.Workbook xlWorkBook;
                Excel.Worksheet xlWorkSheet;
                object misValue = System.Reflection.Missing.Value;
                xlWorkBook = xlApp.Workbooks.Add(misValue);
                xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

                xlWorkSheet.Cells[1, 1] = "Serial No.";
                xlWorkSheet.Cells[1, 2] = "Folder Name";
                xlWorkSheet.Cells[1, 3] = "File Name";
                xlWorkSheet.Cells[1, 4] = "File Path";
                xlWorkSheet.Cells[1, 5] = "Line No.";
                xlWorkSheet.Cells[1, 6] = "Keyword find";

                long serialNo = 1;
                foreach (System.IO.FileInfo file in fileList)
                {
                    long lineCounter = 0;
                    if (System.IO.File.Exists(file.FullName))
                    {
                        string[] lines = File.ReadAllLines(file.FullName);
                        foreach (var keyword in listOfNames)
                        {
                            var currLine = String.Empty;
                            var currWord = String.Empty;
                            foreach (var line in lines)
                            {
                                if (checkBox1.Checked)
                                {
                                    currLine =  line.ToLower();
                                    currWord = keyword.ToLower();
                                }
                                else
                                {
                                    currLine = line;
                                    currWord = keyword;
                                }
                                if (currLine.Trim().Contains(currWord.Trim()))
                                {
                                    lineCounter = Convert.ToInt64(Array.FindIndex(lines, row => row.Trim().ToLower().Contains(keyword.Trim().ToLower())));
                                    if (lineCounter > 0)
                                    {
                                        serialNo++;
                                        lineCounter++;
                                        xlWorkSheet.Cells[serialNo, 1] = serialNo;
                                        xlWorkSheet.Cells[serialNo, 2] = new DirectoryInfo(System.IO.Path.GetDirectoryName(file.FullName)).Name;
                                        xlWorkSheet.Cells[serialNo, 3] = System.IO.Path.GetFileName(file.FullName);
                                        xlWorkSheet.Cells[serialNo, 4] = file.FullName;
                                        xlWorkSheet.Cells[serialNo, 5] = lineCounter;
                                        xlWorkSheet.Cells[serialNo, 6] = keyword;
                                    }
                                }
                            }
                        }
                    }
                }
                xlWorkBook.SaveAs("d:\\KeywordReport_" + DateTime.Now.ToString("yyyyMMdd") + ".xls", Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
                xlWorkBook.Close(true, misValue, misValue);
                xlApp.Quit();

                Marshal.ReleaseComObject(xlWorkSheet);
                Marshal.ReleaseComObject(xlWorkBook);
                Marshal.ReleaseComObject(xlApp);
                MessageBox.Show("Excel Report Prepared !");
                return;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
                throw;
            } 
        }
        #endregion
        private void button2_Click(object sender, EventArgs e)
        {
            label4 = null;
            label5 = null;
            checkBox1.Checked = false;
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void button4_Click(object sender, EventArgs e)
        {
            OpenFileDialog openDialog = new OpenFileDialog();
            openDialog.Title = "Select A File";
            openDialog.Filter = "Text Files (*.txt)|*.txt" + "|" +
                                "Image Files (*.png;*.jpg)|*.png;*.jpg" + "|" +
                                "All Files (*.*)|*.*";
            if (openDialog.ShowDialog() == DialogResult.OK)
            {
                label5.Text = openDialog.FileName;
                strFilePath = openDialog.FileName;
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog folderDlg = new FolderBrowserDialog();
            folderDlg.ShowNewFolderButton = true;
            DialogResult result = folderDlg.ShowDialog();
            if (result == DialogResult.OK)
            {
                Environment.SpecialFolder root = folderDlg.RootFolder;
                strFolderPath =  folderDlg.SelectedPath;
                label4.Text = folderDlg.SelectedPath;
            }
        }
    }
}