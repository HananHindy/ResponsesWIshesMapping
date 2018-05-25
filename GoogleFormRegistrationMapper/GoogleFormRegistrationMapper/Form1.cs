using ExcelDataReader;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace GoogleFormRegistrationMapper
{
    public partial class Form1 : Form
    {
        public List<string> ColumnNames { get; set; }

        private string _readFilePath;
        private DataSet _allFileData;

        private int _idColumnIndex, _choice1ColumnIndex, _choice2ColumnIndex, _choice3ColumnIndex;

        public Form1()
        {
            InitializeComponent();
        }

        private void btnBrowse_Click(object sender, EventArgs e)
        {
            lblResult.Visible = pnl_choices.Visible  = btnStart.Enabled = false;
            OpenFileDialog fileDialog = new OpenFileDialog();
            fileDialog.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm";
            if (fileDialog.ShowDialog() == DialogResult.OK)
            {
                _readFilePath = fileDialog.FileName;
                txtFilePath.Text = _readFilePath;
                btnStart.Enabled = true;
            }
        }


        private void btnStart_Click(object sender, EventArgs e)
        {
            FileInfo file = new FileInfo(_readFilePath);
            FileStream stream = File.Open(_readFilePath, FileMode.Open, FileAccess.Read);

            IExcelDataReader excelReader;
            if (file.Extension.Equals(".xls"))
            {
                excelReader = ExcelReaderFactory.CreateBinaryReader(stream);
            }
            else
            {
                excelReader = ExcelReaderFactory.CreateOpenXmlReader(stream);
            }

            var conf = new ExcelDataSetConfiguration
            {
                ConfigureDataTable = _ => new ExcelDataTableConfiguration
                {
                    UseHeaderRow = true
                }
            };

            _allFileData = excelReader.AsDataSet(conf);
            ColumnNames = _allFileData.Tables[0].Columns.Cast<DataColumn>()
                                    .Select(x => x.ColumnName).ToList();
            _idColumnIndex = 0;
            _choice1ColumnIndex = ColumnNames.Count - 3;
            _choice2ColumnIndex = ColumnNames.Count - 2;
            _choice3ColumnIndex = ColumnNames.Count - 1;


            cmb_id.DataSource = new List<string>(ColumnNames);
            cmb_id.SelectedIndex = _idColumnIndex;

            cmb_choice1.DataSource = new List<string>(ColumnNames);
            cmb_choice1.SelectedIndex = _choice1ColumnIndex;

            cmb_choice2.DataSource = new List<string>(ColumnNames);
            cmb_choice2.SelectedIndex = _choice2ColumnIndex;

            cmb_choice3.DataSource = new List<string>(ColumnNames);
            cmb_choice3.SelectedIndex = _choice3ColumnIndex;

            pnl_choices.Visible = true;

            excelReader.Close();
            stream.Close();
        }

        private void btn_generate_Click(object sender, EventArgs e)
        {
            Dictionary<string, int> specialTrackCount = new Dictionary<string, int>();
            if (dgv_specialCount.RowCount > 0)
            {
                foreach (DataGridViewRow row in dgv_specialCount.Rows)
                {
                    if (row.Cells[0].Value != null)
                    {
                        specialTrackCount.Add(row.Cells[0].Value.ToString(), int.Parse(row.Cells[1].Value.ToString()));
                    }
                }
            }

            int max = int.Parse(txt_max.Text);
            _choice1ColumnIndex = cmb_choice1.SelectedIndex;
            _choice2ColumnIndex = cmb_choice2.SelectedIndex;
            _choice3ColumnIndex = cmb_choice3.SelectedIndex;


            var rows = _allFileData.Tables[0].Rows.Cast<DataRow>();

            Dictionary<string, int> finishedStudents = new Dictionary<string, int>();

            Dictionary<string, List<List<string>>> tracksData = new Dictionary<string, List<List<string>>>();
            foreach (var r in rows)
            {
                int trackMax = max;
                int fullfilledTrack = -1;

                string currentId = r[_idColumnIndex].ToString().ToLower();
                if (finishedStudents.ContainsKey(currentId))
                {
                    continue;
                }

                bool finished = false;
                string choice1 = r[_choice1ColumnIndex].ToString();
                string choice2 = r[_choice2ColumnIndex].ToString();
                string choice3 = r[_choice3ColumnIndex].ToString();

                if (tracksData.ContainsKey(choice1) == false)
                {
                    tracksData.Add(choice1, new List<List<string>>());
                }

                if (specialTrackCount.ContainsKey(choice1))
                {
                    trackMax = specialTrackCount[choice1];
                }

                if (tracksData[choice1].Count < trackMax)
                {
                    tracksData[choice1].Add(r.ItemArray.Select(x => x.ToString()).ToList());
                    finished = true;
                    fullfilledTrack = 1;
                }

                if (finished == false)
                {
                    trackMax = max;
                    if (tracksData.ContainsKey(choice2) == false)
                    {
                        tracksData.Add(choice2, new List<List<string>>());
                    }

                    if (specialTrackCount.ContainsKey(choice2))
                    {
                        trackMax = specialTrackCount[choice2];
                    }

                    if (tracksData[choice2].Count < trackMax)
                    {
                        tracksData[choice2].Add(r.ItemArray.Select(x => x.ToString()).ToList());
                        finished = true;
                        fullfilledTrack = 2;
                    }

                    if (finished == false)
                    {
                        trackMax = max;
                        if (tracksData.ContainsKey(choice3) == false)
                        {
                            tracksData.Add(choice3, new List<List<string>>());
                        }

                        if (specialTrackCount.ContainsKey(choice3))
                        {
                            trackMax = specialTrackCount[choice3];
                        }

                        if (tracksData[choice3].Count < trackMax)
                        {
                            tracksData[choice3].Add(r.ItemArray.Select(x => x.ToString()).ToList());
                            finished = true;
                            fullfilledTrack = 3;
                        }
                    }
                }

                if (finished == false)
                {
                    if (tracksData.ContainsKey("Not Assigned") == false)
                    {
                        tracksData.Add("Not Assigned", new List<List<string>>());
                    }
                    tracksData["Not Assigned"].Add(r.ItemArray.Select(x => x.ToString()).ToList());
                }
                finishedStudents.Add(currentId, fullfilledTrack);
            }

            SaveExcel(tracksData, finishedStudents);
        }

        private void SaveExcel(Dictionary<string, List<List<string>>> tracks, Dictionary<string, int> allStudentsData)
        { 
            SaveFileDialog fileDialog = new SaveFileDialog();
            fileDialog.Filter = "Excel File|*.xlsx";
            if (fileDialog.ShowDialog() == DialogResult.OK)
            {
                using (ExcelPackage excel = new ExcelPackage())
                {
                    foreach (var tr in tracks)
                    {
                        var worksheet = excel.Workbook.Worksheets.Add(tr.Key);

                        string headerRange = "A1:" + Char.ConvertFromUtf32(ColumnNames.Count + 64) + "1";
                        worksheet.Cells[headerRange].LoadFromArrays(new List<string[]>() { ColumnNames.ToArray() });
                        int counter = 2;
                        foreach (var row in tr.Value)
                        {
                            worksheet.Cells[headerRange.Replace("1", counter.ToString())].LoadFromArrays(new List<string[]>() { row.ToArray() });
                            counter++;
                        }
                    }

                    // All Students
                    var worksheet2 = excel.Workbook.Worksheets.Add("ALL STUDENTS");

                    string headerRange2 = "A1:B1";
                    worksheet2.Cells[headerRange2].LoadFromArrays(new List<string[]>() { new[] { "Student ID", "Track Number"} });
                    int counter2 = 2;
                    foreach (var id in allStudentsData.Keys)
                    {
                        worksheet2.Cells[headerRange2.Replace("1", counter2.ToString())].LoadFromArrays(new List<string[]>() { new[] { id, allStudentsData[id].ToString() } });
                        counter2++;
                    }


                    FileInfo excelFile = new FileInfo(fileDialog.FileName);
                    excel.SaveAs(excelFile);
                    lblResult.Text = "File saved with a sheet for each Track";
                    lblResult.Visible = true;
                }
            }
        }
    }
}
