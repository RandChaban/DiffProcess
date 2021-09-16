using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.IO;
using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using System.Diagnostics;
using OfficeOpenXml.Style;

namespace DiffProcess2015
{
    public partial class ProcessForm : Form
    {
        DataSet DataDS;
        public List<string> RejectedStdsList;
        bool editidCapacity;
        public ProcessForm()
        {
            RejectedStdsList = new List<string>();
            InitializeComponent();
        }
        private void exitBtn_Click(object sender, EventArgs e)
        {
            Application.Exit();
            DataDS = null;
            GC.GetTotalMemory(false);
            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
            GC.GetTotalMemory(true);
        }

        private void ProcessForm_Load(object sender, EventArgs e)
        {
        }

        private void startSciBtn_Click(object sender, EventArgs e)
        {
            DateTime start = DateTime.Now;
            DataDS = fillDS(1, "Output\\Scientific.xlsx", "SciSheet");
            SciStdsCountLbl.Text += DataDS.Tables["Stds_tbl"].Rows.Count.ToString() + "/" + DataDS.Tables["Facs_tbl"].Rows.Count.ToString() + "/" + DataDS.Tables["Facs_tbl"].AsEnumerable().Sum((DataRow x) => x.Field<int>("Capacity"));
            this.Refresh();
            int counter = 1;
            Sort();
            while (RejectedStdsList.Count != 0)
            {
                progresslbl.Text = "Loop " + counter.ToString() + " Clearing.." + RejectedStdsList.Count.ToString();
                this.Refresh();
                ClearRejected();
                Sort();
                counter++;
            }
            WriteOutput("Output/SResults.txt");
            WriteLimits("Output/SLimits.txt");
            writeToExcel("Output\\Scientific.xlsx", "SciSheet", DataDS.Tables["Facs_tbl"].Rows.Count + 1);
            if (SciLimitsChkBox.Checked)
                PrintLimits(1, "Output\\Scientific.xlsx", "Limits");
            progresslbl.Text = "Done..";
            this.Refresh();
            AcceptedSciLbl.Text += DataDS.Tables["Stds_tbl"].Select("RecentFaculty<>0").Count().ToString();
            RejectedSciLbl.Text += DataDS.Tables["Stds_tbl"].Select("RecentFaculty=0").Count().ToString();
            this.Refresh();

            DateTime end = DateTime.Now;
            TimeSpan span = end - start;
            exectimelbl.Text = span.Minutes.ToString() + ":" + span.Seconds.ToString();
        }

        private void startLitBtn_Click(object sender, EventArgs e)
        {
            //DateTime start = DateTime.Now;
            //DataDS = fillDS(2, "Output\\Literary.xlsx", "LitSheet");
            //litStdsCountLbl.Text += DataDS.Tables["Stds_tbl"].Rows.Count.ToString() + "/" + DataDS.Tables["Facs_tbl"].Rows.Count.ToString();
            //this.Refresh();
            //int counter = 1;
            //Sort();
            //while (RejectedStdsList.Count != 0)
            //{
            //    progresslbl.Text = "Loop " + counter.ToString() + " Clearing.." + RejectedStdsList.Count.ToString();
            //    this.Refresh();
            //    ClearRejected();
            //    Sort();
            //    counter++;
            //}
            //WriteOutput("Output/LResults.txt");
            //WriteLimits("Output/LLimits.txt");
            //writeToExcel("Output\\Literary.xlsx", "LitSheet", DataDS.Tables["Facs_tbl"].Rows.Count + 1);
            //if (LitLimitsCkhBox.Checked)
            //    PrintLimits(2, "Output\\Literary.xlsx", "Limits");
            //progresslbl.Text = "Done..";
            //this.Refresh();
            //AcceptedLitLbl.Text += DataDS.Tables["Stds_tbl"].Select("RecentFaculty<>0").Count().ToString();
            //RejectedLitLbl.Text += DataDS.Tables["Stds_tbl"].Select("RecentFaculty=0").Count().ToString();
            //this.Refresh();
            //DateTime end = DateTime.Now;
            //TimeSpan span = end - start;
            //exectimelbl.Text = span.Minutes.ToString() + ":" + span.Seconds.ToString();
        }

        private void writeStds(int counter)
        {
            StringBuilder stringBuilder = new StringBuilder();
            foreach (DataRow row in DataDS.Tables["Stds_tbl"].Rows)
                stringBuilder.AppendLine(counter + "#" + row["StudentID"]?.ToString() + "#" + row["recentFaculty"]);
            using (StreamWriter streamWriter = new StreamWriter("Output\\StdsRecentFaculty.txt", append: true))
                streamWriter.Write(stringBuilder.ToString());
        }


        private DataSet fillDS(byte branch,string ExcelFileName,string sheet)
        {
            try
            {
                progresslbl.Text = "Fetching Data..";
                this.Refresh();
                DataSet dataset = new DataSet();

                dataset.Tables.Add("Stds_tbl");
                dataset.Tables["Stds_tbl"].Columns.Add(new DataColumn("StudentID", typeof(int)));
                dataset.Tables["Stds_tbl"].Columns.Add(new DataColumn("DiffTotal", typeof(int)));
                dataset.Tables["Stds_tbl"].Columns.Add(new DataColumn("C1", typeof(int)));
                dataset.Tables["Stds_tbl"].Columns.Add(new DataColumn("C2", typeof(int)));
                dataset.Tables["Stds_tbl"].Columns.Add(new DataColumn("C3", typeof(int)));
                dataset.Tables["Stds_tbl"].Columns.Add(new DataColumn("C4", typeof(int)));
                dataset.Tables["Stds_tbl"].Columns.Add(new DataColumn("C5", typeof(int)));
                dataset.Tables["Stds_tbl"].Columns.Add(new DataColumn("C6", typeof(int)));
                dataset.Tables["Stds_tbl"].Columns.Add(new DataColumn("C7", typeof(int)));
                dataset.Tables["Stds_tbl"].Columns.Add(new DataColumn("C8", typeof(int)));
                dataset.Tables["Stds_tbl"].Columns.Add(new DataColumn("RecentFaculty", typeof(int)));
                dataset.Tables["Stds_tbl"].Columns.Add(new DataColumn("TotalWithReligion", typeof(int)));
                dataset.Tables["Stds_tbl"].Columns.Add(new DataColumn("TotalWithArts", typeof(int)));
                dataset.Tables["Stds_tbl"].Columns.Add(new DataColumn("Choices", typeof(string)));
                dataset.Tables["Stds_tbl"].Columns.Add(new DataColumn("UnivID", typeof(int)));
                dataset.Tables["Stds_tbl"].Columns.Add(new DataColumn("OldChoices", typeof(string)));

                if (File.Exists("Input/Stds.txt"))
                {
                    string[] lines = File.ReadAllLines("Input/Stds.txt");
                    foreach (string line in lines)
                    {
                        string[] values = line.Split('#');
                        string[] choices,Marks;
                        if (values[1] == branch.ToString())
                        {
                            DataRow row = dataset.Tables["Stds_tbl"].NewRow();
                            row["StudentID"] = Convert.ToInt32(values[0]);
                            row["DiffTotal"] = Convert.ToInt32(values[2]);
                            Marks = values[3].Split(',');
                            row["C1"] = Marks[0];
                            row["C2"] = Marks[1];
                            row["C3"] = Marks[2];
                            row["C4"] = Marks[3];
                            row["C5"] = Marks[4];
                            row["C6"] = Marks[5];
                            row["C7"] = Marks[6];
                            row["C8"] = Marks[7];
                            row["TotalWithReligion"] = values[4];
                            row["TotalWithArts"] = values[5];
                            row["Choices"]=values[6];
                            row["UnivID"] = Convert.ToInt32(values[7]);
                            row["OldChoices"] = values[8];
                            choices = values[6].Split(',');
                            row["RecentFaculty"] = choices[0];
                            dataset.Tables["Stds_tbl"].Rows.Add(row);
                        }
                    }
                }
                else
                    MessageBox.Show("Stds.txt File not found");
                
                dataset.Tables.Add("Facs_tbl");
                dataset.Tables["Facs_tbl"].Columns.Add(new DataColumn("FacultyID", typeof(int)));
                dataset.Tables["Facs_tbl"].Columns.Add(new DataColumn("Capacity", typeof(int)));
                dataset.Tables["Facs_tbl"].Columns.Add(new DataColumn("Priorities", typeof(string)));
                dataset.Tables["Facs_tbl"].Columns.Add(new DataColumn("CompGroup", typeof(string)));
                dataset.Tables["Facs_tbl"].Columns.Add(new DataColumn("RejectedMark", typeof(string)));
                dataset.Tables["Facs_tbl"].Columns.Add(new DataColumn("Limit", typeof(string)));
                dataset.Tables["Facs_tbl"].Columns.Add(new DataColumn("FacultyName", typeof(string)));
                dataset.Tables["Facs_tbl"].Columns.Add(new DataColumn("CityName", typeof(string)));
                dataset.Tables["Facs_tbl"].Columns.Add(new DataColumn("InitialLimit", typeof(string)));
                dataset.Tables["Facs_tbl"].Columns.Add(new DataColumn("Follow", typeof(int)));
                dataset.Tables["Facs_tbl"].Columns.Add(new DataColumn("UnivID", typeof(int)));
                if (File.Exists("Input/Facs.txt"))
                {
                    string[] lines = File.ReadAllLines("Input/Facs.txt");
                    int LinesCount = lines.Length;
                    int i = 2;
                    FileInfo existingFile = new FileInfo(Path.GetFullPath(ExcelFileName));
                    using (ExcelPackage package = new ExcelPackage(existingFile))
                    {
                        ExcelWorkbook book = package.Workbook;
                        ExcelWorksheet FirstSheet = book.Worksheets[sheet];
                        if (FirstSheet.Cells["G2"].Value == null)
                        {
                            editidCapacity = false;
                        }
                        else
                        {
                            editidCapacity = true;
                        }

                        foreach (string line in lines)
                        {
                            string[] values = line.Split('#');
                            if (values[1] == branch.ToString())
                            {
                                DataRow row = dataset.Tables["Facs_tbl"].NewRow();
                                row["FacultyID"] = Convert.ToInt32(values[0]);
                                if (!editidCapacity)
                                {
                                    row["Capacity"] = Convert.ToInt32(values[2]);
                                }
                                else
                                {
                                    for (int id = 2; id <= LinesCount + 1; id++)
                                    {
                                        if (FirstSheet.Cells["A" + id.ToString()].Value.ToString() == values[0])
                                        {
                                            row["Capacity"] = Convert.ToInt32(FirstSheet.Cells["G" + id.ToString()].Value);
                                            break;
                                        }
                                    }
                                }
                                i++;
                                row["InitialLimit"] = values[5];
                                row["Priorities"] = values[3];
                                row["CompGroup"] = values[4];
                                row["Follow"] = Convert.ToInt32(values[6]);
                                for (int id = 2; id <= LinesCount+1; id++)
                                {
                                    if (FirstSheet.Cells["A" + id.ToString()].Value.ToString() == values[0])
                                    {
                                        row["FacultyName"] = FirstSheet.Cells["B" + id.ToString()].Value;
                                        row["CityName"] = FirstSheet.Cells["C" + id.ToString()].Value;
                                        break;
                                    }
                                }
                                
                                row["RejectedMark"] = "0,0";
                                row["Limit"] = "0,"+values[5];
                                row["UnivID"] = Convert.ToInt32(values[7]);
                                dataset.Tables["Facs_tbl"].Rows.Add(row);
                            }
                        }
                        package.Save();
                    }
                    
                }
                else
                    MessageBox.Show("Facs.txt File not found");
                return dataset;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
                return null;
            }
        }

        private Int64 ConcateDiffTotal(string facultyID, string studentID)
        {
            string facultyUnivID = DataDS.Tables["Facs_tbl"].Select("FacultyID =" + facultyID).First()["UnivID"].ToString();
            string stdUnivID = DataDS.Tables["Stds_tbl"].Select("StudentID =" + studentID).First()["UnivID"].ToString();
            try
            {
                string[] Priorities = DataDS.Tables["Facs_tbl"].Select("FacultyID =" + facultyID).First()["Priorities"].ToString().Split(',');
                string diffTotal;
                switch (DataDS.Tables["Facs_tbl"].Select("FacultyID =" + facultyID).First()["CompGroup"].ToString())
                {
                    case "-1":
                        diffTotal = DataDS.Tables["Stds_tbl"].Select("StudentID =" + studentID).First()["TotalWithReligion"].ToString() + DataDS.Tables["Stds_tbl"].Select("StudentID =" + studentID).First()["C8"].ToString().PadLeft(3, '0') + DataDS.Tables["Stds_tbl"].Select("StudentID =" + studentID).First()["C6"].ToString().PadLeft(3, '0') + "000";
                        break;
                    case "-2":
                        diffTotal = DataDS.Tables["Stds_tbl"].Select("StudentID =" + studentID).First()["TotalWithReligion"].ToString() + "000000000";
                        break;
                    case "1":
                        diffTotal = DataDS.Tables["Stds_tbl"].Select("StudentID =" + studentID).First()["TotalWithArts"].ToString() + "000000000";
                        break;
                    default:
                        {
                            if (Priorities[0] == "0")
                            {
                                int index = Array.IndexOf(Priorities, "1");
                                diffTotal = DataDS.Tables["Stds_tbl"].Select("StudentID =" + studentID).First()["C" + index].ToString() + "00" + DataDS.Tables["Stds_tbl"].Select("StudentID =" + studentID).First()["DiffTotal"].ToString().PadLeft(4, '0');
                                break;
                            }
                            diffTotal = DataDS.Tables["Stds_tbl"].Select("StudentID =" + studentID).First()["DiffTotal"].ToString();
                            for (int i = 1; i <= 3; i++)
                            {
                                int index = Array.IndexOf(Priorities, i.ToString(), 1);
                                diffTotal = ((index == -1) ? (diffTotal + "000") : (diffTotal + DataDS.Tables["Stds_tbl"].Select("StudentID =" + studentID).First()["C" + index].ToString().PadLeft(3, '0')));
                            }
                            break;
                        }
                }
                diffTotal = ((!(facultyUnivID != "0")) ? ("0" + diffTotal) : ((facultyUnivID == stdUnivID) ? ("2" + diffTotal) : ((Convert.ToInt32(stdUnivID) >= 8) ? ("0" + diffTotal) : ("1" + diffTotal))));
                return Convert.ToInt64(diffTotal);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
                return 0L;
            }
        }

        private void WriteOutput(string filename)
        {
            if (File.Exists(filename))
                File.Delete(filename);
            StringBuilder ResultsSB = new StringBuilder();
            foreach (DataRow row in DataDS.Tables["Stds_tbl"].Rows)
            {
                if (row["RecentFaculty"].ToString() == "0")
                {
                    ResultsSB.AppendLine(row["StudentID"].ToString() + "#0#" + row["RecentFaculty"].ToString() + "#0");
                    continue;
                }
                int index = Array.IndexOf(row["OldChoices"].ToString().Split(','), Convert.ToInt32(DataDS.Tables["Facs_tbl"].Select("FacultyID=" + row["RecentFaculty"].ToString()).First()["Follow"]).ToString()) + 1;
                ResultsSB.AppendLine(row["StudentID"].ToString() + "#" + ConcateDiffTotal(row["RecentFaculty"].ToString(), row["StudentID"].ToString()) + "#" + row["RecentFaculty"].ToString() + "#" + index);
            }
            using (StreamWriter streamWriter = new StreamWriter(filename, append: true))
                streamWriter.Write(ResultsSB.ToString());
        }

        private void WriteLimits(string filename)
        {
            if (File.Exists(filename))
                File.Delete(filename);
            StringBuilder ResultsSB = new StringBuilder();
            foreach (DataRow row in DataDS.Tables["Facs_tbl"].Rows)
                ResultsSB.AppendLine(row["FacultyID"].ToString() + "#" + row["Limit"].ToString().Split(',')[1]);
            using (StreamWriter outfile = new StreamWriter(filename, true))
            {
                outfile.Write(ResultsSB.ToString());
            }
        }

        private void Sort()
        {
            try
            {
                List<string> WaitingList = new List<string>();
                List<string> RejectedList = new List<string>();
                int capacity, addedStds, rejectedAtEqual;
                string facultyID, sort;
                Int64 rejectedMark,Limit;
                foreach (DataRow row in DataDS.Tables["Facs_tbl"].Rows)
                {
                    capacity = Convert.ToInt32(row["Capacity"]);
                    rejectedMark = Convert.ToInt64(row["RejectedMark"].ToString().Split(',')[0]);
                    row["Limit"] = "0," + row["InitialLimit"].ToString();
                    facultyID = row["facultyID"].ToString();
                    addedStds = 1;
                    rejectedAtEqual = 0;
                    WaitingList.Clear();
                    RejectedList.Clear();
                    DataTable dataTable = new DataTable();
                    dataTable.Columns.Add("StudentID");
                    dataTable.Columns.Add("DiffTotal", typeof(Int64));

                    foreach (DataRow dataRow2 in DataDS.Tables["Stds_tbl"].Select("RecentFaculty=" + facultyID))
                    {
                        DataRow dataRow3 = dataTable.NewRow();
                        dataRow3["StudentID"] = dataRow2["StudentID"];
                        dataRow3["DiffTotal"] = Convert.ToInt64(ConcateDiffTotal(facultyID, dataRow2["StudentID"].ToString()));
                        dataTable.Rows.Add(dataRow3);
                    }

                    foreach (DataRow stdrow in dataTable.Select("", "DiffTotal Desc"))
                    {
                        if (addedStds <= capacity && ConcateDiffTotal(facultyID, stdrow["StudentID"].ToString()) > rejectedMark)
                        {
                            WaitingList.Add(stdrow["StudentID"].ToString());
                            addedStds++;
                        }
                        else
                        {
                            RejectedList.Add(stdrow["StudentID"].ToString());
                        }
                    }
                    dataTable.Rows.Clear();
                    if (WaitingList.Count != 0)
                    {
                        if (RejectedList.Count != 0)
                        {
                            rejectedAtEqual++;
                            for (int k = 0; k < RejectedList.Count - 1 && ConcateDiffTotal(facultyID, RejectedList[k]) == ConcateDiffTotal(facultyID, RejectedList[k + 1]); k++)
                            {
                                rejectedAtEqual++;
                            }
                            while (ConcateDiffTotal(facultyID, WaitingList.Last()) == ConcateDiffTotal(facultyID, RejectedList.First()))
                            {
                                rejectedAtEqual++;
                                RejectedList.Insert(0, WaitingList.Last());
                                WaitingList.RemoveAt(WaitingList.Count - 1);
                            }
                            if (rejectedMark < ConcateDiffTotal(facultyID, RejectedList.First().ToString()))
                            {
                                row["RejectedMark"] = ConcateDiffTotal(facultyID, RejectedList.First()) + "," + rejectedAtEqual;
                            }
                        }
                        row["Limit"] = WaitingList.Last() + "," + ConcateDiffTotal(facultyID, WaitingList.Last());
                        RejectedStdsList.AddRange(RejectedList);
                    }
                    else if (RejectedList.Count != 0)
                    {
                        if (rejectedMark < ConcateDiffTotal(facultyID, RejectedList.First().ToString()))
                        {
                            row["RejectedMark"] = ConcateDiffTotal(facultyID, RejectedList.First()) + "," + rejectedAtEqual;
                        }
                        RejectedStdsList.AddRange(RejectedList);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }
        private void ClearRejected()
        {
            try
            {
                if (RejectedStdsList.Count != 0)
                {
                    string student;
                    List<string> ChoicesList;
                    foreach (DataRow stdRow in DataDS.Tables["Stds_tbl"].Rows)
                    {
                        student = stdRow["StudentID"].ToString();
                        if (RejectedStdsList.IndexOf(student) != -1)
                        {
                            ChoicesList = new List<string>(stdRow["Choices"].ToString().Split(','));
                            stdRow["RecentFaculty"] = ChoicesList[ChoicesList.IndexOf(stdRow["RecentFaculty"].ToString()) + 1];
                            RejectedStdsList.Remove(student);
                        }
                    }

                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }
        private void writeToExcel(string filename,string sheet,int FacultiesCount)
        {
            try
            {
                progresslbl.Text = "Writing to Excel..";
                Refresh();
                if (!File.Exists(filename))
                {
                    MessageBox.Show(filename + " not found");
                    return;
                }
                string facultyID;
                int AcceptedAtLimit, TotalAccepted;
                string[] limitArray, RejectedMarkArray, prioritiesArray;
                DataRow facultyRow, StudentRow;
                FileInfo existingFile = new FileInfo(Path.GetFullPath(filename));
                using (ExcelPackage excelPackage = new ExcelPackage(existingFile))
                {
                    ExcelWorkbook workbook = excelPackage.Workbook;
                    ExcelWorksheet PublicSheet = workbook.Worksheets[sheet];
                    Font Font18 = new Font("Areal", 18);
                    Font Font16 = new Font("Areal", 16);
                    Font Font14 = new Font("Areal", 14);
                    Font Font12 = new Font("Areal", 12);
                    bool editedCapacity = true;
                    if (PublicSheet.Cells["G2"].Value == null)
                    {
                        editedCapacity = false;
                    }
                    for (int i = 2; i <= FacultiesCount; i++)
                    {
                        facultyID = PublicSheet.Cells["A" + i].Value.ToString();
                        facultyRow = DataDS.Tables["Facs_tbl"].Select("FacultyID=" + facultyID).First();
                        limitArray = facultyRow["Limit"].ToString().Split(',');
                        RejectedMarkArray = facultyRow["RejectedMark"].ToString().Split(',');
                        prioritiesArray = facultyRow["Priorities"].ToString().Split(',');
                        AcceptedAtLimit = 0;
                        TotalAccepted = 0;
                        if (!editedCapacity)
                        {
                            PublicSheet.Cells["F" + i].Value = Convert.ToInt32(facultyRow["Capacity"]);
                            PublicSheet.Cells["G" + i].Value = Convert.ToInt32(facultyRow["Capacity"]);
                        }
                        foreach (DataRow stdRow in DataDS.Tables["Stds_tbl"].Select("RecentFaculty=" + facultyID))
                        {
                            if (ConcateDiffTotal(facultyID, stdRow["StudentID"].ToString()) == Convert.ToInt64(limitArray[1]))
                            {
                                AcceptedAtLimit++;
                                PublicSheet.Cells["Z" + i].Value = stdRow["StudentID"].ToString();
                            }
                            TotalAccepted++;
                        }
                        PublicSheet.Cells["E" + i].Value = Convert.ToInt64(limitArray[1]);
                        PublicSheet.Cells["H" + i].Value = Convert.ToInt32(TotalAccepted);
                        PublicSheet.Cells["I" + i].Formula = "=G" + i + "-H" + i;
                        PublicSheet.Cells["J" + i].Value = Convert.ToInt32(RejectedMarkArray[1]);
                        PublicSheet.Cells["K" + i].Value = Convert.ToInt64(RejectedMarkArray[0]);
                        PublicSheet.Cells["L" + i].Value = Convert.ToInt32(AcceptedAtLimit);
                        PublicSheet.Cells["M" + i].Value = Convert.ToInt32(facultyRow["InitialLimit"]);
                        if (TotalAccepted != 0)
                        {
                            StudentRow = DataDS.Tables["Stds_tbl"].Select("StudentID=" + limitArray[0]).First();
                            if (prioritiesArray[0] == "0")
                            {
                                for (int m = 1; m <= 8; m++)
                                {
                                    if (m == Array.IndexOf(prioritiesArray, "1"))
                                    {
                                        PublicSheet.Cells["D" + i].Value = Convert.ToInt32(StudentRow["C" + m].ToString());
                                        PublicSheet.Cells[i, m + 13].Value = Convert.ToInt32(StudentRow["C" + m].ToString());
                                        PublicSheet.Cells[i, m + 13].Style.Font.Color.SetColor(Color.Red);
                                        PublicSheet.Cells[i, m + 13].Style.Font.SetFromFont(Font18);
                                    }
                                    else
                                    {
                                        PublicSheet.Cells[i, m + 13].Value = Convert.ToInt32(StudentRow["C" + m].ToString());
                                        PublicSheet.Cells[i, m + 13].Style.Font.Color.SetColor(Color.Black);
                                        PublicSheet.Cells[i, m + 13].Style.Font.SetFromFont(Font12);
                                    }
                                }
                                PublicSheet.Cells["V" + i].Value = Convert.ToInt32(StudentRow["DiffTotal"].ToString());
                                PublicSheet.Cells["V" + i].Style.Font.SetFromFont(Font14);
                                PublicSheet.Cells["V" + i].Style.Font.Color.SetColor(Color.Blue);
                                continue;
                            }
                            switch (facultyRow["CompGroup"].ToString())
                            {
                                default:
                                case "0":
                                    PublicSheet.Cells["D" + i].Value = Convert.ToInt32(StudentRow["DiffTotal"].ToString());
                                    break;
                                case "1":
                                    PublicSheet.Cells["D" + i].Value = Convert.ToInt32(StudentRow["TotalWithArts"].ToString());
                                    break;
                                case "-1":
                                case "-2":
                                    PublicSheet.Cells["D" + i].Value = Convert.ToInt32(StudentRow["TotalWithReligion"].ToString());
                                    break;
                            }
                            for (int j = 1; j <= 8; j++)
                            {
                                PublicSheet.Cells[i, j + 13].Value = Convert.ToInt32(StudentRow["C" + j].ToString());
                                if (prioritiesArray[j] == "1")
                                {
                                    PublicSheet.Cells[i, j + 13].Style.Font.SetFromFont(Font18);
                                    PublicSheet.Cells[i, j + 13].Style.Font.Color.SetColor(Color.Red);
                                }
                                else if (prioritiesArray[j] == "2")
                                {
                                    PublicSheet.Cells[i, j + 13].Style.Font.SetFromFont(Font16);
                                    PublicSheet.Cells[i, j + 13].Style.Font.Color.SetColor(Color.DeepPink);
                                }
                                else if (prioritiesArray[j] == "3")
                                {
                                    PublicSheet.Cells[i, j + 13].Style.Font.SetFromFont(Font14);
                                    PublicSheet.Cells[i, j + 13].Style.Font.Color.SetColor(Color.Blue);
                                }
                                else if (prioritiesArray[j] == "0")
                                {
                                    PublicSheet.Cells[i, j + 13].Style.Font.SetFromFont(Font12);
                                    PublicSheet.Cells[i, j + 13].Style.Font.Color.SetColor(Color.Black);
                                }
                                PublicSheet.Cells["V" + i].Value = Convert.ToInt32(StudentRow["DiffTotal"].ToString());
                                PublicSheet.Cells["V" + i].Style.Font.SetFromFont(Font14);
                            }
                        }
                        else
                        {
                            PublicSheet.Cells["D" + i].Value = Convert.ToInt32(facultyRow["InitialLimit"].ToString());
                            PublicSheet.Cells["D" + i].Style.Font.SetFromFont(Font12);
                            PublicSheet.Cells["D" + i].Style.Font.Color.SetColor(Color.Black);
                            PublicSheet.Cells["N" + i].Value = 0;
                            PublicSheet.Cells["N" + i].Style.Font.SetFromFont(Font12);
                            PublicSheet.Cells["N" + i].Style.Font.Color.SetColor(Color.Black);
                            PublicSheet.Cells["O" + i].Value = 0;
                            PublicSheet.Cells["O" + i].Style.Font.SetFromFont(Font12);
                            PublicSheet.Cells["O" + i].Style.Font.Color.SetColor(Color.Black);
                            PublicSheet.Cells["P" + i].Value = 0;
                            PublicSheet.Cells["P" + i].Style.Font.SetFromFont(Font12);
                            PublicSheet.Cells["P" + i].Style.Font.Color.SetColor(Color.Black);
                            PublicSheet.Cells["Q" + i].Value = 0;
                            PublicSheet.Cells["Q" + i].Style.Font.SetFromFont(Font12);
                            PublicSheet.Cells["Q" + i].Style.Font.Color.SetColor(Color.Black);
                            PublicSheet.Cells["R" + i].Value = 0;
                            PublicSheet.Cells["R" + i].Style.Font.SetFromFont(Font12);
                            PublicSheet.Cells["R" + i].Style.Font.Color.SetColor(Color.Black);
                            PublicSheet.Cells["S" + i].Value = 0;
                            PublicSheet.Cells["S" + i].Style.Font.SetFromFont(Font12);
                            PublicSheet.Cells["S" + i].Style.Font.Color.SetColor(Color.Black);
                            PublicSheet.Cells["T" + i].Value = 0;
                            PublicSheet.Cells["T" + i].Style.Font.SetFromFont(Font12);
                            PublicSheet.Cells["T" + i].Style.Font.Color.SetColor(Color.Black);
                            PublicSheet.Cells["U" + i].Value = 0;
                            PublicSheet.Cells["U" + i].Style.Font.SetFromFont(Font12);
                            PublicSheet.Cells["U" + i].Style.Font.Color.SetColor(Color.Black);
                            PublicSheet.Cells["V" + i].Value = 0;
                            PublicSheet.Cells["V" + i].Style.Font.SetFromFont(Font12);
                            PublicSheet.Cells["V" + i].Style.Font.Color.SetColor(Color.Black);
                        }
                    }
                    PublicSheet.Cells["F" + (FacultiesCount + 1)].Formula = "=SUM(F2:F" + FacultiesCount + ")";
                    PublicSheet.Cells["G" + (FacultiesCount + 1)].Formula = "=SUM(G2:G" + FacultiesCount + ")";
                    PublicSheet.Cells["H" + (FacultiesCount + 1)].Formula = "=SUM(H2:H" + FacultiesCount + ")";
                    PublicSheet.Cells["I" + (FacultiesCount + 1)].Formula = "=SUM(I2:I" + FacultiesCount + ")";
                    excelPackage.Save();
                }
                GC.GetTotalMemory(forceFullCollection: false);
                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();
                GC.GetTotalMemory(forceFullCollection: true);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        private void PrintLimits(int Branch,string ExcelFileName, string Sheet)
        {
            string[] Subjects = ((Branch != 1) ? new string[9] { "المجموع العام", "التاريخ", "الجغرافيا", "الفلسفة", "", "اللغة الأجنبية", "اللغة العربية", "القومية", "الديانة" } : new string[9] { "المجموع العام", "الرياضيات", "الفيزياء", "الكيمياء", "العلوم", "اللغة الأجنبية", "اللغة العربية", "القومية", "الديانة" });
            List<string> PrioritiesList = new List<string>();
            int index, index1, index2, index3;
            int i = 6;
            DataTable DistinctPriorities = new DataTable();
            DistinctPriorities = DataDS.Tables["Facs_tbl"].Select("", "Follow ASC").CopyToDataTable().DefaultView.ToTable(true, "Priorities");
            Font Font18 = new Font("Simplified Arabic", 14f);
            int RowsCount = DistinctPriorities.Rows.Count + DataDS.Tables["Facs_tbl"].Rows.Count;
            FileInfo existingFile = new FileInfo(Path.GetFullPath(ExcelFileName));
            try
            {
                using (ExcelPackage excelPackage = new ExcelPackage(existingFile))
                {
                    ExcelWorkbook book = excelPackage.Workbook;
                    ExcelWorksheet SecondSheet = book.Worksheets[Sheet];
                    ExcelRange range = SecondSheet.Cells[6, 1, RowsCount + 5, 6];
                    range.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                    range.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    range.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                    range.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                    range.Style.Border.Top.Color.SetColor(Color.Black);
                    range.Style.Border.Bottom.Color.SetColor(Color.Black);
                    range.Style.Border.Left.Color.SetColor(Color.Black);
                    range.Style.Border.Right.Color.SetColor(Color.Black);
                    range.Style.Font.SetFromFont(Font18);
                    foreach (DataRow priority in DistinctPriorities.Rows)
                    {
                        SecondSheet.Cells["j" + i].Value = priority["Priorities"].ToString();
                        PrioritiesList = priority["Priorities"].ToString().Split(',').ToList();
                        DataRow[] facs = DataDS.Tables["Facs_tbl"].Select("Priorities='" + priority["Priorities"].ToString() + "'", "Follow ASC,FacultyID DESC");
                        if (PrioritiesList.First() == "0")
                        {
                            index = PrioritiesList.IndexOf("1");
                            SecondSheet.Cells["C" + i].Value = Subjects[index];
                            SecondSheet.Cells["D" + i].Value = Subjects[0];
                            i++;
                            foreach (DataRow row in facs)
                            {
                                SecondSheet.Cells["K" + i].Value = row["FacultyID"].ToString();
                                SecondSheet.Cells["A" + i].Value = row["FacultyName"].ToString();
                                SecondSheet.Cells["B" + i].Value = row["CityName"].ToString();
                                SecondSheet.Cells["C" + i].Value = Convert.ToInt32(row["Limit"].ToString().Split(',')[1].Substring(0, 3));
                                if (row["Limit"].ToString().Split(',')[1].Length > 4)
                                    SecondSheet.Cells["D" + i].Value = Convert.ToInt32(row["Limit"].ToString().Split(',')[1].Substring(5, 4));
                                else
                                    SecondSheet.Cells["D" + i].Value = 870;
                                i++;
                            }
                            continue;
                        }
                        SecondSheet.Cells["C" + i].Value = Subjects[0];
                        index1 = PrioritiesList.IndexOf("1", 1);
                        index2 = PrioritiesList.IndexOf("2", 1);
                        index3 = PrioritiesList.IndexOf("3", 1);
                        if (index1 != -1)
                            SecondSheet.Cells["D" + i].Value = Subjects[index1];
                        if (index2 != -1)
                            SecondSheet.Cells["E" + i].Value = Subjects[index2];
                        if (index3 != -1)
                            SecondSheet.Cells["F" + i].Value = Subjects[index3];
                        i++;
                        foreach (DataRow dataRow3 in facs)
                        {
                            SecondSheet.Cells["K" + i].Value = dataRow3["FacultyID"].ToString();
                            SecondSheet.Cells["A" + i].Value = dataRow3["FacultyName"].ToString();
                            SecondSheet.Cells["B" + i].Value = dataRow3["CityName"].ToString();
                            if (dataRow3["Limit"].ToString().Split(',')[1].Length > 4)
                            {
                                SecondSheet.Cells["C" + i].Value = Convert.ToInt32(dataRow3["Limit"].ToString().Split(',')[1].Substring(0, 4));
                                if (index1 != -1)
                                    SecondSheet.Cells["D" + i].Value = Convert.ToInt32(dataRow3["Limit"].ToString().Split(',')[1].Substring(4, 3));
                                if (index2 != -1)
                                    SecondSheet.Cells["E" + i].Value = Convert.ToInt32(dataRow3["Limit"].ToString().Split(',')[1].Substring(7, 3));
                                if (index3 != -1)
                                    SecondSheet.Cells["F" + i].Value = Convert.ToInt32(dataRow3["Limit"].ToString().Split(',')[1].Substring(10, 3));
                            }
                            else
                            {
                                SecondSheet.Cells["C" + i].Value = Convert.ToInt32(dataRow3["Limit"].ToString().Split(',')[1]);
                                SecondSheet.Cells["D" + i].Value = 0;
                                SecondSheet.Cells["E" + i].Value = 0;
                                SecondSheet.Cells["F" + i].Value = 0;
                            }
                            i++;
                        }
                    }
                    excelPackage.Save();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }
    }
}
