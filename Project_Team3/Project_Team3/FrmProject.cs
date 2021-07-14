using Project_Team4.DAO;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace Project_Team4
{
    public partial class FrmProject : Form
    {
        public FrmProject()
        {
            InitializeComponent();
        }

        public static bool checkData(DataTable tbl1, DataTable tbl2)
        {
            int r1 = tbl1.Rows.Count;
            int c1 = tbl1.Columns.Count;
            int r2 = tbl2.Rows.Count;
            int c2 = tbl2.Columns.Count;

            if (r1 != r2 || c1 != c2)
                return false;
            int count = 0;

            for (int i = 0; i < r1; i++) {
                for (int j = 0; j < r1; j++) {
                    int countt = count;
                    for (int c = 0; c < c1; c++) {
                        if (Equals(tbl1.Rows[i][c], tbl2.Rows[j][c])) {
                            count++;
                        }
                    }
                    if ((count - countt) == c1) break;
                    else count = countt;
                }
            }
            if ((count / c1) == r1)
                return true;
            else
                return false;

        }
        public static bool checkSort(DataTable tbl1, DataTable tbl2)
        {
            int r1 = tbl1.Rows.Count;
            int c1 = tbl1.Columns.Count;
            int r2 = tbl2.Rows.Count;
            int c2 = tbl2.Columns.Count;

            for (int i = 0; i < r1; i++) {
                for (int c = 0; c < c1; c++) {
                    if (!Equals(tbl1.Rows[i][c], tbl2.Rows[i][c]))
                        return false;
                }
            }
            return true;
        }

        public string getFolderPath()
        {
            var fbd = new FolderBrowserDialog();
            DialogResult result = fbd.ShowDialog();
            if (result == DialogResult.OK && !string.IsNullOrEmpty(fbd.SelectedPath)) {
                string path = fbd.SelectedPath;
                return path;
            }
            return "";
        }

        public static DataTable[] getAnswerFile(string directory, string dbName)
        {
            DataTable[] dtAnswer = new DataTable[10];
            string[] sfiles = Directory.GetFiles(directory);
            foreach (string f in sfiles) {
                string script = File.ReadAllText(@f);
                int index = -1;
                //put data of Question Number :X   to data Array index X-1
                //Q10 (last situation)
                if (!int.TryParse(f.Substring(f.Length - 6, 2), out index)) {
                    //Q1-9
                    int.TryParse(f.Substring(f.Length - 5, 1), out index);
                }
                try {
                    dtAnswer[index - 1] = Database.GetBySQL(script, dbName);
                }
                catch (Exception ex) {
                    //MessageBox.Show("Error to load file Q" + (index) + "\n" + ex.Message);
                }
            }
            return dtAnswer;
        }
        public static DataTable[] getSolutionFile(string directory, string dbName)
        {
            DataTable[] dtSolution = new DataTable[10];
            string[] sfiles = Directory.GetFiles(directory);
            int i = 0;
            // put data of Solution to data Array
            foreach (string f in sfiles) {
                string script = File.ReadAllText(@f);
                try {
                    dtSolution[i] = Database.GetBySQL(script, dbName);
                    i++;
                }
                catch (Exception ex) {
                    //MessageBox.Show("Error to load file Q" + (i + 1) + "\n" + ex.Message);
                }
            }
            if (i != 10) return null;
            return dtSolution;
        }

        public static Student GetMarkedStudent(DataTable[] dtsAnswer, DataTable[] dtSolution, string studentID, string studentName, string examPaperCode)
        {
            string mesage = "";
            double mark = 0;
            for (int i = 0; i < 10; i++) {
                mesage += "[Q=" + (i + 1) + "]";
                //student file available
                if (dtsAnswer[i] == null) {
                    mesage += " Empty.";
                }
                else {
                    //each question mark
                    double qMark = 0;
                    if (checkData(dtsAnswer[i], dtSolution[i])) {
                        mesage += "\t- Check Data: Passed => + 0,5";
                        qMark += 0.5;
                        if (checkSort(dtsAnswer[i], dtSolution[i])) {
                            mesage += "\t- Check Sort: Passed => + 0,5";
                            qMark += 0.5;
                        }
                        else {
                            mesage += "\t- Check Sort: Not pass => + 0";
                        }
                    }
                    else {
                        mesage += "\t- Check Data: Not pass => Point = 0\t  Stop checking";
                    }
                    mesage += "\t=> Mark = " + qMark;
                    mark += qMark;
                }
                //mesage += "\n";
            }
            return new Student(studentID, studentName, examPaperCode, mark, mesage);
        }

        DataTable[] dtSAnswer = new DataTable[10];
        DataTable[] dtSolution = new DataTable[10];


        private void putDataToDB(List<Student> list, string tableName)
        {
            if (list.Count != 0) {
                foreach (Student s in list) {
                    try {
                        Student.AddStudent(s, tableName);
                    }
                    catch (SqlException ex) {
                        MessageBox.Show(ex.Message);
                    }
                }
                //lblProcesss.Text = "Done!";
                //prgBar.Value = 100;
                //MessageBox.Show("Marked " + list.Count + " students! Data is put to Database \"DBI202_PE_Result\"");
            }
            else {
                MessageBox.Show("Not found any data! Check your input folder");
            }
        }


        public void LoadDataGridView()
        {
            dgvAll.DataSource = Database.GetAllTable();
            DataGridViewButtonColumn btn = new DataGridViewButtonColumn();
            {
                btn.HeaderText = "Action";
                btn.Name = "btnExport";
                btn.Text = "Export Excel";
                btn.UseColumnTextForButtonValue = true;
                dgvAll.Columns.Add(btn);
            }
        }

        private void FrmProject_Load(object sender, EventArgs e)
        {
            LoadDataGridView();
            dgvAll.AllowUserToAddRows = false;
            cboName.DataSource = Database.GetNameDatabase();
            cboName.DisplayMember = "name";
            cboName.ValueMember = "name";
            txtAnswer.ReadOnly = true;
            txtSolution.ReadOnly = true;
            rtxtMarked.ReadOnly = true;
            btnExcel.Enabled = false;
        }

        // export excel
        public void ExportExcel(string tableName)
        {
            using (SaveFileDialog sfd = new SaveFileDialog() { FileName = tableName, Filter = "Excel File|*.xlsx" }) {
                if (sfd.ShowDialog() == DialogResult.OK) {
                    try {
                        Microsoft.Office.Interop.Excel.ApplicationClass ExcelApp = new Microsoft.Office.Interop.Excel.ApplicationClass();
                        ExcelApp.Application.Workbooks.Add(Type.Missing);
                        DataTable dtMainSQLData = Student.GetResult(tableName);
                        DataColumnCollection dcCollection = dtMainSQLData.Columns;
                        for (int i = 1; i < dtMainSQLData.Rows.Count + 1; i++) {
                            for (int j = 1; j < dtMainSQLData.Columns.Count + 1; j++) {
                                if (i == 1) {
                                    ExcelApp.Cells[i, j] = dcCollection[j - 1].ToString();
                                    ExcelApp.Cells[i + 1, j] = dtMainSQLData.Rows[i - 1][j - 1].ToString();
                                }
                                else
                                    ExcelApp.Cells[i + 1, j] = dtMainSQLData.Rows[i - 1][j - 1].ToString();
                            }
                        }
                        ExcelApp.ActiveWorkbook.SaveCopyAs(sfd.FileName);
                        ExcelApp.ActiveWorkbook.Saved = true;
                        ExcelApp.Quit();
                        MessageBox.Show("Data Exported Successfully into Excel File",
                            "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    catch (Exception ex) {
                        //MessageBox.Show(ex.Message);
                    }
                }
            }
        }

        private void btnExcel_Click(object sender, EventArgs e)
        {
            string[] arr = txtAnswer.Text.Split('-');
            ExportExcel(arr[1]);
        }

        private void btnAnswer_Click(object sender, EventArgs e)
        {
            txtAnswer.Text = getFolderPath();
        }

        private void btnSolution_Click(object sender, EventArgs e)
        {
            txtSolution.Text = getFolderPath();
        }

        private void btnRun_Click(object sender, EventArgs e)
        {
            if (txtAnswer.Text.Equals("") || txtSolution.Text.Equals("")) {
                MessageBox.Show("Please fill all fields", "Warning",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            else {
                btnExcel.Enabled = true;
                rtxtMarked.Text = "";
                //lblProcesss.Text = "";
                //prgBar.Value = 0;
                List<Student> list = new List<Student>();
                // get all Student Folders
                string[] studentAnswer = Directory.GetDirectories(txtAnswer.Text);

                //get all Files from Solution Folder
                dtSolution = getSolutionFile(txtSolution.Text,
                    cboName.SelectedValue.ToString());
                if (dtSolution == null) {
                    MessageBox.Show("Check your solution path!", "Warning",
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
                else {
                    // select each Student Folder
                    foreach (string file in studentAnswer) {
                        //E:\FPT University\2021 Summer\PRN292\New folder\StudentAnswer\HExxxxxx_Name_DBI202_xx
                        string[] studentInfo = file.Split('_');
                        string studentID = "";
                        string studentName = "";
                        string examPaperCode = "";
                        try {
                            studentID = studentInfo[0].Substring(studentInfo[0].Length - 8, 8);
                            studentName = studentInfo[1];
                            examPaperCode = studentInfo[3];

                            //get all Files from Student Folder
                            dtSAnswer = getAnswerFile(file,
                                cboName.SelectedValue.ToString());

                            rtxtMarked.Text += studentID + "\n";

                            dtSolution = getSolutionFile(txtSolution.Text,
                                cboName.SelectedValue.ToString());
                            //Marking
                            Student student = GetMarkedStudent(dtSAnswer,
                                dtSolution, studentID, studentName, examPaperCode);
                            list.Add(student);
                        }
                        catch (Exception ex) {
                            //rtxtMarked.Text += "Wrong format file name found!\n";
                            //MessageBox.Show(ex.Message);
                        }
                    }
                    string[] arr = txtAnswer.Text.Split('-');

                    //create DB TB
                    Database.CreateDB();
                    Database.CreateTable(arr[1]);

                    //put data to DB
                    putDataToDB(list, arr[1]);
                    dgvAll.Columns.Remove("btnExport");
                    LoadDataGridView();
                }
            }
        }

        private void dgvAll_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            DataGridViewRow row = this.dgvAll.Rows[e.RowIndex];
            if (dgvAll.Columns[e.ColumnIndex].Name == "btnExport") {
                string tableName = row.Cells[1].Value.ToString();
                if (MessageBox.Show("Do you want export excel " + tableName + "?", "Comfirm",
                    MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes) {
                    ExportExcel(tableName);
                }
            }
        }
    }
}
