using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;


namespace Tamrin8
{
    public partial class MainForm : Form
    {
        #region Constructor
        /// <summary>
        /// Default Constructor
        /// </summary>
        public MainForm()
        {
            InitializeComponent();
        }
        #endregion

        #region Fields
        /// <summary>
        /// Inpput File Path
        /// </summary>
        public static string excelFilePath =
            @"C:\Users\faranam\Desktop\Exam\08 - Excel read\branches.xlsx";

        /// <summary>
        /// Opening Excel For use
        /// </summary>
        public static Microsoft.Office.Interop.Excel.Application
            Excel = new Microsoft.Office.Interop.Excel.Application();

        /// <summary>
        /// Excel Work Book
        /// </summary>
        public static Workbook workBook = Excel.Workbooks.Open(
            Filename: excelFilePath, IgnoreReadOnlyRecommended: true);

        /// <summary>
        /// Excel Work Sheet
        /// </summary>
        public static Worksheet workSheet = 
            (Worksheet)workBook.Worksheets.get_Item(1);

        /// <summary>
        /// Work Sheet Range
        /// </summary>
        public static Range usedRange = workSheet.UsedRange;
        #endregion

        #region AddGridColumns Function
        /// <summary>
        /// Add Columns to Grid
        /// </summary>
        public void AddGridColumns()
        {
            for (int i = 1; i <= usedRange.Columns.Count; i++)
            {
                mainDataGird.Columns.Add(workSheet.Cells[1, i].Value2, workSheet.Cells[1, i].Value2);
            }
        }
        #endregion

        #region Add Rows Function
        /// <summary>
        /// Add Rows to Grid
        /// </summary>
        public void AddRows()
        {
            string context;
            for (int i = 2; i <= usedRange.Rows.Count; i++)
            {
                mainDataGird.Rows.Add();
                for (int j = 1; j <= usedRange.Columns.Count; j++)
                {
                    context = Convert.ToString((usedRange.Cells[i, j] as Range).Value2);
                    mainDataGird.Rows[i - 2].Cells[j - 1].Value = context;
                }
            }
        }
        #endregion

        #region Close Excel Function and End Program
        /// <summary>
        /// Close Work Book and Quit Excel
        /// Last Step Of Program
        /// </summary>
        public void CloseExcel()
        {
            workBook.Close(true);
            Excel.Quit();
            MessageBox.Show("Jobs Done!");
        }
        #endregion

        #region MainFrom_Load EvENT Rasie
        /// <summary>
        /// Form Loading Time Cycle
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void MainForm_Load(object sender, EventArgs e)
        {
            AddGridColumns();
            AddRows();
            CloseExcel();
        }
        #endregion
    }
}
