/* Title:           Main Window - Disable Employees
 * Date:            5-30-18
 * Author:          Terry Holmes   
 * 
 * Description:     This program is designed to disable employees in a bulk mode */

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using NewEmployeeDLL;
using NewEventLogDLL;

using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;

namespace DisableEmployees
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        WPFMessagesClass TheMessagesClass = new WPFMessagesClass();
        EmployeeClass TheEmployeeClass = new EmployeeClass();
        EventLogClass TheEventLogClass = new EventLogClass();

        ImportDataSet TheImportDataSet = new ImportDataSet();

        int gintColumnCounter;

        public MainWindow()
        {
            InitializeComponent();
        }

        private void mitClose_Click(object sender, RoutedEventArgs e)
        {
            TheMessagesClass.CloseTheProgram();
        }

        private void mitImportExcel_Click(object sender, RoutedEventArgs e)
        {
            bool blnFatalError = false;
            Excel.Application xlDropOrder;
            Excel.Workbook xlDropBook;
            Excel.Worksheet xlDropSheet;
            Excel.Range range;

            string strInformation = "";
            int intColumnRange = 0;
            int intCounter;
            int intNumberOfRecords;
            char[] chaInformation;
            int intCharCounter;
            int intCharLength;
            string strCompleteWord = "";
            int intLength;

            try
            {
                TheImportDataSet.employees.Rows.Clear();

                Microsoft.Win32.OpenFileDialog dlg = new Microsoft.Win32.OpenFileDialog();
                dlg.FileName = "Document"; // Default file name
                //dlg.DefaultExt = ".csv"; // Default file extension
               // dlg.Filter = "Excel (.csv)|*.csv"; // Filter files by extension

                // Show open file dialog box
                Nullable<bool> result = dlg.ShowDialog();

                // Process open file dialog box results
                if (result == true)
                {
                    // Open document
                    string filename = dlg.FileName;
                }

                PleaseWait PleaseWait = new PleaseWait();
                PleaseWait.Show();

                xlDropOrder = new Excel.Application();
                xlDropBook = xlDropOrder.Workbooks.Open(dlg.FileName, 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
                xlDropSheet = (Excel.Worksheet)xlDropOrder.Worksheets.get_Item(1);

                range = xlDropSheet.UsedRange;
                intNumberOfRecords = range.Rows.Count;
                intColumnRange = range.Columns.Count;

                for (intCounter = 1; intCounter <= intNumberOfRecords; intCounter++)
                {
                    strInformation = Convert.ToString((range.Cells[intCounter, 1] as Excel.Range).Value2);

                    gintColumnCounter = 0;

                    chaInformation = strInformation.ToCharArray();

                    intCharLength = strInformation.Length - 1;

                    ImportDataSet.employeesRow ActiveEmployeeRow = TheImportDataSet.employees.NewemployeesRow();

                    ActiveEmployeeRow.EmployeeID = Convert.ToInt32(strInformation);

                    strInformation = Convert.ToString((range.Cells[intCounter, 2] as Excel.Range).Value2);

                    ActiveEmployeeRow.FirstName = strInformation;

                    strInformation = Convert.ToString((range.Cells[intCounter, 3] as Excel.Range).Value2);

                    ActiveEmployeeRow.LastName = strInformation;

                    strInformation = Convert.ToString((range.Cells[intCounter, 4] as Excel.Range).Value2);

                    ActiveEmployeeRow.PhoneNumber = strInformation;

                    TheImportDataSet.employees.Rows.Add(ActiveEmployeeRow);
                }

               

                PleaseWait.Close();
                dgrResults.ItemsSource = TheImportDataSet.employees;
            }
            catch (Exception Ex)
            {
                TheEventLogClass.InsertEventLogEntry(DateTime.Now, "Disable Employees // Main Window // Import Excel Menu Item " + Ex.Message);

                TheMessagesClass.ErrorMessage(Ex.ToString());
            }
        }

        private void mitDisableEmployee_Click(object sender, RoutedEventArgs e)
        {
            int intEmployeeID;
            int intCounter;
            int intNumberOfRecords;
            bool blnFatalError;

            try
            {
                intNumberOfRecords = TheImportDataSet.employees.Rows.Count - 1;

                for(intCounter = 0; intCounter <= intNumberOfRecords; intCounter++)
                {
                    intEmployeeID = TheImportDataSet.employees[intCounter].EmployeeID;

                    blnFatalError = TheEmployeeClass.DeactivateEmployee(intEmployeeID);

                    if (blnFatalError == true)
                        throw new Exception();
                }
            }
            catch (Exception Ex)
            {
                TheEventLogClass.InsertEventLogEntry(DateTime.Now, "Disable Employees // Main window // Disable Employee Menu Item " + Ex.Message);

                TheMessagesClass.ErrorMessage(Ex.ToString());
            }
        }
    }
}
