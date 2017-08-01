using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;
using DataTable = System.Data.DataTable;

namespace ExcelDedupe
{
    internal class ExcelDedupe
    {
        private static void Main(string[] args)
        {
            List<DataTable> sheets = new List<DataTable>();

            var fullPathToExcel = args[0];
            var connString = string.Format(
                "Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties='Excel 12.0';", fullPathToExcel);

            foreach (var arg in args)
            {
                if (arg == args.First())
                    continue;

                DataTable table =
                    ExcelReadWrite.ExcelReadWrite.GetDataTable("SELECT * from [" + arg + "$]", connString);
                table.TableName = arg;

                sheets.Add(table);
            }


            var firstNameIndex = 0;
            var lastNameIndex = 0;
            var secondaryFirstNameIndex = 0;
            var secondaryLastNameIndex = 0;
            var secondarySalutationIndex = 0;
            var secondaryMiddleIndex = 0;
            var secondarySuffixIndex = 0;
            var secondaryBusinessTitleIndex = 0;
            var organizationNameIndex = 0;

            var columnIndex = 0;

            foreach (DataColumn column in sheets[0].Columns)
            {
                if (column.ColumnName.Contains("Applicant First"))
                    firstNameIndex = columnIndex;
                else if (column.ColumnName.Contains("Applicant Last"))
                    lastNameIndex = columnIndex;
                else if (column.ColumnName.Contains("Secondary Contact First"))
                    secondaryFirstNameIndex = columnIndex;
                else if (column.ColumnName.Contains("Secondary Contact Last"))
                    secondaryLastNameIndex = columnIndex;
                else if (column.ColumnName.Contains("Secondary Contact Salutation"))
                    secondarySalutationIndex = columnIndex;
                else if (column.ColumnName.Contains("Secondary Contact Middle"))
                    secondaryMiddleIndex = columnIndex;
                else if (column.ColumnName.Contains("Secondary Contact Suffix"))
                    secondarySuffixIndex = columnIndex;
                else if (column.ColumnName.Contains("Secondary Contact Business Title"))
                    secondaryBusinessTitleIndex = columnIndex;
                else if (column.ColumnName.Contains("Organization Name"))
                    organizationNameIndex = columnIndex;
                columnIndex++;
            }


            foreach (DataTable sheet in sheets)
            foreach (DataRow line in sheet.Rows)
            {
                var secondaryFirstName = line[secondaryFirstNameIndex].ToString();
                var secondaryLastName = line[secondaryLastNameIndex].ToString();
                var organizationName = line[organizationNameIndex].ToString();

                var found = FoundDuplicate(secondaryFirstName, secondaryLastName, organizationName, sheets,
                    firstNameIndex, lastNameIndex, organizationNameIndex);

                if (found)
                {
                    line[secondaryFirstNameIndex] = "";
                    line[secondaryLastNameIndex] = "";
                    line[secondarySalutationIndex] = "";
                    line[secondaryMiddleIndex] = "";
                    line[secondarySuffixIndex] = "";
                    line[secondaryBusinessTitleIndex] = "";
                }
            }

            WriteExcelFile(sheets, args[0]);
        }

        private static void WriteExcelFile(List<DataTable> data, string originalFileName)
        {
            Application xlApp = new Application();

            object misValue = Missing.Value;

            Workbook xlWorkBook = xlApp.Workbooks.Add(misValue);

            var sheetIndex = 1;

            foreach (DataTable sheet in data)
            {
                Worksheet xlWorkSheet;
                try
                {
                    xlWorkSheet = (Worksheet) xlWorkBook.Worksheets.get_Item(sheetIndex);
                }
                catch (COMException)
                {
                    xlWorkSheet = (Worksheet) xlWorkBook.Worksheets.Add();
                }

                var columnIndex = 1;

                foreach (DataColumn column in sheet.Columns)
                {
                    xlWorkSheet.Cells[1, columnIndex] = column.ColumnName;
                    columnIndex++;
                }

                var rowIndex = 2;

                foreach (DataRow row in sheet.Rows)
                {
                    for (var i = 0; i < sheet.Columns.Count; i++)
                        xlWorkSheet.Cells[rowIndex, i + 1] = row[i].ToString();

                    rowIndex++;
                }
                xlWorkSheet.Name = sheet.TableName;

                sheetIndex++;
            }

            xlWorkBook.SaveAs(originalFileName.Replace(".xlsx", ".xls"), XlFileFormat.xlWorkbookNormal, misValue,
                misValue, misValue, misValue, XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue,
                misValue);
            xlWorkBook.Close(true, misValue, misValue);
            xlApp.Quit();
        }

        private static bool FoundDuplicate(string secondaryFirstName, string secondaryLastName, string organizationName,
            List<DataTable> sheets, int firstNameIndex, int lastNameIndex, int organizationNameIndex)
        {
            foreach (DataTable sheet in sheets)
            foreach (DataRow line in sheet.Rows)
            {
                var firstName = line[firstNameIndex].ToString();
                var lastName = line[lastNameIndex].ToString();
                var currentOrganizationName = line[organizationNameIndex].ToString();

                if (currentOrganizationName == organizationName && firstName == secondaryFirstName &&
                    lastName == secondaryLastName)
                    return true;
            }

            return false;
        }
    }
}