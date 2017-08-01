using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using DataProvider.MySQL;
using Microsoft.Office.Interop.Excel;
using MySql.Data.MySqlClient;
using DataTable = System.Data.DataTable;

namespace Foundations
{
    internal class FoundationInfo
    {
        private static void Main(string[] args)
        {
            Command command = new Command
            {
                SqlStatementId = "SELECT_FOUNDATION_INFO"
            };

            DataAccess access = new DataAccess();

            List<Foundations> foundationInfo = new List<Foundations>();

            using (MySqlDataReader reader = access.GetReader(command))
            {
                while (reader.Read())
                    if (!reader.IsDBNull(0))
                    {
                        Foundations foundations = new Foundations
                        {
                            FoundationId = reader.IsDBNull(0) ? -1 : reader.GetInt32(0),
                            FoundationName = reader.GetString(1),
                            ScholarshipValue = reader.IsDBNull(2) ? "False" : reader.GetString(2)
                        };

                        foundationInfo.Add(foundations);
                    }
            }

            List<int> foundationIds = foundationInfo.Select(f => f.FoundationId).ToList();
            List<Foundations> contactFoundations = new List<Foundations>();


            foreach (var id in foundationIds)
            {
                ParameterSet parameters = new ParameterSet();
                parameters.Add(DbType.Int32, "FOUNDATION_ID", id);

                command = new Command
                {
                    SqlStatementId = "SELECT_FOUNDATION_CONTACT_EMAIL",
                    ParameterCollection = parameters
                };

                List<string> contactEmails = new List<string>();

                using (MySqlDataReader reader = access.GetReader(command))
                {
                    while (reader.Read())
                        if (!reader.IsDBNull(0))
                            contactEmails.Add(reader.GetString(0));
                }

                Foundations foundation = foundationInfo.First(f => f.FoundationId == id);
                foundation.Contacts = new List<string>();
                foreach (var contact in contactEmails)
                    foundation.Contacts.Add(contact);

                contactFoundations.Add(foundation);
            }

            var fullPathToMasterExcel = args[0];
            var connString = string.Format(
                "Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties='Excel 12.0';",
                fullPathToMasterExcel);

            DataTable newTable =
                ExcelReadWrite.ExcelReadWrite.GetDataTable("SELECT * from [To Chris March$]", connString);
            newTable.TableName = "To Chris March";

            newTable.Columns.Add("Scholarship");


            foreach (DataRow line in newTable.Rows)
            {
                if (line[0].ToString() != "")
                    continue;

                List<Foundations> foundFoundations =
                    contactFoundations.Where(f => f.FoundationName.ToLower() == line[2].ToString().ToLower()).ToList();

                if (foundFoundations.Count() == 1)
                {
                    line[0] = foundFoundations.First().FoundationId;
                    line[10] = foundFoundations.First().ScholarshipValue;
                }
                else if (foundFoundations.Count > 1)
                {
                    List<Foundations> nonScholarshipfoundFoundations =
                        foundFoundations.Where(f => f.ScholarshipValue != "True").ToList();

                    if (nonScholarshipfoundFoundations.Count() == 1)
                    {
                        line[0] = nonScholarshipfoundFoundations.First().FoundationId;
                        line[10] = nonScholarshipfoundFoundations.First().ScholarshipValue;
                    }
                    else if (nonScholarshipfoundFoundations.Any())
                    {
                        nonScholarshipfoundFoundations = nonScholarshipfoundFoundations
                            .Where(f => f.Contacts.Contains(line[6].ToString()))
                            .ToList();

                        if (nonScholarshipfoundFoundations.Count == 1)
                        {
                            line[0] = nonScholarshipfoundFoundations.First().FoundationId;
                            line[10] = nonScholarshipfoundFoundations.First().ScholarshipValue;
                        }
                    }
                    else
                    {
                        foundFoundations = foundFoundations
                            .Where(f => f.Contacts.Contains(line[6].ToString()))
                            .ToList();

                        if (foundFoundations.Count >= 1)
                        {
                            line[0] = nonScholarshipfoundFoundations.First().FoundationId;
                            line[10] = nonScholarshipfoundFoundations.First().ScholarshipValue;
                        }
                    }
                }
                else
                {
                    List<Foundations> contactfoundFoundations = contactFoundations
                        .Where(f => f.Contacts.Contains(line[6].ToString()))
                        .ToList();

                    if (contactfoundFoundations.Count == 1)
                    {
                        line[0] = contactfoundFoundations.First().FoundationId;
                        line[10] = contactfoundFoundations.First().ScholarshipValue;
                    }
                    else
                    {
                        List<Foundations> contactNonScholarshipfoundFoundations = contactfoundFoundations
                            .Where(f => f.ScholarshipValue != "True")
                            .ToList();

                        if (contactNonScholarshipfoundFoundations.Count == 1)
                        {
                            line[0] = contactfoundFoundations.First().FoundationId;
                            line[10] = contactfoundFoundations.First().ScholarshipValue;
                        }
                    }
                }
            }

            List<DataTable> tableList = new List<DataTable>();
            tableList.Add(newTable);

            WriteExcelFile(tableList, args[0]);
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

                xlWorkSheet.Cells[1, 1] = "Id";
                xlWorkSheet.Cells[1, 2] = "GLM Name";
                xlWorkSheet.Cells[1, 3] = "Account Name";
                xlWorkSheet.Cells[1, 4] = "2017 permissions";
                xlWorkSheet.Cells[1, 5] = "First Name";
                xlWorkSheet.Cells[1, 6] = "Last Name";
                xlWorkSheet.Cells[1, 7] = "Email";
                xlWorkSheet.Cells[1, 8] = "Billing City";
                xlWorkSheet.Cells[1, 9] = "Billing State/Province";
                xlWorkSheet.Cells[1, 10] = "Type of Grantmaker";
                xlWorkSheet.Cells[1, 11] = "Scholarship";

                for (var i = 1; i <= 11; i++)
                    xlWorkSheet.Cells[1, i].BorderAround(XlLineStyle.xlContinuous, XlBorderWeight.xlThin);

                var rowIndex = 2;

                foreach (DataRow row in sheet.Rows)
                {
                    for (var i = 0; i < 11; i++)
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

        public struct Foundations
        {
            public string FoundationName;
            public int FoundationId;
            public string ScholarshipValue;
            public List<string> Contacts;
        }
    }
}