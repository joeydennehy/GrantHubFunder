using System;
using System.Collections.Generic;
using System.Data;
using System.Net.Mail;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using Microsoft.Office.Interop.Excel;
using DataTable = System.Data.DataTable;

namespace GrantHubFunders
{
    internal class GrantHubFunders
    {
        private static void Main(string[] args)
        {
            var fullPathToMasterExcel = args[0];
            var connString = string.Format(
                "Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties='Excel 12.0';",
                fullPathToMasterExcel);

            DataTable masterTable = ExcelReadWrite.ExcelReadWrite.GetDataTable("SELECT * from [Master$]", connString);
            masterTable.TableName = args[0];

            var fullPathToNewExcel = args[1];
            connString = string.Format(
                "Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties='Excel 12.0';",
                fullPathToNewExcel);

            DataTable newTable = ExcelReadWrite.ExcelReadWrite.GetDataTable("SELECT * from [Sheet1$]", connString);
            newTable.TableName = args[1];

            var organizationNameIndex = 0;
            var organizationTaxIdIndex = 0;
            var organizationWebsiteIndex = 0;
            var organizationPhoneIndex = 0;
            var organizationEmailIndex = 0;

            var columnIndex = 0;

            foreach (DataColumn column in newTable.Columns)
            {
                if (column.ColumnName.Contains("Name"))
                    organizationNameIndex = columnIndex;
                else if (column.ColumnName.Contains("Tax Id"))
                    organizationTaxIdIndex = columnIndex;
                else if (column.ColumnName.Contains("Web Site"))
                    organizationWebsiteIndex = columnIndex;
                else if (column.ColumnName.Contains("Phone Number"))
                    organizationPhoneIndex = columnIndex;
                else if (column.ColumnName.Contains("Email Address"))
                    organizationEmailIndex = columnIndex;

                columnIndex++;
            }

            var masterOrganizationNameIndex = 0;
            var masterTaxIdIndex = 0;
            var masterWebSiteIndex = 0;
            var masterPhoneIndex = 0;
            var masterEmailIndex = 0;

            columnIndex = 0;

            foreach (DataColumn column in masterTable.Columns)
            {
                if (column.ColumnName.Contains("Name"))
                    masterOrganizationNameIndex = columnIndex;
                else if (column.ColumnName.Contains("Tax Id"))
                    masterTaxIdIndex = columnIndex;
                else if (column.ColumnName.Contains("Web site"))
                    masterWebSiteIndex = columnIndex;
                else if (column.ColumnName.Contains("Phone"))
                    masterPhoneIndex = columnIndex;
                else if (column.ColumnName.Contains("Email"))
                    masterEmailIndex = columnIndex;

                columnIndex++;
            }

            DataTable dedupedTaxIdOrWebsite = new DataTable();
            SetTableColumns(dedupedTaxIdOrWebsite);

            DataTable dedupedNonTaxIdOrWebsite = new DataTable();
            SetTableColumns(dedupedNonTaxIdOrWebsite);

            DataTable taxIdOrWebsiteOrganizations = new DataTable();
            SetTableColumns(taxIdOrWebsiteOrganizations);

            DataTable nonTaxIdOrWebsiteOrganizations = new DataTable();
            SetTableColumns(nonTaxIdOrWebsiteOrganizations);

            DataTable masterTaxIdOrWebsite = new DataTable();
            SetTableColumns(masterTaxIdOrWebsite);

            DataTable masterNonTaxIdOrWebsite = new DataTable();
            SetTableColumns(masterNonTaxIdOrWebsite);

            DataTable masterDeduped = new DataTable();
            SetTableColumns(masterDeduped);


            foreach (DataRow line in masterTable.Rows)
            {
                Organizations organization = new Organizations
                {
                    Name = line[masterOrganizationNameIndex].ToString(),
                    TaxId = line[masterTaxIdIndex].ToString(),
                    Website = line[masterWebSiteIndex].ToString(),
                    Phone = line[masterPhoneIndex].ToString(),
                    Email = line[masterEmailIndex].ToString()
                };

                if (!string.IsNullOrWhiteSpace(organization.TaxId) || !string.IsNullOrWhiteSpace(organization.Website))
                    GetTableRow(masterTaxIdOrWebsite, organization);
                else
                    GetTableRow(masterNonTaxIdOrWebsite, organization);
            }

            foreach (DataRow line in newTable.Rows)
            {
                Organizations organization = new Organizations
                {
                    Name = line[organizationNameIndex].ToString(),
                    TaxId = line[organizationTaxIdIndex].ToString(),
                    Website = line[organizationWebsiteIndex].ToString(),
                    Phone = line[organizationPhoneIndex].ToString(),
                    Email = line[organizationEmailIndex].ToString()
                };

                if (!string.IsNullOrWhiteSpace(organization.TaxId) || !string.IsNullOrWhiteSpace(organization.Website))
                    GetTableRow(taxIdOrWebsiteOrganizations, organization);
                else
                    GetTableRow(nonTaxIdOrWebsiteOrganizations, organization);
            }

            foreach (DataRow organization in taxIdOrWebsiteOrganizations.Rows)
                if (!string.IsNullOrWhiteSpace(organization[1].ToString()))
                {
                    if (IsDuplicateOrganizationByTaxId(masterTaxIdOrWebsite, organization) == null)
                        if (IsDuplicateOrganizationByTaxId(dedupedTaxIdOrWebsite, organization) == null)
                        {
                            GetTableRow(dedupedTaxIdOrWebsite, organization);
                            GetTableRow(masterTaxIdOrWebsite, organization);
                        }
                }
                else
                {
                    if (IsDuplicateOrganization(masterTaxIdOrWebsite, organization) == null)
                        if (IsDuplicateOrganization(dedupedTaxIdOrWebsite, organization) == null)
                        {
                            GetTableRow(dedupedTaxIdOrWebsite, organization);
                            GetTableRow(masterTaxIdOrWebsite, organization);
                        }
                }

            dedupedTaxIdOrWebsite.TableName = "EINs and-or Websites";


            foreach (DataRow organization in nonTaxIdOrWebsiteOrganizations.Rows)
                if (IsDuplicateOrganization(masterNonTaxIdOrWebsite, organization) == null)
                    if (IsDuplicateOrganization(dedupedNonTaxIdOrWebsite, organization) == null)
                    {
                        GetTableRow(dedupedNonTaxIdOrWebsite, organization);
                        GetTableRow(masterNonTaxIdOrWebsite, organization);
                    }

            dedupedNonTaxIdOrWebsite.TableName = "No EINs, No Website";

            List<DataTable> sheets = new List<DataTable>();

            sheets.Add(dedupedNonTaxIdOrWebsite);
            sheets.Add(dedupedTaxIdOrWebsite);

            WriteExcelFile(sheets, args[1]);

            DataTable dedupeMaster = new DataTable();
            SetTableColumns(dedupeMaster);

            foreach (DataRow organization in masterTaxIdOrWebsite.Rows)
            {
                if (!string.IsNullOrWhiteSpace(organization[1].ToString()))
                {
                    if (IsDuplicateOrganizationByTaxId(dedupeMaster, organization) == null)
                    {

                        DataRow row = dedupeMaster.NewRow();
                        row[0] = organization[0];
                        row[1] = organization[1];
                        row[2] = organization[2];
                        row[3] = organization[3];
                        row[4] = organization[4];
                        dedupeMaster.Rows.Add(row);

                    }
                }
                else
                {
                    GetTableRow(masterNonTaxIdOrWebsite, organization);
                }
            }

            foreach (DataRow organization in masterNonTaxIdOrWebsite.Rows)
            {
              
                    if (IsDuplicateOrganization(dedupeMaster, organization) == null)
                    {

                        DataRow row = dedupeMaster.NewRow();
                        row[0] = organization[0];
                        row[1] = organization[1];
                        row[2] = organization[2];
                        row[3] = organization[3];
                        row[4] = organization[4];
                        dedupeMaster.Rows.Add(row);

                    }
                
            }

            dedupeMaster.TableName = "Master";

            List<DataTable> masterSheets = new List<DataTable>();
            masterSheets.Add(dedupeMaster);

            WriteExcelFile(masterSheets,
                "Master_" + DateTime.Now.Year + DateTime.Now.Month + DateTime.Now.Day + ".xls");
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

                xlWorkSheet.Cells[1, 1] = "Name";
                xlWorkSheet.Cells[1, 2] = "Tax Id";
                xlWorkSheet.Cells[1, 3] = "Web site";
                xlWorkSheet.Cells[1, 4] = "Phone";
                xlWorkSheet.Cells[1, 5] = "Email";
                xlWorkSheet.Cells[1, 6] = "Initials";
                xlWorkSheet.Cells[1, 7] = "SF";
                xlWorkSheet.Cells[1, 8] = "GLM";
                xlWorkSheet.Cells[1, 9] = "Phone";
                xlWorkSheet.Cells[1, 10] = "Email";

                for (var i = 1; i <= 10; i++)
                    xlWorkSheet.Cells[1, i].BorderAround(XlLineStyle.xlContinuous, XlBorderWeight.xlMedium);

                var rowIndex = 2;

                foreach (DataRow row in sheet.Rows)
                {
                    for (var i = 0; i < 5; i++)
                    {
                        xlWorkSheet.Cells[rowIndex, i + 1] = row[i].ToString();

                        if (i + 1 == 4 || i + 1 == 5)
                            xlWorkSheet.Cells[rowIndex, i + 1]
                                .BorderAround(XlLineStyle.xlContinuous, XlBorderWeight.xlThin);
                    }

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

        private static void SetTableColumns(DataTable table)
        {
            table.Columns.Add(new DataColumn("Name"));
            table.Columns.Add(new DataColumn("TaxId"));
            table.Columns.Add(new DataColumn("Website"));
            table.Columns.Add(new DataColumn("Phone"));
            table.Columns.Add(new DataColumn("Email"));
        }

        private static void GetTableRow(DataTable table, Organizations organization)
        {
            DataRow row = table.NewRow();
            row["Name"] = organization.Name;
            row["TaxId"] = organization.TaxId;
            row["Website"] = organization.Website;
            row["Phone"] = organization.Phone;
            row["Email"] = organization.Email;

            table.Rows.Add(row);
        }

        private static void GetTableRow(DataTable table, DataRow organization)
        {
            DataRow row = table.NewRow();
            row["Name"] = organization["Name"];
            row["TaxId"] = organization["TaxId"];
            row["Website"] = organization["Website"];
            row["Phone"] = organization["Phone"];
            row["Email"] = organization["Email"];
            table.Rows.Add(row);
        }

        private static string IsDuplicateOrganizationByTaxId(DataTable deduped, DataRow compare)
        {
            foreach (DataRow dedupe in deduped.Rows)
                if (dedupe[1].ToString() == compare[1].ToString())
                    return string.Format("Tax Id: {0} - {1}", dedupe[1], compare[1]);


            return null;
        }

        private static string IsDuplicateOrganization(DataTable deduped, DataRow organization)
        {
            foreach (DataRow dedupe in deduped.Rows)
            {
                if (string.IsNullOrWhiteSpace(organization[0].ToString())) continue;
                if (dedupe[0].ToString().ToLower() == organization[0].ToString().ToLower())
                    return string.Format("Name: {0} - {1}", dedupe[0], organization[0]);
                var s = dedupe[0].ToString().ToLower();
                var t = organization[0].ToString().ToLower();

                if (string.IsNullOrEmpty(s))
                {
                    if (string.IsNullOrEmpty(t))
                        return null;
                    return null;
                }

                if (string.IsNullOrEmpty(t))
                    return null;

                var n = s.Length;
                var m = t.Length;
                int[,] d = new int[n + 1, m + 1];

                // initialize the top and right of the table to 0, 1, 2, ...
                for (var i = 0; i <= n; d[i, 0] = i++)
                {
                }
                for (var j = 1; j <= m; d[0, j] = j++)
                {
                }

                for (var i = 1; i <= n; i++)
                for (var j = 1; j <= m; j++)
                {
                    var cost = t[j - 1] == s[i - 1] ? 0 : 1;
                    var min1 = d[i - 1, j] + 1;
                    var min2 = d[i, j - 1] + 1;
                    var min3 = d[i - 1, j - 1] + cost;
                    d[i, j] = Math.Min(Math.Min(min1, min2), min3);
                }

                var distance = d[n, m];
                var bigger = Math.Max(s.Length, t.Length);
                var percent = (int) ((bigger - distance) / (double) bigger * 100);

                if (percent >= 95)
                    return string.Format("Name: {0} - {1}", dedupe[0], organization[0]);

                if (percent >= 90 && percent < 95)
                {
                    Uri uriResult;
                    Uri.TryCreate(organization[2].ToString(), UriKind.Absolute, out uriResult);
                    if (uriResult != null)
                    {
                        if (dedupe[2].ToString().ToLower() == organization[2].ToString().ToLower())
                            return string.Format("Webstie: {0} - {1}", dedupe[2], organization[2]);
                    }
                    else
                    {
                        MailAddress email;
                        try
                        {
                            email = new MailAddress(organization[4].ToString());
                        }
                        catch
                        {
                            email = null;
                        }
                        if (email != null)
                        {
                            if (dedupe[4].ToString().ToLower() == organization[4].ToString().ToLower())
                                return string.Format("Email: {0} - {1}", dedupe[4], organization[4]);
                        }
                        else
                        {
                            var dedupePhone = Regex.Replace(dedupe[3].ToString(), "[^0-9 _]", "");
                            var organizationPhone = Regex.Replace(organization[3].ToString(), "[^0-9 _]", "");
                            if (!string.IsNullOrWhiteSpace(organizationPhone))
                                if (dedupePhone == organizationPhone)
                                    return string.Format("Phone: {0} - {1}", dedupe[3], organization[3]);
                        }
                    }
                }
            }
            return null;
        }

        private struct Organizations
        {
            public string Name;
            public string TaxId;
            public string Website;
            public string Phone;
            public string Email;
        }
    }
}