using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;
using System.Web;
using System.Windows.Forms;

namespace Appointment
{
    public partial class Form1 : Form
    {
        private DataTable excelData;
        private readonly HttpClient httpClient = new HttpClient();
        public Form1()
        {
            InitializeComponent();
            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
        }
        private void btnUpload_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Excel Files|*.xlsx;*.xls";
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                var filePath = openFileDialog.FileName;
                excelData = ReadExcelToDataTable(filePath);
                dataGridView1.DataSource = excelData;
            }
        }

        private DataTable ReadExcelToDataTable(string filePath)
        {
            FileInfo fileInfo = new FileInfo(filePath);
            using (var package = new ExcelPackage(fileInfo))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets[0];
                DataTable dt = new DataTable();

                // Load headers
                for (int col = worksheet.Dimension.Start.Column; col <= worksheet.Dimension.End.Column; col++)
                {
                    dt.Columns.Add(worksheet.Cells[1, col].Text);
                }

                // Load data
                for (int row = worksheet.Dimension.Start.Row + 1; row <= worksheet.Dimension.End.Row; row++)
                {
                    DataRow dr = dt.NewRow();
                    for (int col = worksheet.Dimension.Start.Column; col <= worksheet.Dimension.End.Column; col++)
                    {
                        dr[col - 1] = worksheet.Cells[row, col].Text;
                    }
                    dt.Rows.Add(dr);
                }

                return dt;
            }
        }

        private void btnCheckEnc_Click(object sender, EventArgs e)
        {
            //FetchEncValue();
            if (excelData == null)
            {
                MessageBox.Show("Please upload an Excel file first.");
                return;
            }
            for (int i = 0; i < excelData.Rows.Count; i++)
            {
                var url = excelData.Rows[i][13].ToString(); // Column N = index 13
                var enc = excelData.Rows[i][14].ToString(); // Column O = index 14

                lblCurrentRow.Text = $"Processing Row {i + 1} ..."; // +2 to account for header row

                if (string.IsNullOrWhiteSpace(url) || !string.IsNullOrWhiteSpace(enc))
                    continue;

                //listBoxLog.Items.Add($"Processing Row {i + 2}...");
                var resultEnc = FetchEncValue(url).Result;
                excelData.Rows[i][14] = resultEnc;
                dataGridView1.Refresh(); // Show updates in UI

                //// 2. Call SaveBooking API if enc is valid
                //var row = excelData.Rows[i];
                //if (!string.IsNullOrWhiteSpace(resultEnc) && !resultEnc.StartsWith("ERROR"))
                //{
                //    string saveResult = CallSaveBookingApiAsync(row).Result;
                //    row["Save dsta"] = "save";
                //    dataGridView1.Refresh();
                //}
                //else
                //{
                //    row["Save dsta"] = "";
                //}
            }
            
            //string newColumnName = "Enc Value";
            //if (excelData.Columns.Contains(newColumnName))
            //    newColumnName += "_" + Guid.NewGuid().ToString("N").Substring(0, 4);

            //excelData.Columns.Add(newColumnName, typeof(string));

            //foreach (DataRow row in excelData.Rows)
            //{
            //    row[newColumnName] = "Default";
            //}

            //dataGridView1.DataSource = null; // Reset to apply
            //dataGridView1.DataSource = excelData;
        }
        private async Task<string> CallSaveBookingApiAsync(DataRow row)
        {
            try
            {
                string url = row["Url"]?.ToString();
                var parsed = ParseUrlParameters(url);

                // Extract time and format to HH:mm
                string rawTime = parsed["time"]; // e.g. 11m00
                string formattedTime = rawTime.Replace("m", ":"); // 11:00

                // Construct full date
                string bookDate = $"{parsed["year"]}-{parsed["month"].PadLeft(2, '0')}-{parsed["day"].PadLeft(2, '0')}"; // e.g. 2025-08-21

                var payload = new
                {
                    bookdate = bookDate,
                    booktime = formattedTime,
                    adults = int.Parse(parsed["adults"]),
                    children = int.Parse(parsed["children"]),
                    customers = 1,
                    couponcode = "",
                    cid = 12,
                    bid = int.Parse(parsed["bid"]),
                    pid = 0,
                    paymod = "",
                    ofirstname = row["First-name"]?.ToString(),
                    olastname = row["Last-name"]?.ToString(),
                    oemail = "test@example.com", // or get from Excel
                    ocountry = "PK",
                    ocity = row["City"]?.ToString(),
                    oaddress = "",
                    opostalcode = "",
                    ophone = row["Mobile*"]?.ToString(),
                    omobile = "",
                    ccustom1 = row["Number-of-the-Greek-Decision-(Apofasi)*"]?.ToString(),
                    ccustom2 = row["Employer's-Name-in-Greece*"]?.ToString(),
                    ccustom3 = row["Father's-name*"]?.ToString(),
                    ccustom4 = row["Passport-Number*"]?.ToString(),
                    ccustom5 = row["Date-of-Expiry-of-Passport*"]?.ToString(),
                    p1firstname = row["First-name"]?.ToString(),
                    p1lastname = row["Last-name"]?.ToString(),
                    ocomments = "",
                    invoice = 0,
                    enc = row["Enc"]?.ToString(),
                    rnd = new Random().Next(1, 100)
                };

                string json = Newtonsoft.Json.JsonConvert.SerializeObject(payload);
                var content = new StringContent(json, Encoding.UTF8, "application/json");

                using (var client = new HttpClient())
                {
                    HttpResponseMessage response = null;
                    int retries = 0;

                    do
                    {
                        response = await client.PostAsync("https://appointment.mfa.gr/inner.php/en/reservations/aero/makebook", content);
                        retries++;
                        btnSaveForm.Text = $"Save Booking {retries}";
                        if (!response.IsSuccessStatusCode)
                            await Task.Delay(3000);
                    } while (!response.IsSuccessStatusCode && retries < 5);

                    if (response.IsSuccessStatusCode)
                        return "Success";
                    else
                        return $"Failed: {response.StatusCode}";
                }
            }
            catch (Exception ex)
            {
                return $"Error: {ex.Message}";
            }
        }

        private void btnDownloadExcel_Click(object sender, EventArgs e)
        {
            if (excelData == null || excelData.Rows.Count == 0)
            {
                MessageBox.Show("No data available to export.");
                return;
            }

            SaveFileDialog sfd = new SaveFileDialog
            {
                Filter = "Excel Files|*.xlsx",
                FileName = "UpdatedReport.xlsx"
            };

            if (sfd.ShowDialog() == DialogResult.OK)
            {
                using (var package = new ExcelPackage())
                {
                    var ws = package.Workbook.Worksheets.Add("Sheet1");

                    // Load data from DataTable including headers
                    ws.Cells["A1"].LoadFromDataTable(excelData, true);
                    ws.Cells[ws.Dimension.Address].AutoFitColumns();

                    // Save to selected file
                    package.SaveAs(new FileInfo(sfd.FileName));
                }

                MessageBox.Show("Excel file downloaded successfully!");
            }
        }

        private async Task<string> FetchEncValue(string url)
        {
            //string url = "https://appointment.mfa.gr/en/reservations/aero/book/?bid=32&day=21&month=8&year=2025&time=11m00&adults=1&children=0";

            int retryCount = 0;
            int maxRetries = Convert.ToInt32(txtRateLimit.Text);
            double timeoutSeconds = Convert.ToDouble(txtInterval.Text);

            var httpClient1 = new HttpClient()
            {
                Timeout = TimeSpan.FromSeconds(timeoutSeconds)
            };

            while (retryCount < maxRetries)
            {
                try
                {
                    btnCheckEnc.Text = $"Check Enc {retryCount + 1}";
                    var response = httpClient1.GetAsync(url).Result;


                    if (response.StatusCode == System.Net.HttpStatusCode.ServiceUnavailable)
                    {
                        retryCount++;
                        await Task.Delay(1000); // wait before retrying
                        continue;
                    }

                    response.EnsureSuccessStatusCode();
                    string html = await response.Content.ReadAsStringAsync();

                    // Parse HTML to extract enc value using HtmlAgilityPack
                    HtmlAgilityPack.HtmlDocument doc = new HtmlAgilityPack.HtmlDocument();
                    doc.LoadHtml(html);

                    var encNode = doc.DocumentNode.SelectSingleNode("//input[@name='enc']");

                    if (encNode != null)
                    {
                        string encValue = encNode.GetAttributeValue("value", "");
                        MessageBox.Show("enc: " + encValue);
                        return encValue;
                    }
                    else
                    {
                        MessageBox.Show("enc input field not found.");
                        return "";
                    }
                }
                catch (Exception ex)
                {
                    retryCount++;
                    if (retryCount >= maxRetries)
                    {
                        //MessageBox.Show("Failed after "+maxRetries+" attempts. Error: " + ex.Message);
                        btnCheckEnc.Text = "Check Enc";
                        return "";
                    }
                    Task.Delay(1000).Wait(); // delay before retry
                }
            }
            return "";
        }
        private Dictionary<string, string> ParseUrlParameters(string url)
        {
            var uri = new Uri(url);
            var query = ParseQueryString(uri.Query);

            return new Dictionary<string, string>
            {
                ["bid"] = query["bid"] ?? "",
                ["day"] = query["day"] ?? "",
                ["month"] = query["month"] ?? "",
                ["year"] = query["year"] ?? "",
                ["time"] = query["time"] ?? "",
                ["adults"] = query["adults"] ?? "1",
                ["children"] = query["children"] ?? "0"
            };
        }

        private Dictionary<string, string> ParseQueryString(string url)
        {
            var uri = new Uri(url);
            var query = uri.Query.TrimStart('?').Split('&', (char)StringSplitOptions.RemoveEmptyEntries);

            return query
                .Select(q => q.Split('='))
                .Where(q => q.Length == 2)
                .ToDictionary(q => q[0], q => Uri.UnescapeDataString(q[1]));
        }

        private void btnSaveForm_Click(object sender, EventArgs e)
        {
            if (excelData == null)
            {
                MessageBox.Show("Please upload an Excel file first.");
                return;
            }
            for (int i = 0; i < excelData.Rows.Count; i++)
            {
                var url = excelData.Rows[i][13].ToString(); // Column N = index 13
                var enc = excelData.Rows[i][14].ToString(); // Column O = index 14
                var saveData = excelData.Rows[i][15].ToString(); // Column O = index 15

                lblCurrentRow.Text = $"Processing Row {i + 1} ..."; // +2 to account for header row

                if (string.IsNullOrWhiteSpace(url) || !string.IsNullOrWhiteSpace(enc) || !string.IsNullOrWhiteSpace(saveData))
                    continue;

                var row = excelData.Rows[i];
                if (!string.IsNullOrWhiteSpace(enc) && !enc.StartsWith("ERROR"))
                {
                    string saveResult = CallSaveBookingApiAsync(row).Result;
                    row["Save data"] = "save";
                    dataGridView1.Refresh();
                }
                else
                {
                    row["Save dsta"] = "";
                }
            }

            //string newColumnName = "Enc Value";
            //if (excelData.Columns.Contains(newColumnName))
            //    newColumnName += "_" + Guid.NewGuid().ToString("N").Substring(0, 4);

            //excelData.Columns.Add(newColumnName, typeof(string));

            //foreach (DataRow row in excelData.Rows)
            //{
            //    row[newColumnName] = "Default";
            //}

            //dataGridView1.DataSource = null; // Reset to apply
            //dataGridView1.DataSource = excelData;
        }
    }
}
