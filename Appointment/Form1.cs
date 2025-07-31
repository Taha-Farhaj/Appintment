using Newtonsoft.Json.Linq;
using OfficeOpenXml;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Security.Policy;
using System.Text;
using System.Text.Json;
using System.Threading;
using System.Threading.Tasks;
using System.Web;
using System.Windows.Forms;

namespace Appointment
{
    public partial class Form1 : Form
    {
        private DataTable excelData;
        private readonly HttpClient httpClient = new HttpClient();
        private string cid = "";
        private CancellationTokenSource cancellationTokenSource;
        private bool isProcessing = false;
        private ChromeDriver driver = new ChromeDriver();

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

        private string Chromejob(int row = 2)
        {
            if (cancellationTokenSource.IsCancellationRequested)
            {
                btnCheckEnc.Text = "Start Booking";
                isProcessing = false;
                cancellationTokenSource.Dispose();
                return "Cancelled";
            }
            try
            {
                int rowCount = excelData.Rows.Count;

                // Start Chrome
               

                string username = excelData.Rows[row]["Email"].ToString();
                string password = excelData.Rows[row][2].ToString();
                string phone = excelData.Rows[row][3].ToString();
                string mobile = excelData.Rows[row][4].ToString();
                string greeceRegion = excelData.Rows[row][5].ToString();
                string apofasiNumber = excelData.Rows[row][6].ToString();
                string employeeName = excelData.Rows[row][7].ToString();
                string passportNumber = excelData.Rows[row][8].ToString();
                string passportExp = excelData.Rows[row][9].ToString();

                // Navigate to login page
                driver.Navigate().GoToUrl(txtWebsiteUrl.Text);

                // Fill login form
                driver.FindElement(By.Id("name")).Clear();
                driver.FindElement(By.Id("name")).SendKeys(username);
                driver.FindElement(By.Id("password")).SendKeys(password);
                driver.FindElement(By.ClassName("i-login-g")).Click();

                Thread.Sleep(2000); // Wait for login

                // Check for availability (adjust logic based on actual DOM)
                if (!driver.PageSource.Contains("No available space found") && !driver.PageSource.Contains("Invalid password"))
                {
                    Console.WriteLine("Slot available");

                    // Fill appointment form
                    //driver.FindElement(By.Id("appointment_field")).SendKeys(appointmentData);
                    //driver.FindElement(By.Id("submit_button")).Click();
                    //Console.WriteLine($"Appointment submitted for {username}");
                    // Click the first available row
                    //driver.FindElement(By.XPath("//ul[@class='ts']/li[text()='Available']")).Click();

                    var wait = new WebDriverWait(driver, TimeSpan.FromSeconds(10));
                    IWebElement availableBtn = wait.Until(d => d.FindElement(By.CssSelector("ul.ts li[onclick=\"switch_view('free')\"]")));

                    try
                    {
                        // Try standard click
                        availableBtn.Click();
                    }
                    catch (ElementNotInteractableException)
                    {
                        // Fallback to JS click
                        IJavaScriptExecutor js = (IJavaScriptExecutor)driver;
                        js.ExecuteScript("arguments[0].click();", availableBtn);
                    }

                    var slotRows = driver.FindElements(By.CssSelector("table.table tbody tr"));
                    if (slotRows.Count == 0)
                    {
                        Console.WriteLine($"No slots found for user: {username}");
                    }

                    for (int i = 0; i < slotRows.Count; i++)
                    {
                        try
                        {
                            slotRows[i].Click();
                            Thread.Sleep(1000);

                            // Click "New Appointment" icon
                            //var newApptIcon = driver.FindElement(By.CssSelector("span.i.i-p.col0"));
                            //newApptIcon.Click();
                            //Thread.Sleep(1000);

                            // Fill popup form fields(replace with actual field names)
                            driver.FindElement(By.Id("reservation_phone")).SendKeys(phone);
                            driver.FindElement(By.Id("reservation_mobile")).Clear();
                            driver.FindElement(By.Id("reservation_mobile")).SendKeys(mobile);

                            // Click submit
                            var createBtn = wait.Until(d => d.FindElement(By.XPath("//button[contains(text(), 'Create appointment')]")));

                            // Scroll into view (optional, to avoid not-interactable errors)
                            ((IJavaScriptExecutor)driver).ExecuteScript("arguments[0].scrollIntoView(true);", createBtn);
                            Thread.Sleep(200); // brief pause to ensure visibility

                            // Click the button
                            createBtn.Click();
                            Console.WriteLine($"Appointment submitted for {username}");
                            break;

                        }
                        catch (Exception)
                        {
                            slotRows[i].Click();
                        }
                    }


                    IWebElement dropdown = driver.FindElement(By.Id("form_1"));

                    driver.FindElement(By.Id("form_1")).Click();
                    //driver.FindElement(By.XPath("//div[text()='"+greeceRegion+"']")).Click();
                    driver.FindElement(By.Id("form_2")).SendKeys(apofasiNumber);
                    driver.FindElement(By.Id("form_3")).SendKeys(employeeName);
                    driver.FindElement(By.Id("form_4")).SendKeys(passportNumber);
                    driver.FindElement(By.Id("form_5")).SendKeys(passportExp);

                    //driver.FindElement(By.Name("form_commit")).Click();

                    Thread.Sleep(2000);
                    driver.Navigate().GoToUrl("https://www.supersaas.com/users/logout/Saimways?form=form_2&return=Work");
                    //driver.Quit();
                    //Thread.Sleep(2000);
                }
                else
                {
                    Console.WriteLine("No appointment slot available");
                    return "Ivalid Passord";
                    //driver.Quit();
                }
                //Thread.Sleep(2000);
                //driver.Navigate().GoToUrl("https://www.supersaas.com/users/logout/Saimways?form=form_2&return=Work");
                

            }
            catch (Exception ex)
            {
                return ex.Message;
            }
            return "";
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

        private async void btnCheckEnc_Click(object sender, EventArgs e)
        {

            if (isProcessing)
            {
                MessageBox.Show("Process is already running.");
                return;
            }
            isProcessing = true;
            cancellationTokenSource = new CancellationTokenSource();
            btnStop.Enabled = true;

            if (excelData == null)
            {
                MessageBox.Show("Please upload an Excel file first.");
                return;
            }
            for (int i = 0; i < excelData.Rows.Count; i++)
            {
                //var url = excelData.Rows[i]["URL"].ToString(); // Column N = index 13
                //var enc = excelData.Rows[i]["Enc"].ToString(); // Column O = index 14

                //lblCurrentRow.Text = $"Processing Row {i + 1} ..."; // +2 to account for header row

                //if (string.IsNullOrWhiteSpace(url) || !string.IsNullOrWhiteSpace(enc))
                //    continue;

                //var resultEnc = await Task.Run(() => FetchEncValue(url, cancellationTokenSource.Token));
                //cancellationTokenSource.Token.ThrowIfCancellationRequested();
                //var resultEnc = FetchEncValue(url).Result;
                var status = Chromejob(i);
                excelData.Rows[i]["Status"] = status;
                dataGridView1.Refresh(); // Show updates in UI

                if (status == "Cancelled")
                    break;

            }
            driver.Quit();
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
                    //ws.Cells[ws.Dimension.Address]?.AutoFitColumns();

                    // Save to selected file
                    package.SaveAs(new FileInfo(sfd.FileName));
                }

                MessageBox.Show("Excel file downloaded successfully!");
            }
        }


        private void btnSaveForm_Click(object sender, EventArgs e)
        {
            if (excelData == null)
            {
                MessageBox.Show("Please upload an Excel file first.");
                return;
            }
        }

        private void btnStop_Click(object sender, EventArgs e)
        {
            if (cancellationTokenSource != null && isProcessing)
            {
                cancellationTokenSource.Cancel();
                MessageBox.Show("Cancellation requested...");
            }
        }
    }
}
