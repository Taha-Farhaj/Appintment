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
        public ChromeDriver driver = new ChromeDriver();

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

        private async Task<string> ChromeJobAsync(int row, CancellationToken token)
        {
            if (cancellationTokenSource.IsCancellationRequested)
            {
                btnCheckEnc.Text = "Start Booking";
                isProcessing = false;
                cancellationTokenSource.Dispose();
                return "Cancelled";
            }
            string status = "";
            try
            {


                string username = excelData.Rows[row]["Email"].ToString();
                string password = excelData.Rows[row][1].ToString();
                string phone = excelData.Rows[row][2].ToString();
                string mobile = excelData.Rows[row][3].ToString();
                string greeceRegion = excelData.Rows[row][4].ToString();
                string apofasiNumber = excelData.Rows[row][5].ToString();
                string employeeName = excelData.Rows[row][6].ToString();
                string passportNumber = excelData.Rows[row][7].ToString();
                string passportExp = excelData.Rows[row][8].ToString();

                driver.Navigate().GoToUrl(txtWebsiteUrl.Text);

                var wait = new WebDriverWait(driver, TimeSpan.FromSeconds(10));

                // Login
                wait.Until(d => d.FindElement(By.Id("name"))).Clear();
                driver.FindElement(By.Id("name")).SendKeys(username);
                driver.FindElement(By.Id("password")).SendKeys(password);
                driver.FindElement(By.ClassName("i-login-g")).Click();

                // small delay replaced with wait for something that indicates login complete
                await Task.Delay(500, token); // still cancellable

                if (token.IsCancellationRequested)
                    return "Cancelled";

                string pageSource = driver.PageSource;
                if (pageSource.Contains("Invalid password"))
                {
                    return "Invalid Password or no slot";
                }

                // Click available view
                var availableBtn = wait.Until(d => d.FindElement(By.CssSelector("ul.ts li[onclick=\"switch_view('free')\"]")));
                try
                {
                    availableBtn.Click();
                }
                catch (ElementNotInteractableException)
                {
                    ((IJavaScriptExecutor)driver).ExecuteScript("arguments[0].click();", availableBtn);
                }

                TryBookAppointmentAsync(driver,wait,phone,mobile,greeceRegion,apofasiNumber,employeeName,passportNumber,passportExp,token).GetAwaiter();

                //var slotRows = wait.Until(d => d.FindElements(By.CssSelector("table.table tbody tr")));
                //if (slotRows.Count == 0)
                //    return "No slots found";

                //bool submitted = false;
                //foreach (var slot in slotRows)
                //{
                //    if (token.IsCancellationRequested)
                //        return "Cancelled";

                //    try
                //    {
                //        slot.Click();
                //        await Task.Delay(500, token); // small wait for UI

                //        // Fill form fields
                //        wait.Until(d => d.FindElement(By.Id("reservation_phone"))).SendKeys(phone);
                //        var mobileEl = driver.FindElement(By.Id("reservation_mobile"));
                //        mobileEl.Clear();
                //        mobileEl.SendKeys(mobile);

                //        var createBtn = wait.Until(d => d.FindElement(By.XPath("//button[contains(text(), 'Create appointment')]")));
                //        ((IJavaScriptExecutor)driver).ExecuteScript("arguments[0].scrollIntoView(true);", createBtn);
                //        await Task.Delay(500, token);
                //        createBtn.Click();

                //        submitted = true;
                //        break;
                //    }
                //    catch
                //    {
                //        // try next slot
                //    }
                //}

                //if (!submitted)
                //    return "Could not submit appointment";

                //// Fill additional details
                ////driver.FindElement(By.Id("form_1")).Click();
                ////driver.FindElement(By.XPath("//div[text()='"+greeceRegion+"']")).Click();
                //await Task.Delay(500, token);
                //IWebElement selectElem = wait.Until(d => d.FindElement(By.Id("form_1")));
                //var select = new SelectElement(selectElem);
                //select.SelectByText(greeceRegion);

                //driver.FindElement(By.Id("form_2")).SendKeys(apofasiNumber);
                //driver.FindElement(By.Id("form_3")).SendKeys(employeeName);
                //driver.FindElement(By.Id("form_4")).SendKeys(passportNumber);
                //driver.FindElement(By.Id("form_5")).SendKeys(passportExp);


                //driver.FindElement(By.Name("form_commit")).Click();
                await Task.Delay(500, token);

                driver.Navigate().GoToUrl("https://www.supersaas.com/users/logout/Saimways?form=form_2&return=Work");

                return "Success";
            }
            catch (OperationCanceledException)
            {
                return "Cancelled";
            }
            catch (Exception ex)
            {
                return ex.Message;
            }
        }

        async Task<string> TryBookAppointmentAsync(IWebDriver driver, WebDriverWait wait, string phone, string mobile,
string greeceRegion, string apofasiNumber, string employeeName, string passportNumber, string passportExp,
CancellationToken token)
        {
            bool submitted = false;

            while (!submitted)
            {
                var slotRows = driver.FindElements(By.CssSelector("table.table tbody tr"));
                if (!slotRows.Any())
                    return "No slots available";

                foreach (var slot in slotRows)
                {
                    if (token.IsCancellationRequested)
                        return "Cancelled";

                    try
                    {
                        Console.WriteLine("Trying slot, current URL before click: " + driver.Url);
                        slot.Click();
                        Thread.Sleep(500);
                        
                        Console.WriteLine("After slot click URL: " + driver.Url);

                        // Fill fields
                        wait.Until(d => d.FindElement(By.Id("reservation_phone"))).Clear();
                        wait.Until(d => d.FindElement(By.Id("reservation_phone"))).SendKeys(phone);

                        var mobileEl = wait.Until(d => d.FindElement(By.Id("reservation_mobile")));
                        mobileEl.Clear();
                        mobileEl.SendKeys(mobile);

                        var createAppointBtn = wait.Until(d => d.FindElement(By.XPath("//button[contains(text(), 'Create appointment')]")));
                        ((IJavaScriptExecutor)driver).ExecuteScript("arguments[0].scrollIntoView(true);", createAppointBtn);
                        createAppointBtn.Click();


                        string pageSource = driver.PageSource;

                        if (pageSource.Contains("An error prohibited this appointment"))
                        {
                            Console.WriteLine("Form submission error, closing dialog...");
                            driver.FindElement(By.XPath("//a[@onclick=\"hideDialog('reservation')\"]")).Click();
                            await Task.Delay(500, token);
                            continue; // Break inner loop, re-fetch slotRows
                        }

                        // Fill next form
                        var selectElem = wait.Until(d => d.FindElement(By.Id("form_1")));
                        var select = new SelectElement(selectElem);
                        select.SelectByText(greeceRegion);

                        driver.FindElement(By.Id("form_2")).Clear();
                        driver.FindElement(By.Id("form_2")).SendKeys(apofasiNumber);
                        driver.FindElement(By.Id("form_3")).Clear();
                        driver.FindElement(By.Id("form_3")).SendKeys(employeeName);
                        driver.FindElement(By.Id("form_4")).Clear();
                        driver.FindElement(By.Id("form_4")).SendKeys(passportNumber);
                        driver.FindElement(By.Id("form_5")).Clear();
                        driver.FindElement(By.Id("form_5")).SendKeys(passportExp);

                        driver.FindElement(By.Name("form_commit")).Click();

                        
                        pageSource = driver.PageSource;
                        if (pageSource.Contains("An error prohibited this appointment"))
                        {
                            driver.FindElement(By.PartialLinkText("Cancel")).Click();
                            slotRows = driver.FindElements(By.CssSelector("table.table tbody tr"));
                            continue; // Try next slot (refetch)
                        }

                        if (pageSource.Contains("No available space found") || pageSource.Contains("error"))
                        {
                            Console.WriteLine("Booking failed, trying next.");
                            continue;
                        }

                        if (pageSource.Contains("Appointment confirmed"))
                        {
                            Console.WriteLine("Appointment successfully submitted. Final URL: " + driver.Url);
                            submitted = true;
                            break;
                        }
                    }
                    catch (StaleElementReferenceException)
                    {
                        Console.WriteLine("DOM updated. Will retry.");
                        break;
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Unexpected error on slot: {ex.Message}");
                        continue;
                    }
                }

                // Optional wait between full slot refresh cycles
                await Task.Delay(1000, token);
            }

            return submitted ? "Submitted" : "No slot succeeded";
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
            if (excelData == null)
            {
                MessageBox.Show("Please upload an Excel file first.");
                return;
            }

            isProcessing = true;
            btnStop.Enabled = true;
            cancellationTokenSource = new CancellationTokenSource();

            try
            {
                for (int i = 0; i < excelData.Rows.Count; i++)
                {
                    if (cancellationTokenSource.Token.IsCancellationRequested)
                        break;

                    var status = await ChromeJobAsync(i, cancellationTokenSource.Token);
                    excelData.Rows[i]["Status"] = status;
                    dataGridView1.Refresh();
                    if (status == "Cancelled")
                        break;
                }
            }
            finally
            {
                driver?.Quit();
                isProcessing = false;
                btnStop.Enabled = false;
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
