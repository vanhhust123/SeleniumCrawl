using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.IO;
using OpenQA.Selenium;
using HtmlAgilityPack;
using System.Net;
namespace SeleniumHaiQuan
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.OutputEncoding = Encoding.UTF8;
            Console.InputEncoding = Encoding.UTF8;
            string Url = @"https://www.customs.gov.vn/Lists/ThongKeHaiQuan/SoLieuDinhKy.aspx?Group=S%u1ed1+li%u1ec7u+th%u1ed1ng+k%u00ea";
            Extract extract = new Extract(Url);
            extract.Navigate();
           
            int pdfs = 0;
            var DescribeTextFile = File.AppendText(AppContext.BaseDirectory + @"\Data\Describe.txt");
            foreach (var item in extract.ExtractAll())
            {
                //Console.WriteLine(item[0] + "\t" + item[1] + "\t" + item[2]);
                // Item 0: Mô tả 
                // Item 1: Ngày
                // Item 2: Link
                // Datatype: string
                if (item[2] == "0")
                {
                    continue;
                }
                else
                {
                   
                    using (WebClient client = new WebClient())
                    {
                        try
                        {
                            client.DownloadFile(item[2], AppContext.BaseDirectory + @"\Data\PDFs\" + String.Format("Data{0}.pdf", pdfs));
                            
                            var line = item[0] + "," + item[1].ToString()+","+pdfs.ToString();
                            DescribeTextFile.WriteLine(line);
                            Console.WriteLine(line);
                            Console.Write("Đã ghi {0}", pdfs);
                            pdfs++;
                        }
                        catch
                        {
                            continue;
                        }
                    }
                    Console.WriteLine();
                }
                Console.WriteLine();
            }
            DescribeTextFile.Dispose();
            Console.ReadKey();

        }
    }

    // Class này dùng để trích xuất trang web Thuế Hải Quan.
    public class Extract
    {
        private int Options;
        public string Url;
        public IWebDriver webDriver;
        public HtmlDocument htmlDoc;

        //Contructer
        public Extract(string Url)
        {
            this.webDriver = new OpenQA.
                Selenium.Chrome.
                ChromeDriver(AppContext.BaseDirectory);
            this.Url = Url;
            htmlDoc = new HtmlDocument();
            Options = 0;
        }

        // Điều hướng trình duyệt đến địa chỉ
        public void Navigate()
        {
            this.webDriver.Url = this.Url;
            this.webDriver.Navigate();
           
            Options = webDriver.FindElements(
                By.XPath("/html/body/form/div[3]/div[4]/div[1]/div[3]/div[2]/table/tbody/tr/td[3]/select/option")).
                Count;
            Console.WriteLine("Done Navigate");
        }

        // Output của hàm này là list<tr> mỗi khi trình duyệt giả lập được load lại khi bấm nút kế tiếp.
        public IEnumerable<HtmlNode> ExtractRows(IWebDriver web, HtmlDocument html)
        {
            html = new HtmlDocument();
            html.LoadHtml(web.PageSource);
            var Table = html.
                DocumentNode.
                SelectNodes("/html/body/form/div[3]/div[4]/div[1]/div[3]/div[3]/table/tbody")[0];
            return Table.Elements("tr");
        }

        // Input của hàm này là List<tr> khi load trang mới, output List<[Nội dung, Ngày, Link Pdf]> 
        // của mỗi lần load 
        public IEnumerable<string[]> ExtractLinkPDF(IEnumerable<HtmlNode> nodes)
        {
            List<string[]> list = new List<string[]>();
            var Rows = nodes.ToList();
            foreach (var Row in Rows)
            {
                if (Rows.IndexOf(Row) == 0)
                {
                    continue;
                }
                else
                {
                    var Tds = Row.Elements("td").ToList();
                    string[] info = new string[3];
                    info[0] = Tds[0].InnerText;
                    info[1] = Tds[4].InnerText;
                    try
                    {
                        var link = Tds[4].Element("a");
                        info[2] = link.Attributes["href"].Value;
                    }
                    catch
                    {
                        info[2] = "0";
                    }
                    list.Add(info);
                    yield return info;
                }
            }

        }

        // Hàm trích xuất hàng loạt khi load liên tục web
        public IEnumerable<string[]> ExtractAll()
        {

            for (int i = 0; i < this.Options; i++)
            {
                if (i == 0)
                {
                    foreach (var item in this.ExtractLinkPDF(this.ExtractRows(this.webDriver, this.htmlDoc)))
                    {
                        yield return item;
                    }
                }
                else
                {
                    this.webDriver.
                        FindElement(By.XPath("/html/body/form/div[3]/div[4]/div[1]/div[3]/div[2]/table/tbody/tr/td[4]/a[1]")).
                        Click();
                 
                    foreach (var item in this.ExtractLinkPDF(this.ExtractRows(this.webDriver, this.htmlDoc)))
                    {
                        yield return item;
                    }

                }
            }
        }

    }
}
