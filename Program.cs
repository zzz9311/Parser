using ClosedXML.Excel;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using RevocationParser.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;

namespace RevocationParser
{
    class Program
    {
        public static IWebDriver Driver = new ChromeDriver(AppDomain.CurrentDomain.BaseDirectory);
        public static List<ParseModel> ParsedInfo = new List<ParseModel>();
        static void Main(string[] args)
        {
            int ID;
            string fileName = "Parsed.xlsx";
            using (var excelWorkbook = new XLWorkbook(fileName))
            {
                var nonEmptyDataRows = excelWorkbook.Worksheet(1).RowsUsed();

                foreach (var dataRow in nonEmptyDataRows)
                {
                    if (dataRow.RowNumber() < 2)
                    {
                        continue;
                    }
                    var cell = dataRow.Cell(2).Value;
                    ID = Convert.ToInt32(dataRow.Cell(2).Value);
                    var Yandex = dataRow.Cell(3).Value;
                    var Ozon = dataRow.Cell(4).Value;
                    var Otzovik = dataRow.Cell(5).Value;
                    var IRecommend = dataRow.Cell(6).Value;
                    if (Yandex.ToString().Contains("yandex"))
                    {
                        ParsedInfo.AddRange(ParseYandex(Yandex.ToString(), ID));
                    }
                    if (Ozon.ToString().Contains("ozon"))
                    {
                        ParsedInfo.AddRange(ParseOzon(Ozon.ToString(), ID));
                    }
                    if (Otzovik.ToString().Contains("otzovik"))
                    {
                        ParsedInfo.AddRange(ParseOrzovik(Otzovik.ToString(), ID));
                    }
                    if (IRecommend.ToString().Contains("irecommend"))
                    {
                        ParsedInfo.AddRange(ParseIRecommend(IRecommend.ToString(), ID));
                    }
                }
                Driver.Close();
                Save();
                Console.ReadKey();

            }
        }

        public static void Change()
        {
            string fileName = "02.09.2021 - 170.xlsx";
            using (var excelWorkbook = new XLWorkbook(fileName))
            {
                var nonEmptyDataRows = excelWorkbook.Worksheet(1).RowsUsed();
                foreach (var dataRow in nonEmptyDataRows)
                {
                    if (dataRow.RowNumber() < 2)
                    {
                        continue;
                    }
                    Console.WriteLine(dataRow.Cell(2).Value);
                    if (dataRow.Cell(2).Value!= "" && Convert.ToInt64(dataRow.Cell(2).Value)==0)
                    {
                        dataRow.Cell(2).Value = "";
                    }
                    if ((string)dataRow.Cell(6).Value == "Нет комментария")
                    {
                        dataRow.Cell(6).Value = "";
                    }
                    if ((string)dataRow.Cell(7).Value == "Нет плюсов")
                    {
                        dataRow.Cell(7).Value = "";
                    }
                    if ((string)dataRow.Cell(8).Value == "Нет минусов")
                    {
                        dataRow.Cell(8).Value = "";
                    }
                }
                excelWorkbook.SaveAs($"sss{DateTime.Now.ToShortDateString()} - {ParsedInfo.Count}.xlsx");
            }
        }

        public static void Save()
        {
            IXLWorkbook wb = new XLWorkbook();
            IXLWorksheet ws = wb.Worksheets.Add("Sample Sheet");

            ws.Cell(1, 1).Value = "ID товара";
            ws.Cell(1, 2).Value = "id";
            ws.Cell(1, 3).Value = "date";
            ws.Cell(1, 4).Value = "author";
            ws.Cell(1, 5).Value = "rating";
            ws.Cell(1, 6).Value = "description";
            ws.Cell(1, 7).Value = "pros";
            ws.Cell(1, 8).Value = "cons";
            ws.Cell(1, 9).Value = "link";
            int Counter = 2;
            foreach (var el in ParsedInfo)
            {
                ws.Cell(Counter, 1).Value = el.ID;
                ws.Cell(Counter, 2).Value = el.id == 0? "":el.id.ToString();
                ws.Cell(Counter, 3).Value = el.date;
                ws.Cell(Counter, 4).Value = el.author;
                ws.Cell(Counter, 5).Value = el.rating;
                ws.Cell(Counter, 6).Value = el.description ?? "";
                ws.Cell(Counter, 7).Value = el.pros ?? "";
                ws.Cell(Counter, 8).Value = el.cons ?? "";
                ws.Cell(Counter, 9).Value = el.link;
                Counter++;
            }
            wb.SaveAs($"ozon{DateTime.Now.ToShortDateString()} - {ParsedInfo.Count}.xlsx");
        }
        public static List<ParseModel> ParseOzon(string path, int id) // main
        {
            int j = 0;
            try
            {
                Driver.Navigate().GoToUrl(path);
                Thread.Sleep(8000);
            }
            catch (Exception)
            {

            }
            Thread.Sleep(1000);
            for (int i = 0; i < 500; i++)
            {
                //window.scrollTo({0}, {1})
                ((IJavaScriptExecutor)Driver).ExecuteScript($"window.scrollTo(4100, 4100)");
                j = j + 1500;
            }
            IWebElement Element = null;
            List<IWebElement> Elements = new List<IWebElement>();
            try
            {
                Element = Driver.FindElement(By.ClassName("ga9"));
                Elements = Element.FindElements(By.ClassName("gb4")).ToList();
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error");
            }
            return ParserOzonComment(Elements, path, id);
        }

        public static List<ParseModel> ParserOzonComment(List<IWebElement> Comments, string path, int id)
        {
            List<ParseModel> ParsedModels = new List<ParseModel>();
            foreach (var el in Comments)
            {
                ParseModel ParsedComment = new ParseModel();
                ParsedComment.ID = id;
                ParsedComment.id = Convert.ToInt32(el.FindElement(By.ClassName("e2q9")).GetAttribute("data-review-id"));
                ParsedComment.link = path;
                ParsedComment.date = el.FindElement(By.ClassName("e2v8")).Text;
                ParsedComment.author = el.FindElement(By.ClassName("e2w5")).Text;
                ParsedComment.rating = GetRatingFromOzon(el.FindElement(By.ClassName("_3xol")).GetCssValue("width"));
                var CommentDescription = el.FindElements(By.ClassName("e2u6"));
                var AboutDescription = el.FindElements(By.ClassName("e2u7"));
                for (int i = 0; i < AboutDescription.Count; i++)
                {
                    if (AboutDescription[i].Text.Equals("Достоинства"))
                    {
                        ParsedComment.pros = CommentDescription[i].Text;
                    }
                    else if (AboutDescription[i].Text.Equals("Недостатки"))
                    {
                        ParsedComment.cons = CommentDescription[i].Text;
                    }
                    else if (AboutDescription[i].Text.Equals("Комментарий"))
                    {
                        ParsedComment.description = CommentDescription[i].Text;
                    }
                }
                ParsedModels.Add(ParsedComment);
            }
            return ParsedModels;
        }

        public static int GetRatingFromOzon(string rating)
        {
            var SubstrungRating = rating.Substring(0, rating.Length - 2);
            var IntRating = Convert.ToInt32(SubstrungRating);
            switch (IntRating)
            {
                case 90: return 5;
                case 72: return 4;
                case 54: return 3;
                case 36: return 2;
                case 18: return 1;
            }
            return 0;
        }


        public static List<ParseModel> ParseYandex(string path, int id)
        {
            try
            {
                Driver.Navigate().GoToUrl(path);
                Thread.Sleep(5000);
                var Element = Driver.FindElement(By.Id("scroll-to-reviews-list"));
                var Elements = Element.FindElements(By.CssSelector("div[data-zone-name='product-review']")).ToList();
                return ParserYandexComment(Elements, path, id);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
                return new List<ParseModel>();
            }
        } // main
        public static List<ParseModel> ParserYandexComment(List<IWebElement> Comments, string path, int id)
        {
            List<ParseModel> ParsedModels = new List<ParseModel>();
            foreach (var el in Comments)
            {
                ParseModel ParsedComment = new ParseModel();
                ParsedComment.ID = id;
                ParsedComment.id = GetNumbersFromText(el.GetAttribute("data-zone-data"));
                ParsedComment.link = path;
                try
                {
                    ParsedComment.date = el.FindElement(By.ClassName("kx7am")).Text.Split(',')[0];
                }
                catch (Exception)
                {

                    ParsedComment.date = "Без даты";
                }

                try
                {
                    ParsedComment.author = el.FindElement(By.CssSelector("div[data-zone-name='name']")).Text;
                }
                catch (Exception)
                {

                    ParsedComment.author = "Имя скрыто";
                }
                var Ratings = el.FindElement(By.ClassName("Ymh5F")).GetAttribute("data-rate");
                ParsedComment.rating = Convert.ToInt32(Ratings);

                var CommentDescription = el.FindElements(By.ClassName("_1yULR"));
                var AboutDescription = el.FindElements(By.ClassName("_272cj"));
                for (int i = 0; i < AboutDescription.Count; i++)
                {
                    if (AboutDescription[i].Text.Equals("Достоинства:"))
                    {
                        ParsedComment.pros = CommentDescription[i].Text;
                    }
                    else if (AboutDescription[i].Text.Equals("Недостатки:"))
                    {
                        ParsedComment.cons = CommentDescription[i].Text;
                    }
                    else if (AboutDescription[i].Text.Equals("Комментарий:"))
                    {
                        ParsedComment.description = CommentDescription[i].Text;
                    }
                }
                ParsedModels.Add(ParsedComment);
            }
            return ParsedModels;
        }
        public static int GetNumbersFromText(string text)
        {
            return Int32.Parse(Regex.Match(text, @"\d+").Value);
        }

        public static List<ParseModel> ParseOrzovik(string path, int id)
        {
            List<ParseModel> ParsedModels = new List<ParseModel>();
            List<string> Urls = new List<string>();
            Driver.Navigate().GoToUrl(path);
            var Elements = Driver.FindElements(By.ClassName("mshow0"));
            foreach (var el in Elements)
            {
                string Path = el.FindElement(By.ClassName("review-read-link")).GetAttribute("href");
                Urls.Add(Path);
            }
            foreach (var el in Urls)
            {
                ParsedModels.Add(OtzovikReviewParser(el, id));
                Thread.Sleep(1000);
            }
            return ParsedModels;
        } // main

        public static ParseModel OtzovikReviewParser(string path, int id)
        {
            Driver.Navigate().GoToUrl(path);
            ParseModel ParsedReview = new ParseModel();
            ParsedReview.id = 0;
            ParsedReview.ID = id;
            ParsedReview.link = path;
            ParsedReview.author = Driver.FindElement(By.XPath("/html/body/div[2]/div/div/div/div/div[3]/div[1]/div[1]/div[1]/a[2]/span")).Text;
            ParsedReview.date = Driver.FindElement(By.XPath("/html/body/div[2]/div/div/div/div/div[3]/div[1]/div[2]/span/span")).Text;
            ParsedReview.cons = Driver.FindElement(By.ClassName("review-minus")).Text;
            ParsedReview.description = Driver.FindElement(By.ClassName("description")).Text;
            ParsedReview.pros = Driver.FindElement(By.ClassName("review-plus")).Text;
            ParsedReview.rating = Convert.ToInt32(Driver.FindElement(By.XPath("//*[@id=\"content\"]/div/div/div/div/div[3]/div[1]/table/tbody[2]/tr[2]/td[2]/div")).GetAttribute("title").Replace("Общий рейтинг: ",""));
            return ParsedReview;
        }


        public static List<ParseModel> ParseIRecommend(string path, int id) // main
        {
            List<ParseModel> ParsedModels = new List<ParseModel>();
            try
            {
                Driver.Navigate().GoToUrl(path);
            }
            catch (Exception)
            {
                return ParsedModels;
            }
            List<string> Urls = new List<string>();
            var Elements = Driver.FindElement(By.ClassName("list-comments")).FindElements(By.ClassName("item"));
            foreach (var el in Elements.Where(i => !String.IsNullOrEmpty(i.Text)))
            {
                string Path = el.FindElement(By.ClassName("reviewTitle")).FindElement(By.TagName("a")).GetAttribute("href");
                Urls.Add(Path);
            }

            foreach (var el in Urls)
            {
                ParsedModels.Add(ParseIRecommendReview(el, id));
            }
            return ParsedModels;
        }

        public static ParseModel ParseIRecommendReview(string path, int id)
        {
            ParseModel ParsedReview = new ParseModel();
            try
            {
                Driver.Navigate().GoToUrl(path);
                Thread.Sleep(1000);
                ParsedReview.id = 0;
                ParsedReview.ID = id;
                ParsedReview.link = path;// made
                ParsedReview.author = Driver.FindElement(By.ClassName("reviewer")).Text;// made
                ParsedReview.date = Driver.FindElement(By.ClassName("dtreviewed")).Text;// made
                try
                {
                    ParsedReview.cons = Driver.FindElement(By.XPath("//*[@id=\"content\"]/div[3]/div[3]/div[4]/div[1]/ul")).Text;
                }
                catch (Exception)
                {
                    ParsedReview.cons = "";

                } //made
                try
                {
                    ParsedReview.pros = Driver.FindElement(By.XPath("//*[@id=\"content\"]/div[3]/div[3]/div[4]/div[2]/ul")).Text; // made
                }
                catch (Exception)
                {
                    ParsedReview.pros = "";
                } //made
                ParsedReview.description = Driver.FindElement(By.ClassName("hasinlineimage")).Text; // made
                ParsedReview.rating = Convert.ToInt32(Driver.FindElement(By.XPath("//*[@id=\"content\"]/div[3]/div[3]/div[1]/div[1]/div[2]/div[2]/meta[3]")).GetAttribute("content")); // made
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
            }

            return ParsedReview;
        }

    }
}
