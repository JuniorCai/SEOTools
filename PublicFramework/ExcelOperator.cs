using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;

namespace PublicFramework
{
    public class ExcelOperator
    {
        public string FilePath { get; set; }

        public string SaveDirectoryPath { get; set; }

        public string FileName { get; set; }

        public int RowCount { get; set; }

        public int ColumnCount { get; set; }

        public List<OriginalWord> OriginalWords { get; set; }

        public List<string> RootWords { get; set; }

        public List<string> PrefixWords { get; set; }

        public List<string> SuffixWords { get; set; }


        public Microsoft.Office.Interop.Excel.Application app;
        public Microsoft.Office.Interop.Excel.Workbooks wbs;
        public Microsoft.Office.Interop.Excel.Workbook wb;
        public Microsoft.Office.Interop.Excel.Worksheets wss;
        public Microsoft.Office.Interop.Excel.Worksheet ws;


        public ExcelOperator(string directoryPath, string filePath, string rootWords)
        {
            this.FilePath = filePath;
            this.SaveDirectoryPath = directoryPath.LastIndexOf(@"\") == (directoryPath.Length - 1)
                ? directoryPath.Substring(0, directoryPath.Length - 1)
                : directoryPath;
            this.RootWords = rootWords.Split(',').ToList();
        }

        public void Open() //打开一个Microsoft.Office.Interop.Excel文件
        {
            app = new Microsoft.Office.Interop.Excel.Application();
            wbs = app.Workbooks;
            wb = wbs.Add(FilePath);
            ws = (Microsoft.Office.Interop.Excel.Worksheet) wb.Worksheets["Sheet1"];

            this.RowCount = ws.UsedRange.Rows.Count;
            this.ColumnCount = ws.UsedRange.Columns.Count;
            FileName = FilePath;
        }


        #region 分词

        public void Read()
        {
            OriginalWords = new List<OriginalWord>();
            Array rangeValue = (this.ws.Range[ws.Cells[1, 1], ws.Cells[this.RowCount, 3]]).Value2 as Array;
            //Microsoft.Office.Interop.Excel.Range range2 = this.ws.Range[ws.Cells[2, 3], ws.Cells[this.RowCount, 3]];

            for (int index = 1; index <= RowCount; index++)
            {
                OriginalWords.Add(new OriginalWord()
                {
                    Word = rangeValue.GetValue(index, 1).ToString(),
                    DailySearchCount = int.Parse(rangeValue.GetValue(index, 3).ToString())
                });
            }
            Close();
        }

        public List<OriginalWord> ApartOriginalWords()
        {
            foreach (string rootWord in RootWords)
            {
                int rootWordLength = rootWord.Length;
                if (rootWordLength == 0)
                    break;
                foreach (OriginalWord originalWord in OriginalWords)
                {
                    int originalWordLength = originalWord.Word.Length;
                    originalWord.RootWord = rootWord;

                    if (originalWordLength > 0 && originalWord.Word.Contains(rootWord))
                    {
                        int index = originalWord.Word.IndexOf(rootWord);
                        originalWord.PrefixWord = originalWord.Word.Substring(0, index + rootWordLength);
                        originalWord.SuffixWord = originalWord.Word.Substring(index);

                    }
                    else
                    {
                        string item = originalWord.Word;
                    }
                }

                OriginalWords = OriginalWords.Where(o => o.PrefixWord != null && o.SuffixWord != null).ToList();
            }
            return OriginalWords;
        }


        public void FlushApartResultToFile(List<OriginalWord> apartResult)
        {
            app = new Microsoft.Office.Interop.Excel.Application();
            wbs = app.Workbooks;
            wb = wbs.Add(true);

            Microsoft.Office.Interop.Excel.Worksheet ws =
                (Microsoft.Office.Interop.Excel.Worksheet) wb.Worksheets["Sheet1"];

            for (int i = 1; i <= apartResult.Count; i++)
            {
                ws.Cells[i, 1] = apartResult[i - 1].Word;
                ws.Cells[i, 2] = apartResult[i - 1].DailySearchCount;
                ws.Cells[i, 3] = apartResult[i - 1].PrefixWord;
                ws.Cells[i, 4] = apartResult[i - 1].SuffixWord;
                ws.Cells[i, 5] = apartResult[i - 1].RootWord;

            }
            SaveAs(RootWords[0] + "分词");
            Close();
        }


        #endregion


        #region 统计

        public void ReadApartResult()
        {
            OriginalWords = new List<OriginalWord>();
            Array rangeValue = (this.ws.Range[ws.Cells[1, 1], ws.Cells[this.RowCount, 5]]).Value2 as Array;
            //Microsoft.Office.Interop.Excel.Range range2 = this.ws.Range[ws.Cells[2, 3], ws.Cells[this.RowCount, 3]];

            for (int index = 1; index <= RowCount; index++)
            {
                OriginalWords.Add(new OriginalWord()
                {
                    Word = rangeValue.GetValue(index, 1).ToString(),
                    DailySearchCount = int.Parse(rangeValue.GetValue(index, 2).ToString()),
                    PrefixWord = rangeValue.GetValue(index, 3).ToString(),
                    SuffixWord = rangeValue.GetValue(index, 4).ToString(),
                    RootWord = rangeValue.GetValue(index, 5).ToString()
                });
            }
            Close();
        }


        public void CountSearch()
        {
            List<CountSearch> countResultList = new List<CountSearch>();

            var beforeQuery = from originalWord in OriginalWords
                group originalWord by originalWord.PrefixWord
                into g
                select new CountSearch()
                {
                    WordType = "Before",
                    FixWord = g.Key,
                    WordCount = g.Count(),
                    WordSearch = g.Sum(item => item.DailySearchCount)
                };

            foreach (CountSearch countSearch in beforeQuery)
            {
                var result = beforeQuery.Where(q => q.FixWord.Contains(countSearch.FixWord));
                if (result.Any())
                {
                    countResultList.Add(new CountSearch()
                    {
                        FixWord = countSearch.FixWord,
                        WordCount = result.Sum(r => r.WordCount),
                        WordSearch = result.Sum(r => r.WordSearch),
                        WordType = countSearch.WordType
                    });
                }
            }

          //  var beforeList = beforeQuery.ToList();

            var afterQuery = from originalWord in OriginalWords
                group originalWord by originalWord.SuffixWord
                into g
                select new CountSearch()
                {
                    WordType = "After",
                    FixWord = g.Key,
                    WordCount = g.Count(),
                    WordSearch = g.Sum(item => item.DailySearchCount)
                };
            foreach (CountSearch countSearch in afterQuery)
            {
                var result = afterQuery.Where(q => q.FixWord.Contains(countSearch.FixWord));
                if (result.Any())
                {
                    countResultList.Add(new CountSearch()
                    {
                        FixWord = countSearch.FixWord,
                        WordCount = result.Sum(r => r.WordCount),
                        WordSearch = result.Sum(r => r.WordSearch),
                        WordType = countSearch.WordType
                    });
                }
            }
            //var afterList = afterQuery.ToList();
            //beforeList.AddRange(afterList);
            //beforeList.Sort((a, b) => -a.WordSearch.CompareTo(b.WordSearch));
            countResultList.Sort((a, b) => -a.WordSearch.CompareTo(b.WordSearch));

            FlushToFile(countResultList);

        }

        private void FlushToFile(List<CountSearch> finalList)
        {
            app = new Microsoft.Office.Interop.Excel.Application();
            wbs = app.Workbooks;
            wb = wbs.Add(true);

            Microsoft.Office.Interop.Excel.Worksheet ws =
                (Microsoft.Office.Interop.Excel.Worksheet) wb.Worksheets["Sheet1"];

            for (int i = 1; i <= finalList.Count; i++)
            {
                ws.Cells[i, 1] = finalList[i - 1].WordType;
                ws.Cells[i, 2] = finalList[i - 1].FixWord;
                ws.Cells[i, 3] = finalList[i - 1].WordCount;
                ws.Cells[i, 4] = finalList[i - 1].WordSearch;

            }
            SaveAs("词缀统计");
            Close();
        }

        #endregion



        //文档另存为
        public bool SaveAs(object fileName)
        {
            try
            {
                wb.SaveAs(SaveDirectoryPath + @"\" + fileName, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    Type.Missing, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlExclusive, Type.Missing,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                return true;

            }

            catch (Exception ex)
            {
                return false;

            }
        }


        public void Close()
            //关闭一个Microsoft.Office.Interop.Excel对象，销毁对象
        {
            //wb.Save();
            wb.Close(Type.Missing, Type.Missing, Type.Missing);
            wbs.Close();
            app.Quit();
            wb = null;
            wbs = null;
            app = null;
            GC.Collect();
        }

    }

    public class OriginalWord
    {
        public string Word { get; set; }

        public int DailySearchCount { get; set; }

        public string PrefixWord { get; set; }

        public string SuffixWord { get; set; }

        public string RootWord { get; set; }
    }

    public class CountSearch
    {
        public string WordType { get; set; }

        public string FixWord { get; set; }

        public int WordCount { get; set; }

        public int WordSearch { get; set; }


    }
}
