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


        public ExcelOperator(string filePath,string rootWords)
        {
            this.FilePath = filePath;
            this.RootWords = rootWords.Split(',').ToList();
        }

        public void Open()//打开一个Microsoft.Office.Interop.Excel文件
        {
            app = new Microsoft.Office.Interop.Excel.Application();
            wbs = app.Workbooks;
            wb = wbs.Add(FilePath);
            ws = (Microsoft.Office.Interop.Excel.Worksheet)wb.Worksheets["Sheet0"];

            this.RowCount = ws.UsedRange.Rows.Count;
            this.ColumnCount = ws.UsedRange.Columns.Count;
            FileName = FilePath;
        }

        public void Read()
        {
            OriginalWords = new List<OriginalWord>();
            Array rangeValue = (this.ws.Range[ws.Cells[2, 1], ws.Cells[this.RowCount, 3]]).Value2 as Array;
            //Microsoft.Office.Interop.Excel.Range range2 = this.ws.Range[ws.Cells[2, 3], ws.Cells[this.RowCount, 3]];

            for (int index = 1; index < RowCount-1; index++)
            {
                OriginalWords.Add(new OriginalWord(){Word = rangeValue.GetValue(index,1).ToString(),DailySearchCount = int.Parse(rangeValue.GetValue(index,3).ToString())});
            }
        }

        public void ApartOriginalWords()
        {
            foreach (string rootWord in RootWords)
            {
                int rootWordLength = rootWord.Length;
                if(rootWordLength==0)
                    break;
                foreach (OriginalWord originalWord in OriginalWords)
                {
                    int originalWordLength = originalWord.Word.Length;
                    originalWord.RootWord = rootWord;

                    if (originalWordLength > 0 && originalWord.Word.Contains(rootWord))
                    {
                        int index = originalWord.Word.IndexOf(rootWord);
                        if (index == 0)
                        {
                            originalWord.PrefixWord = "";
                            originalWord.SuffixWord = originalWord.Word;
                        }else if (originalWordLength - index == rootWordLength)
                        {
                            originalWord.PrefixWord = originalWord.Word;
                            originalWord.SuffixWord = "";
                        }
                        else
                        {
                            originalWord.PrefixWord = originalWord.Word.Substring(0,index+rootWordLength);
                            originalWord.SuffixWord = originalWord.Word.Substring(index);
                        }
                    }
                }

                OriginalWords = OriginalWords.Where(o => o.PrefixWord != null&&o.SuffixWord!=null).ToList();
            }
        }

        public void CountSearch()
        {

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
