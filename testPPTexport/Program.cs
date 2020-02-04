using Spire.Presentation;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace testPPTexport
{
    internal class Program
    {
        private static void Main(string[] args)
        {
            Presentation ppt = new Presentation();
            ppt.SlideSize.Type = SlideSizeType.Screen16x9;

            double[] widths = new double[] { 100, 100, 100, 100, 100 };
            double[] heights = new double[] { 15, 15, 15, 15, 15 };
            ITable table = ppt.Slides[0].Shapes.AppendTable(80, 80, widths, heights);

            table.StylePreset = TableStylePreset.LightStyle1Accent2;

            string[,] data = new string[,]
            {
                {"排名","姓名", "銷售額","回款額","工號"},
                {"1","李彪","18270","18270","0011"},
                {"2","李娜","18105","18105","0025"},
                {"3","張麗","17987","17987","0008"},
                {"4","黃豔","17790","17790","0017"},
            };

            for (int i = 0; i < 5; i++)
            {
                for (int j = 0; j < 5; j++)
                {
                    table[j, i].TextFrame.Text = data[i, j];
                    table[j, i].TextFrame.Paragraphs[0].TextRanges[0].LatinFont = new TextFont("Arial");
                }
            }

            ppt.SaveToFile("建立表格.pptx", FileFormat.Pptx2010);
        }
    }
}