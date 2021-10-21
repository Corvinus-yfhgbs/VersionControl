using SummerOlympic.Entities;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using System.Reflection;

namespace SummerOlympic
{
    public partial class Form1 : Form
    {
        List<OlympicResult> results = new List<OlympicResult>();
        Excel.Application xlApp;
        Excel.Workbook xlWB;
        Excel.Worksheet xlSheet;
        public Form1()
        {
            InitializeComponent();
            LoadData("Summer_olympic_Medals.csv");
            CreateYearFilter();
            LoadPosition();
        }


        private void LoadData(string fileName)
        {
            using (var sr = new StreamReader(fileName, Encoding.Default))
            {
                sr.ReadLine();
                while (!sr.EndOfStream)
                {
                    var line = sr.ReadLine().Split(',');
                    var or = new OlympicResult()
                    {
                        Year = int.Parse(line[0]),
                        Country = line[3],
                        Medals = new int[]
                        {
                            int.Parse(line[5]),
                            int.Parse(line[6]),
                            int.Parse(line[7])
                        }
                    };

                    results.Add(or);
                }

            }

        }

        private void CreateYearFilter()
        {
            var years = (from r in results
                         orderby r.Year
                         select r.Year).Distinct();

            comboBox1.DataSource = years.ToList();
        }

        private int CalculatePosition(OlympicResult Oresult)
        {
            var betterResults = 0;
            var filteredResults = (from r in results
                             where r.Year == Oresult.Year & r.Country != Oresult.Country
                             select r);

            foreach (var result in filteredResults)
            {
                if (Oresult.Medals[0] < result.Medals[0])
                {
                    betterResults++;
                }
                else if (Oresult.Medals[0] == result.Medals[0])
                {
                    if (Oresult.Medals[1] < result.Medals[1])
                    {
                        betterResults++;
                    }
                    else if (Oresult.Medals[1] == result.Medals[1])
                    {
                        if (Oresult.Medals[2] < result.Medals[2])
                        {
                            betterResults++;
                        }
                    }
                }
            }
            return betterResults + 1;
        }

        private void LoadPosition()
        {
            foreach (var result in results)
            {
                result.Position = CalculatePosition(result);
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                xlApp = new Excel.Application();
                xlWB = xlApp.Workbooks.Add(Missing.Value);
                xlSheet = xlWB.ActiveSheet;

                CreateExcel();

                xlApp.Visible = true;
                xlApp.UserControl = true;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                xlWB.Close(false, Type.Missing, Type.Missing);
                xlApp.Quit();
                xlWB = null;
                xlApp = null;
            }
        }


        private void CreateExcel()
        {
            var headers = new string[]
            {
                "Helyezés",
                "Ország",
                "Arany",
                "Ezüst",
                "Bronz"
            };

            for (int i = 0; i < headers.Length; i++)
            {
                xlSheet.Cells[1, i + 1] = headers[i];
            }

            var filteredResult = (from r in results
                                  where r.Year == (int)comboBox1.SelectedItem
                                  orderby r.Position
                                  select r);
            var counter = 2;
            foreach (var result in filteredResult)
            {
                xlSheet.Cells[counter, 1] = result.Position;
                xlSheet.Cells[counter, 2] = result.Country;
                for (int i = 0; i <= 2; i++)
                {
                    xlSheet.Cells[counter, i + 3] = result.Medals[i];
                }
                counter++;
            }
        }
    }
}
