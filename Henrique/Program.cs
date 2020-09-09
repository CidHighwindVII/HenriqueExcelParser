using System.IO;
using System.Threading;
using System.Windows.Forms;
using System.Linq;
using System.Collections.Generic;

namespace Henrique
{
    public class Program
    {
        static void Main(string[] args)
        {
            ReadCsvFile();
            //ReadExcelFile2();
        }

        private static void ReadCsvFile()
        {
            string filePath = string.Empty;
            Thread t = new Thread((ThreadStart)(() =>
            {
                using (OpenFileDialog openFileDialog = new OpenFileDialog())
                {
                    openFileDialog.InitialDirectory = "c:\\";
                    openFileDialog.Filter = "Csv Files|*.csv";
                    openFileDialog.FilterIndex = 1;
                    openFileDialog.RestoreDirectory = true;

                    if (openFileDialog.ShowDialog() == DialogResult.OK)
                    {
                        filePath = openFileDialog.FileName;
                    }
                }
            }));

            // Run your code from a thread that joins the STA Thread
            t.SetApartmentState(ApartmentState.STA);
            t.Start();
            t.Join();

            List<List<string>> table = new List<List<string>>();
            if (!string.IsNullOrEmpty(filePath))
            {
                var lines = File.ReadLines(filePath);
                
                if(lines.Count() > 0) {
                    foreach (string line in lines)
                    {
                        string[] lineValues = line.Split(';');

                        if (lineValues[4] != null && !string.IsNullOrEmpty(lineValues[4])){
                            List<string> row = new List<string>();

                            if (lineValues.Count() > 10)
                            {
                                List<string> list = lineValues.ToList();
                                list[6] += "," + list[7];
                                list[8] += "," + list[9];
                                list[11] += "," + list[12];
                                list.RemoveAt(12);
                                list.RemoveAt(9);
                                list.RemoveAt(7);
                                lineValues = list.ToArray();
                            }

                            foreach (string value in lineValues)
                            {
                                row.Add(value.Replace(';', ','));
                            }

                            table.Add(row);
                        } 
                        else
                        {
                            table[table.Count - 1][1] += " " + lineValues[1];
                            table[table.Count - 1][2] += " " + lineValues[2];
                        }
                    }
                }
            }

            table[0].Aggregate((a, b) => a + ";" + b);

            File.WriteAllLines(filePath.Replace(".csv", "fixed.csv"), table.Select(row => row.Aggregate((a, b) => a + ";" + b)));
        }
    }
}
