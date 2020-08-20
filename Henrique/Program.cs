using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Common;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace Henrique
{
    public class Program
    {
        static void Main(string[] args)
        {
            ReadExcelFile();
            //ReadExcelFile2();
        }

        /// <summary>
        /// Criar uma linha na linha de comandos com o progresso
        /// </summary>
        /// <param name="progress">progresso</param>
        /// <param name="total">número total de elementos a processar</param>
        private static void drawTextProgressBar(int progress, int total)
        {
            //draw empty progress bar
            Console.CursorLeft = 0;
            Console.Write("["); //start
            Console.CursorLeft = 32;
            Console.Write("]"); //end
            Console.CursorLeft = 1;
            float onechunk = 30.0f / total;

            //draw filled part
            int position = 1;
            for (int i = 0; i < onechunk * progress; i++)
            {
                Console.BackgroundColor = ConsoleColor.Gray;
                Console.CursorLeft = position++;
                Console.Write(" ");
            }

            //draw unfilled part
            for (int i = position; i <= 31; i++)
            {
                Console.BackgroundColor = ConsoleColor.Green;
                Console.CursorLeft = position++;
                Console.Write(" ");
            }

            //draw totals
            Console.CursorLeft = 35;
            Console.BackgroundColor = ConsoleColor.Black;
            Console.Write("A processar página " + progress.ToString() + " de " + total.ToString() + "    "); //blanks at the end remove any excess
        }

        /// <summary>
        /// Se a linha do excel não tiver data passa os valores do artigo e decrição para a linha de cima
        /// </summary>
        /// <param name="worksheet">Pagina do Excel</param>
        /// <param name="row">Linha de onde vão passar valores</param>
        private static void ParseValues(ref Excel.Worksheet worksheet, ref int row)
        {
            if (worksheet.Cells[row, 1].value == null)
            {
                worksheet.Cells[row - 1, 3].value += worksheet.Cells[row, 3].value;
                worksheet.Cells[row - 1, 5].value += " " + worksheet.Cells[row, 5].value;

                //Tirar "Enters" a mais
                string type = worksheet.Cells[row - 1, 5].value.GetType().ToString();
                if (type == "System.String")
                {
                    worksheet.Cells[row - 1, 3].value = ((string)worksheet.Cells[row - 1, 3].value).Replace("\n","");
                    worksheet.Cells[row - 1, 5].value = ((string)worksheet.Cells[row - 1, 5].value).Replace("\n", "");
                }

                worksheet.Rows[row].Delete();
                row -= 1;
            }

            row += 1;
        }

        private static Excel.Worksheet GetSheetsUnion(ref Excel.Workbook excelWorkbook)
        {
            Excel.Worksheet result = excelWorkbook.Sheets[1];

            if (excelWorkbook.Sheets.Count > 1)
            {
                // Reformata a primeira página
                int row = 1;
                string artigo = result.Cells[row, 3].value;

                drawTextProgressBar(1, excelWorkbook.Sheets.Count);

                while (!string.IsNullOrEmpty(artigo))
                {
                    ParseValues(ref result, ref row);
                    try
                    {
                        artigo = (string)result.Cells[row, 3].value;
                    }
                    catch (Exception e)
                    {
                        double val = result.Cells[row, 3].value;
                        artigo = val.ToString();
                    }
                }

                drawTextProgressBar(2, excelWorkbook.Sheets.Count);

                // Reformata as restantes páginas e passa-las para a primeira página
                for (int i = 2; i <= excelWorkbook.Sheets.Count; i++)
                {
                    Excel.Worksheet nextSheet = excelWorkbook.Sheets[i];
                    int nextSheetRow = 2;

                    artigo = (string)nextSheet.Cells[nextSheetRow, 3].value;
                    while (!string.IsNullOrEmpty(artigo))
                    {
                        ParseValues(ref nextSheet, ref nextSheetRow);
                        try
                        {
                            artigo = (string)nextSheet.Cells[nextSheetRow, 3].value;
                        }
                        catch (Exception e)
                        {
                            double val = nextSheet.Cells[nextSheetRow, 3].value;
                            artigo = val.ToString();
                        }
                    }

                    // Faz a cópia das restantes páginas para a primeira
                    string fromCells = string.Format("A2:S{0}", nextSheetRow - 1);
                    string toCells = string.Format("A{0}:S{1}", row, row + nextSheetRow - 1);
                    nextSheet.Range[fromCells].Copy(result.Range[toCells]);
                    result.Range[toCells].Insert(Excel.XlInsertShiftDirection.xlShiftDown);

                    Excel.Range dest = result.Range[toCells];
                    nextSheet.Range[fromCells].Copy(dest);

                    Excel.Range copyFrom = nextSheet.Range[string.Format("A2:S{0}", nextSheetRow - 1)];

                    // Incrementa a linha onde estámos atualmente na primeira página
                    row += nextSheetRow;
                    drawTextProgressBar(i, excelWorkbook.Sheets.Count);
                }
            }

            return result;
        }

        private static void ReadExcelFile()
        {
            Console.WriteLine("Escolha um ficheiro a processar. (Pressione qualquer tecla para continuar)");
            Console.ReadLine();

            // Visitar site https://ourcodeworld.com/articles/read/890/how-to-solve-csharp-exception-current-thread-must-be-set-to-single-thread-apartment-sta-mode-before-ole-calls-can-be-made-ensure-that-your-main-function-has-stathreadattribute-marked-on-it 
            string filePath = string.Empty;
            Thread t = new Thread((ThreadStart)(() =>
            {
                using (OpenFileDialog openFileDialog = new OpenFileDialog())
                {
                    openFileDialog.InitialDirectory = "c:\\";
                    openFileDialog.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm";
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

            if (!string.IsNullOrEmpty(filePath))
            {
                Excel.Application excelApp = new Excel.Application();
                Excel.Workbook excelWorkbook = excelApp.Workbooks.Open(filePath);
                Excel.Worksheet excelSheet = GetSheetsUnion(ref excelWorkbook);
                
                // for(int i = excelWorkbook.Sheets.Count; i > 1; i--)
                // {
                //     ((Excel.Worksheet)excelWorkbook.Sheets[i]).Delete();
                // }

                try
                {
                    excelWorkbook.SaveAs(filePath.Replace(".xlsx", "Corrigido.xlsx"));
                }
                catch (Exception ex)
                {

                }
                finally
                {
                    excelApp.Quit();
                }
            }
        }

        private static void ReadExcelFile2()
        {
            string filePath = getFilePath();           

            using (OleDbConnection conn = new OleDbConnection(string.Format("provider=Microsoft.Jet.OLEDB.4.0;Data Source='{0}';Extended Properties=Excel 8.0;", filePath)))
            {
                conn.Open();
                DataTable dtSheet = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);

                for(int i = 0; i < dtSheet.Columns[2].Table.Rows.Count; i++)
                {
                    string pageName = (string)dtSheet.Rows[i][2];

                    using(OleDbDataAdapter command = new OleDbDataAdapter(string.Format("select * from [{0}]", pageName), conn))
                    {
                        var selectCommand = command.SelectCommand;
                        var reader = selectCommand.ExecuteReader();
                        while (reader.NextResult())
                        {
                            
                        }
                    }
                }
                //dtSheet.Rows[0][2];
                //dtSheet.Columns.List[2];
                //MyCommand = new System.Data.OleDb.OleDbDataAdapter("select * from [Sheet1$]", conn);
            }

            //MyCommand.TableMappings.Add("Table", "TestTable");
            //DtSet = new System.Data.DataSet();
            //MyCommand.Fill(DtSet);
        }

        private static string getFilePath()
        {
            string filePath = string.Empty;
            Thread t = new Thread((ThreadStart)(() =>
            {
                using (OpenFileDialog openFileDialog = new OpenFileDialog())
                {
                    openFileDialog.InitialDirectory = "c:\\";
                    openFileDialog.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm";
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

            return filePath;
        }
    }
}
