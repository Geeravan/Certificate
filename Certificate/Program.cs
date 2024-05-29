using ClosedXML.Excel;
using DocumentFormat.OpenXml.Packaging;
using System.Text;

namespace CertificateGenerator
{
    class Certificate
    {
        static void Main(string[] args)
        {
            string excelFilePath = @"C:\Users\prema\OneDrive\Desktop\ZertifikateTest\Anwesenheit.xlsx";         //Pfad Exel ändern
            string templateFilePath = @"C:\Users\prema\OneDrive\Desktop\ZertifikateTest\Vorlage.docx";          //Pfad Word ändern
            string outputDirectory = @"C:\Users\prema\OneDrive\Desktop\ZertifikateTest\output";                 //Pfad Zielordner
            
            if (!Directory.Exists(outputDirectory))
            {
                Directory.CreateDirectory(outputDirectory);
            }
            
            using (var workbook = new XLWorkbook(excelFilePath))
            {
                var worksheet = workbook.Worksheet(1);                                                          // Nummer des Arbeitsblattes angeben
                var rows = worksheet.RangeUsed().RowsUsed();

                foreach (var row in rows.Skip(1))
                {
                    string vorname = row.Cell(1).GetValue<string>();
                    string nachname = row.Cell(2).GetValue<string>();
                    var projektleistungen = new StringBuilder();


                    for (int i = 3; i <= row.CellCount(); i++)                                                     // Excel Tabelle muss immer oben links sein
                    {
                        string project = worksheet.Cell(1, i).GetValue<string>();
                        string value = row.Cell(i).GetValue<string>();
                        if (value.Equals("Ja", StringComparison.OrdinalIgnoreCase))
                        {
                            projektleistungen.AppendLine("• " + project + '\n');
                        }
                    }
                    
                    if (projektleistungen.Length > 0)
                    {
                        string outputFilePath = Path.Combine(outputDirectory, $"{vorname}_{nachname}.docx");        //Filename
                        File.Copy(templateFilePath, outputFilePath, true);
                        using (var wordDoc = WordprocessingDocument.Open(outputFilePath, true))
                        {
                            string docText = null;
                            using (StreamReader sr = new StreamReader(wordDoc.MainDocumentPart.GetStream()))
                            {
                                docText = sr.ReadToEnd();
                            }

                            docText = docText.Replace("Vorname", vorname);
                            docText = docText.Replace("Nachname", nachname);
                            docText = docText.Replace("Projektleistungen", projektleistungen.ToString());

                            using (StreamWriter sw = new StreamWriter(wordDoc.MainDocumentPart.GetStream(FileMode.Create)))
                            {
                                sw.Write(docText);
                            }
                        }
                    }
                }
            }

            Console.WriteLine("Zertifikate wurden erfolgreich erstellt.");
        }
    }
}