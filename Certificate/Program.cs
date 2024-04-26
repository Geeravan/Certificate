using System;
using Xceed.Document.NET;
using Xceed.Words.NET;

class Program
{
    static void Main()
    {
        string filePath = @"\Users\prema\OneDrive\Desktop\ZertifikateTest\TN-Zertifikat_Vorlage.docx";  //PFAD ändern
        if (!System.IO.File.Exists(filePath))
        {
            Console.WriteLine("Die Datei existiert nicht. Bitte überprüfe den Pfad und versuche es erneut.");
            return;
        }

        using (var document = DocX.Load(filePath))
        {


            Console.WriteLine("Gib den Voramen ein:");

            string name = Console.ReadLine();

            Console.WriteLine("Gib den Nachnamen ein:");
            string nachname = Console.ReadLine();

            Console.WriteLine("Gib die neuen Projekte ein (durch Komma getrennt):");
            string projectsInput = Console.ReadLine();
            string[] projects = projectsInput.Split(',');
            string projectsWithLineBreaks = string.Join("\n", projects);

            document.ReplaceText("Name", name);
            document.ReplaceText("Nachname", nachname);
            document.ReplaceText("Projekte", projectsWithLineBreaks.Trim());

            Console.WriteLine("Gib den Namen für die neue Datei ein (ohne die Dateiendung .docx):");

            string newFileName = Console.ReadLine();
            string newFilePath = $@"\Users\prema\OneDrive\Desktop\ZertifikateTest\{newFileName}.docx";  // PFAD ÄNDERN

            string directoryPath = Path.GetDirectoryName(newFilePath);
            if (!Directory.Exists(directoryPath))
            {
                Directory.CreateDirectory(directoryPath);
            }
            document.SaveAs(newFilePath);

            Console.WriteLine("Daten wurden erfolgreich in das Word-Dokument eingetragen!");
        }
    } 
}

