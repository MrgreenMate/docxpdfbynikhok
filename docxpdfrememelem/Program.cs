using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace docxpdfrememelem
{
    internal class Program
    {
        static void Main(string[] args)
        {
            // Az aktuális mappa elérési útja, ahol a program fut
           string folderPath = AppDomain.CurrentDomain.BaseDirectory;
            ;

            // Betöltjük a Word alkalmazás COM objektumot
            Application wordApp = new Application();
            wordApp.Visible = false;

            try
            {
                // Végigmegyünk minden .docx fájlon a mappában
                foreach (string docxFile in Directory.GetFiles(folderPath, "*.docx"))
                {
                    string pdfFile = Path.ChangeExtension(docxFile, ".pdf");

                    // Megnyitjuk a .docx fájlt
                    Document document = wordApp.Documents.Open(docxFile);

                    // Mentjük .pdf formátumban
                    document.SaveAs2(pdfFile, WdSaveFormat.wdFormatPDF);

                    // Bezárjuk a dokumentumot
                    document.Close();
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Hiba történt: {ex.Message}");
            }
            finally
            {
                // Kilépünk a Word alkalmazásból
                wordApp.Quit();
            }

            Console.WriteLine("Konverzió befejezve.");
            Console.ReadKey();
        }
    }
}
