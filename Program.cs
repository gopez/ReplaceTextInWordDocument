using Microsoft.Office.Interop.Word;
using System.IO;

namespace ReplaceTextInWordDocument
{
    class Program
    {
        //========================================
        //
        //  
        //========================================
        static void Main(string[] args)
        {
            string filename    = "Template.docx";
            string projectPath = Directory.GetParent(Directory.GetCurrentDirectory()).Parent.FullName;
            string dataPath    = Path.Combine(projectPath, filename);

            Application wordApp = new Application();
            Document document   = wordApp.Documents.Open(dataPath, ReadOnly: true);

            document.Activate();

            FindAndReplace(wordApp, "{{FULL_NAME}}", "Mr. SMith2");
            FindAndReplace(wordApp, "{{EMAIL}}",     "myemail2@server.com");

            string newfilename = "Test2.docx";
            document.SaveAs2(Path.Combine(projectPath, newfilename));

            // quit
            document.Close();
            wordApp.Quit();
        }

        //========================================
        //
        //  https://stackoverflow.com/questions/19252252/c-sharp-word-interop-find-and-replace-everything
        //========================================
        private static void FindAndReplace(Application doc, object findText, object replaceWithText)
        {
            //options
            object matchCase = false;
            object matchWholeWord = true;
            object matchWildCards = false;
            object matchSoundsLike = false;
            object matchAllWordForms = false;
            object forward = true;
            object format = false;
            object matchKashida = false;
            object matchDiacritics = false;
            object matchAlefHamza = false;
            object matchControl = false;
            object read_only = false;
            object visible = true;
            object replace = 2;
            object wrap = 1;
            //execute find and replace
            doc.Selection.Find.Execute(ref findText, ref matchCase, ref matchWholeWord,
                ref matchWildCards, ref matchSoundsLike, ref matchAllWordForms, ref forward, ref wrap, ref format, ref replaceWithText, ref replace,
                ref matchKashida, ref matchDiacritics, ref matchAlefHamza, ref matchControl);
        }
    }
}


