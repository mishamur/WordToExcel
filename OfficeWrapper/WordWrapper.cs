using System.Runtime.InteropServices;
using System.Text;
using Word = Microsoft.Office.Interop.Word;


namespace OfficeWrapper
{
    public class WordWrapper : IDisposable
    {
        //поля
        private Word.Application application = null;
        private Word.Document document = null;

        //конструктор
        public WordWrapper()
        {
            
        }

        //считывание с файла
        public StringBuilder ReadOneLine(string filePath)
        {

            try
            {
                application = new Word.Application();
                document = application.Documents.Open(FileName:filePath, ReadOnly:true);
            }
            catch
            {
                this.RealeseComObjects();
                throw;
            }

            StringBuilder result = new StringBuilder();
            var wordEnumerator = document.Words.GetEnumerator();

            while (wordEnumerator.MoveNext())
            {
                string text = ((Word.Range)wordEnumerator.Current).Text;
                if (Char.TryParse(text, out char ch))
                {
                    //control char [carriage return]
                    if (ch == 13)
                    {
                        break;
                    }      
                }
                result.Append(text);

            }

            return result;
        }

        private void RealeseComObjects()
        {
            if (this.document != null)
            {
                this.document.Close();
                while(Marshal.ReleaseComObject(this.document) != 0);
                document = null;
            }

            if (this.application != null)
            {
                this.application.Quit();
                while(Marshal.ReleaseComObject(this.application) != 0);
                application = null;
            }
        }


        public void Dispose()
        {
            this.RealeseComObjects();
        }


    }
}
