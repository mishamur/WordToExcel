using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;


namespace OfficeWrapper
{
    public class ExcelWrapper : IDisposable
    {
        //поля
        Excel.Application application = null;
        Excel.Workbook workbook = null;
        Excel.Worksheet worksheet = null;

        //конструктор
        public ExcelWrapper()
        {

        }

        //запись в файл
        public void WriteToB2Cell(string filePath, string content)
        {
            

            try
            {
                this.application = new Excel.Application();
                this.workbook = application.Workbooks.Open(Filename: filePath, ReadOnly: false); ;
                this.worksheet = workbook.ActiveSheet;

            }
            catch
            {
                RealeseComObjects();
                throw;
            }

            worksheet.Cells[2, "B"] = content;
            ((Excel.Range)worksheet.Cells[2, "B"]).Font.Bold = true;
            application.Columns[2].AutoFit();
            workbook.Save();

        }

        private void RealeseComObjects()
        {
            if (this.worksheet != null)
            {
                while (Marshal.ReleaseComObject(this.worksheet) != 0);
                worksheet = null;
            }

            if (this.workbook != null)
            {
                this.workbook.Close();
                while(Marshal.ReleaseComObject(this.workbook) != 0);
                workbook = null;
            }

            if(application.Workbooks != null)
            {
                application.Workbooks.Close();
                while (Marshal.ReleaseComObject(this.application.Workbooks) != 0);
            }
        
            if (this.application != null)
            {
                this.application.Quit();
                while (Marshal.ReleaseComObject(this.application) != 0);
                application = null;
            }
        }

        public void Dispose()
        {
            this.RealeseComObjects();
        }

    }
}
