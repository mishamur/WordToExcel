/*
     * ДЗ
     * есть документ word  с олной строчкой текста, приложение которое через ком подключится Doc -> excel запишет в ячейку b2 полужирным
     * Документ Word, Excel
     * позднее связывание   dynamic
     * Word.Application  метод RealeaseComeObject
     * Excel.Application
     * 
     * полужирный шрифт
     */

using OfficeWrapper;

public class Program
{
    public static void Main(string[] args)
    {
        string result = "";
        try
        {
           
            using (WordWrapper wordWrapper = new WordWrapper())
            {
                result = wordWrapper.ReadOneLine(@"C:\Users\User\Documents\mveuC#\excelWordHW\wordFile.docx").ToString();
            }
        }
        catch(Exception ex)
        {
            Console.WriteLine(ex.Message);
            return;
        }


        try
        {
            using (ExcelWrapper excelWrapper = new ExcelWrapper())
            {
                excelWrapper.WriteToB2Cell(@"C:\Users\User\Documents\mveuC#\excelWordHW\excelFile.xlsx", result);
            }
        }
        catch(Exception ex)
        {
            Console.WriteLine(ex.Message);
            return;
        }

    }
}