using System;
using System.Globalization;
using Xceed.Words.NET;

namespace ReplaceDocxExample
{
    class Program
    {
        private static readonly string caminho = @"C:\Users\Vinicius Mussak\Desktop\DocxReplaceExample\DocxReplaceExample\";

        static void Main(string[] args)
        {
            Replace();

            Console.WriteLine("Foi :)");

            Console.ReadKey();
        }

        static void Replace()
        {
            using (var documento = DocX.Load(caminho + "documento.docx"))
            {
                documento.ReplaceText("#nome", "Vinicius Mussak");

                string mes = DateTime.Now.ToString("MMMM", CultureInfo.CreateSpecificCulture("pt-br"));
                documento.ReplaceText("#mes", mes);

                documento.SaveAs(caminho + "novo-documento.docx");
            }
        }

    }
}
