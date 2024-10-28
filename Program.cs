using System;
using System.IO;
using OfficeOpenXml;
using RedesLAN;

class Program
{
    static void Main(string[] args)
    {
        Console.WriteLine("Por favor ingresa la ruta completa del archivo Excel:");
        string filePath = Console.ReadLine();

        // Eliminar comillas del inicio y fin
        if (filePath != null)
        {
            filePath = filePath.Trim('\"');
        }



        if (!File.Exists(filePath))
        {
            Console.WriteLine("El archivo no existe en la ruta especificada. Verifica la ruta e inténtalo de nuevo.");
            return;
        }
        else {
            // Habilita la licencia para usar EPPlus
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            Console.WriteLine("El archivo ha sido cargado");
            Proceso proceso = new Proceso(filePath);
            proceso.AccederExcel(); 

        }

        




     
    }

}
