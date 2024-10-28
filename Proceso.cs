using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeOpenXml;


namespace RedesLAN
{
    public class Proceso
    {
        private string filePath;

        public Proceso(string rutaArchivo)
        {
            filePath = rutaArchivo;
        }


        public void AccederExcel()
        {
            try
            {
                using (var package = new ExcelPackage(new FileInfo(filePath)))
                {
                    var Hoja1 = package.Workbook.Worksheets[0];
                    int filasCont = Hoja1.Dimension.Rows; 

                    // Recorre cada fila en la columna A
                    for (int fila = 2; fila <= filasCont; fila++)
                    {
                        var ipCelda = Hoja1.Cells[fila, 1].Text; 
                        if (!string.IsNullOrEmpty(ipCelda))
                        {
                            if (confirmarIP(ipCelda)==true)
                            {
                                string ipClass = claseIP(ipCelda);
                                string ipStructure = estructuraIP(ipCelda);
                                string ipType = tipoIp(ipCelda);
                                string netAddress = direccionRed(ipCelda);
                                string broadcastNet = broadcastRed(ipCelda);
                                string defaultMask = mascaraXDefecto(ipCelda);
                                string hostAddress = direccionHosts(ipCelda);


                                Hoja1.Cells[fila, 2].Value = ipClass;
                                Hoja1.Cells[fila, 3].Value = ipStructure;
                                Hoja1.Cells[fila, 4].Value = ipType;
                                Hoja1.Cells[fila, 5].Value = netAddress;
                                Hoja1.Cells[fila, 6].Value = broadcastNet;
                                Hoja1.Cells[fila, 7].Value = defaultMask;
                                Hoja1.Cells[fila, 8].Value = hostAddress;
                            }
                            else
                            {
                                Hoja1.Cells[fila, 2].Value = "IP Inválida";
                                Hoja1.Cells[fila, 3].Value = "IP Inválida";
                                Hoja1.Cells[fila, 4].Value = "IP Inválida";
                                Hoja1.Cells[fila, 5].Value = "IP Inválida";
                                Hoja1.Cells[fila, 6].Value = "IP Inválida";
                                Hoja1.Cells[fila, 7].Value = "IP Inválida";
                                Hoja1.Cells[fila, 8].Value = "IP Inválida";
                            }
                        }
                    }

                    // Guarda los cambios en el archivo Excel
                    package.Save();
                    Console.WriteLine("Proceso completado. El análisis de cada IP se ha agregado al archivo.");
                }

                // Abre el archivo Excel después de guardar
                Process.Start(new ProcessStartInfo(filePath) { UseShellExecute = true });

            }
            catch (Exception ex)
            {
                Console.WriteLine($"Ocurrió un error al procesar el archivo: {ex.Message}");
            }
        }

        //Confirmar si la IP es válida
        public bool confirmarIP(string ip)
        {
            var octetosIP = ip.Split('.');
            int primerOcteto = int.Parse(octetosIP[0]);
            int segundoOcteto = int.Parse(octetosIP[1]);
            int tercerOcteto = int.Parse(octetosIP[2]);
            int cuartoOcteto = int.Parse(octetosIP[3]);

            if(primerOcteto <=255 && segundoOcteto <=255 && tercerOcteto <=255 && cuartoOcteto <= 255)
            {
                return true;
            }else
                return false;
        }


        // Método para determinar la clase de la IP
        public string claseIP(string ip)
        {
            var octetosIP = ip.Split('.');
            int primerOcteto = int.Parse(octetosIP[0]);

            if (primerOcteto >= 0 && primerOcteto <= 127)
            {
                return "Clase A";
            }
            else if (primerOcteto >= 128 && primerOcteto <= 191)
            {
                return "Clase B";
            }
            else if (primerOcteto >= 192 && primerOcteto <= 223)
            {
                return "Clase C";
            }
            else if (primerOcteto >= 224 && primerOcteto <= 239)
            {
                return "Clase D";
            }
            else
            {
                return "Clase E";
            }
        }


        // Método para determinar la estructura de la IP
        public string estructuraIP(string ip)
        {
            var octetosIP = ip.Split('.');
            int primerOcteto = int.Parse(octetosIP[0]);

            if (primerOcteto >= 0 && primerOcteto <= 127)
            {
                return "R.H.H.H";  //Clase A
            }
            else if (primerOcteto >= 128 && primerOcteto <= 191)
            {
                return "R.R.H.H"; //Clase B
            }
            else if (primerOcteto >= 192 && primerOcteto <= 223)
            {
                return "R.R.R.H"; //Clase C
            }
            else
            {
                return "N/A"; //Clases D y E
            }
        }

        //Averiguar qué tipo de IP es
        public string tipoIp(string ip)
        {
            var octetosIP = ip.Split('.');
            int primerOcteto = int.Parse(octetosIP[0]);
            int segundoOcteto = int.Parse(octetosIP[1]);
            int tercerOcteto = int.Parse(octetosIP[2]);
            int cuartoOcteto = int.Parse(octetosIP[3]);

            if (primerOcteto == 127 || primerOcteto == 0)
            {
                return "Reservada"; //Tipo 1 RESERVADA, 127 es LoopBack.
            }
            else if (primerOcteto >= 224 && primerOcteto <= 239)
            {
                return "Reservada Multicast"; 
            }
            else if (primerOcteto >= 240 && primerOcteto <= 255)
            {
                return "Reservada Experimental";
            }
            else if ((segundoOcteto==0 && tercerOcteto==0 && cuartoOcteto==0) || (tercerOcteto == 0 && cuartoOcteto == 0) || (cuartoOcteto == 0))
            {
                return "Reservada - Dirección de Red"; 
            }
            else if ((segundoOcteto == 255 && tercerOcteto == 255 && cuartoOcteto == 255) || (tercerOcteto == 255 && cuartoOcteto == 255) || (cuartoOcteto == 255 && (primerOcteto >= 192 && primerOcteto <= 223))) //El último me asegura que sea clase C para que el 4to octeto sea el único 255.
            {
                return "Reservada - Dirección de Broadcast"; 
            }
            else if (primerOcteto == 10 || (primerOcteto == 172 && (segundoOcteto>=16 && segundoOcteto<=31)) || (primerOcteto == 192 && segundoOcteto==168))
            {
                return "Privada"; //Tipo 2 PRIVADAS.
            }           
            else
            {
                return "Pública"; 
            }
        }

        //Dirección de Red
        public string direccionRed(string ip)
        {            
            var octetosIP = ip.Split('.');
            int primerOcteto = int.Parse(octetosIP[0]);
            int segundoOcteto = int.Parse(octetosIP[1]);
            int tercerOcteto = int.Parse(octetosIP[2]);
            int cuartoOcteto = int.Parse(octetosIP[3]);            

            if (primerOcteto >= 0 && primerOcteto <= 127)
            {
                return $"{primerOcteto}.0.0.0";  //Clase A: R.0.0.0
            }
            else if (primerOcteto >= 128 && primerOcteto <= 191)
            {
                return $"{primerOcteto}.{segundoOcteto}.0.0"; //Clase B: R.R.0.0
            }
            else if (primerOcteto >= 192 && primerOcteto <= 223)
            {
                return $"{primerOcteto}.{segundoOcteto}.{tercerOcteto}.0"; //Clase C: R.R.R.0
            }
            else
            {
                return "N/A"; //Clases D y E
            }
        }

        //Broadcast
        public string broadcastRed(string ip)
        {
            var octetosIP = ip.Split('.');
            int primerOcteto = int.Parse(octetosIP[0]);
            int segundoOcteto = int.Parse(octetosIP[1]);
            int tercerOcteto = int.Parse(octetosIP[2]);
            int cuartoOcteto = int.Parse(octetosIP[3]);

            if (primerOcteto >= 0 && primerOcteto <= 127)
            {
                return $"{primerOcteto}.255.255.255";  //Clase A
            }
            else if (primerOcteto >= 128 && primerOcteto <= 191)
            {
                return $"{primerOcteto}.{segundoOcteto}.255.255"; //Clase B
            }
            else if (primerOcteto >= 192 && primerOcteto <= 223)
            {
                return $"{primerOcteto}.{segundoOcteto}.{tercerOcteto}.255"; //Clase C
            }
            else
            {
                return "N/A"; //Clases D y E
            }
        }

        //Máscara por defecto
        public string mascaraXDefecto(string ip)
        {
            var octetosIP = ip.Split('.');
            int primerOcteto = int.Parse(octetosIP[0]);
            int segundoOcteto = int.Parse(octetosIP[1]);
            int tercerOcteto = int.Parse(octetosIP[2]);
            int cuartoOcteto = int.Parse(octetosIP[3]);

            if (primerOcteto >= 0 && primerOcteto <= 127)
            {
                return "255.0.0.0";  //Clase A
            }
            else if (primerOcteto >= 128 && primerOcteto <= 191)
            {
                return "255.255.0.0"; //Clase B
            }
            else if (primerOcteto >= 192 && primerOcteto <= 223)
            {
                return "255.255.255.0"; //Clase C
            }
            else
            {
                return "N/A"; //Clases D y E
            }
        }

        //Dirección de Hosts
        public string direccionHosts(string ip)
        {
            var octetosIP = ip.Split('.');
            int primerOcteto = int.Parse(octetosIP[0]);
            int segundoOcteto = int.Parse(octetosIP[1]);
            int tercerOcteto = int.Parse(octetosIP[2]);
            int cuartoOcteto = int.Parse(octetosIP[3]);

            if (primerOcteto >= 0 && primerOcteto <= 127)
            {
                return $"0.{segundoOcteto}.{tercerOcteto}.{cuartoOcteto}";  //Clase A: 0.H.H.H
            }
            else if (primerOcteto >= 128 && primerOcteto <= 191)
            {
                return $"0.0.{tercerOcteto}.{cuartoOcteto}"; //Clase B: 0.0.H.H

            }
            else if (primerOcteto >= 192 && primerOcteto <= 223)
            {
                return $"0.0.0.{cuartoOcteto}";  //Clase C: 0.0.0.H
            }
            else
            {
                return "N/A"; //Clases D y E
            }
        }


    }
}

