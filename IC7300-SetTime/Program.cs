using System;
using System.IO.Ports;
using System.Threading;
using System.Management;
using System.Globalization;
using IC7300_SetTime.Languages;

namespace IC7300_SetTime
{
    class Program
    {
        static void Main()
        {

            int Baut = 0;
            string PortIcom = "";

            string DatahoraSTR = "FEFE94E01A050095";
            string DatafechaSTR = "FEFE94E01A050094";
            byte[] Datahora;
            byte[] Datafecha;



            //******************************************* Detectar Puerto del ICOM ***************************************************
            PortIcom = GetPortIcom();

            if (PortIcom != string.Empty)
            {
                Console.WriteLine(Languages.Strings.detection, PortIcom);
                Console.WriteLine();

                //******************************************* Detectar Baut Rate ***************************************************
                Baut = GetBaud(PortIcom);
                if(Baut!=0)
                {
                    Console.WriteLine(Languages.Strings.baud, Baut);
                    Console.WriteLine();


                    //*************************************** Iniciamos la puesta en hora *******************************************

                    Console.Write(Languages.Strings.wait);

                    //esperamos que llegue al segundo 00
                    int seg = DateTime.Now.Second;
                    Console.Write((59 - seg).ToString().PadLeft(2, '0') + Languages.Strings.seconds);
                    while (DateTime.Now.Second != 0)
                    {
                        if (DateTime.Now.Second != seg)
                        {
                            seg = DateTime.Now.Second;
                            Console.SetCursorPosition(Console.CursorLeft - (Languages.Strings.seconds.Length + 2), Console.CursorTop);
                            Console.Write((59 - seg).ToString().PadLeft(2, '0') + Languages.Strings.seconds);
                        }
                        Thread.Sleep(50);
                    }
                    Console.WriteLine();
                    Console.WriteLine();

                    //Creamos la matriz de configuracion
                    string hora_actual = DateTime.Now.ToString("HHmm");
                    DatahoraSTR += hora_actual;
                    DatahoraSTR += "FD";
                    Datahora = ToByteArray(DatahoraSTR);

                    string fecha_actual = DateTime.Now.ToString("yyyyMMdd");
                    DatafechaSTR += fecha_actual;
                    DatafechaSTR += "FD";
                    Datafecha = ToByteArray(DatafechaSTR);


                    //abrimos puerto y enviamos los datos de configuración
                    using (var port = new SerialPort(PortIcom, Baut))
                    {
                        port.Open();
                        port.Write(Datahora, 0, Datahora.Length);
                        Thread.Sleep(100);
                        port.Write(Datafecha, 0, Datafecha.Length);
                        Thread.Sleep(200);
                        port.Close();
                    }

                    //****** comprobar escritura ******    
                    if(Comprobar(fecha_actual, hora_actual, PortIcom, Baut))
                    {
                        //escritura realizada correctamente
                        Console.WriteLine();
                        Console.WriteLine(Strings.ok);
                        Console.WriteLine();
                        Console.WriteLine(Languages.Strings.fecha + DateTime.Now.ToString("yyyy/MM/dd"));
                        Console.WriteLine(Languages.Strings.hora + DateTime.Now.ToString("HH:mm"));
                        Console.WriteLine();
                    }
                    else
                    {
                        //error de escritura fecha / hora
                        Console.WriteLine(Strings.error);                        
                    }
                }
                else
                {
                    //puerto no detectado
                    Console.WriteLine(Languages.Strings.noBaud);
                }

            }
            else
            {
                //Icom no detectado
                Console.WriteLine(Languages.Strings.no7300);
            }


            Console.WriteLine(Languages.Strings.salir);
            Console.ReadKey();
        }


        //Convierte una cadena String en formato hexadecimal a una matriz de bytes
        public static byte[] ToByteArray(String hexString)
        {
            byte[] retval = new byte[hexString.Length / 2];
            for (int i = 0; i < hexString.Length; i += 2)
                retval[i / 2] = Convert.ToByte(hexString.Substring(i, 2), 16);
            return retval;
        }


        //Busca el puerto donde está conectado el Icom7300
        static string GetPortIcom()
        {
            ManagementClass processClass = new ManagementClass("Win32_PnPEntity");
            ManagementObjectCollection Ports = processClass.GetInstances();
            foreach (ManagementObject property in Ports)
            {
                var name = property.GetPropertyValue("Name");
                if (name != null && name.ToString().Contains("COM"))
                {
                    var portInfo = new SerialPortInfo(property);
                    if (portInfo.DeviceID.Contains("IC-7300"))
                    {
                        return "COM" + portInfo.Name[portInfo.Name.IndexOf("COM") + 3].ToString();
                    }
                }
            }
            return string.Empty;
        }


        //Busca el baudrate configurado en el icom7300
        static int GetBaud(string Port)
        {
            byte[] DATO = ToByteArray("FEFE94E01A050075FD");
            int[] BautR = { 4800, 9600, 19200, 38400, 57600, 115200 };
            byte[] lectura = new byte[20];

            for (int i = 0; i < BautR.Length; i++)
            {
                using (var port = new SerialPort(Port, BautR[i]))
                {
                    port.Open();

                    port.Write(DATO, 0, DATO.Length);
                    Thread.Sleep(50);
                    int leido = port.Read(lectura, 0, port.BytesToRead);
                    port.Close();

                    if (leido > 0 && leido != DATO.Length)
                    {
                        return BautR[i];
                    }  
                    
                }
            }
            return 0;
        }


        static bool Comprobar(string fecha, string hora, string Port, int Baut)
        {
            byte[] DATOfecha = ToByteArray("FEFE94E01A050094FD");
            byte[] DATOhora = ToByteArray("FEFE94E01A050095FD");
            byte[] lectura = new byte[30];
            string fechaleida;
            string horaleida;

            using (var port = new SerialPort(Port, Baut))
            {
                port.Open();

                //solicitamos fecha escrita en el icom
                port.Write(DATOfecha, 0, DATOfecha.Length);
                Thread.Sleep(50);
                int leido = port.Read(lectura, 0, port.BytesToRead);
                
                fechaleida =
                    lectura[leido - 5].ToString("X").PadLeft(2, '0') +
                    lectura[leido - 4].ToString("X").PadLeft(2, '0') +
                    lectura[leido - 3].ToString("X").PadLeft(2, '0') +
                    lectura[leido - 2].ToString("X").PadLeft(2, '0');

                //solicitamos hora escrita en el icom
                port.Write(DATOhora, 0, DATOhora.Length);
                Thread.Sleep(50);
                leido = port.Read(lectura, 0, port.BytesToRead);

                horaleida =
                    lectura[leido - 3].ToString("X").PadLeft(2, '0') +
                    lectura[leido - 2].ToString("X").PadLeft(2, '0');

                port.Close();
            }

            if (fecha == fechaleida && hora == horaleida) return true;
            else return false;
        }

    }


}
