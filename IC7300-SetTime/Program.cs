using System;
using System.IO.Ports;
using System.Threading;
using System.Management;
using System.Globalization;

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
                    DatahoraSTR += DateTime.Now.ToString("HHmm");
                    DatahoraSTR += "FD";
                    Datahora = ToByteArray(DatahoraSTR);

                    DatafechaSTR += DateTime.Now.ToString("yyyyMMdd");
                    DatafechaSTR += "FD";
                    Datafecha = ToByteArray(DatafechaSTR);


                    //abrimos puerto y enviamos los datos de configuración
                    using (var port = new SerialPort(PortIcom, Baut))
                    {
                        port.StopBits = StopBits.Two;

                        port.Open();
                        port.Write(Datahora, 0, Datahora.Length);
                        Thread.Sleep(100);
                        port.Write(Datafecha, 0, Datafecha.Length);
                        Thread.Sleep(200);
                        port.Close();
                    }

                    Console.WriteLine(Languages.Strings.fecha + DateTime.Now.ToString("yyyy/MM/dd"));
                    Console.WriteLine(Languages.Strings.hora + DateTime.Now.ToString("HH:mm"));
                    Console.WriteLine();



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
    }


}
