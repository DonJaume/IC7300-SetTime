using System;
using System.IO.Ports;
using System.Threading;
using System.Management;

namespace IC7300_SetTime
{
    class Program
    {
        static void Main()
        {
            string DatahoraSTR = "FEFE94E01A050095";
            string DatafechaSTR = "FEFE94E01A050094";
            byte[] Datahora;
            byte[] Datafecha;


            //detectamos Icom 7300
            string PortIcom = GetPortIcom();
            if (PortIcom != string.Empty)
            {
                Console.WriteLine("Icom 7300 detectado en puerto {0}.", PortIcom);
                Console.WriteLine();
                Console.Write("Esperando minuto exacto, tiempo restante: ");

                //esperamos que llegue al segundo 00
                int seg = DateTime.Now.Second;
                Console.Write((59 - seg).ToString().PadLeft(2, '0') + " seconds.");
                while (DateTime.Now.Second != 0)
                {
                    if (DateTime.Now.Second != seg)
                    {
                        seg = DateTime.Now.Second;
                        Console.SetCursorPosition(Console.CursorLeft - 11, Console.CursorTop);
                        Console.Write((59 - seg).ToString().PadLeft(2, '0') + " seconds.");
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
                using (var port = new SerialPort(PortIcom, 9600))
                {
                    port.StopBits = StopBits.Two;

                    port.Open();
                    port.Write(Datahora, 0, Datahora.Length);
                    Thread.Sleep(100);
                    port.Write(Datafecha, 0, Datafecha.Length);
                    Thread.Sleep(200);
                    port.Close();
                }

                Console.WriteLine("Fecha establecida: " + DateTime.Now.ToString("yyyy/MM/dd"));
                Console.WriteLine("Hora establecida: " + DateTime.Now.ToString("HH:mm"));
                Console.WriteLine();

            }
            else
            {
                Console.WriteLine("Icom 7300 no detectado.");
            }


            Console.WriteLine("Pulse una tecla para salir.");
            Console.ReadKey();
        }


        public static byte[] ToByteArray(String hexString)
        {
            byte[] retval = new byte[hexString.Length / 2];
            for (int i = 0; i < hexString.Length; i += 2)
                retval[i / 2] = Convert.ToByte(hexString.Substring(i, 2), 16);
            return retval;
        }



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
    }


}
