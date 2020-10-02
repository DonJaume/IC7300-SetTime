using System.Management;

namespace IC7300_SetTime
{
    class SerialPortInfo
    {
        public SerialPortInfo(ManagementObject property)
        {

            this.DeviceID = property.GetPropertyValue("DeviceID") as string ?? string.Empty;
            this.Name = property.GetPropertyValue("Name") as string ?? string.Empty;
        }

        public string DeviceID;
        public string Name;
    }
}
