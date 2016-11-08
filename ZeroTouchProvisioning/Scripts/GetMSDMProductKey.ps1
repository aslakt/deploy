$Source = @"
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Text;

namespace Get.MSDMProductKey
{
    public class Program
    {
        [DllImport("kernel32")]
        private static extern uint EnumSystemFirmwareTables(uint FirmwareTableProviderSignature, IntPtr pFirmwareTableBuffer, uint BufferSize);
        [DllImport("kernel32")]
        private static extern uint GetSystemFirmwareTable(uint FirmwareTableProviderSignature, uint FirmwareTableID, IntPtr pFirmwareTableBuffer, uint BufferSize);

        private static bool checkMSDM(out byte[] buffer)
        {
            uint firmwareTableProviderSignature = 0x41435049; // 'ACPI' in Hexadecimal
            uint bufferSize = EnumSystemFirmwareTables(firmwareTableProviderSignature, IntPtr.Zero, 0);
            IntPtr pFirmwareTableBuffer = Marshal.AllocHGlobal((int)bufferSize);
            buffer = new byte[bufferSize];
            EnumSystemFirmwareTables(firmwareTableProviderSignature, pFirmwareTableBuffer, bufferSize);
            Marshal.Copy(pFirmwareTableBuffer, buffer, 0, buffer.Length);
            Marshal.FreeHGlobal(pFirmwareTableBuffer);
            if (Encoding.ASCII.GetString(buffer).Contains("MSDM"))
            {
                uint firmwareTableID = 0x4d44534d; // Reversed 'MSDM' in Hexadecimal
                bufferSize = GetSystemFirmwareTable(firmwareTableProviderSignature, firmwareTableID, IntPtr.Zero, 0);
                buffer = new byte[bufferSize];
                pFirmwareTableBuffer = Marshal.AllocHGlobal((int)bufferSize);
                GetSystemFirmwareTable(firmwareTableProviderSignature, firmwareTableID, pFirmwareTableBuffer, bufferSize);
                Marshal.Copy(pFirmwareTableBuffer, buffer, 0, buffer.Length);
                Marshal.FreeHGlobal(pFirmwareTableBuffer);
                return true;
            }
            return false;
        }

        public string GetProductKey()
        {
            byte[] buffer;
            if (checkMSDM(out buffer))
            {
                Encoding encoding = Encoding.GetEncoding(0x4e4);
                string signature = encoding.GetString(buffer, 0x0, 0x4);
                int length = BitConverter.ToInt32(buffer, 0x4);
                byte revision = (byte)buffer.GetValue(0x8);
                byte checksum = (Byte)buffer.GetValue(0x9);
                string oemid = encoding.GetString(buffer, 0xa, 0x6);
                string oemtableid = encoding.GetString(buffer, 0x10, 0x8);
                int oemrev = BitConverter.ToInt32(buffer, 0x18);
                string creatorid = encoding.GetString(buffer, 0x1c, 0x4);
                int creatorrev = BitConverter.ToInt32(buffer, 0x20);
                int sls_version = BitConverter.ToInt32(buffer, 0x24);
                int sls_reserved = BitConverter.ToInt32(buffer, 0x28);
                int sls_datatype = BitConverter.ToInt32(buffer, 0x2C);
                int sls_datareserved = BitConverter.ToInt32(buffer, 0x30);
                int sls_datalength = BitConverter.ToInt32(buffer, 0x34);
                string sls_data = encoding.GetString(buffer, 0x38, sls_datalength);
                string result = "Signature         : " + signature +
                    "\nLength            : " + length +
                    "\nRevison           : " + revision.ToString("X") +
                    "\nChecksum          : " + checksum.ToString("X") +
                    "\nOEM ID            : " + oemid +
                    "\nOEM Table ID      : " + oemtableid +
                    "\nOEM Revision      : " + oemrev +
                    "\nCreator ID        : " + creatorid +
                    "\nCreator Revision  : " + creatorrev +
                    "\nSLS Version       : " + sls_version +
                    "\nSLS Reserved      : " + sls_reserved +
                    "\nSLS Datatype      : " + sls_datatype +
                    "\nSLS Data Reserved : " + sls_datareserved +
                    "\nSLS Data Length   : " + sls_datalength +
                    "\nKey               : " + sls_data;
                
                if(!String.IsNullOrEmpty(sls_data))
                {
                    return sls_data;
                }
                else
                {
                    return null;
                }
            }
            else
            {
                return null;
            }
        }
    }
}
"@

Add-Type -TypeDefinition $Source -Language CSharp

Try {
    $TSEnvironment = New-Object -COMObject Microsoft.SMS.TSEnvironment
} Catch {
    Write-Host "Unable to load TS Environment"
}

$MSDM = New-Object -TypeName Get.MSDMProductKey.Program
$ProductKey = $MSDM.GetProductKey()

If ($ProductKey -ne $null) {
    Write-Host "Found product key in BIOS"
    Try {
        $TSEnvironment.value("FirmwareProductKey") = $ProductKey
        Write-Host "Property FirmwareProductKey is now: $ProductKey"
    } Catch {
        Write-Host "Unable to store Product Key to TS Environment"
    }
} Else {
    Write-Host "Did not find a product key in BIOS"
}