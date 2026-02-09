using System;
using System.Runtime.InteropServices;
using System.Net.Sockets;
using System.Text;
using System.Linq;

namespace Posnet1CComponent
{
    // Генерируем интерфейс, чтобы 1С видела методы
    [Guid("D4A2B1E5-8C4F-4E2A-9B3D-7F1A6C8E9D0A")] 
    [InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    [ComVisible(true)]
    public interface IPosnetPrinter
    {
        [DispId(1)]
        string SendCommand(string ip, int port, string command);
    }

    // Реализация класса
    [Guid("F1B2C3D4-E5F6-4A5B-8C7D-9E0F1A2B3C4D")]
    [ClassInterface(ClassInterfaceType.None)]
    [ComVisible(true)]
    [ProgId("AddIn.PosnetPrinter")]
    public class PosnetPrinter : IPosnetPrinter
    {
        private static readonly ushort[] crcTable = new ushort[256];

        static PosnetPrinter()
        {
            ushort poly = 0x1021;
            for (ushort i = 0; i < 256; i++)
            {
                ushort value = 0;
                ushort temp = (ushort)(i << 8);
                for (byte j = 0; j < 8; j++)
                {
                    if (((value ^ temp) & 0x8000) != 0)
                        value = (ushort)((value << 1) ^ poly);
                    else
                        value = (ushort)(value << 1);
                    temp <<= 1;
                }
                crcTable[i] = value;
            }
        }

        public string SendCommand(string ip, int port, string command)
        {
            try
            {
                Encoding posnetEncoding = Encoding.GetEncoding(1250);
                
                byte[] stx = { 0x02 };
                string bodyStr = command + (char)0x03;
                byte[] body = posnetEncoding.GetBytes(bodyStr);
                
                // Расчет CRC16
                ushort crc = 0xFFFF;
                foreach (byte b in body)
                {
                    crc = (ushort)((crc << 8) ^ crcTable[((crc >> 8) ^ b) & 0xFF]);
                }
                byte[] crcBytes = Encoding.ASCII.GetBytes(crc.ToString("X4"));

                byte[] fullFrame = stx.Concat(body).Concat(crcBytes).ToArray();

                using (TcpClient client = new TcpClient())
                {
                    // Таймаут на подключение
                    var connectResult = client.BeginConnect(ip, port, null, null);
                    if (!connectResult.AsyncWaitHandle.WaitOne(TimeSpan.FromSeconds(3))) 
                        return "ERR_TIMEOUT_CONNECT";

                    client.EndConnect(connectResult); 
                     
                    client.SendTimeout = 3000;
                    client.ReceiveTimeout = 5000;

                    using (NetworkStream stream = client.GetStream())
                    {
                        stream.Write(fullFrame, 0, fullFrame.Length);
                        stream.Flush();

                        byte[] buffer = new byte[2048];
                        int bytesRead = stream.Read(buffer, 0, buffer.Length);
                        
                        if (bytesRead > 0)
                        {
                            if (buffer[0] == 0x06) return "OK";
                            if (buffer[0] == 0x15) return "ERR_CRC";
                            if (buffer[0] == 0x02) // Начинается с STX
                            {
                                // Ищем конец текста (ETX)
                                int etxIndex = Array.IndexOf(buffer, (byte)0x03, 0, bytesRead);
                                if (etxIndex > 1)
                                {
                                    // Извлекаем строку между STX (индекс 0) и ETX (etxIndex)
                                    return posnetEncoding.GetString(buffer, 1, etxIndex - 1).Trim();
                                }
                            }
                            return Encoding.GetEncoding(1250).GetString(buffer, 0, bytesRead).Trim();
                        }
                    }
                }
                return "ERR_NO_RESPONSE";
            }
            catch (Exception ex)
            {
                return "ERR_SYSTEM: " + ex.Message;
            }
        }
    }
}