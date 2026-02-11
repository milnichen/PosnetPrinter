using System;
using System.Runtime.InteropServices;
using System.Net.Sockets;
using System.Text;
using System.Linq;
using System.Globalization;

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
        private const int ConnectTimeoutMs = 3000;
        private const int SendTimeoutMs = 3000;
        private const int ReceiveTimeoutMs = 5000;

        public string SendCommand(string ip, int port, string command)
        {
            try
            {
                // Базова валідація параметрів
                if (string.IsNullOrWhiteSpace(ip))
                    return "ERR_PARAM_IP";

                if (port <= 0 || port > 65535)
                    return "ERR_PARAM_PORT";

                if (command == null)
                    return "ERR_PARAM_COMMAND";

                Encoding posnetEncoding = Encoding.GetEncoding(1250);
                byte[] fullFrame = BuildFrame(command, posnetEncoding);

                using (TcpClient client = new TcpClient())
                {
                    // Таймаут на подключение
                    var connectResult = client.BeginConnect(ip, port, null, null);
                    if (!connectResult.AsyncWaitHandle.WaitOne(TimeSpan.FromMilliseconds(ConnectTimeoutMs)))
                        return "ERR_TIMEOUT_CONNECT";

                    client.EndConnect(connectResult); 
                     
                    client.SendTimeout = SendTimeoutMs;
                    client.ReceiveTimeout = ReceiveTimeoutMs;

                    using (NetworkStream stream = client.GetStream())
                    {
                        stream.Write(fullFrame, 0, fullFrame.Length);
                        stream.Flush();

                        string response = ReadAndParseResponse(stream, posnetEncoding);
                        if (response != null)
                            return response;
                    }
                }
                return "ERR_NO_RESPONSE";
            }
            catch (Exception ex)
            {
                return "ERR_SYSTEM: " + ex.Message;
            }
        }

        /// <summary>
        /// Формирует кадр Posnet: STX + BODY (command + ETX) + CRC(HEX4)
        /// </summary>
        private static byte[] BuildFrame(string command, Encoding encoding)
        {
            byte[] stx = { 0x02 };
            string bodyStr = command + (char)0x03; // ETX
            byte[] body = encoding.GetBytes(bodyStr);

            ushort crc = PosnetCrc16.Compute(body);
            byte[] crcBytes = Encoding.ASCII.GetBytes(crc.ToString("X4"));

            return stx.Concat(body).Concat(crcBytes).ToArray();
        }

        /// <summary>
        /// Читает ответ от принтера, обрабатывая ACK/NAK и полный кадр STX...ETX+CRC.
        /// Возвращает строку ответа или null, если данных нет.
        /// </summary>
        private static string ReadAndParseResponse(NetworkStream stream, Encoding encoding)
        {
            var buffer = new byte[512];
            var accumulated = new System.Collections.Generic.List<byte>();

            while (true)
            {
                int bytesRead = stream.Read(buffer, 0, buffer.Length);
                if (bytesRead <= 0)
                    break;

                accumulated.AddRange(buffer.Take(bytesRead));

                if (accumulated.Count == 0)
                    continue;

                byte first = accumulated[0];

                // ACK / NAK
                if (first == 0x06)
                    return "OK";

                if (first == 0x15)
                    return "ERR_CRC";

                // Кадр STX ... ETX + CRC(4 байта в HEX)
                if (first == 0x02)
                {
                    byte[] current = accumulated.ToArray();
                    int etxIndex = Array.IndexOf(current, (byte)0x03, 1); // ищем ETX после STX
                    if (etxIndex > 1 && current.Length >= etxIndex + 1 + 4)
                    {
                        // Есть ETX и как минимум 4 байта CRC
                        int crcStart = etxIndex + 1;
                        string crcHex = Encoding.ASCII.GetString(current, crcStart, 4);

                        if (!ushort.TryParse(crcHex, NumberStyles.HexNumber, CultureInfo.InvariantCulture, out ushort crcReceived))
                        {
                            return "ERR_CRC_FORMAT";
                        }

                        // Тело для проверки CRC: от индекса 1 (после STX) включительно до ETX включительно
                        int bodyLength = etxIndex; // с 1 до etxIndex включительно => etxIndex байт
                        byte[] bodyWithEtx = new byte[bodyLength];
                        Array.Copy(current, 1, bodyWithEtx, 0, bodyLength);

                        ushort crcCalculated = PosnetCrc16.Compute(bodyWithEtx);
                        if (crcCalculated != crcReceived)
                        {
                            return "ERR_CRC_RESPONSE";
                        }

                        // Строка между STX и ETX (без ETX)
                        string text = encoding.GetString(current, 1, etxIndex - 1).Trim();
                        return text;
                    }
                }
            }

            if (accumulated.Count > 0)
            {
                // Непротокольный/частичный ответ — возвращаем как есть в CP1250
                return encoding.GetString(accumulated.ToArray(), 0, accumulated.Count).Trim();
            }

            return null;
        }
    }

    internal static class PosnetCrc16
    {
        private static readonly ushort[] CrcTable = new ushort[256];

        static PosnetCrc16()
        {
            const ushort poly = 0x1021;
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
                CrcTable[i] = value;
            }
        }

        public static ushort Compute(byte[] data)
        {
            if (data == null || data.Length == 0)
                return 0;

            ushort crc = 0xFFFF;
            foreach (byte b in data)
            {
                crc = (ushort)((crc << 8) ^ CrcTable[((crc >> 8) ^ b) & 0xFF]);
            }

            return crc;
        }
    }
}