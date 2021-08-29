using System;
using System.Text;
using System.Net.Sockets;
using System.Net;

namespace cmdZebra2
{
    class Program
    {
        private IPEndPoint printerIP;

        public string IP { get; private set; }

        static void Main(string[] args)
        {
            string zpl = "^XA^MNW^FT5,5^FH^A0N,30,30^FDLine1^FS^XZ";
            Program n = new Program();
            n.SendData(zpl);
            Console.ReadKey();
        }

        public void SendData(string zpl)
        {
            NetworkStream ns = null;
            Socket socket = null;

            try
            {
                string ipAddress = "192.168.254.254";
                printerIP = new IPEndPoint(IPAddress.Parse(ipAddress), 6101);
                socket = new Socket(printerIP.AddressFamily,
                  SocketType.Stream,
                ProtocolType.Tcp);     
                Console.WriteLine(printerIP); // test
                socket.Connect(printerIP);
                Console.WriteLine("Connected"); // test
                ns = new NetworkStream(socket);

                byte[] toSend = Encoding.ASCII.GetBytes(zpl);
                //byte[] toSend = { 0x1B, 0x74, 0x00, 0x0A, 0x00, 0x14, 0x00, 0x50, 0x61, 0x72, 0x6B, 0x4E, 0x61, 0x6D, 0x65, 0x00, 0x0A, 0x0A, 0x0D, 0x1E };
                Console.WriteLine(zpl);
                ns.Write(toSend, 0, toSend.Length);
                Console.WriteLine(toSend.ToString());
                Console.WriteLine("Byte Sent");
                Console.ReadKey();
            }
            finally
            {
                if (ns != null)
                    ns.Close();

                if (socket != null && socket.Connected)
                    socket.Close();
            }
        }
    }
}

