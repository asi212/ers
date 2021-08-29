using System;
using System.Text;
using System.Net.Sockets;
using System.Net;
using System.Windows.Forms;

namespace Zebra
{
    class Send_printer
    {
        private IPEndPoint printerIP;

        public string IP { get; private set; }

        public void sn_to_hex(string sn, string num_copies, string label_type)
        {
            /// convert sn to string zpl data here //
            ///
            string zpl = "^XA^MNW^FT5,5^FH^A0N,30,30^FDLine1^FS^XZ";
            Send_printer n = new Send_printer();
            n.SendData(zpl, num_copies, label_type);
        }

        public void SendData(string zpl, string num_copies, string label_type)
        {
            NetworkStream ns = null;
            Socket socket = null;

            int i;
            for (i = 1; i <= Convert.ToInt32(num_copies); i++) // column 1 fills in LF
            {
                try
                {
                    string ipAddress = "192.168.254.254";
                    printerIP = new IPEndPoint(IPAddress.Parse(ipAddress), 6101);
                    socket = new Socket(printerIP.AddressFamily,
                      SocketType.Stream,
                    ProtocolType.Tcp);
                    socket.Connect(printerIP);
                    ns = new NetworkStream(socket);

                    byte[] toSend = Encoding.ASCII.GetBytes(zpl);
                    ns.Write(toSend, 0, toSend.Length);
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
}

