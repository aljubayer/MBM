

using System;
using System.Collections.Generic;
using System.Text;
using System.Net.Sockets;

namespace MinimalisticTelnet
{
    enum Verbs
    {
        WILL = 251,
        WONT = 252,
        DO = 253,
        DONT = 254,
        IAC = 255
    }

    enum Options
    {
        SGA = 3
    }

    public class TelnetConnection
    {
        TcpClient tcpSocket;

        int TimeOutMs = 500;

        public TelnetConnection(string Hostname, int Port)
        {
            tcpSocket = new TcpClient(Hostname, Port);
        }

        public string Login(string Username, string Password, int LoginTimeOutMs, out bool successFullyLoggedIn)
        {
            successFullyLoggedIn = false;
            int oldTimeOutMs = TimeOutMs;
            TimeOutMs = LoginTimeOutMs;
            string s = ReadUntilPromptFoundOrTimedOut(1, "No Login Prompt Found");

            WriteLine(Username);

            s += ReadUntilPromptFoundOrTimedOut(1, "No Password Prompt Found");

            WriteLine(Password);

            s += ReadUntilPromptFoundOrTimedOut(1, "Not Connected");
            TimeOutMs = oldTimeOutMs;

            if (s.ToLower().Contains("welcome"))
            {
                successFullyLoggedIn = true;
            }

            return s;
        }

        public void WriteLine(string cmd)
        {
            Write(cmd + "\r");
        }

        private void Write(string cmd)
        {
            if (!tcpSocket.Connected) return;
            byte[] buf = System.Text.ASCIIEncoding.ASCII.GetBytes(cmd.Replace("\0xFF", "\0xFF\0xFF"));
            tcpSocket.GetStream().Write(buf, 0, buf.Length);
        }

        public string Interrupt()
        {
            StringBuilder stringBuilder = new StringBuilder();
            byte[] buf = BitConverter.GetBytes(25);
            tcpSocket.GetStream().Write(buf, 0, buf.Length);
            stringBuilder.Append(ReadUntilPromptFoundOrTimedOut(10, "CONNECTION FAILED"));
            return ((object)stringBuilder).ToString();
        }

        public string ExecuteCommand(string commandText)
        {
            SetTimeOutLimit(commandText);
            StringBuilder stringBuilder = new StringBuilder();
            WriteLine(commandText);
            stringBuilder.Append(ReadUntilPromptFoundOrTimedOut(10, "CONNECTION FAILED"));
            return ((object)stringBuilder).ToString();

        }

        private void SetTimeOutLimit(string commandText)
        {

            if (!commandText.Contains(",") && !commandText.Contains(":") && !commandText.Contains("="))
            {
                TimeOutMs = 1000;
            }
            else
            {
                TimeOutMs = 150;
            
            }
        }

        public string ReadUntilPromptFoundOrTimedOut(int timeOutMinutes, string errorMessageToBeShown)
        {
            if (!tcpSocket.Connected)
                return "NOT CONNECTED\n";
            string str = "" + this.Read();
            DateTime now = DateTime.Now;
            str.TrimEnd();
            string asd = string.Empty;
            for (; !str.TrimEnd().EndsWith("\b") && !str.TrimEnd().EndsWith(":") && !str.TrimEnd().EndsWith("=") && !str.TrimEnd().EndsWith(","); str = str + this.Read())
            {
                asd += str;
                if ((DateTime.Now - now).TotalMilliseconds >= timeOutMinutes * 60000.0)
                    throw new Exception(errorMessageToBeShown);
            }
            return str;

        }

        public string Read()
        {
            if (!tcpSocket.Connected) return null;
            StringBuilder sb = new StringBuilder();
            do
            {
                ParseTelnet(sb);
                System.Threading.Thread.Sleep(TimeOutMs);
            } while (tcpSocket.Available > 0);
            return sb.ToString();
        }

        public bool IsConnected
        {
            get { return tcpSocket.Connected; }
        }

        //Update the Input object.
        void ParseTelnet(StringBuilder sb)
        {
            while (tcpSocket.Available > 0)
            {
                int input = tcpSocket.GetStream().ReadByte();
                switch (input)
                {
                    case -1:
                        break;
                    case (int)Verbs.IAC:
                        // interpret as command
                        int inputverb = tcpSocket.GetStream().ReadByte();
                        if (inputverb == -1) break;
                        switch (inputverb)
                        {
                            case (int)Verbs.IAC:
                                //literal IAC = 255 escaped, so append char 255 to string
                                sb.Append(inputverb);
                                break;
                            case (int)Verbs.DO:
                            case (int)Verbs.DONT:
                            case (int)Verbs.WILL:
                            case (int)Verbs.WONT:
                                // reply to all commands with "WONT", unless it is SGA (suppres go ahead)
                                int inputoption = tcpSocket.GetStream().ReadByte();
                                if (inputoption == -1) break;
                                tcpSocket.GetStream().WriteByte((byte)Verbs.IAC);
                                if (inputoption == (int)Options.SGA)
                                    tcpSocket.GetStream().WriteByte(inputverb == (int)Verbs.DO ? (byte)Verbs.WILL : (byte)Verbs.DO);
                                else
                                    tcpSocket.GetStream().WriteByte(inputverb == (int)Verbs.DO ? (byte)Verbs.WONT : (byte)Verbs.DONT);
                                tcpSocket.GetStream().WriteByte((byte)inputoption);
                                break;
                            default:
                                break;
                        }
                        break;
                    default:
                        sb.Append((char)input);
                        break;
                }
            }
        }
    }
}
