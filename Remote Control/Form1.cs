using System;
using System.ComponentModel;
using System.Net.Sockets;
using System.Text;
using System.Windows.Forms;

namespace PowerPointAddIn2
{
    internal partial class Form1 : Form
    {
        string v, res;
        TcpListener server;
        TcpClient client;


        public Form1()
        {
            InitializeComponent();
        }

        public Form1(string v, TcpListener server, TcpClient client)
        {
            this.v = v;
            this.server = server;
            this.client = client;
            InitializeComponent();
            backgroundWorker1.RunWorkerAsync();
            done.Visible = false;
        }

        private void Form1_FormClosed(object sender, FormClosedEventArgs e)
        {
            if (backgroundWorker1.IsBusy)
                backgroundWorker1.CancelAsync();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            ip.Text = v;
        }

        private void backgroundWorker1_DoWork(object sender, DoWorkEventArgs e)
        {
            while (true)
            {
                try
                {
                    client = server.AcceptTcpClient();
                    client.NoDelay = true;
                    byte[] receivedData = new byte[1024];
                    NetworkStream stream = client.GetStream();
                    stream.Read(receivedData, 0, receivedData.Length);

                    System.Diagnostics.Debug.WriteLine("External Client");
                    StringBuilder msg = new StringBuilder();
                    foreach (byte b in receivedData)
                    {
                        if (b.Equals(59))
                            break;
                        else
                            msg.Append(Convert.ToChar(b).ToString());
                    }

                    System.Diagnostics.Debug.WriteLine(msg.ToString());
                    res = msg.ToString();
                    if (msg.ToString().Trim().Equals("start"))
                    {
                        break;
                    }
                }
                catch (SocketException)
                {
                    server.Stop();
                    break;
                }
            }
        }

        private void done_Click(object sender, EventArgs e)
        {
            // if (msg.ToString().Length > 0)
            // {
            System.IO.StreamWriter writer = new System.IO.StreamWriter(client.GetStream());
            writer.Write("command:{true}");
            writer.Flush();
            writer.Close();
            //   res = msg.ToString();
            //}

            client.Close();
            this.Close();
        }

        private void backgroundWorker1_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            try
            {
                if (res.Trim().Equals("start"))
                {
                    label2.Visible = false;
                    done.Visible = true;
                    ip.Text = "Connected";
                }
            }
            catch (NullReferenceException)
            {
                this.Close();
            }
        }
    }

    /// <summary>
    /// Your custom message box helper.
    /// </summary>
    public static class CustomMessageBox
    {
        public static void Show(string title, TcpListener server, TcpClient client)
        {
            // using construct ensures the resources are freed when form is closed
            using (var form = new Form1(title, server, client))
            {
                form.ShowDialog();
            }
        }
    }
}
