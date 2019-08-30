using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.Office.Tools.Ribbon;
using System.Net;
using System.Net.Sockets;
using System.Drawing;
using System.IO;
using System.Runtime.InteropServices;
using Microsoft.Office.Core;
using System.Text.RegularExpressions;
using System.Threading;


namespace Remote_Control
{
    public partial class RibbonControl
    {
        Application powerPoint;
        MsoTriState activeAlready = MsoTriState.msoFalse;
        TcpListener server;
        TcpClient client;
        static int route = 1;
        String slides = "";

        private void RibbonControl_Load(object sender, RibbonUIEventArgs e)
        {
            powerPoint = new Application();
            Directory.CreateDirectory("C:\\Users\\" + getUsername() + "\\.tenue\\");
            GetIPAddress();
        }

        static string IPAddress = "localhost";
        static IPAddress ipSub;

        public static IPAddress GetIPAddress()
        {
            IPHostEntry Host = default(IPHostEntry);
            string Hostname = null;
            Hostname = System.Environment.MachineName;
            Host = Dns.GetHostEntry(Hostname);

            foreach (IPAddress IP in Host.AddressList)
            {
                if (IP.AddressFamily == AddressFamily.InterNetwork && (!System.Net.IPAddress.IsLoopback(IP)))
                {
                    IPAddress = Convert.ToString(IP);
                    ipSub = IP;
                }
            }
            return ipSub;
        }

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            System.Diagnostics.Debug.WriteLine("===Responded Here===");
            try
            {
                if (!backgroundWorker1.IsBusy)
                {
                    backgroundWorker1.RunWorkerAsync();
                    backgroundWorker2.RunWorkerAsync();
                }
                else
                {
                    System.Windows.Forms.MessageBox.Show("Connection already established on: " + GetIPAddress() + "",
                        "Information",
                        System.Windows.Forms.MessageBoxButtons.OK,
                        System.Windows.Forms.MessageBoxIcon.Information);
                };
            }
            catch (COMException)
            {
                System.Windows.Forms.MessageBox.Show("Open a presentation first.",
                       "Information",
                       System.Windows.Forms.MessageBoxButtons.OK,
                       System.Windows.Forms.MessageBoxIcon.Information);
            }

        }

        private void stop_Click(object sender, RibbonControlEventArgs e)
        {
            if (activeAlready == MsoTriState.msoTrue || backgroundWorker1.IsBusy)
            {
                backgroundWorker1.CancelAsync();
                activeAlready = MsoTriState.msoCTrue;
                System.Windows.Forms.MessageBox.Show("Remote disconnected.",
                        "Information",
                        System.Windows.Forms.MessageBoxButtons.OK,
                        System.Windows.Forms.MessageBoxIcon.Information);
            }
            else
            {
                System.Windows.Forms.MessageBox.Show("Not connected.",
                        "Information",
                        System.Windows.Forms.MessageBoxButtons.OK,
                        System.Windows.Forms.MessageBoxIcon.Information);
            }
        }

        private void backgroundWorker1_DoWork(object sender, System.ComponentModel.DoWorkEventArgs e)
        {
           if (Dns.GetHostEntry(GetIPAddress()).AddressList.Count() > 1)
            {
                server = new TcpListener(ipSub, 2522);
                client = default(TcpClient);

                try
                {
                    server.Start();
                    System.Diagnostics.Debug.WriteLine("===Server Started===");
                    PowerPointAddIn2.CustomMessageBox.Show(Convert.ToString(GetIPAddress()), server, client);
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine(ex.ToString());
                }

                while (true)
                {
                    if (backgroundWorker1.CancellationPending)
                    {
                        e.Cancel = true;
                        server.Stop();
                        System.Diagnostics.Debug.WriteLine("===Server Stopped===");
                        return;
                    }

                    try
                    {
                        client = server.AcceptTcpClient();
                        client.NoDelay = true;
                        byte[] receivedData = new byte[1024];
                        NetworkStream stream = client.GetStream();
                        stream.Read(receivedData, 0, receivedData.Length);

                        System.Diagnostics.Debug.WriteLine(Encoding.ASCII.GetString(receivedData, 0, receivedData.Length));
                        StringBuilder msg = new StringBuilder();
                        foreach (byte b in receivedData)
                        {
                            if (b.Equals(59))
                                break;
                            else
                                msg.Append(Convert.ToChar(b).ToString());
                        }

                        System.Diagnostics.Debug.WriteLine(msg.ToString());
                        if (activeAlready == MsoTriState.msoCTrue)
                        {
                            StreamWriter writer = new StreamWriter(client.GetStream());
                            writer.Write("command:{stopped}");
                            writer.Flush();
                            writer.Close();
                            client.Close();
                            activeAlready = MsoTriState.msoFalse;
                            break;
                        }
                        else
                        {
                            if (msg.ToString().Length > 0)
                            {
                                System.Diagnostics.Debug.WriteLine(msg.ToString());
                                if (msg.ToString().Equals("needImages"))
                                {
                                    for (int i = 1; i <= powerPoint.ActivePresentation.Slides.Count; i++)
                                    {
                                        Thread t = new Thread(() => createFTPClient(i));
                                        t.Start();
                                        Thread.Sleep(2000);
                                    }
                                    StreamWriter writer = new StreamWriter(client.GetStream());
                                    writer.Write("command:{true}");
                                    writer.Write(Environment.NewLine);
                                    writer.Flush();
                                    writer.Close();
                                }
                                else
                                {
                                    if (!msg.ToString().Equals("-1"))
                                        performCommand(msg.ToString());

                                    StreamWriter writer = new StreamWriter(client.GetStream());
                                    writer.Write("command:{true}");
                                    if (!msg.ToString().Equals("-1"))
                                    {
                                        if (!slides.Equals(""))
                                        {
                                            writer.Write("\nslides:[" + slides + "]");
                                        }
                                        try
                                        {
                                            writer.Write("\nposition:" +
                                                powerPoint.ActivePresentation.SlideShowWindow.View.Slide.SlideIndex +
                                                ";");
                                        }
                                        catch (COMException) { }
                                    }
                                    writer.Flush();
                                    writer.Close();
                                }
                                client.Close();
                            }
                        }
                    }
                    catch (SocketException)
                    {
                        activeAlready = MsoTriState.msoFalse;
                        if (client != null)
                            client.Close();
                        server.Stop();
                        break;
                    }

                    if (activeAlready == MsoTriState.msoCTrue)
                        break;
                }
            }
            else
            {
                System.Windows.Forms.MessageBox.Show("Unable to establish connection.\nPlease ensure you are connected to a network",
                    "Error",
                    System.Windows.Forms.MessageBoxButtons.OK,
                    System.Windows.Forms.MessageBoxIcon.Error);
            }
        }

        public void createFTPClient(int i)
        {
            IPAddress ip = ipSub;
            TcpListener server = new TcpListener(ip, (4240 + route));
            System.Diagnostics.Debug.WriteLine("Route: " + 4240 + route);
            route = route + 1;
            TcpClient client = default(TcpClient);

            try
            {
                server.Start();
                System.Diagnostics.Debug.WriteLine("===FTPServer Started===");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine(ex.StackTrace);
            }

            while (true)
            {
                client = server.AcceptTcpClient();
                try
                {
                    var fileIO = File.OpenRead("C:\\Users\\" + getUsername() + "\\.tenue\\" + i + ".png");
                    using (var clientSocket = client.GetStream())
                    {
                        System.Diagnostics.Debug.WriteLine("C:\\Users\\" + getUsername() + "\\.tenue\\" + i + ".png: writing...");
                        var buffer = new byte[1024 * 16];
                        int count;
                        while ((count = fileIO.Read(buffer, 0, buffer.Length)) > 0)
                        {
                            clientSocket.Write(buffer, 0, count);
                            System.Diagnostics.Debug.WriteLine("writing... " + count);
                            if ((count / 16) != 1024)
                                break;
                        }
                        clientSocket.Flush();
                        System.Diagnostics.Debug.WriteLine("C:\\Users\\" + getUsername() + "\\.tenue\\" + i + ".png: dONE writing...");
                        clientSocket.Close();
                        fileIO.Close();
                        client.Close();
                    }
                }
                catch (FileNotFoundException) { }
            }
        }

        public void exportAndSlideNames()
        {
            int position = 1;
            foreach (Slide slide in powerPoint.ActivePresentation.Slides)
            {
                slide.Export("C:\\Users\\" + getUsername() + "\\.tenue\\" + position + ".png", "PNG", 468, 256);
                position = position + 1;
            }
        }

        public static string getUsername()
        {
            String name = "";
            DirectoryInfo[] f = new DirectoryInfo("C:\\Users").GetDirectories();
            foreach (DirectoryInfo d in f)
            {
                if (!d.Name.Equals("Public"))
                    name = d.Name;
            }
            if (!Environment.UserName.Equals(""))
                return Environment.UserName;
            else
                return name;
        }

        private void performCommand(string v)
        {
            try
            {
                activeAlready = powerPoint.ActivePresentation.SlideShowWindow.Active;
            }
            catch (COMException)
            {
                activeAlready = MsoTriState.msoFalse;
            }

            if (v.Equals("startPresentation"))
            {
                if (activeAlready == MsoTriState.msoFalse)
                {
                    powerPoint.ActivePresentation.SlideShowSettings.Run();
                    activeAlready = MsoTriState.msoTrue;
                }
                else
                {
                    powerPoint.ActivePresentation.SlideShowSettings.Run();
                    activeAlready = MsoTriState.msoTrue;
                }
            }
            else if (v.Equals("stopPresentation"))
            {
                if (activeAlready == MsoTriState.msoTrue)
                {
                    powerPoint.ActivePresentation.SlideShowWindow.View.Exit();
                    activeAlready = MsoTriState.msoFalse;
                }
            }
            else if (v.Equals("0"))
                if (activeAlready == MsoTriState.msoTrue)
                    powerPoint.ActivePresentation.SlideShowWindow.View.Previous();
                else
                    performCommand("startPresentation");
            else if (v.Equals("1"))
                if (activeAlready == MsoTriState.msoTrue)
                    powerPoint.ActivePresentation.SlideShowWindow.View.Next();
                else
                    performCommand("startPresentation");
            else if (v.Contains("index"))
            {
                if (activeAlready == MsoTriState.msoTrue)
                {
                    int index = int.Parse(v.ToString().Split(new Char[] { ':' })[1]);
                    System.Diagnostics.Debug.WriteLine("" + index);
                    powerPoint.ActivePresentation.SlideShowWindow.View.GotoSlide(index);
                }
                else
                {
                    performCommand("startPresentation");
                    int index = int.Parse(v.ToString().Split(new Char[] { ':' })[1]);
                    System.Diagnostics.Debug.WriteLine("" + index);
                    powerPoint.ActivePresentation.SlideShowWindow.View.GotoSlide(index);
                }
            }
            else if (v.Equals("stop"))
            {
                if (activeAlready == MsoTriState.msoTrue)
                {
                    performCommand("stopPresentation");
                }

                if (activeAlready == MsoTriState.msoTrue || backgroundWorker1.IsBusy)
                {
                    backgroundWorker1.CancelAsync();
                    activeAlready = MsoTriState.msoCTrue;
                    System.Windows.Forms.MessageBox.Show("Connection stopped.",
                            "Information",
                            System.Windows.Forms.MessageBoxButtons.OK,
                            System.Windows.Forms.MessageBoxIcon.Information);
                }
                else
                {
                    System.Windows.Forms.MessageBox.Show("Not connected.",
                            "Information",
                            System.Windows.Forms.MessageBoxButtons.OK,
                            System.Windows.Forms.MessageBoxIcon.Information);
                }
            }
        }

        private void backgroundWorker2_DoWork(object sender, System.ComponentModel.DoWorkEventArgs e)
        {
            slides = "";
            foreach (Slide slide in powerPoint.ActivePresentation.Slides)
            {
                if (slide.Shapes.HasTitle == MsoTriState.msoTrue)
                {
                    if (slide.Shapes.Title.TextFrame.HasText == MsoTriState.msoTrue)
                    {
                        if (!slide.Shapes.Title.TextFrame.TextRange.Text.Equals(""))
                        {
                            int count = Regex.Matches(slides, slide.Shapes.Title.TextFrame.TextRange.Text).Count;
                            if (count > 0)
                            {
                                slides = slides + slide.Shapes.Title.TextFrame.TextRange.Text.Replace(";", " ").Trim() + "(" + count + ")";
                            }
                            else
                            {
                                slides = slides + slide.Shapes.Title.TextFrame.TextRange.Text.Replace(";", " ").Trim();
                            }
                        }
                    }
                }
                else
                {
                    slides = slides + slide.Name;
                }
                slides = slides + ";";
            }
            exportAndSlideNames();
            //new Thread(exportAndSlideNames).Start();
        }
    }
}
