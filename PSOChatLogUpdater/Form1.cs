using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Linq.Expressions;
using System.Runtime.CompilerServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;

namespace PSOChatLogUpdater
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        private void Form1_Load(object sender, EventArgs e)
        {
            //PSOChatLog_ver0.55.zip 0.55
            //string[] strArgs = new string[15];
            long lArgv = 0;
            do
            {
                //string CommandStr = System.Environment.CommandLine;
                string[] strArgs = System.Environment.GetCommandLineArgs();
                lArgv = strArgs.Length;

                stop();
                if (lArgv == 1)
                {
                    ReadyToRunUpdateManual();
                    break;
                }
                stop();
                var strPathArchive = System.Environment.CurrentDirectory + "\\Archive\\" + strArgs[1];
                var strPathExtract = System.Environment.CurrentDirectory + "\\Extract\\" + strArgs[2];
                doExtractArchive(strPathArchive, strPathExtract);
                timer1.Interval = 1000;
                timer1.Start();
            } while (false);
        }
        public int GetProcessesByWindowTitle(string windowTitle, List<string> strProcessName)
        {
            System.Collections.ArrayList list = new System.Collections.ArrayList();
            //System.Diagnostics.Debugger.Break();
            int iRet = 0;
            //すべてのプロセスを列挙する
            System.Diagnostics.Process[] p;
            p = System.Diagnostics.Process.GetProcesses();
            for (int i = 0; i < p.Length; i++)
            {
                //指定された文字列がメインウィンドウのタイトルに含まれているか調べる
                if (0 <= p[i].MainWindowTitle.IndexOf(windowTitle))
                {
                    //System.Diagnostics.Debugger.Break();
                    strProcessName.Add(p[i].MainWindowTitle);
                    listBox1.Items.Add("PSO Chat Log Process = " + iRet + ", " + strProcessName[iRet]);
                    listBox1.TopIndex = listBox1.Items.Count - 1;
                    iRet++;
                }
            }
            return (iRet);
        }
        private void timer1_Tick(object sender, EventArgs e)
        {
            var bBreakFlg = false;
            do
            {
                var strProcessName = new List<string>();
                //ウィンドウのタイトルに「PSOChatlog 」を含むプロセスをすべて取得する
                var iCount = GetProcessesByWindowTitle("PSO Chat Log", strProcessName);
                //System.Diagnostics.Debugger.Break();
                if (iCount == 0)
                {
                    break;
                }
                for (int i = 0; i < iCount; i++)
                {
                    listBox1.Items.Add("PSO Chat Log Process Live Count = " + iCount.ToString());
                    listBox1.TopIndex = listBox1.Items.Count - 1;
                    listBox1.Items.Add("Waiting...");
                    listBox1.TopIndex = listBox1.Items.Count - 1;
                    bBreakFlg = true;
                    break;
                }
            } while (false);
            if (bBreakFlg == false)
            {
                doUpdate(Value.folderFrom, Value.folderTo);
            }
        }
        private void doExtractArchive(string strPathArchive, string strPathExtract)
        {
            Value.folderFrom = strPathExtract + "\\PSOChatLog";
            Value.folderTo = System.Environment.CurrentDirectory;

            var strMsg = "Extract Success";
            var directoryInfo = Directory.CreateDirectory(strPathExtract);
            do
            {
                //ファイルを解凍する。
                try
                {
                    System.IO.Compression.ZipFile.ExtractToDirectory(strPathArchive, strPathExtract, System.Text.Encoding.GetEncoding("shift_jis"));
                }
                catch (System.Exception eZipFile)
                {
                    strMsg = "";
                    listBox1.Items.Add(eZipFile.ToString().Replace(System.Environment.CurrentDirectory, ""));
                    listBox1.TopIndex = listBox1.Items.Count - 1;
                }
                if (strMsg == "")
                {
                    break;
                }
                listBox1.Items.Add("Extract Success !");
                listBox1.TopIndex = listBox1.Items.Count - 1;
            } while (false);
        }
        private void ReadyToRunUpdateManual()
        {
            stop();
            //string[] strPathVer = System.IO.Directory.GetDirectories(System.Environment.CurrentDirectory + "\\Extract\\", "*", System.IO.SearchOption.AllDirectories);
            button2.Enabled = false;
            do
            {
                //if (File.Exists(System.Environment.CurrentDirectory + "\\Archive\\") == false)
                if (System.IO.Directory.Exists(System.Environment.CurrentDirectory + "\\Archive\\") == false)
                {
                    break;
                }
                string[] strPathArchive = System.IO.Directory.GetFiles(System.Environment.CurrentDirectory + "\\Archive\\", "*.zip", System.IO.SearchOption.TopDirectoryOnly);
                for (int i = 0; i < strPathArchive.Length; i++)
                {
                    var strPathExtract = "";
                    strPathExtract = strPathArchive[i].Replace(System.Environment.CurrentDirectory + "\\Archive\\PSOChatLog_ver", "");
                    strPathExtract = strPathExtract.Replace(".zip", "");
                    doExtractArchive(strPathArchive[i], System.Environment.CurrentDirectory + "\\Extract\\" + strPathExtract);
                }
                stop();
                string[] strPathVercheck = System.IO.Directory.GetDirectories(System.Environment.CurrentDirectory + "\\Extract\\", "*", System.IO.SearchOption.TopDirectoryOnly);
                for (int i = 0; i < strPathVercheck.Length; i++)
                {
                    strPathVercheck[i] = strPathVercheck[i].Replace(System.Environment.CurrentDirectory + "\\Extract\\", "");
                }
                comboBox1.DataSource = strPathVercheck;
                comboBox1.Text = strPathVercheck[strPathVercheck.Length - 1];
                button2.Enabled = true;
            } while (false);
        }
        private void doUpdate(string folderFrom, string folderTo)
        {
            //autofrom//"C:\\Users\\admin\\source\\repos\\PSOChatLogUpdater\\PSOChatLogUpdater\\bin\\Debug\\Extract\\0.55\\PSOChatLog"
            //autoto//"C:\\Users\\admin\\source\\repos\\PSOChatLogUpdater\\PSOChatLogUpdater\\bin\\Debug"
            listBox1.Items.Add("Update");
            listBox1.TopIndex = listBox1.Items.Count - 1;

            FileCopy(folderFrom, folderTo);
            string[] pathDir = System.IO.Directory.GetDirectories(folderFrom, "*", System.IO.SearchOption.AllDirectories);
            for (int i = 0; i < pathDir.Length; i++)
            {
                //var strDifference = pathDir[i].Replace(folderFrom + "\\", "");
                var strDifference = pathDir[i].Replace(folderFrom, "");

                FileCopy(pathDir[i], folderTo + strDifference);
            }
            listBox1.Items.Add("Update Finish !");
            listBox1.TopIndex = listBox1.Items.Count - 1;
            this.button1.Enabled = true;
            timer1.Stop();
        }
        private void FileCopy(string folderFrom, string folderTo)
        {
            var strMsgFrom = "";
            var strMsgTo = "";
            var myCheck = false;
            var bForceUpdateFlg = true;

            string[] pathFrom = System.IO.Directory.GetFiles(folderFrom);
            for (int j = 0; j < pathFrom.Count(); j++)
            {
                //コピー先のフォルダーが存在するか確認し、なければ作成します。
                string pathTo = pathFrom[j].Replace(folderFrom, folderTo);
                string targetFolder = System.IO.Path.GetDirectoryName(pathTo);
                if (System.IO.Directory.Exists(targetFolder) == false)
                {
                    listBox1.Items.Add("CreateDirectory = " + targetFolder);
                    listBox1.TopIndex = listBox1.Items.Count - 1;
                    System.IO.Directory.CreateDirectory(targetFolder);
                }

                //１ファイルのコピー実行。同名のファイルがある場合上書きします。
                System.Diagnostics.Debug.WriteLine("Copy = " + pathFrom + " → " + pathTo);
                strMsgFrom = pathFrom[j].Replace(System.Environment.CurrentDirectory, "");
                strMsgTo = pathTo.Replace(System.Environment.CurrentDirectory + "\\", "");
                listBox1.Items.Add("copying ..." + strMsgFrom + " → " + strMsgTo);
                listBox1.TopIndex = listBox1.Items.Count - 1;

                //CSVファイルなら上書きしない
                myCheck = Regex.IsMatch(pathFrom[j], ".+.csv", RegexOptions.Singleline);
                if (myCheck == true)
                {
                    listBox1.Items.Add("csv = " + strMsgFrom + " → " + strMsgTo);
                    bForceUpdateFlg = false;
                }

                //ファイルをコピー
                try
                {
                    System.IO.File.Copy(pathFrom[j], pathTo, bForceUpdateFlg);
                }
                catch
                {
                    System.Diagnostics.Debugger.Break();

                    listBox1.Items.Add("error = " + strMsgFrom + " → " + strMsgTo);
                    listBox1.TopIndex = listBox1.Items.Count - 1;
                }
            }
        }
        private void DoResize()
        {
            //リサイズ終了後、再描画とレイアウトロジックを実行する
            //520x480 : 480x400
            listBox1.Width = this.Width - 40;//640時600
            listBox1.Height = this.Height - 188;//480時292
            button1.Top = this.Height - 110;//480時370
            //button1.Left = 12;
            button1.Width = this.Width - 132;//520時388
            //button1.Height = 59;//19
        }
        private void button2_Click(object sender, EventArgs e)
        {
            button2.Enabled = false;
            doUpdate(System.Environment.CurrentDirectory + "\\Extract\\" + comboBox1.Text + "\\PSOChatLog", System.Environment.CurrentDirectory);
            button2.Enabled = true;
        }
        private void button1_Click(object sender, EventArgs e)
        {
            this.Close();
        }
        private void stop()
        {
            //System.Diagnostics.Debugger.Break();
        }
        private void Form1_Resize(object sender, EventArgs e)
        {
            DoResize();
        }
        private void Form1_ResizeEnd(object sender, EventArgs e)
        {
            DoResize();
        }
    }
    public static class Value
    {
        public static string folderFrom;
        public static string folderTo;
    }
}
