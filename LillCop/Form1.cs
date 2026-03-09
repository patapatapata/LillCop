using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using System.IO;
using System.Reflection;
using System.Diagnostics;
using System.Runtime.CompilerServices;
using System.Threading;
using Microsoft.VisualBasic;
using Microsoft.Win32;
using Microsoft.VisualBasic.CompilerServices;

namespace LillCop
{
    public partial class Form1 : Form
    {
        static public Type type;// = Type.GetTypeFromProgID("Illustrator.Application." + illVer, true);
        static public dynamic illApp;// = Activator.CreateInstance(type, true);

        public Form1()
        {
            InitializeComponent();
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            timer1.Start();
            Properties.Settings.Default.Reload();
            buttonApp1.Text = Properties.Settings.Default.app1name;
            buttonApp2.Text = Properties.Settings.Default.app2name;
            buttonApp3.Text = Properties.Settings.Default.app3name;
            buttonApp4.Text = Properties.Settings.Default.app4name;
            Form1.CheckForIllegalCrossThreadCalls = false;
            label4.Text = "";
            startSetting f = new startSetting();
            f.ShowDialog(this);
            f.Dispose();
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            fileCheck();
        }


        #region 汎用
        private void linkPathCheck(ref String myPath)
        {
            dynamic illSel;
            illSel = illApp.ActiveDocument.Selection[0];
            if (illSel.Typename == "PlacedItem")
            {
                myPath = illSel.File;
                return;
            }
            else if (illSel.Typename == "RasterItem")
            {
                try
                {
                    myPath = "【埋込】" + illSel.File;
                }
                catch
                {
                    myPath = @"【埋込】\";
                }
                return;
            }
            else if (illSel.Typename == "GroupItem")
            {
                if (illSel.PlacedItems.Count == 1)
                {
                    myPath = illSel.PlacedItems[1].File;
                    return;
                }
                else if (illSel.RasterItems.Count == 1)
                {
                    try
                    {
                        myPath = "【埋込】" + illSel.RasterItems[1].File;
                    }
                    catch
                    {
                        myPath = @"【埋込】\";
                    }
                    return;
                }
                while (illSel.GroupItems[1].GroupItems.Count > 0)
                {
                    illSel = illSel.GroupItems[1];
                    System.Windows.Forms.Application.DoEvents();
                }
                if (illSel.GroupItems[1].PlacedItems.Count == 1)
                {
                    myPath = illSel.GroupItems[1].PlacedItems[1].File;
                }
                else if (illSel.GroupItems[1].RasterItems.Count == 1)
                {
                    try
                    {
                        myPath = "【埋込】" + illSel.GroupItems[1].RasterItems[1].File;
                    }
                    catch
                    {
                        myPath = @"【埋込】\";
                    }
                }
            }
        }

        private void illReLink()
        {
            dynamic illSel;
            illSel = illApp.ActiveDocument.Selection[0];
            if (illSel.Typename == "PlacedItem")
            {
                illSel.File = textBox3.Text + "\\" + textBox2.Text;
                return;
            }
            else if (illSel.Typename == "RasterItem")
            {
                rasterRelink(illSel);
                return;
            }
            else if (illSel.Typename == "GroupItem")
            {
                if (illSel.PlacedItems.Count == 1)
                {
                    illSel.PlacedItems[1].File = textBox3.Text + "\\" + textBox2.Text;
                    return;
                }
                else if (illSel.RasterItems.Count == 1)
                {
                    rasterRelink(illSel.RasterItems[1]);
                    return;
                }
                while (illSel.GroupItems[1].GroupItems.Count > 0)
                {
                    illSel = illSel.GroupItems[1];
                }
                if (illSel.GroupItems[1].PlacedItems.Count == 1)
                {
                    illSel.GroupItems[1].PlacedItems[1].File = textBox3.Text + "\\" + textBox2.Text;
                }
                else if (illSel.GroupItems[1].RasterItems.Count == 1)
                {
                    rasterRelink(illSel.GroupItems[1].RasterItems[1]);
                }
            }
        }

        private void rasterRelink(dynamic rasterTarget)
        {
            rasterTarget.Embedded = false;
            rasterTarget.Copy();
            dynamic tmpGroup = rasterTarget.parent;
            tmpGroup.Paste();
            tmpGroup.PlacedItems[1].Position = rasterTarget.Position;
            tmpGroup.PlacedItems[1].Move(rasterTarget, illApp.AiElementPlacement.aiPlaceAfter);
            tmpGroup.PlacedItems[1].File = textBox3.Text + "\\" + textBox2.Text;
            rasterTarget.Delete();
            /*
            dynamic replacedImage = illApp.ActiveDocument.PlacedItems.Add();
            replacedImage.File = textBox3.Text + "\\" + textBox2.Text;
            replacedImage.Position = rasterTarget.Position;
            replacedImage.Width = rasterTarget.Width;
            replacedImage.Height = rasterTarget.Height;
            replacedImage.Move(rasterTarget, AiElementPlacement.aiPlaceAfter);
            rasterTarget.Delete();*/
        }

        private void fileCheck()
        {
            timer1.Stop();
            try
            {
                String myPath = "";
                linkPathCheck(ref myPath);
                if (textBox1.Text != myPath)
                {
                    pictureBox1.Visible = false;
                    pictureBox2.Visible = false;
                    pictureSame.Visible = false;
                    textBox1.Text = myPath;
                    textBox2.Text = System.IO.Path.GetFileName(myPath);
                    show.Visible = true;
                    ExpShow.Visible = true;
                    label4.Text = "";
                    if (System.IO.File.Exists(textBox1.Text) == true)
                    {
                        if (!backgroundWorker1.IsBusy)
                        {
                            // バックグラウンド処理を開始
                            backgroundWorker1.RunWorkerAsync();
                        }
                        else
                        {
                            backgroundWorker1.CancelAsync();
                            backgroundWorker1.RunWorkerAsync();
                        }
                    }
                }
                if (System.IO.File.Exists(myPath) == false && myPath != "")
                {
                    pictureBox1.Visible = true;
                    label4.Text = "";
                }
                else
                {
                    pictureBox1.Visible = false;
                }
                SameOrWarning();
            }
            catch
            {
                textBox1.Text = "";
                textBox2.Text = "";
                label4.Text = "";
                show.Visible = false;
                ExpShow.Visible = false;
                pictureBox1.Visible = false;
                pictureBox2.Visible = false;
                pictureSame.Visible = false;
            }
            timer1.Start();
        }
        #endregion;

        private void textBox3_DragDrop(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                pictureBox2.Visible = false;
                pictureSame.Visible = false;
                string[] fileName = (string[])e.Data.GetData(DataFormats.FileDrop);
                textBox3.Text = fileName[0].ToString();
                if (System.IO.File.Exists(fileName[0]) == true)
                {
                    textBox3.Text = System.IO.Path.GetDirectoryName(fileName[0]);
                    SameOrWarning();
                }
                else
                {
                    SameOrWarning();
                }
            }
        }

        private void textBox3_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                e.Effect = DragDropEffects.Copy;
            }
        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {
            SameOrWarning();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (textBox3.Text == "")
            {
                System.Media.SystemSounds.Beep.Play();
                MessageBox.Show("コピー先を設定してください。");
                return;
            }
            //if (textBox1.Text == textBox3.Text + "\\" + textBox2.Text)
            if (String.Compare(textBox1.Text, textBox3.Text + "\\" + textBox2.Text, true) == 0)
            {
                System.Media.SystemSounds.Beep.Play();
                MessageBox.Show("コピー元とコピー先が同じです。");
                return;
            }
            timer1.Stop();
            if (System.IO.File.Exists(textBox1.Text) == true && System.IO.File.Exists(textBox3.Text + "\\" + textBox2.Text) == false)
            {
                System.IO.File.Copy(textBox1.Text, textBox3.Text + "\\" + textBox2.Text);
                illReLink();
            }
            else if (System.IO.File.Exists(textBox1.Text) == false && System.IO.File.Exists(textBox3.Text + "\\" + textBox2.Text) == true)
            {
                System.Media.SystemSounds.Beep.Play();
                DialogResult result = MessageBox.Show("現在のリンクファイルは見つかりませんが\r\n同名のファイルがコピー先に存在しています。\r\n既存ファイルにリンク先を変更しますか？", "",
                                                      MessageBoxButtons.YesNo,
                                                      MessageBoxIcon.Exclamation,
                                                      MessageBoxDefaultButton.Button2);
                //何が選択されたか調べる
                if (result == DialogResult.Yes)
                {
                    illReLink();
                }
                if (result == DialogResult.No)
                {
                    //「いいえ」が選択された時
                }
            }
            else if (System.IO.File.Exists(textBox1.Text) == true && System.IO.File.Exists(textBox3.Text + "\\" + textBox2.Text) == true)
            {
                DateTime dtUpdate = System.IO.File.GetLastWriteTime(textBox3.Text + "\\" + textBox2.Text);
                DateTime dtUpdate2 = System.IO.File.GetLastWriteTime(textBox1.Text);
                System.Media.SystemSounds.Beep.Play();
                DialogResult result = MessageBox.Show("コピー先に同名ファイルが存在しています。\r\n既存ファイルにリンク先を変更しますか？\r\n\r\n既存ファイル：" + dtUpdate + "\r\n現在のリンク：" + dtUpdate2, "",
                                                      MessageBoxButtons.YesNo,
                                                      MessageBoxIcon.Exclamation,
                                                      MessageBoxDefaultButton.Button2);
                //何が選択されたか調べる
                if (result == DialogResult.Yes)
                {
                    illReLink();
                }
                if (result == DialogResult.No)
                {
                    //「いいえ」が選択された時
                }
            }
            else if (System.IO.File.Exists(textBox1.Text) == false && System.IO.File.Exists(textBox3.Text + "\\" + textBox2.Text) == false)
            {
                System.Media.SystemSounds.Beep.Play();
                MessageBox.Show("リンクファイルが見つかりません。");
            }
            timer1.Start();
        }

        private void backgroundWorker1_DoWork(object sender, DoWorkEventArgs e)
        {
            string tmpEx = System.IO.Path.GetExtension(textBox1.Text);
            tmpEx = tmpEx.ToUpper();
            if (tmpEx == ".AI")
            {
                string tmpStr = "";
                string tmpStr2 = "";
                System.Text.RegularExpressions.Regex findStr1 = new System.Text.RegularExpressions.Regex(@"^%%Creator: ");
                System.Text.RegularExpressions.Regex findStr2 = new System.Text.RegularExpressions.Regex(@"^%%AI8_CreatorVersion: ");
                System.Text.RegularExpressions.Regex findStr3 = new System.Text.RegularExpressions.Regex(@"^/Producer \(Adobe PDF library ");
                System.Text.RegularExpressions.Regex findStr4 = new System.Text.RegularExpressions.Regex(@"^%AI9_ColorModel: ");
                System.Text.RegularExpressions.Regex findStr5 = new System.Text.RegularExpressions.Regex(@"^%%DocumentCustomColors: .");
                System.Text.RegularExpressions.Regex findStr6 = new System.Text.RegularExpressions.Regex(@"^%%DocumentFiles:.");
                System.Text.RegularExpressions.Regex findAIStr2 = new System.Text.RegularExpressions.Regex(@">>stream$");
                System.Text.RegularExpressions.Regex findAIStr3 = new System.Text.RegularExpressions.Regex(@"^stream$");
                System.Text.RegularExpressions.Regex findAIStr4 = new System.Text.RegularExpressions.Regex(@"^%!");
                System.Text.RegularExpressions.Regex findAIStr5 = new System.Text.RegularExpressions.Regex(@"^<\?");
                using (FileStream fs = new FileStream(textBox1.Text, FileMode.Open, FileAccess.Read, FileShare.Read, 1024, FileOptions.SequentialScan))
                {
                    using (StreamReader sr = new StreamReader(fs, Encoding.GetEncoding("Shift_JIS")))
                    {
                        string strLine;
                        while ((strLine = sr.ReadLine()) != null)
                        {
                            if (findAIStr2.IsMatch(strLine) || findAIStr3.IsMatch(strLine))
                            {
                                strLine = sr.ReadLine();
                                if (findAIStr4.IsMatch(strLine))
                                {
                                    while (strLine != "%%EndComments")
                                    {
                                        strLine = sr.ReadLine();
                                        if (findStr1.IsMatch(strLine))
                                        {
                                            tmpStr = findStr1.Replace(strLine, "");
                                            tmpStr = tmpStr.Replace("Adobe ", "");
                                            tmpStr = tmpStr.Replace("(R)", "");
                                            strLine = sr.ReadLine();
                                            tmpStr2 = findStr2.Replace(strLine, "");
                                            if (tmpStr2.IndexOf("%%Title: ") > -1)
                                            {
                                                tmpStr2 = "EPS？";
                                            }
                                        }
                                    }
                                }
                                else
                                {
                                    if (findAIStr5.IsMatch(strLine) == false)
                                    {
                                        while (strLine != "endobj")
                                        {
                                            strLine = sr.ReadLine();
                                        }
                                    }
                                }
                            }
                            if (findStr1.IsMatch(strLine))
                            {
                                tmpStr = findStr1.Replace(strLine, "");
                                tmpStr = tmpStr.Replace("Adobe ", "");
                                tmpStr = tmpStr.Replace("(R)", "");
                                strLine = sr.ReadLine();
                                tmpStr2 = findStr2.Replace(strLine, "");
                                if (tmpStr2.IndexOf("%%Title: ") > -1)
                                {
                                    tmpStr2 = "EPS？";
                                }
                            }
                        }
                        label4.Text = tmpStr + " / " + tmpStr2;
                    }
                }
            }
            if (tmpEx == ".EPS")
            {
                string strLine;
                string tmpStr = "";
                string tmpStr2 = "";
                int errCheckEps = 1;
                System.Text.RegularExpressions.Regex findStr2 = new System.Text.RegularExpressions.Regex(@"^%%Creator: ");
                System.Text.RegularExpressions.Regex findStr3 = new System.Text.RegularExpressions.Regex(@"^%%AI8_CreatorVersion: ");
                System.Text.RegularExpressions.Regex findStr5 = new System.Text.RegularExpressions.Regex(@"^%%Creator: Adobe Illustrator");
                System.Text.RegularExpressions.Regex findStr6 = new System.Text.RegularExpressions.Regex(@"^%%Creator: Adobe Photoshop");
                System.Text.RegularExpressions.Regex findStr8 = new System.Text.RegularExpressions.Regex(@"^%%Creator: Adobe InDesign");
                using (FileStream fs = new FileStream(textBox1.Text, FileMode.Open, FileAccess.Read, FileShare.Read, 1024, FileOptions.SequentialScan))
                {
                    using (StreamReader sr = new StreamReader(fs, Encoding.GetEncoding("Shift_JIS")))
                    {
                        while ((strLine = sr.ReadLine()) != null)
                        {
                            try
                            {
                                if (findStr5.IsMatch(strLine))
                                {
                                    errCheckEps = 0;
                                    tmpStr = findStr2.Replace(strLine, "");
                                    tmpStr = tmpStr.Replace("Adobe ", "");
                                    tmpStr = tmpStr.Replace("(R)", "");
                                    strLine = sr.ReadLine();
                                    tmpStr2 = findStr3.Replace(strLine, "");
                                    break;
                                }
                                if (findStr6.IsMatch(strLine))
                                {
                                    errCheckEps = 2;
                                    tmpStr = strLine.Replace("%%Creator: Adobe ", "");
                                    tmpStr = tmpStr.Replace("Version ", "");
                                    break;
                                }
                                if (findStr8.IsMatch(strLine))
                                {
                                    tmpStr = strLine.Replace("%%Creator: Adobe ", "");
                                    label4.Text = tmpStr;
                                    return;
                                }
                            }
                            catch
                            {
                                label4.Text = "不明";
                                return;
                            }
                        }
                        if (tmpStr == "")
                        {
                            label4.Text = "不明";
                            return;
                        }
                        //Illustratorの場合
                        if (errCheckEps == 0)
                        {
                            System.Text.RegularExpressions.Regex findStr1 = new System.Text.RegularExpressions.Regex(@"^userdict /ai9_skip_data get exec");
                            System.Text.RegularExpressions.Regex findStr9 = new System.Text.RegularExpressions.Regex(@"^%%DocumentCustomColors: .");
                            System.Text.RegularExpressions.Regex findStr11 = new System.Text.RegularExpressions.Regex(@"^[\d\.]+ [\d\.]+ [\d\.]+ rgb$");
                            System.Text.RegularExpressions.Regex findStr12 = new System.Text.RegularExpressions.Regex(@"^%%BeginDocument: ");
                            System.Text.RegularExpressions.Regex findStr15 = new System.Text.RegularExpressions.Regex(@"^%%DocumentProcessColors:  ");
                            System.Text.RegularExpressions.Regex findAIEPSStr1 = new System.Text.RegularExpressions.Regex(@"^%%BeginBinary: ");
                            System.Text.RegularExpressions.Regex findAIEPSStr2 = new System.Text.RegularExpressions.Regex(@"^%BeginPhotoshop: ");
                            System.Text.RegularExpressions.Regex findAIEPSStr3 = new System.Text.RegularExpressions.Regex(@"^%%BeginICCProfile: ");
                            System.Text.RegularExpressions.Regex findAIEPSStr4 = new System.Text.RegularExpressions.Regex(@"^%%BeginResource: ");
                            while ((strLine = sr.ReadLine()) != null)
                            {
                                if (findAIEPSStr1.IsMatch(strLine))
                                {
                                    while (strLine != "%%EndBinary")
                                    {
                                        strLine = sr.ReadLine();
                                    }
                                }
                                if (findAIEPSStr2.IsMatch(strLine))
                                {
                                    while (strLine != "%EndPhotoshop")
                                    {
                                        strLine = sr.ReadLine();
                                    }
                                }
                                if (findAIEPSStr3.IsMatch(strLine))
                                {
                                    while (strLine != "%%EndICCProfile")
                                    {
                                        strLine = sr.ReadLine();
                                    }
                                }
                                if (findAIEPSStr4.IsMatch(strLine))
                                {
                                    while (strLine != "%%EndResource")
                                    {
                                        strLine = sr.ReadLine();
                                    }
                                }
                                if (findStr1.IsMatch(strLine))
                                {
                                    sr.ReadLine();
                                    sr.ReadLine();
                                    strLine = sr.ReadLine();
                                    tmpStr = findStr2.Replace(strLine, "");
                                    tmpStr = tmpStr.Replace("Adobe ", "");
                                    tmpStr = tmpStr.Replace("(R)", "");
                                    tmpStr = tmpStr.Replace("(TM)", "");
                                    strLine = sr.ReadLine();
                                    tmpStr2 = findStr3.Replace(strLine, "");
                                    while (strLine != "%AI9_PrivateDataEnd")
                                    {
                                        strLine = sr.ReadLine();
                                    }
                                }
                            }
                            label4.Text = tmpStr + " / " + tmpStr2;
                            return;
                        }
                        //PhotoShopの場合
                        if (errCheckEps == 2)
                        {
                            bool tmpPnum = true;
                            System.Text.RegularExpressions.Regex findStr4 = new System.Text.RegularExpressions.Regex(@"^ +<stEvt:softwareAgent>");
                            System.Text.RegularExpressions.Regex findStr7 = new System.Text.RegularExpressions.Regex(@"^ +<xap:CreatorTool>");
                            System.Text.RegularExpressions.Regex findStr13 = new System.Text.RegularExpressions.Regex(@"^ +<photoshop:ColorMode>");
                            System.Text.RegularExpressions.Regex findStr14 = new System.Text.RegularExpressions.Regex(@"^ +<photoshop:ICCProfile>");
                            System.Text.RegularExpressions.Regex findPhStr1 = new System.Text.RegularExpressions.Regex(@"^%%BeginBinary:");
                            while ((strLine = sr.ReadLine()) != null)
                            {
                                if (findStr4.IsMatch(strLine) && tmpPnum)
                                {
                                    tmpStr = findStr4.Replace(strLine, "");
                                    tmpStr = tmpStr.Replace("</stEvt:softwareAgent>", "");
                                    tmpStr = tmpStr.Replace("Adobe ", "");
                                    tmpPnum = false;
                                }
                                if (findStr7.IsMatch(strLine) && tmpPnum)
                                {
                                    tmpStr = findStr7.Replace(strLine, "");
                                    tmpStr = tmpStr.Replace("</xap:CreatorTool>", "");
                                    tmpStr = tmpStr.Replace("Adobe ", "");
                                    tmpPnum = false;
                                }
                                if (findPhStr1.IsMatch(strLine))
                                {
                                    label4.Text = tmpStr;
                                    return;
                                }
                            }
                            label4.Text = tmpStr;
                        }
                    }
                }
            }
        }

        private void button6_Click(object sender, EventArgs e)
        {
            label3.Focus();
            if (textBox1.Text != "")
            {
                string tmpText = System.Text.RegularExpressions.Regex.Replace(textBox1.Text, "^【埋込】", "");
                string folderName = System.IO.Path.GetDirectoryName(tmpText);
                if (System.IO.Directory.Exists(folderName) == true)
                {
                    textBox3.Text = folderName;
                }
            }
        }

        private void button5_Click(object sender, EventArgs e)
        {
            if (textBox3.Text == "")
            {
                System.Media.SystemSounds.Beep.Play();
                MessageBox.Show("コピー先を設定してください。");
                return;
            }
            //if (textBox1.Text == textBox3.Text + "\\" + textBox2.Text)
            if (String.Compare(textBox1.Text, textBox3.Text + "\\" + textBox2.Text, true) == 0)
            {
                System.Media.SystemSounds.Beep.Play();
                MessageBox.Show("コピー元とコピー先が同じです。");
                return;
            }
            timer1.Stop();
            if (System.IO.File.Exists(textBox1.Text) == true && System.IO.File.Exists(textBox3.Text + "\\" + textBox2.Text) == false)
            {
                System.IO.File.Copy(textBox1.Text, textBox3.Text + "\\" + textBox2.Text);
            }
            else if (System.IO.File.Exists(textBox1.Text) == false)
            {
                System.Media.SystemSounds.Beep.Play();
                MessageBox.Show("ファイルが見つかりません。");
            }
            else if (System.IO.File.Exists(textBox1.Text) == true && System.IO.File.Exists(textBox3.Text + "\\" + textBox2.Text) == true)
            {
                System.Media.SystemSounds.Beep.Play();
                MessageBox.Show("同名のファイルが存在しています。");
            }
            timer1.Start();
        }

        private void Folder_Click(object sender, EventArgs e)
        {
            //FolderBrowserDialogクラスのインスタンスを作成
            FolderBrowserDialog fbd = new FolderBrowserDialog();

            //上部に表示する説明テキストを指定する
            //fbd.Description = "コピー先を選択してください。";
            //ルートフォルダを指定する
            //デフォルトでDesktop
            fbd.RootFolder = Environment.SpecialFolder.Desktop;
            //最初に選択するフォルダを指定する
            //RootFolder以下にあるフォルダである必要がある
            fbd.SelectedPath = textBox3.Text;
            //ユーザーが新しいフォルダを作成できるようにする
            //デフォルトでTrue
            fbd.ShowNewFolderButton = true;

            //ダイアログを表示する
            if (fbd.ShowDialog(this) == DialogResult.OK)
            {
                //選択されたフォルダを表示する
                textBox3.Text = fbd.SelectedPath;
                //Console.WriteLine(fbd.SelectedPath);
            }
        }

        private void button9_Click(object sender, EventArgs e)
        {
            label3.Focus();
            dynamic illSel;
            illSel = Form1.illApp.ActiveDocument.Selection[0];
            if (illSel.Typename == "PlacedItem")
            {
                linkItemShow(illSel);
            }
            else if (illSel.Typename == "GroupItem")
            {
                if (illSel.PlacedItems.Count > 0)
                {
                    linkItemShow(illSel.PlacedItems[1]);
                }
                else if (illSel.RasterItems.Count > 0)
                {
                    linkItemShow(illSel.RasterItems[1]);
                }
                else if (illSel.GroupItems.Count > 0)
                {
                    linkItemShow(illSel.GroupItems[1].RasterItems[1]);
                }
            }
        }

        private void linkItemShow(dynamic linkItemTarget)
        {
            dynamic[] cntrPos = { linkItemTarget.VisibleBounds[0] + (linkItemTarget.VisibleBounds[2] - linkItemTarget.VisibleBounds[0])/2, linkItemTarget.VisibleBounds[1] + (linkItemTarget.VisibleBounds[3]- linkItemTarget.VisibleBounds[1])/2 };
            illApp.ActiveDocument.Views[1].CenterPoint = cntrPos;
        }

        private void ExpShow_Click(object sender, EventArgs e)
        {
            label3.Focus();
            System.Diagnostics.Process.Start("EXPLORER.EXE", @"/select," + textBox1.Text);
        }

        private void button2_Click_1(object sender, EventArgs e)
        {
            if (textBox3.Text == "")
            {
                System.Media.SystemSounds.Beep.Play();
                MessageBox.Show("フォルダを選択してください。");
                return;
            }
            //if (textBox1.Text == textBox3.Text + "\\" + textBox2.Text)
            if (String.Compare(textBox1.Text, textBox3.Text + "\\" + textBox2.Text, true) == 0)
            {
                System.Media.SystemSounds.Beep.Play();
                MessageBox.Show("同じ名前です。");
                return;
            }
            timer1.Stop();
            if (System.IO.File.Exists(textBox1.Text) == true && System.IO.File.Exists(textBox3.Text + "\\" + textBox2.Text) == false)
            {
                string folderName = System.IO.Path.GetDirectoryName(textBox1.Text);
                if (folderName != textBox3.Text)
                {
                    System.Media.SystemSounds.Beep.Play();
                    DialogResult tmpResult = MessageBox.Show("リネーム先のフォルダ設定は正しいですか？\r\n現在の設定ではファイルの移動が行われます。\r\nこのまま続行しますか？", "",
                                                          MessageBoxButtons.YesNo,
                                                          MessageBoxIcon.Exclamation,
                                                          MessageBoxDefaultButton.Button2);
                    //何が選択されたか調べる
                    if (tmpResult == DialogResult.Yes)
                    {

                    }
                    if (tmpResult == DialogResult.No)
                    {
                        timer1.Start();
                        return;
                    }
                }
                dynamic illGraphics;
                illGraphics = illApp.ActiveDocument.PlacedItems;
                int tmpVal = 0;
                for (int i = 1; i < illGraphics.Count+1; i++)
                {
                    try
                    {
                        //if (textBox1.Text == illGraphics[i].File)
                        if (String.Compare(textBox1.Text, illGraphics[i].File, true) == 0)
                        {
                            tmpVal++;
                            if (tmpVal > 1)
                            {
                                break;
                            }
                        }
                    }
                    catch
                    { }
                }
                if (tmpVal == 1)
                {
                    System.IO.File.Move(textBox1.Text, textBox3.Text + "\\" + textBox2.Text);
                    illReLink();
                }
                else if (tmpVal > 1)
                {
                    System.Media.SystemSounds.Beep.Play();
                    DialogResult result = MessageBox.Show("ドキュメント内で他にもこのファイルをリンクしています。\r\nこのままリネームしますか？", "",
                                                          MessageBoxButtons.YesNo,
                                                          MessageBoxIcon.Exclamation,
                                                          MessageBoxDefaultButton.Button2);
                    //何が選択されたか調べる
                    if (result == DialogResult.Yes)
                    {
                        string tmpFileName = System.IO.Path.GetFileName(textBox1.Text);
                        try
                        {
                            System.IO.File.Move(textBox1.Text, textBox3.Text + "\\" + textBox2.Text);
                        }
                        catch
                        {
                            System.Media.SystemSounds.Beep.Play();
                            MessageBox.Show("ファイルが使用中のためリネームできませんでした。\r\n処理を中止します。");
                            timer1.Start();
                            return;
                        }
                        for (int i = 1; i < illGraphics.Count+1; i++)
                        {
                            try
                            {
                                if (illGraphics[i].File != "")
                                {

                                }
                            }
                            catch
                            {
                                if (tmpFileName == illGraphics[i].Name)
                                {
                                    
                                    illGraphics[i].Relink(textBox3.Text + "\\" + textBox2.Text);
                                }
                            }
                        }
                    }
                    if (result == DialogResult.No)
                    {
                        //「いいえ」が選択された時
                    }
                }
            }
            else if (System.IO.File.Exists(textBox1.Text) == false)
            {
                System.Media.SystemSounds.Beep.Play();
                MessageBox.Show("ファイルが見つかりません。");
            }
            else if (System.IO.File.Exists(textBox1.Text) == true && System.IO.File.Exists(textBox3.Text + "\\" + textBox2.Text) == true)
            {
                System.Media.SystemSounds.Beep.Play();
                MessageBox.Show("同名のファイルが存在しています。");
            }
            timer1.Start();
        }

        #region アプリケーション起動;
        private void buttonApp1_Click(object sender, EventArgs e)
        {
            try
            {
                String myPath = "";
                linkPathCheck(ref myPath);
                if (System.IO.File.Exists(myPath) == true)
                {
                    System.Diagnostics.ProcessStartInfo psi = new System.Diagnostics.ProcessStartInfo();
                    psi.FileName = Properties.Settings.Default.app1path;
                    psi.Arguments = myPath;
                    System.Diagnostics.Process.Start(psi);
                }
            }
            catch
            { }
        }

        private void buttonApp2_Click(object sender, EventArgs e)
        {
            try
            {
                String myPath = "";
                linkPathCheck(ref myPath);
                if (System.IO.File.Exists(myPath) == true)
                {
                    System.Diagnostics.ProcessStartInfo psi = new System.Diagnostics.ProcessStartInfo();
                    psi.FileName = Properties.Settings.Default.app2path;
                    psi.Arguments = myPath;
                    System.Diagnostics.Process.Start(psi);
                }
            }
            catch
            { }
        }

        private void buttonApp3_Click(object sender, EventArgs e)
        {
            try
            {
                String myPath = "";
                linkPathCheck(ref myPath);
                if (System.IO.File.Exists(myPath) == true)
                {
                    System.Diagnostics.ProcessStartInfo psi = new System.Diagnostics.ProcessStartInfo();
                    psi.FileName = Properties.Settings.Default.app3path;
                    psi.Arguments = myPath;
                    System.Diagnostics.Process.Start(psi);
                }
            }
            catch
            { }
        }

        private void buttonApp4_Click(object sender, EventArgs e)
        {
            try
            {
                String myPath = "";
                linkPathCheck(ref myPath);
                if (System.IO.File.Exists(myPath) == true)
                {
                    System.Diagnostics.ProcessStartInfo psi = new System.Diagnostics.ProcessStartInfo();
                    psi.FileName = Properties.Settings.Default.app4path;
                    psi.Arguments = myPath;
                    System.Diagnostics.Process.Start(psi);
                }
            }
            catch
            { }
        }
        #endregion;

        private void Setting_Click(object sender, EventArgs e)
        {
            label3.Focus();
            Setting f = new Setting();
            f.ShowDialog(this);
            f.Dispose();
            buttonApp1.Text = Properties.Settings.Default.app1name;
            buttonApp2.Text = Properties.Settings.Default.app2name;
            buttonApp3.Text = Properties.Settings.Default.app3name;
            buttonApp4.Text = Properties.Settings.Default.app4name;
        }

        private void textBox2_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == (char)Keys.Enter)
            {
                ReLink.Focus();
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            label3.Focus();
            try
            {
                string docPath = Form1.illApp.ActiveDocument.FullName;
                System.Diagnostics.Process.Start("EXPLORER.EXE", @"/select," + docPath);
            }
            catch { }
        }

        private void button1_Click_1(object sender, EventArgs e)
        {
            label3.Focus();
            try
            {
                string docPath = Form1.illApp.ActiveDocument.FullName;
                string folderName = System.IO.Path.GetDirectoryName(docPath);
                dynamic illGraphics;
                illGraphics = Form1.illApp.ActiveDocument.PlacedItems;
                string tmpList = "";
                for (int i = 1; i < illGraphics.Count + 1; i++)
                {
                    try
                    {
                        if (illGraphics[i].File.ToLower().IndexOf(folderName.ToLower()) != 0)
                        {
                            tmpList = tmpList + System.IO.Path.GetFileName(illGraphics[i].File) + " (" + System.IO.Path.GetDirectoryName(illGraphics[i].File) + ")\r\n";
                        }
                    }
                    catch
                    { }
                }
                if (tmpList != "")
                {
                    System.Media.SystemSounds.Beep.Play();
                    MessageBox.Show(tmpList + "\r\n以上のリンクファイルが、Illustratorドキュメントのある階層内(" + folderName + ")に存在していません。");
                }
                else if (illGraphics.Count > 0)
                {
                    System.Media.SystemSounds.Beep.Play();
                    MessageBox.Show("問題は見つかりませんでした。");
                }
            }
            catch 
            {
            }
        }

        private void SameOrWarning()
        {
            if (textBox3.Text != "" && System.IO.File.Exists(textBox3.Text + "\\" + textBox2.Text))
            {
                //if (textBox1.Text == textBox3.Text + "\\" + textBox2.Text)
                if (String.Compare(textBox1.Text, textBox3.Text + "\\" + textBox2.Text, true) == 0)
                {
                    pictureBox2.Visible = false;
                    pictureSame.Visible = true;
                }
                else
                {
                    pictureBox2.Visible = true;
                    pictureSame.Visible = false;
                }
            }
            else
            {
                pictureBox2.Visible = false;
                pictureSame.Visible = false;
            }
        }

    }
}
