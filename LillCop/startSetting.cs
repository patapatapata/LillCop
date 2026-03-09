using System;
using System.Linq;
using System.Windows.Forms;
using Microsoft.Win32;
using System.Security.Permissions;

namespace LillCop
{
    public partial class startSetting : Form
    {
        public startSetting()
        {
            InitializeComponent();
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            listBox1.Sorted = false;
            PopulateListBox();
        }

        private void PopulateListBox()
        {
            listBox1.Items.Clear();
            var versions = new System.Collections.Generic.List<string>();
            using (var hkcr = Registry.ClassesRoot)
            {
                foreach (var name in hkcr.GetSubKeyNames())
                {
                    if (name.StartsWith("Illustrator.Application.", StringComparison.OrdinalIgnoreCase))
                    {
                        var ver = name.Substring("Illustrator.Application.".Length);
                        if (!versions.Contains(ver)) versions.Add(ver);
                    }
                }
            }
            // 数値として降順ソート（未解析は末尾）
            var sorted = versions
                .OrderByDescending(s => { int v; return int.TryParse(s, out v) ? v : int.MinValue; })
                .ToArray();
            listBox1.Items.AddRange(sorted);

            listBox1.SelectedIndex = Properties.Settings.Default.ver;
        }

        private void listBox1_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == (char)Keys.Enter) LetsIllustrator();
        }
        private void listBox1_Click(object sender, EventArgs e) => LetsIllustrator();
        private void button1_Click(object sender, EventArgs e) => LetsIllustrator();

        private void LetsIllustrator()
        {
            if (listBox1.SelectedItem == null) { MessageBox.Show("バージョンを選択してください。"); return; }
            var ver = listBox1.SelectedItem.ToString();
            try
            {
                var progId = "Illustrator.Application." + ver;
                var t = Type.GetTypeFromProgID(progId, false);
                if (t == null) { MessageBox.Show(progId + " は登録されていません。"); return; }
                Form1.type = t;
                Form1.illApp = Activator.CreateInstance(t, true);
                this.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Illustrator " + ver + " を起動できませんでした。\n" + ex.Message);
            }
            Properties.Settings.Default.ver = listBox1.SelectedIndex;
            Properties.Settings.Default.Save();
        }

        [SecurityPermission(SecurityAction.Demand, Flags = SecurityPermissionFlag.UnmanagedCode)]
        protected override void WndProc(ref Message m)
        {
            const int WM_SYSCOMMAND = 0x112;
            const long SC_CLOSE = 0xF060L;
            if (m.Msg == WM_SYSCOMMAND && (m.WParam.ToInt64() & 0xFFF0L) == SC_CLOSE) Application.Exit();
            base.WndProc(ref m);
        }
    }
}