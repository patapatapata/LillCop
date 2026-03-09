using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Runtime.ConstrainedExecution;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml.Linq;

namespace LillCop
{
    public partial class Setting : Form
    {
        public Setting()
        {
            InitializeComponent();
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            Properties.Settings.Default.Reload();
            textBox1.Text = Properties.Settings.Default.app1name;
            textBox4.Text = Properties.Settings.Default.app2name;
            textBox6.Text = Properties.Settings.Default.app3name;
            textBox8.Text = Properties.Settings.Default.app4name;
            textBox2.Text = Properties.Settings.Default.app1path;
            textBox3.Text = Properties.Settings.Default.app2path;
            textBox5.Text = Properties.Settings.Default.app3path;
            textBox7.Text = Properties.Settings.Default.app4path;
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

        private void button2_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                var ver = listBox1.SelectedItem.ToString();
                var progId = "Illustrator.Application." + ver;
                var t = Type.GetTypeFromProgID(progId, false);
                if (t == null) { MessageBox.Show(progId + " は登録されていません。"); return; }
                Form1.type = t;
                Form1.illApp = Activator.CreateInstance(t, true);
                Properties.Settings.Default.app1name = textBox1.Text;
                Properties.Settings.Default.app2name = textBox4.Text;
                Properties.Settings.Default.app3name = textBox6.Text;
                Properties.Settings.Default.app4name = textBox8.Text;
                Properties.Settings.Default.app1path = textBox2.Text;
                Properties.Settings.Default.app2path = textBox3.Text;
                Properties.Settings.Default.app3path = textBox5.Text;
                Properties.Settings.Default.app4path = textBox7.Text;
                Properties.Settings.Default.ver = listBox1.SelectedIndex;
                Properties.Settings.Default.Save();
                this.Close();
            }
            catch
            {
                MessageBox.Show("Illustrator is uncontrollable.");
            }
        }
    }
}
