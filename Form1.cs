using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using TXTextControl;

namespace tx_labels
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void loadToolStripMenuItem_Click(object sender, EventArgs e)
        {
            TXTextControl.LoadSettings ls = new TXTextControl.LoadSettings();
            ls.ApplicationFieldFormat = TXTextControl.ApplicationFieldFormat.MSWord;

            textControl1.Load("template.docx", TXTextControl.StreamType.WordprocessingML, ls);
        }

        private void mergeToolStripMenuItem_Click(object sender, EventArgs e)
        {
            DataSet ds = new DataSet();
            ds.ReadXml("data.xml");

            mailMerge1.MergeBlocks(ds);
            mailMerge1.Merge(ds.Tables[0], false);
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            textControl1.Tables.GridLines = false;
        }
    }
}
