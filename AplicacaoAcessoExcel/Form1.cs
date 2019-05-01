using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

// USANDO
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Ribbon;
using Microsoft.Office.Tools.Excel;
//USANDO
namespace AplicacaoAcessoExcel
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Excel.Worksheet.activeWorksheet = Globals.ThisAddIn.Application.ActiveSheet;
        }
    }//CLASS FORM
}
