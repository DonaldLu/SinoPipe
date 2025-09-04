using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace SinoPipe_2020.UserControls
{
    public partial class SettingUC : UserControl
    {
        //private static Bitmap Image;

        public SettingUC()
        {
            InitializeComponent();
        }

        private void stylechange_y_Click(object sender, EventArgs e)
        {
            this.BackgroundImage = Properties.Resources.榮耀黃;
        }
    }
}
