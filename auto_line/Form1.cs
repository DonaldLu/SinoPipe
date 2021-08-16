using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System.IO;
using Auto_ManHole;
using System.Runtime.InteropServices;

namespace auto_line
{
    public partial class Form1 : System.Windows.Forms.Form
    {
        [DllImport("Gdi32.dll", EntryPoint = "CreateRoundRectRgn")]
        private static extern IntPtr CreateRoundRectRgn
        (
            int nLeftRect,     // x-coordinate of upper-left corner
            int nTopRect,      // y-coordinate of upper-left corner
            int nRightRect,    // x-coordinate of lower-right corner
            int nBottomRect,   // y-coordinate of lower-right corner
            int nWidthEllipse, // height of ellipse
            int nHeightEllipse // width of ellipse
        );

        public Form2 Read_form
        {
            get;
            set;
        }

        //外部事件處理:讀取其他的cs檔
        ExternalEvent externalEvent_automh;
        Auto_ManHole.Handler_AutoMH handler_automh = new Auto_ManHole.Handler_AutoMH();
        
        ExternalEvent externalEvent_autopipe;
        PipeV2.Handler_autoPipe handler_autopipe = new PipeV2.Handler_autoPipe();

        ExternalEvent externalEvent_ssd;
        SweepSideDitch handler_ssd = new SweepSideDitch();

        ExternalEvent externalEvent_alloutput;
        OutputCounting handler_alloutput = new OutputCounting();

        ExternalEvent ExternalEvent_PartSel;
        ParticalSelection handler_PartSel = new ParticalSelection();

        ExternalEvent ExternalEvent_callback;
        Callback handler_callback = new Callback();

        ExternalEvent externalEvent_make_sheet;
        MakeSheet handler_make_sheet = new MakeSheet();

        ExternalEvent externalEvent_output_excel;
        OutputExcel handler_output_excel = new OutputExcel();

        ExternalEvent externalEvent_Zmove;
        Zmove handler_Zmove = new Zmove();

        public Form1(UIDocument uIDocument, Form2 form2)
        {
            InitializeComponent();
            Read_form = form2;
            this.FormBorderStyle = FormBorderStyle.None;
            Region = System.Drawing.Region.FromHrgn(CreateRoundRectRgn(0, 0, Width, Height, 20, 20));

            //建立外部事件
            externalEvent_automh = ExternalEvent.Create(handler_automh);
            externalEvent_autopipe = ExternalEvent.Create(handler_autopipe);
            externalEvent_ssd = ExternalEvent.Create(handler_ssd);
            externalEvent_alloutput = ExternalEvent.Create(handler_alloutput);
            ExternalEvent_PartSel = ExternalEvent.Create(handler_PartSel);
            ExternalEvent_callback = ExternalEvent.Create(handler_callback);
            externalEvent_make_sheet = ExternalEvent.Create(handler_make_sheet);
            externalEvent_output_excel = ExternalEvent.Create(handler_output_excel);
            externalEvent_Zmove = ExternalEvent.Create(handler_Zmove);

            //初始偏移量xyz
            textBox_x.Text = "294820.3124";
            textBox_y.Text = "2762885.975";
            textBox_z.Text = "0";

            //剖面數
            picture_comboBox1.Items.Add(1);
            picture_comboBox1.Items.Add(2);
            picture_comboBox1.Items.Add(4);
            picture_comboBox1.Items.Add(6);

        }

        private void Form1_Load(object sender, EventArgs e)
        {
            //找出字體大小,並算出比例
            float dpiX, dpiY;
            Graphics graphics = this.CreateGraphics();
            dpiX = graphics.DpiX;
            dpiY = graphics.DpiY;
            int intPercent = (dpiX == 96) ? 100 : (dpiX == 120) ? 125 : 150;

            // 針對字體變更Form的大小
            this.Height = this.Height * intPercent / 100;
            this.Width = this.Width * intPercent / 100;
            this.Size = new System.Drawing.Size(this.header.Size.Width, this.header.Size.Height+this.homeleftpanel.Size.Height+this.panel2.Height);

            //將基礎介面之大小縮為(600,385)
            Region = System.Drawing.Region.FromHrgn(CreateRoundRectRgn(0, 0, Width, Height, 20, 20));
            //將基礎介面之四個角落進行半徑為20之圓滑處理

            int start_position = leftpanel.Location.X + leftpanel.Size.Width;
            int start_position_y = panel2.Location.Y + panel2.Height;

            main.Location = new System.Drawing.Point(start_position, start_position_y);
            //上傳檔案初始介面位置移到(150,125)
            outputpanel.Location = new System.Drawing.Point(start_position, start_position_y);
            //建置介面位置移到(150,125)
            picturepanel.Location = new System.Drawing.Point(start_position, start_position_y);
            //圖資產出介面移到(150,125)

            ToolTip toolTip = new ToolTip();
            toolTip.SetToolTip(ParticalSelBtm, "使用者選取管線，並查詢管線相關資訊");
            toolTip.SetToolTip(CallBackBtm, "使用者輸入數量計算書之管線ID，並將其獨立顯示");
        }

        //數量計算樣板檔選擇按鍵
        private void button1_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.ShowDialog();
            upload_textbox.Text = openFileDialog.FileName;
            handler_automh.file_path = upload_textbox.Text;
            handler_autopipe.file_path = upload_textbox.Text;
            handler_ssd.file_path = upload_textbox.Text;
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        //建立人孔按鍵
        private void button2_Click(object sender, EventArgs e)
        {
            handler_automh.xyz_shift = new List<double>();
            handler_automh.xyz_shift.Add(double.Parse(textBox_x.Text));
            handler_automh.xyz_shift.Add(double.Parse(textBox_y.Text));
            handler_automh.xyz_shift.Add(double.Parse(textBox_z.Text));

            externalEvent_automh.Raise();
            
        }

        //建立管線按鍵
        private void button3_Click(object sender, EventArgs e)
        {
            handler_autopipe.xyz_shift = new List<double>();
            handler_autopipe.xyz_shift.Add(double.Parse(textBox_x.Text));
            handler_autopipe.xyz_shift.Add(double.Parse(textBox_y.Text));
            handler_autopipe.xyz_shift.Add(double.Parse(textBox_z.Text));
            externalEvent_autopipe.Raise();
        }

        private void tabPage1_Click(object sender, EventArgs e)
        {

        }

        private void label4_Click(object sender, EventArgs e)
        {

        }

        //視窗縮小按鍵
        private void mini_Click(object sender, EventArgs e)
        {
            this.WindowState = FormWindowState.Minimized;
        }

        //關閉視窗按鍵
        private void leavebutton_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }
        
        //讀取資料頁籤按鍵
        private void uploadbutton_Click(object sender, EventArgs e)
        {
            leftpanel.Height = uploadbutton.Height;
            leftpanel.Top = uploadbutton.Top;
            main.Show();
            outputpanel.Hide();
            picturepanel.Hide();
        }

        //建立模型頁籤按鍵
        private void outputbutton_Click(object sender, EventArgs e)
        {
            leftpanel.Height = outputbutton.Height;
            leftpanel.Top = outputbutton.Top;
            main.Hide();
            outputpanel.Show();
            picturepanel.Hide();

        }

        private void label3_Click(object sender, EventArgs e)
        {

        }

        private void outputUC1_Load(object sender, EventArgs e)
        {

        }

        private void outputpanel_Paint(object sender, PaintEventArgs e)
        {

        }

        private void homebutton_Click(object sender, EventArgs e)
        {
           

        }

        private void settingbutton_Click(object sender, EventArgs e)
        {
  
        }
        
        private void button1_Click_1(object sender, EventArgs e)
        {
            outputpanel.BackgroundImage = Properties.Resources.榮耀黃;
            this.BackgroundImage = Properties.Resources.榮耀黃;

            
        }

        private void button2_Click_1(object sender, EventArgs e)
        {
            outputpanel.BackgroundImage = Properties.Resources.淵海藍;
            this.BackgroundImage = Properties.Resources.淵海藍;
        }

        private void button3_Click_1(object sender, EventArgs e)
        {
            outputpanel.BackgroundImage = Properties.Resources.玫瑰紅;
            this.BackgroundImage = Properties.Resources.玫瑰紅;
        }

        //建置側溝按鍵
        private void ditchbutton_Click(object sender, EventArgs e)
        {
            handler_ssd.xyz_shift = new List<double>();
            handler_ssd.xyz_shift.Add(double.Parse(textBox_x.Text));
            handler_ssd.xyz_shift.Add(double.Parse(textBox_y.Text));
            handler_ssd.xyz_shift.Add(double.Parse(textBox_z.Text));
            externalEvent_ssd.Raise();
        }

        //圖資產出夜間按鍵
        private void pictureout_Click(object sender, EventArgs e)
        {
            leftpanel.Height = pictureout.Height;
            leftpanel.Top = pictureout.Top;
            main.Hide();
            outputpanel.Hide();
            picturepanel.Show();
        }

        //數量計算樣板檔選擇按鍵
        private void button1_Click_2(object sender, EventArgs e)
        {
            OpenFileDialog openFile = new OpenFileDialog();
            openFile.ShowDialog();
            ExceltextBox.Text = openFile.FileName;
            handler_alloutput.file_path = ExceltextBox.Text;
        }

        //數量計算按鈕
        private void AllOutputBtm_Click(object sender, EventArgs e)
        {
           
            externalEvent_alloutput.Raise();
        }

        //查詢管線資訊按鍵
        private void ParticalSelBtm_Click(object sender, EventArgs e)
        {
            ExternalEvent_PartSel.Raise();
        }

        //獨立顯示按鍵
        private void CallBackBtm_Click(object sender, EventArgs e)
        {
            handler_callback.callValues = CallBacktextBox.Text;
            ExternalEvent_callback.Raise();
        }

        //圖框選擇按鈕
        private void picture_button_Click(object sender, EventArgs e)
        {
            //讀取圖框資料
            picture_textBox.Text = "";
            OpenFileDialog openFileDialogFrame = new OpenFileDialog();
            openFileDialogFrame.Multiselect = true;
            openFileDialogFrame.ShowDialog();
            handler_make_sheet.openFileDialog = openFileDialogFrame;
            picture_textBox.Text = String.Format("已選取 {0} 個圖框", openFileDialogFrame.FileNames.Count());
        }

        private void picture_textBox_TextChanged(object sender, EventArgs e)
        {

        }

        //剖面:開始出圖按鍵
        private void picture_Button1_Click(object sender, EventArgs e)
        {
            handler_make_sheet.sectionLineNumber = Convert.ToInt32(picture_comboBox1.Text);
            handler_make_sheet.sectionName = Read_form.selected;//載入form2目前所選取的圖面
            externalEvent_make_sheet.Raise();
        }

       
        private void picture_comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        //讀取匯出Excel的.xlxs檔
        /*private void button2_Click_2(object sender, EventArgs e)
        {

            //讀取.xlxs資料
            output_excel_textBox.Text = "";
            OpenFileDialog openFileDialogFrame = new OpenFileDialog();
            openFileDialogFrame.Multiselect = true;
            openFileDialogFrame.ShowDialog();
            handler_output_excel.Filepath = openFileDialogFrame.FileName;
            output_excel_textBox.Text = openFileDialogFrame.FileName;
            //String.Format("已選取 {0} 個圖框", openFileDialogFrame.FileNames.Count());
        }*/

        //模型倒出按鍵
        private void button3_Click_2(object sender, EventArgs e)
        {
            externalEvent_output_excel.Raise();
            handler_output_excel.xyz_shift = new List<double>();
            handler_output_excel.xyz_shift.Add(double.Parse(textBox_x.Text));
            handler_output_excel.xyz_shift.Add(double.Parse(textBox_y.Text));
            handler_output_excel.xyz_shift.Add(double.Parse(textBox_z.Text));
        }

        //選取頗面按鈕
        private void SectionViewBtn_Click(object sender, EventArgs e)
        {
            Read_form.Visible = true;
        }

        //修正埋管深度按鍵
        private void button4_Click(object sender, EventArgs e)
        {
            externalEvent_Zmove.Raise();
        }

        //數量計算:選擇欲儲存之路徑按鍵
        private void save_btn_Click(object sender, EventArgs e)
        {
            save_textbox.Text = "";
            SaveFileDialog dialog = new SaveFileDialog();
            dialog.FileName = "";
            dialog.ShowDialog();
            handler_alloutput.path = dialog.FileName;
            save_textbox.Text = dialog.FileName;
        }

        private void homebuttompanel_Paint(object sender, PaintEventArgs e)
        {

        }

        private void textBox_x_TextChanged(object sender, EventArgs e)
        {

        }

        private void tabPage2_Click(object sender, EventArgs e)
        {

        }

        //模型倒出樣板檔選擇按鍵
        private void button2_Click_2(object sender, EventArgs e)
        {
            //讀取.xlxs資料
            textBox1.Text = "";
            OpenFileDialog openFileDialogFrame = new OpenFileDialog();
            openFileDialogFrame.Multiselect = true;
            openFileDialogFrame.ShowDialog();
            handler_output_excel.Filepath = openFileDialogFrame.FileName;
            textBox1.Text = openFileDialogFrame.FileName;
            //String.Format("已選取 {0} 個圖框", openFileDialogFrame.FileNames.Count());
        }

        private void textBox1_TextChanged_1(object sender, EventArgs e)
        {

        }

        private void textBox1_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyData == (Keys.Control | Keys.V))
                (sender as System.Windows.Forms.TextBox).Paste();
        }
    }
}
