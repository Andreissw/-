using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using ModuleConnect;


namespace ЖурналВходногоКонтроля
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        int id;
        private void button1_Click(object sender, EventArgs e)
        {
            int k = 0;
            for (int i = 0; i < Ad.ColumnCount; i++)
                if ((string)Ad.Rows[0].Cells[i].Value == null)
                    k++;

            if (k == 43)
                return;

           id = (int)Class1.SelectStringInt(@"use SMDCOMPONETS SELECT top(1) id  FROM [SMDCOMPONETS].[dbo].[Журнал]  order by [id] desc ");

         
            string[] Столбцы = new string[46];
            for (int i = 0; i < Столбцы.Length; i++)
                   Столбцы[i] = GrInf.Columns[i+1].HeaderText;
           

            for (int i = 0; i < Ad.RowCount; i++)
            {
                     Class1.SelectString(@"  Insert into [SMDCOMPONETS].[dbo].[Журнал]  (id)  Values  ('" + (id + Convert.ToInt32(i + 1)) + "')");
                string sql = "";
                for (int b = 0; b < 46; b++)
                {
                    switch (b) { 
                         case 1:
                            sql = sql + " [" + Столбцы[b] + "] = '" + Convert.ToDateTime(Ad.Rows[i].Cells[b].Value).ToString("yyyy-MM-dd") + "',  ";
                         break;

                         case 45:
                            sql = sql + " [" + Столбцы[b] + "] = '" + Ad.Rows[i].Cells[b].Value + "'  ";
                         break;
                        default:
                            sql = sql + " [" + Столбцы[b] + "] = '" + Ad.Rows[i].Cells[b].Value + "',  ";
                            break;
                    }
                    
                }
                
                Class1.SelectString(@"  update [SMDCOMPONETS].[dbo].[Журнал] set "+ sql +"  where id = '" + (id + Convert.ToInt32(i + 1)) + "'");
            }

            Class1.loadgrid(GrInf, @"use SMDCOMPONETS SELECT top(250) *  FROM [SMDCOMPONETS].[dbo].[Журнал]   order by [Дата поставки] desc ");
            GrInf.ClearSelection();
            row = 0;



        }
        string name;

        private void RFID_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                IN();
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            IN();
        }

        private void IN()
        {
            name = Class1.SelectString(@"use FAS Select UserName FROM [FAS].[dbo].[FAS_Users] where RFID = '" + RFID.Text + "'").ToString();
            if (name != "")
                methodLog(); //Вход в программу
            else
                methodRefresh(); //Не правильный пароль
        }

        void methodLog()
        {
            if (name == "Черепанова Ж.В.")
            { delBT.Enabled = true; UpdateBT.Enabled = true; }

            this.Text = "Пользователь - " + name;
            label8.Visible = true;
            //GrLogin.Location = new Point(0, 0);
            GrLogin.Visible = false;
            GrForm.Visible = true;
            GrForm.Location = new Point(12, 12);
            GrForm.Size = new Size(1152, 864);
            this.Size = new Size(1152, 864);
            Ad.Rows.Add(1);
            Class1.loadgrid(GrInf, @"use SMDCOMPONETS SELECT  TOP(250) *  FROM [SMDCOMPONETS].[dbo].[Журнал]  order by [Дата поставки] desc ");
            row = (int) GrInf.Rows[GrInf.CurrentRow.Index].Cells[0].Value;
            GrInf.ClearSelection();
            row = 0;
            Class1.loadgrid(ГридProject, "use SMDCOMPONETS SELECT distinct(ФИО)  FROM[SMDCOMPONETS].[dbo].[Журнал]");
            цикл(ФИОCB);
            Class1.loadgrid(ГридProject, "use SMDCOMPONETS SELECT distinct(Заказчик)  FROM[SMDCOMPONETS].[dbo].[Журнал]");
            цикл(ЗаказчикCB);
        }

        void methodRefresh()
        {
            RFID.Clear();
            RFID.Select();
        }

        void цикл(ComboBox CB)
        {

            for (int i = 0; i < ГридProject.RowCount - 1; i++)
            {
                CB.Items.Add(ГридProject.Rows[i].Cells[0].Value);
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            

            //this.Size = new Size(290, 160);
            GrLogin.Location = new Point(10, 10);
            GrLogin.Size = new Size(300, 200);
            RFID.Select();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            if (CH.Checked == true & ФИОCB.Text == "" & ЗаказчикCB.Text == "")
                запрос_фильтр("[Дата Поставки] between('" + picst.Value.ToString("yyyy-MM-dd") + "') and('" + picend.Value.ToString("yyyy-MM-dd") + "')");
            else if(CH.Checked == true & ФИОCB.Text != "" & ЗаказчикCB.Text == "")
                запрос_фильтр("[Дата Поставки] between('" + picst.Value.ToString("yyyy-MM-dd") + "') and('" + picend.Value.ToString("yyyy-MM-dd") + "') and ФИО = '"+ ФИОCB.Text +"'");
            else if (CH.Checked == true & ФИОCB.Text != "" & ЗаказчикCB.Text != "")
                запрос_фильтр("[Дата Поставки] between('" + picst.Value.ToString("yyyy-MM-dd") + "') and('" + picend.Value.ToString("yyyy-MM-dd") + "') and ФИО = '" + ФИОCB.Text + "' and Заказчик = '" + ЗаказчикCB.Text + "' ");
            else if (CH.Checked == true & ФИОCB.Text == "" & ЗаказчикCB.Text != "")
                запрос_фильтр("[Дата Поставки] between('" + picst.Value.ToString("yyyy-MM-dd") + "') and('" + picend.Value.ToString("yyyy-MM-dd") + "')  and Заказчик = '" + ЗаказчикCB.Text + "' ");
            else if (CH.Checked == false & ФИОCB.Text != "" & ЗаказчикCB.Text != "")
                запрос_фильтр("ФИО = '" + ФИОCB.Text + "' and Заказчик = '" + ЗаказчикCB.Text + "' ");
            else if (CH.Checked == false & ФИОCB.Text == "" & ЗаказчикCB.Text != "")
                запрос_фильтр("Заказчик = '" + ЗаказчикCB.Text + "' ");
            else if (CH.Checked == false & ФИОCB.Text != "" & ЗаказчикCB.Text == "")
                запрос_фильтр("ФИО = '" + ФИОCB.Text + "' ");

            GrInf.ClearSelection();
            row = 0;

        }

        void запрос_фильтр(string name)
        {
            Class1.loadgrid(GrInf, @"use SMDCOMPONETS SELECT TOP(250) * FROM [SMDCOMPONETS].[dbo].[Журнал]   where "+ name + " order by [Дата поставки] desc ");
        }

        private void button4_Click(object sender, EventArgs e) //выход
        {
            UpdateBT.Enabled = false;
            delBT.Enabled = false;
            
            label8.Visible = false;
            HeaderPB.Value = 0;
            RowPB.Value = 0;
            Ad.RowCount = 0;
            GrLogin.Visible = true;
            GrForm.Visible = false;
            RFID.Clear();
            ФИОCB.Text = "";
            ЗаказчикCB.Text = "";
            RFID.Select();
            ФИОCB.Items.Clear();
            ЗаказчикCB.Items.Clear();
            //clearGrid();
        }

        private void button5_Click(object sender, EventArgs e)
        {
            clearGrid();
            GrInf.ClearSelection();
            row = 0;

        }

        private void clearGrid()
        {
            for (int i = 0; i < Ad.ColumnCount; i++)
                Ad.Rows[0].Cells[i].Value = "";
        }

        int row = 0;
        private void GrInf_CellClick(object sender, DataGridViewCellEventArgs e)
        {

            row = Convert.ToInt32(GrInf.Rows[GrInf.CurrentRow.Index].Cells[0].Value);
        }

        private void button6_Click(object sender, EventArgs e) //Удалить строку
        {
            if (name != "Черепанова Ж.В.")
                return;
           

            if (row != 0)
            {
                Class1.SelectString(@" USE SMDCOMPONETS  delete [SMDCOMPONETS].[dbo].[Журнал]  where id = '" + row + "'"); // Запрос удаление
                Class1.loadgrid(GrInf, @"use SMDCOMPONETS SELECT TOP(250) *  FROM [SMDCOMPONETS].[dbo].[Журнал]   order by [Дата поставки] desc "); //Запрос на выборку данных
            }
            GrInf.ClearSelection();
            row = 0;
        }

        private void button7_Click(object sender, EventArgs e)
        {
            if (this.name != "Черепанова Ж.В.")
                return;

            string nameid;
                    

            if (row != 0)
            {
                for (int i = 0; i < 44; i++) {
                    nameid = GrInf.Rows[GrInf.CurrentRow.Index].Cells[i].Value.ToString();

                    if (i == 31 || i == 2)
                    { 
                        if (nameid != "")
                           Class1.SelectString("use SMDCOMPONETS update  [SMDCOMPONETS].[dbo].[Журнал]  set [" + GrInf.Columns[i].HeaderText + "] = '" + Convert.ToDateTime(nameid).ToString("yyyy-MM-dd") + "'  where id = '" + row + "'");
                    }
                    else
                        Class1.SelectString("use SMDCOMPONETS update  [SMDCOMPONETS].[dbo].[Журнал]  set [" + GrInf.Columns[i].HeaderText + "] = '" + nameid + "'  where id = '" + row + "'");                                                                                                                // Запрос Изменение

                }
                Class1.loadgrid(GrInf, @"use SMDCOMPONETS SELECT TOP(250) *  FROM [SMDCOMPONETS].[dbo].[Журнал]   order by [Дата поставки] desc   "); //Запрос на выборку данных
            }
           
            GrInf.ClearSelection();
            row = 0;
        }

         void button8_Click(object sender, EventArgs e) // Выгрузка в Excel
        {
            GrInf.ClearSelection();
            row = 0;
            button8.Enabled = false;
            button10.Enabled = false;
            ExcelAsync();
           




        }

        async void ExcelAsync()
        {
            await Task.Run(() => ExcelMethod());
        }


     

        void progressbarHeader(int k, int max, Label lb, int pr)
        {
          
            Invoke((Action)(() =>
            {
                lb.Text = pr.ToString() + " %";
                HeaderPB.Maximum = max;
                HeaderPB.Value = k;

            }));

        }

        void ExcelMethod() // Вызов метода в ассинхроном режиме
        {
            Invoke((Action)(() => { HeaderPB.Value = 0; }));

            Microsoft.Office.Interop.Excel.Application ExcelApp = new Microsoft.Office.Interop.Excel.Application();
            ExcelApp.Application.Workbooks.Add(Type.Missing);
            ExcelApp.Columns.ColumnWidth = 15;

            int k = 0;
            for (int i = 0; i < 44; i++)
            {
                k++;
                ExcelApp.Cells[1, i + 1] = GrInf.Columns[i].HeaderText;

                int pr1 = (Int32)(Convert.ToInt32(i+1) / (44 / 100M));

                progressbarHeader(k,44, HeaderPBLB,pr1);
            }
            k = 0;

            for (int i = 0; i < GrInf.ColumnCount; i++)
                for (int j = 0; j < GrInf.RowCount; j++)
                {
                    //double pr = Convert.ToDouble(k / (GrInf.RowCount * GrInf.ColumnCount));
                    int pr = (Int32) (k/(GrInf.RowCount * GrInf.ColumnCount/100M));
                    k++;
                    ExcelApp.Cells[j + 2, i + 1] = (GrInf[i, j].Value).ToString();
                    progressbarRow(k, GrInf.RowCount * GrInf.ColumnCount, pr,RowPB);
                }

            ExcelApp.Visible = true;
         
            
          
                GC.Collect();
             ExcelApp = null;
           
         

        }

        void progressbarRow(int k, int maxval, int pr,ProgressBar PR)
        {
            Invoke((Action)(() =>
            {
                Procent.Text = pr.ToString() + " %";
                PR.Maximum = maxval;
                PR.Value = k;
                if (pr == 99) { Procent.Text = "100 %"; button8.Enabled = true; button10.Enabled = true; }
            }));
        }

        void progressbarRow(int k, int maxval, int pr, ProgressBar PR, Label lb)
        {
            Invoke((Action)(() =>
            {
                lb.Text = pr.ToString() + " %";
                PR.Maximum = maxval;
                PR.Value = k;
                if (pr == 99) { lb.Text = "100 %"; button8.Enabled = true; }
            }));
        }

        private void GrForm_Enter(object sender, EventArgs e)
        {

        }

        private void CH_CheckedChanged(object sender, EventArgs e)
        {
            if (CH.Checked == true)
                GRDate.Enabled = true;
            else
                GRDate.Enabled = false;
        }

        private void button9_Click(object sender, EventArgs e)
        {
            Class1.loadgrid(GrInf, @"use SMDCOMPONETS SELECT  TOP(250) *  FROM [SMDCOMPONETS].[dbo].[Журнал]  order by [Дата поставки] desc ");
            GrInf.ClearSelection();
            row = 0;
        }

        string path;

        private void button10_Click(object sender, EventArgs e)
        {
            //button8.Enabled = false;
            //button10.Enabled = false;
            using (var file = new OpenFileDialog()) //Выбрать файл Excel
            {
                file.Filter = "All files (*.*)|*.*";
                file.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                try
                {
                    if (file.ShowDialog() == DialogResult.OK)
                        path = file.FileName;
                    else
                    { 
                        button8.Enabled = true;
                        button10.Enabled = true;
                        return;
                    }
                }
                catch (Exception t)
                { MessageBox.Show(t.ToString()); }
            };

            ExcelIn(path);            
        }

        async void ExcelIn(string path)
        {
            await Task.Run(() => ExcelInMethod(path));
        }

        void ExcelInMethod(string path)
        {
            Invoke((Action)(() => { HeaderPB.Value = 0; }));
            Invoke((Action)(() => { RowPB.Value = 0; }));


            Microsoft.Office.Interop.Excel.Application Excel = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook Book = Excel.Workbooks.Open(path);
            Microsoft.Office.Interop.Excel.Worksheet Sheet = (Microsoft.Office.Interop.Excel.Worksheet)Book.Sheets[1];
            var lastCell = Sheet.Cells.SpecialCells(Microsoft.Office.Interop.Excel.XlCellType.xlCellTypeLastCell);  //1 ячейку


            for (int i = 0; i < 46; i++) //Проверка шаблона и таблицы грида 
            {
                if (Sheet.Cells[1, Convert.ToInt32(i + 1)].Text.ToString() != Ad.Columns[i].HeaderText)
                {
                    MessageBox.Show("Произошла ошибка! Несоответствие столбцов шаблона | Столбец - " + Sheet.Cells[1, Convert.ToInt32(i + 1)].Text.ToString(), "Что то не так", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    Book.Close(false, Type.Missing, Type.Missing); //закрыть не сохраняя
                    Excel.Quit(); // выйти из экселя
                    GC.Collect();
                    return;
                }
            }

            Invoke((Action)(() => { Ad.RowCount = lastCell.Row - 1; }));


            int b = 1;

            for (int k = 0; k < lastCell.Row - 1; k++)
            {
                int pr = (Int32)(Convert.ToInt32(k +1) / ((lastCell.Row - 1) / 100M));

                progressbarRow(Convert.ToInt32(k + 1), lastCell.Row - 1, pr, HeaderPB, HeaderPBLB);

                for (int i = 0; i < lastCell.Column; i++)
                {
                    Invoke((Action)(() => { Ad.Rows[k].Cells[i].Value = Sheet.Cells[Convert.ToInt32(k + 2), Convert.ToInt32(i + 1)].Text.ToString(); }));

                    int pr1 = (Int32)(b / ((lastCell.Column * (lastCell.Row - 1)) / 100M));

                    progressbarRow(b, lastCell.Column * (lastCell.Row - 1), pr1, RowPB);
                    b++;


                }

            }
            Invoke((Action)(() => {    button8.Enabled = true; button10.Enabled = true; }));
           
            Book.Close(false, Type.Missing, Type.Missing); //закрыть не сохраняя
            Excel.Quit(); // выйти из экселя
            GC.Collect();

        }

      
    }
}
