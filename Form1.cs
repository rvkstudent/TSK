using ClosedXML.Excel;
using Excel = Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using System.Threading;
using System.Text.RegularExpressions;
using System.Text;

namespace Зарплата
{

    public partial class Form1 : Form
    {
        DataClasses1DataContext db = new DataClasses1DataContext(@"Data Source=ROMAN-ПК\SQLEXPRESS;Initial Catalog=Portal_base;Integrated Security=True");
         NumberStyles style;
         CultureInfo culture;

                
        public Form1()
        {
            InitializeComponent();

            
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            string[] dirs = Directory.GetFiles("c:\\XLTest\\", "*.txt");
            listBox2.Items.Clear();
            listBox2.DataSource = dirs;
        }
        private void AddToBase(String filename, String TableName, String Period,  List<int> XLRows, List<String> ToBaseRows, List<String> rowType)
        {
            style = NumberStyles.Number;
            culture = CultureInfo.CreateSpecificCulture("en-GB");
            int TableRowsCount = 0;
            var wb = new XLWorkbook(filename);
            var ws = wb.Worksheet(1);

            System.Data.SqlClient.SqlConnection sqlConnection1 =
                                  new System.Data.SqlClient.SqlConnection(@"Data Source=ROMAN-ПК\SQLEXPRESS;Initial Catalog=Portal_base;Integrated Security=True");

            System.Data.SqlClient.SqlCommand cmd = new System.Data.SqlClient.SqlCommand();
            cmd.CommandType = System.Data.CommandType.Text;

            cmd.Connection = sqlConnection1;

            sqlConnection1.Open();

            if (TableName != "Zakr")
            { 
            cmd.CommandText = "DELETE FROM [dbo].[" + TableName + "] WHERE Period ='" + Period + "';";
            cmd.CommandText = cmd.CommandText + "DELETE FROM [dbo].[" + TableName + "] WHERE Period IS NULL;";

               

                try
                {
                    TableRowsCount = cmd.ExecuteNonQuery();
                }
                catch (System.Data.SqlClient.SqlException e)
                {
                    MessageBox.Show(e.Message.ToString());
                    TableRowsCount = 0;
                }

            }
            

            

            string[] tempString = new string[20]; 
            double[] tempDouble =  new double[20];
            Int32[] tempInt = new Int32[20];

            String rows = "(";
            String values = "(";

            if(TableRowsCount >0)
                cmd.CommandText = "DELETE FROM [dbo].[" + TableName + "] WHERE Period = '" + Period + "';";

            int excelRow = ws.RowsUsed().Count();

            progressBar1.Maximum = excelRow;

            for (int i = 2; i <= excelRow; i++)
            {
                rows = "(";
                values = "(";
                DateTime result;

                progressBar1.Value = i;

                for (int j = 0; j < ToBaseRows.Count; j++)
                {

                    if (rowType[j] == "date")
                    {

                        if (DateTime.TryParse(ws.Cell(i, XLRows[j]).Value.ToString(), out result )== true)
                            values = values + "'" + result.ToString() + "'";
                        else
                            values = values + "'" + "01-01-2016" + "'";

                    }
                    if (rowType[j] == "string")
                    {
                        string tmp = "";

                        tmp = ws.Cell(i, XLRows[j]).Value.ToString();

                        if (ws.Cell(i, XLRows[j]).Value.ToString().Contains("-Восток") || ws.Cell(i, XLRows[j]).Value.ToString().Contains("Санкт-Петербург Восток"))
                            tmp =  "Санкт-Петербург-Восток";
                       
                        if (ws.Cell(i, XLRows[j]).Value.ToString().Contains("В.Нов") || ws.Cell(i, XLRows[j]).Value.ToString().Contains("Великий Новгород"))
                            tmp = "Великий-Новгород" ;
                      
                        if (ws.Cell(i, XLRows[j]).Value.ToString().Contains("-Кондер") || ws.Cell(i, XLRows[j]).Value.ToString().Contains("Хабаровск Кондер"))
                            tmp = "Хабаровск-Кондер" ;
                      
                            values = values + "'" + tmp + "'";

                    }
                    if (rowType[j] == "float")
                    {
                                            

                        if (Double.TryParse(ws.Cell(i, XLRows[j]).Value.ToString().Replace(",", "."), style, culture, out tempDouble[j]) == true)
                            tempDouble[j] = Convert.ToDouble(String.Format("{0:f}", tempDouble[j]));
                        else
                            tempDouble[j] = 0;

                        values = values + tempDouble[j].ToString().Replace(",", ".");
                    }
                    if (rowType[j] == "zp_float")
                    {

                        if (Double.TryParse(ws.Cell(i, XLRows[j]).Value.ToString().Substring(2), style, culture, out tempDouble[j]) == true)
                            tempDouble[j] = Convert.ToDouble(String.Format("{0:f}", tempDouble[j]));

                         else
                            tempDouble[j] = 0;

                        values = values + tempDouble[j].ToString().Replace(",", ".");
                    }

                    if (rowType[j] == "ktu_float")
                    {

                        if (Double.TryParse(ws.Cell(i, XLRows[j]).Value.ToString().Replace(",", "."), style, culture, out tempDouble[j]) == true)
                            tempDouble[j] = Convert.ToDouble(String.Format("{0:f}", tempDouble[j]));
                        else
                            tempDouble[j] = 1;

                        if(tempDouble[j]>1 || tempDouble[j]==0 || ws.Cell(i, XLRows[j-1]).Value.ToString().Contains("Механ") == true)
                            tempDouble[j] = 1;

                        values = values + tempDouble[j].ToString().Replace(",", ".");
                    }

                    if (rowType[j] == "int")
                    {


                        if (Int32.TryParse(ws.Cell(i, XLRows[j]).Value.ToString(), style, culture, out tempInt[j]) == true)
                            tempInt[j] = Convert.ToInt32(tempInt[j]);
                        else
                            tempInt[j] = 0;

                        values = values + tempInt[j].ToString();
                    }

                    if (rowType[j] == "parse_int")
                    {


                      if (!ws.Cell(i, XLRows[j]).Value.ToString().Equals(""))
                        { 
                        if (Int32.TryParse(ws.Cell(i, XLRows[j]).Value.ToString().Split(new Char[] { '/','(' })[1].Trim(), style, culture, out tempInt[j]) == true)
                            tempInt[j] = Convert.ToInt32(tempInt[j]);
                        else
                            tempInt[j] = 0;
                        }
                        else
                            tempInt[j] = 0;

                        values = values + tempInt[j].ToString();
                    }




                    rows = rows + ToBaseRows[j];

                    // UPDATE Employees SET HireDate = '20131101' WHERE ID = 1000
                    if (j < ToBaseRows.Count-1)
                    {
                        rows = rows + ",";
                        values = values+ ",";
                    }
                        
                   

                }

                if (TableName != "Zakr")
                { 
                    rows = rows + ",Period)";
                values = values + ",'" + Period + "')";
                }
                else
                {
                    rows = rows + ")";
                    values = values + ")";
                }
                cmd.CommandText = "INSERT " + TableName + rows + " VALUES " + values + ";";
                cmd.ExecuteNonQuery();
               
            }

           

            sqlConnection1.Close();

        }

        private void button1_Click(object sender, EventArgs e)
        {
            style = NumberStyles.Number;
            culture = CultureInfo.CreateSpecificCulture("en-GB");

            Double summa_bonusov = 0;

            label3.Text = "Процесс: чтение исходных данных";
            label3.Update();

           var wb = new XLWorkbook(textBox1.Text + comboBox1.SelectedItem.ToString() + ".xlsx");
            var ws = wb.Worksheet(1);

       
           
            int key = (from c in db.Bonus_za_ZNR select c.ID).Count();

            progressBar1.Maximum = ws.RowsUsed().Count();


             if (checkBox1.Checked == true) // затираем период в базе если галка стоит
             { 

             System.Data.SqlClient.SqlConnection sqlConnection1 =
                                   new System.Data.SqlClient.SqlConnection(@"Data Source=ROMANNB-ПК;Initial Catalog=Zarplata;Integrated Security=True");

             System.Data.SqlClient.SqlCommand cmd = new System.Data.SqlClient.SqlCommand();
             cmd.CommandType = System.Data.CommandType.Text;
             cmd.CommandText = "DELETE FROM [dbo].[Bonus_za_ZNR] WHERE Period = '"+ comboBox1.SelectedItem.ToString() + "';";
             cmd.CommandText = cmd.CommandText + "DELETE FROM [dbo].[FOT_analise] WHERE Period = '" + comboBox1.SelectedItem.ToString() + "';";

                 cmd.Connection = sqlConnection1;

             sqlConnection1.Open();
             cmd.ExecuteNonQuery();
                
             sqlConnection1.Close();

             }

            Double[] maxPersent = new double[10];
            Double[] minPersent = new double[10];

            List<Int32> list_ZNR = new List<Int32>();
            List<Int32> list_ZNR_new = new List<Int32>();
            List<String> list_Role = new List<string>();
            List<Double> list_Bonus = new List<Double>();
            List<Double> list_Davnost = new List<Double>();

            Double _ktu = 0, _sum_klient = 0, _sum_tsk = 0, _davnost = 0, _bonus = 0, _persent = 0;
            String _srok = "", _dolzhnost = "", _role = "";
            Int32 _zapros = 0, _remont = 0;


            maxPersent[1] = 0; // внешние механик
            minPersent[1] = 100; // внешние механик

            maxPersent[2] = 0; // внутр механик
            minPersent[2] = 100; // внутр механик

            

             for (int i = 2; i <= progressBar1.Maximum; i++)
             {
                 progressBar1.Value = i;

                 label3.Text = "Процесс: обработка таблицы Зарплата " + i + " из " + progressBar1.Maximum;
                 label3.Update();



                 if (Double.TryParse(ws.Cell(i, 17).Value.ToString(), style, culture, out _ktu) == true)
                     _ktu = Convert.ToDouble(String.Format("{0:f}", _ktu));
                 else
                     _ktu = 1;

                 if (_ktu > 1) _ktu = 1;

                 if (Double.TryParse(ws.Cell(i, 12).Value.ToString(), style, culture, out _sum_klient) == true)
                     _sum_klient = Convert.ToDouble(String.Format("{0:f}", _sum_klient));
                 else
                     _sum_klient = 1;

                 if (Double.TryParse(ws.Cell(i, 13).Value.ToString(), style, culture, out _sum_tsk) == true)
                     _sum_tsk = Convert.ToDouble(String.Format("{0:f}", _sum_tsk));
                 else
                     _sum_tsk = 1;

                 if (Double.TryParse(ws.Cell(i, 18).Value.ToString(), style, culture, out _davnost) == true)
                     _davnost = Convert.ToDouble(String.Format("{0:f}", _davnost));
                 else
                     _davnost = 1;

                 if (_davnost == 0)
                     _davnost = 1;
                 if (_ktu == 0)
                     _ktu = 1;

                 if (ws.Cell(i, 16).Value.ToString().Contains("Механик"))
                     _ktu = 1;

                 if (Double.TryParse(ws.Cell(i, 21).Value.ToString().Substring(2), style, culture, out _bonus) == true)
                     _bonus = Convert.ToDouble(String.Format("{0:f}", _bonus));
                 else
                     _bonus = 1;

                 summa_bonusov = summa_bonusov + _bonus;

                 Int32.TryParse(ws.Cell(i, 7).Value.ToString(), style, culture, out _zapros);

                 Int32.TryParse(ws.Cell(i, 10).Value.ToString(), style, culture, out _remont);

                 _srok = ws.Cell(i, 20).Value.ToString().Trim();

                 _dolzhnost = ws.Cell(i, 4).Value.ToString().Trim();

                 _role = ws.Cell(i, 16).Value.ToString().Trim();

                
                 _persent = _bonus*100/(_sum_tsk+ _sum_klient)/_davnost/_ktu;
                 _persent = Convert.ToDouble(String.Format("{0:0.##}", _persent));

               

                 if (_role.Contains("Механик") && !_dolzhnost.Contains("Финанс") && !_srok.Contains("Y") && _sum_klient > 0) // механик внешние
                 {
                     if (maxPersent[1] < _persent)
                         maxPersent[1] = _persent;
                     if (minPersent[1] > _persent)
                         minPersent[1] = _persent;
                 }

                 if (_role.Contains("Механик") && !_dolzhnost.Contains("Финанс") && !_srok.Contains("Y") && _sum_tsk > 0) // механик внешние
                 {
                     if (maxPersent[2] < _persent)
                         maxPersent[2] = _persent;
                     if (minPersent[2] > _persent)
                         minPersent[2] = _persent;
                 }

                 list_ZNR.Add(_remont);
                 list_Role.Add(_role);
                 list_Bonus.Add(_bonus);
                 list_Davnost.Add(_davnost);

                 if (checkBox1.Checked == true) // если разрешено обновление БД
                 {
                     Bonus_za_ZNR Temp = new Bonus_za_ZNR
                     {
                         ID = key++,

                         Period = comboBox1.Text.ToString(),
                         FIO = ws.Cell(i, 3).Value.ToString().Trim(),
                         Dolzhnost = _dolzhnost,
                         Filial = ws.Cell(i, 6).Value.ToString().Trim(),
                         Zapros = _zapros,
                         Remont = _remont,
                         Sum_klient = _sum_klient,
                         Sum_TSK = _sum_tsk,
                         Role = _role,
                         KTU = _ktu,
                         Davnost = _davnost,
                         Srok = _srok,
                         Bonus = _bonus,
                       //  Percent = _persent

                     };



                     db.Bonus_za_ZNR.InsertOnSubmit(Temp);
                 }
             }

             label3.Text = "Процесс: обработка таблицы Зарплата завершена";
             label3.Update();

             if (checkBox1.Checked == true)
             {// если разрешено обновление БД
                 db.SubmitChanges();
                 label3.Text = "Процесс: сохранение таблицы Зарплата в БД";
             }

             summa_bonusov = Convert.ToDouble(String.Format("{0:0.##}", summa_bonusov));
             label2.Text = summa_bonusov.ToString();

             listBox1.Items.Add("Внешние работы макс. % механика: " + maxPersent[1]);
             listBox1.Items.Add("Внешние работы мин. % механика: " + minPersent[1]);
             listBox1.Items.Add("------------------------------------------------");
             listBox1.Items.Add("Внешние работы макс. % механика: " + maxPersent[2]);
             listBox1.Items.Add("Внешние работы мин. % механика: " + minPersent[2]);
            
            list_ZNR_new = list_ZNR.Distinct().ToList();
              
            int All = list_Role.Count();

             listBox1.Items.Add("------------------------------------------------");
              listBox1.Items.Add("Проверка дубликатов ЗнР");
              int count_dubles = 0;

            if(checkBox3.Checked == true) // проверка премии ПП
            {
                   progressBar1.Value = 0;

                    var Proverka_PP = (from c in db.Bonus_za_ZNR where (c.Period!= comboBox1.Text.ToString()) select c).ToList();

                    var CurrentPeriod = (from c in db.Bonus_za_ZNR where (c.Period == comboBox1.Text.ToString()) select c).ToList();

                    progressBar1.Maximum = CurrentPeriod.Count();



                    foreach (var c in CurrentPeriod)
                    {
                        progressBar1.Value++;

                        foreach (var d in Proverka_PP)
                        {
                            if(c.Remont == d.Remont && c.Role == d.Role)
                            {
                                listBox1.Items.Add(c.Remont + " " + c.Role + "уже было в " + d.Period); count_dubles++;
                            }

                        }
                    }

                    listBox1.Items.Add("Найдено оплаченных ЗнР в ПП " + count_dubles++);
            }

            if (checkBox2.Checked == true) // Подсчет таблицы ФОТ
            {
           
            Double _FOT_prod_mat = 0;
            Double _FOT_prod_trud = 0;
            Double _FOT_brigad_mat = 0;
            Double _FOT_brigad_trud = 0;
            Double _FOT_oform_mat = 0;
            Double _FOT_oform_trud = 0;
            Double _FOT_mehan_mat = 0;
            Double _FOT_mehan_trud = 0;
            Double _FOT_mehan_rashod = 0;

            Double _Percent_prod_mat = 0;
            Double _Percent_prod_trud = 0;
            Double _Percent_brigad_mat = 0;
            Double _Percent_brigad_trud = 0;
            Double _Percent_oform_mat = 0;
            Double _Percent_oform_trud = 0;
           
            Double _Percent_mehan_trud = 0;
            Double _Percent_mehan_rashod = 0;

            Double _Summa_mat = 1, _Summa_trud = 1, _Summa_rashod = 1;
            String _Truck = "", _klient = "";

            key = (from c in db.FOT_analise select c.ID).Count();

            var temp_FOT_analise = (from c in db.Remont select c).ToList();

            int temp_FOT_analise_count = (from c in db.Remont select c).Count();

            int k = 0;

                       progressBar1.Maximum = list_ZNR_new.Count();
            progressBar1.Value = 0;

            //var Remont_num = (from c in db.Remont where select c.);

            int currentZNR;

            foreach (Int32 c in list_ZNR_new)
            {
                progressBar1.Value++;

                k++;

                label3.Text = "Процесс: обработка аналитики ФОТ " + k + " из " + progressBar1.Maximum;
                label3.Update();

                currentZNR = c;

               

                for (int i = 0; i < temp_FOT_analise_count; i++ )
                { 

                    if (c == temp_FOT_analise[i].Remont_num)
                    {


                                 _Summa_mat = (Double)temp_FOT_analise[i].Summa_mat;

                                          _Summa_trud = (Double)temp_FOT_analise[i].Summa_trud;

                                  _Summa_rashod = (Double)temp_FOT_analise[i].Summa_rashod; 


                                  _Truck = temp_FOT_analise[i].Truck_model;

                          _klient = temp_FOT_analise[i].Klient;

                                  break;
                        }

                    }

                for (int j = 0; j < All; j++)
                {
                    
                    if (list_ZNR[j].Equals(c) && (list_Role[j].Contains("Продавец труд") || list_Role[j].Contains("Контракт труд")))
                    {
                        _FOT_prod_trud = _FOT_prod_trud + list_Bonus[j];
                        _davnost = list_Davnost[j];
                    }
                    if (list_ZNR[j].Equals(c) && (list_Role[j].Contains("Продавец материалы") || list_Role[j].Contains("Контракт материалы")))
                    {
                        _FOT_prod_mat = _FOT_prod_mat + list_Bonus[j];
                        _davnost = list_Davnost[j];
                    }
                    if (list_ZNR[j].Equals(c) && list_Role[j].Contains("Бригадир труд"))
                    {
                        _FOT_brigad_trud = _FOT_brigad_trud + list_Bonus[j];
                        _davnost = list_Davnost[j];
                    }
                    if (list_ZNR[j].Equals(c) && list_Role[j].Contains("Бригадир материалы"))
                    {
                        _FOT_brigad_mat = _FOT_brigad_mat + list_Bonus[j];
                        _davnost = list_Davnost[j];
                    }
                    if (list_ZNR[j].Equals(c) && list_Role[j].Contains("Оформитель труд"))
                    {
                        _FOT_oform_trud = _FOT_oform_trud + list_Bonus[j];
                        _davnost = list_Davnost[j];
                    }
                    if (list_ZNR[j].Equals(c) && list_Role[j].Contains("Оформитель материалы"))
                    {
                        _FOT_oform_mat = _FOT_oform_mat + list_Bonus[j];
                        _davnost = list_Davnost[j];
                    }
                    if (list_ZNR[j].Equals(c) && (list_Role[j].Contains("Механик труд") || list_Role[j].Contains("Механик труд док") || list_Role[j].Contains("Труд ручной ФОТ")))
                    {
                        _FOT_mehan_trud = _FOT_mehan_trud + list_Bonus[j];
                    }
                    if (list_ZNR[j].Equals(c) && (list_Role[j].Contains("Механик расходы") || list_Role[j].Contains("Механик расходы док") || list_Role[j].Contains("Расходы ручной ФОТ")))
                    {
                        _FOT_mehan_rashod = _FOT_mehan_rashod + list_Bonus[j];
                    }
                                     

                    if (_davnost!=0 && _Summa_trud != 0)
                    { 
                    _Percent_prod_trud = _FOT_prod_trud * 100 / _Summa_trud / _davnost;
                    _Percent_prod_trud = Convert.ToDouble(String.Format("{0:0.##}", _Percent_prod_trud));

                    _Percent_brigad_trud = _FOT_brigad_trud * 100 / _Summa_trud / _davnost;
                    _Percent_brigad_trud = Convert.ToDouble(String.Format("{0:0.##}", _Percent_brigad_trud));

                    _Percent_oform_trud = _FOT_oform_trud * 100 / _Summa_trud / _davnost;
                    _Percent_oform_trud = Convert.ToDouble(String.Format("{0:0.##}", _Percent_oform_trud));

                    _Percent_mehan_trud = _FOT_mehan_trud * 100 / _Summa_trud;
                    _Percent_mehan_trud = Convert.ToDouble(String.Format("{0:0.##}", _Percent_mehan_trud));
                                            
                    }

                    if (_davnost != 0 && _Summa_mat != 0)
                    {
                        _Percent_prod_mat = _FOT_prod_mat * 100 / _Summa_mat / _davnost;
                        _Percent_prod_mat = Convert.ToDouble(String.Format("{0:0.###}", _Percent_prod_mat));

                        _Percent_brigad_mat = _FOT_brigad_mat * 100 / _Summa_mat / _davnost;
                        _Percent_brigad_mat = Convert.ToDouble(String.Format("{0:0.###}", _Percent_brigad_mat));

                        _Percent_oform_mat = _FOT_oform_mat * 100 / _Summa_mat / _davnost;
                        _Percent_oform_mat = Convert.ToDouble(String.Format("{0:0.###}", _Percent_oform_mat));

                    }

                    if (_davnost != 0 && _Summa_rashod != 0)
                    {
                        _Percent_mehan_rashod = _FOT_mehan_rashod * 100 / _Summa_rashod;
                        _Percent_mehan_rashod = Convert.ToDouble(String.Format("{0:0.##}", _Percent_mehan_rashod));
                    }
                }

                FOT_analise Temp = new FOT_analise
                {
                    ID = key++,

                    Period = comboBox1.Text.ToString(),
                    Remont = c,
                    FOT_prod_mat = _FOT_prod_mat,
                    FOT_prod_trud = _FOT_prod_trud,
                    FOT_brigad_mat = _FOT_brigad_mat,
                    FOT_brigad_trud = _FOT_brigad_trud,
                    FOT_oform_mat = _FOT_oform_mat,
                    FOT_oform_trud = _FOT_oform_trud,
                    FOT_mehan_mat = _FOT_mehan_mat,
                    FOT_mehan_trud = _FOT_mehan_trud,
                    FOT_mehan_rashod = _FOT_mehan_rashod,
                    Truck = _Truck,
                    Summa_mat = _Summa_mat,
                    Summa_trud = _Summa_trud,
                    Summa_rashod = _Summa_rashod,
                    Klient = _klient,
                    Davnost = _davnost,
                    Percent_prod_trud = _Percent_prod_trud,
                    Percent_prod_mat = _Percent_prod_mat,
                    Percent_brigad_mat = _Percent_brigad_mat,
                    Percent_brigad_trud = _Percent_brigad_trud,
                    Percent_mehan_trud = _Percent_mehan_trud,
                    Percent_mehan_rashod = _Percent_mehan_rashod,
                    Percent_oform_mat = _Percent_oform_mat,
                    Percent_oform_trud = _Percent_oform_trud
                    
                };

                db.FOT_analise.InsertOnSubmit(Temp);

                 _FOT_prod_mat = 0;
                 _FOT_prod_trud = 0;
                 _FOT_brigad_mat = 0;
                 _FOT_brigad_trud = 0;
                 _FOT_oform_mat = 0;
                 _FOT_oform_trud = 0;
                 _FOT_mehan_mat = 0;
                 _FOT_mehan_trud = 0;
                 _FOT_mehan_rashod = 0;
                _Summa_mat = 1;
                _Summa_trud = 1;
                _Summa_rashod = 1;
                _Truck = "";
                _klient = "";
                _Percent_prod_trud = 0;
                    _Percent_prod_mat = 0;
                _Percent_brigad_mat = 0;
                 _Percent_brigad_trud = 0;
                 _Percent_mehan_trud = 0;
                 _Percent_mehan_rashod = 0;
                 _Percent_oform_mat = 0;
                 _Percent_oform_trud = 0;
            }

            label3.Text = "Процесс: запись аналитики ФОТ в БД";
            label3.Update();

          db.SubmitChanges();

            wb.Dispose();

            label3.Text = "Процесс: Завершено";
            label3.Update();
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {

            style = NumberStyles.Number;
            culture = CultureInfo.CreateSpecificCulture("en-GB");

            var wb = new XLWorkbook(textBox2.Text);
            var ws = wb.Worksheet(1);

            int key = (from c in db.Remont select c.Remont_num).Count();
            bool Exist = false;
            int RowsCount = 0;

            
            Int32 _zapros = 0, _remont = 0;
            Double _Summa_mat = 0, _Summa_trud = 0, _Summa_rashod = 0;

            RowsCount = ws.RowsUsed().Count();

            var Remont_num = (from c in db.Remont select c.Remont_num).ToList();

            for (int i = 2; i <= RowsCount; i++)
            {

                if (Double.TryParse(ws.Cell(i, 29).Value.ToString().Replace(",","."), style, culture, out _Summa_mat) == true)
                    _Summa_mat = Convert.ToDouble(String.Format("{0:f}", _Summa_mat));
                else
                    _Summa_mat = 0;

                if (Double.TryParse(ws.Cell(i, 30).Value.ToString().Replace(",", "."), style, culture, out _Summa_trud) == true)
                    _Summa_trud = Convert.ToDouble(String.Format("{0:f}", _Summa_trud));
                else
                    _Summa_trud = 0;

                if (Double.TryParse(ws.Cell(i, 31).Value.ToString().Replace(",", "."), style, culture, out _Summa_rashod) == true)
                    _Summa_rashod = Convert.ToDouble(String.Format("{0:f}", _Summa_rashod));
                else
                    _Summa_rashod = 0;

               
                Int32.TryParse(ws.Cell(i, 10).Value.ToString(), style, culture, out _zapros);

                Int32.TryParse(ws.Cell(i, 20).Value.ToString(), style, culture, out _remont);

                foreach (int c in Remont_num)
                {
                    if (_remont.Equals(c))
                    {
                        Exist = true; break;
                    }
                    else { Exist = false; }
                }

                if (Exist == false)
                {
                    

                    Remont Temp = new Remont
                    {

                        Zapros = _zapros,
                        Remont_num = _remont,
                        Filial = ws.Cell(i, 2).Value.ToString(),
                        Klient = ws.Cell(i, 4).Value.ToString(),
                        Truck_model = ws.Cell(i, 7).Value.ToString(),
                        Prichina = ws.Cell(i, 9).Value.ToString(),
                        ZNR_Date_Open = ws.Cell(i, 22).Value.ToString(),
                        ZNR_Date_Close = ws.Cell(i, 24).Value.ToString(),
                        Summa_mat = _Summa_mat,
                        Summa_trud = _Summa_trud,
                        Summa_rashod = _Summa_rashod,
                        Status = ws.Cell(i, 23).Value.ToString()
                    };

                    db.Remont.InsertOnSubmit(Temp);
                }
           
        }

            db.SubmitChanges();
       
    }

        private void button3_Click(object sender, EventArgs e)
        {
            List<int> list1 = new List<int>();

            List<String> list2 = new List<string>();
            List<String> list3 = new List<string>();

           

            //new comment

            try
            {
                // Only get files that begin with the letter "c."
                string[] dirs = Directory.GetFiles(textBox1.Text, "*.xlsx");

                listBox1.Items.Add("Обнаружено " + dirs.Length + " файлов в каталоге.");

                foreach (string dir in dirs)
                {
                    var wb = new XLWorkbook(dir);
                    var ws = wb.Worksheet(1);

                    int excelRow = ws.RowsUsed().Count();

                  
                    for (int i = 1; i <= 20; i++)
                    {
                        for (int j = 1; j <= 20; j++)
                        {
                            if (ws.Cell(i, j).Value.ToString().Equals("Проверка факт"))
                            {// Занесение CRM
                                listBox1.Items.Add("Обнаружен отчет CRM. Имя " + dir);
                                listBox1.Update();

                                list1.Clear(); list2.Clear(); list3.Clear();

                                list1.Add(2);
                                list1.Add(5);
                                list1.Add(12);
                                list1.Add(15);


                                list2.Add("Tab_num"); 
                                list2.Add("Filial");
                                list2.Add("viezd_pers");
                                list2.Add("viezd_vsego");


                                list3.Add("int"); 
                                list3.Add("string");
                                list3.Add("int");
                                list3.Add("int");


                                AddToBase(dir, "crm_max", comboBox1.SelectedItem.ToString(), list1, list2, list3); //CRM максимов
                            }

                            if (ws.Cell(i, j).Value.ToString().Equals("Звонки (пл/фт)"))
                            {
                                listBox1.Items.Add("Обнаружен отчет ККДК. Имя " + dir);
                                listBox1.Update();

                                list1.Clear(); list2.Clear(); list3.Clear();

                                list1.Add(3);
                                list1.Add(4);
                                list1.Add(7);
                                list1.Add(8);
                                list1.Add(9);

                                list2.Add("crm_filial"); 
                                list2.Add("prod_count");
                                list2.Add("zvonok_count");
                                list2.Add("viezd_count");
                                list2.Add("smeta_count");


                                list3.Add("string"); 
                                list3.Add("int");
                                list3.Add("parse_int");
                                list3.Add("parse_int");
                                list3.Add("parse_int");

                                AddToBase(dir, "Kkdk", comboBox1.SelectedItem.ToString(), list1, list2, list3); //CRM

                            }
                            if (ws.Cell(i, j).Value.ToString().Equals("Отв. Куратор"))
                            {
                                listBox1.Items.Add("Обнаружен список кураторов. Имя " + dir);
                                listBox1.Update();


                                list1.Clear(); list2.Clear(); list3.Clear();

                                list1.Add(1);
                                list1.Add(2);
                                list1.Add(3);


                                list2.Add("kurator_fio"); 
                                list2.Add("kurator_id");
                                list2.Add("kurator_filial");

                                list3.Add("string"); 
                                list3.Add("int");
                                list3.Add("string");

                                AddToBase(dir, "Motivation", comboBox1.SelectedItem.ToString(), list1, list2, list3); //куратор


                            }

                            if (ws.Cell(i, j).Value.ToString().Equals("представление: ЗО задача ЗнР"))
                            {
                                listBox1.Items.Add("Обнаружен отчет ЗО задача ЗнР. Имя " + dir);
                                listBox1.Update();

                                list1.Clear(); list2.Clear(); list3.Clear();

                                list1.Add(2);
                                list1.Add(4);
                                list1.Add(28);
                                list1.Add(29);
                                list1.Add(30);
                                list1.Add(31);
                                list1.Add(12);
                                list1.Add(23);
                                list1.Add(20);
                                list1.Add(14);
                                list1.Add(24);
                                list1.Add(7);
                                list1.Add(9);

                                list2.Add("Filial"); 
                                list2.Add("Klient");
                                list2.Add("Summa_vsego");
                                list2.Add("Summa_mat");
                                list2.Add("Summa_trud");
                                list2.Add("Summa_rashod");
                                list2.Add("Status_ZO");
                                list2.Add("Status_ZNR");
                                list2.Add("Nomer_ZNR");
                                list2.Add("ZO_zakr_date");
                                list2.Add("ZNR_zakr_date");
                                list2.Add("Truck_model");
                                list2.Add("Prichina");

                                list3.Add("string"); 
                                list3.Add("string");
                                list3.Add("float");
                                list3.Add("float");
                                list3.Add("float");
                                list3.Add("float");
                                list3.Add("string");
                                list3.Add("string");
                                list3.Add("int");
                                list3.Add("date");
                                list3.Add("date");
                                list3.Add("string");
                                list3.Add("string");

                                AddToBase(dir, "Zakr", comboBox1.SelectedItem.ToString(), list1, list2, list3); //закрывашки


                            }

                            if (ws.Cell(i, j).Value.ToString().Equals("Остаток не поступивших оригиналов документов на конец периода"))
                            {
                                listBox1.Items.Add("Обнаружен отчет ОД. Имя " + dir);
                                listBox1.Update();


                                list1.Clear(); list2.Clear(); list3.Clear();

                                list1.Add(1);
                                list1.Add(9);
                                list1.Add(16);


                                list2.Add("OD_date"); 
                                list2.Add("summa_doc");
                                list2.Add("Filial");


                                list3.Add("date"); 
                                list3.Add("float");
                                list3.Add("string");


                                AddToBase(dir, "net_OD", comboBox1.SelectedItem.ToString(), list1, list2, list3); //непоступившие ОД


                            }

                            if (ws.Cell(i, j).Value.ToString().Equals("Остаток - Просроченная задолженность на конец периода"))
                            {
                                listBox1.Items.Add("Обнаружен отчет дебеторка. Имя " + dir);
                                listBox1.Update();


                                list1.Clear(); list2.Clear(); list3.Clear();

                                list1.Add(4);
                                list1.Add(16);
                                list1.Add(25);


                                list2.Add("deb_filial"); 
                                list2.Add("deb_date");
                                list2.Add("deb_saldo");


                                list3.Add("string"); 
                                list3.Add("date");
                                list3.Add("float");


                                AddToBase(dir, "Debitor", comboBox1.SelectedItem.ToString(), list1, list2, list3); //дебиторка


                            }

                            if (ws.Cell(i, j).Value.ToString().Equals("Выработка механиков всего за период"))
                            {
                                listBox1.Items.Add("Обнаружен отчет выработка. Имя " + dir);
                                listBox1.Update();


                                list1.Clear(); list2.Clear(); list3.Clear();

                                list1.Add(4);
                                list1.Add(5);                                
                                list1.Add(12);
                                list1.Add(6);

                                list2.Add("Tab_num"); 
                                list2.Add("FIO");
                                list2.Add("time_vsego");
                                list2.Add("Data_priema");


                                list3.Add("int"); 
                                list3.Add("string");
                                list3.Add("float");
                                list3.Add("date");

                                AddToBase(dir, "Virabotka", comboBox1.SelectedItem.ToString(), list1, list2, list3); //выработка


                            }

                            if (ws.Cell(i, j).Value.ToString().Equals("Ведомость"))
                            {

                                listBox1.Items.Add("Обнаружен отчет Ведомость. Имя " + dir);
                                listBox1.Update();


                                list1.Clear(); list2.Clear(); list3.Clear();

                                list1.Add(2);
                                list1.Add(3);
                                list1.Add(4);
                                list1.Add(5);


                                list2.Add("Filial"); 
                                list2.Add("tab");
                                list2.Add("FIO");
                                list2.Add("Dolzhnost");

                                list3.Add("string"); 
                                list3.Add("int");
                                list3.Add("string");                                
                                list3.Add("string");


                                AddToBase(dir, "Shtat", comboBox1.SelectedItem.ToString(), list1, list2, list3); //сотрудники


                            }

                            if (ws.Cell(i, j).Value.ToString().Equals("Уволенные"))
                            {

                                listBox1.Items.Add("Обнаружен отчет Уволенные. Имя " + dir);
                                listBox1.Update();


                                list1.Clear(); list2.Clear(); list3.Clear();

                                list1.Add(3);
                                list1.Add(4);
                                list1.Add(5);
                              


                                list2.Add("Tab_num"); 
                                list2.Add("FIO");
                                list2.Add("Date_uvol");
                              

                                list3.Add("int"); 
                                list3.Add("string");
                                list3.Add("date");
                                


                                AddToBase(dir, "Uvolen", comboBox1.SelectedItem.ToString(), list1, list2, list3); //уволенные


                            }

                            if (ws.Cell(i, j).Value.ToString().Equals("Планы"))
                            {

                                listBox1.Items.Add("Обнаружен отчет Планы. Имя " + dir);
                                listBox1.Update();


                                list1.Clear(); list2.Clear(); list3.Clear();

                                list1.Add(1);
                                list1.Add(2);
                            
                            
                                list2.Add("Filial"); 
                                list2.Add("Sum_plan");
                               
                                
                                list3.Add("string"); 
                                list3.Add("float");
                               

                                AddToBase(dir, "Plann", comboBox1.SelectedItem.ToString(), list1, list2, list3); //уволенные


                            }

                            if (ws.Cell(i, j).Value.ToString().Equals("ФЗП"))
                            {

                                listBox1.Items.Add("Обнаружен отчет ФЗП. Имя " + dir);
                                listBox1.Update();


                                list1.Clear(); list2.Clear(); list3.Clear();

                                list1.Add(1);
                                list1.Add(2);
                                list1.Add(3);

                                list2.Add("FIO"); 
                                list2.Add("Tab_num");
                                list2.Add("Oklad");


                                list3.Add("string"); 
                                list3.Add("int");
                                list3.Add("int");

                                AddToBase(dir, "Stavki", comboBox1.SelectedItem.ToString(), list1, list2, list3); //уволенные


                            }

                            if (ws.Cell(i, j).Value.ToString().Equals("Истек срок"))
                            {

                                listBox1.Items.Add("Обнаружен отчет Расчет ЗП. Имя " + dir);
                                listBox1.Update();


                                list1.Clear(); list2.Clear(); list3.Clear();

                                list1.Add(3);
                                list1.Add(6);
                                list1.Add(7);
                                list1.Add(8);
                                list1.Add(10);
                                list1.Add(12);
                                list1.Add(13);
                                list1.Add(16);
                                list1.Add(17);
                                list1.Add(18);
                               
                               list1.Add(20);
                               list1.Add(21);

                                list2.Add("FIO"); 
                                list2.Add("Filial");
                                list2.Add("Nomer_ZO");
                                list2.Add("Tip_ZO");
                                list2.Add("Nomer_ZNR");
                                list2.Add("Sum_klient");
                                list2.Add("Sum_tsk");
                                list2.Add("Role");
                                list2.Add("KTU");
                                list2.Add("Davnost");
                               list2.Add("Srok");
                              list2.Add("Raschet");

                                list3.Add("string"); 
                                list3.Add("string");
                                list3.Add("int");
                                list3.Add("string");
                                list3.Add("int");
                                list3.Add("float");
                                list3.Add("float");
                                list3.Add("string");
                                list3.Add("ktu_float");
                                list3.Add("ktu_float");
                                list3.Add("string");
                                list3.Add("zp_float");

                                AddToBase(dir, "ZP_po_FIO", comboBox1.SelectedItem.ToString(), list1, list2, list3); //уволенные


                            }


                        }
                    }


                }
            }
            catch (Exception f)
            {
                string s = String.Concat("The process failed: ",f.ToString());
                listBox1.Items.Add(s); listBox1.Update();
            }

            string Period = "'%" + comboBox1.SelectedItem.ToString() + "%'";

            System.Data.SqlClient.SqlConnection sqlConnection1 =
                                  new System.Data.SqlClient.SqlConnection(@"Data Source=ROMAN-ПК\SQLEXPRESS;Initial Catalog=Portal_base;Integrated Security=True");

            System.Data.SqlClient.SqlCommand cmd = new System.Data.SqlClient.SqlCommand();
            cmd.CommandType = System.Data.CommandType.Text;

            cmd.Connection = sqlConnection1;

            sqlConnection1.Open();

            try
            {   // Open the text file using a stream reader.
                using (StreamReader sr = new StreamReader("C:\\Users\\RomanNB\\Documents\\zarplata.sql"))
                {
                    // Read the stream to a string, and write the string to the console.
                    cmd.CommandText = sr.ReadToEnd();

                    cmd.ExecuteNonQuery();

                }

            }
            catch (Exception ee)
            {
                listBox1.Items.Add("Ошибка чтения обработки ЗП: " + ee.Message);

            }

            //C:\Users\RomanNB\Documents

            cmd.CommandText = "update Motivation SET Motivation.prod_count = Kkdk.prod_count, plan_viezd = Kkdk.prod_count * 20, fact_viezd = Kkdk.viezd_count, plan_zvonok = Kkdk.prod_count * 100, fact_zvonok = Kkdk.zvonok_count, plan_smeta = Kkdk.prod_count * 50, fact_smeta = Kkdk.smeta_count from Motivation inner join Kkdk on Kkdk.crm_filial = Motivation.kurator_filial AND Kkdk.Period like " + Period + "; ";
            cmd.ExecuteNonQuery();
            cmd.CommandText = "update Motivation SET fact_viezd_max = v2.v_vsego from Motivation inner join (select Filial,SUM(viezd_vsego) v_vsego, Period from crm_max group by Filial, Period) v2 on v2.Filial = Motivation.kurator_filial AND v2.Period like "+Period+";";
            cmd.ExecuteNonQuery();
            cmd.CommandText = "update Motivation SET fact_zakr = z.summa from Motivation inner join (select Filial,SUM(Summa_vsego) summa from Zakr where Klient != 'Техстройконтракт' and ZO_zakr_date >= '" + textBox3.Text.Trim() +"' and ZO_zakr_date < dateadd(month, +1, '" + textBox3.Text.Trim() + "') group by Filial) z on z.Filial like '%' + Motivation.kurator_filial + '%' where Period like " + Period + "; ";
            cmd.ExecuteNonQuery();
            cmd.CommandText = "update Motivation SET debitora = d.saldo from Motivation inner join (select deb_filial,SUM(deb_saldo) saldo, Period from Debitor where deb_date < '" + textBox3.Text.Trim() + "' and deb_date > dateadd(year, -3, '" + textBox3.Text.Trim() + "') group by deb_filial, Period) d on d.deb_filial = Motivation.kurator_filial AND d.Period like " + Period + ";";
            cmd.ExecuteNonQuery();
            cmd.CommandText = "update Motivation SET net_od = od.doc from Motivation inner join (select Filial,SUM(summa_doc) doc, Period from net_OD where OD_date < '" + textBox3.Text.Trim() + "' and OD_date > dateadd(year, -3, '" + textBox3.Text.Trim() + "') group by Filial, Period) od on od.Filial = Motivation.kurator_filial AND od.Period like " + Period + ";";
            cmd.ExecuteNonQuery();
            cmd.CommandText = "update Motivation SET plan_zakr = pl.Sum_plan from Motivation inner join (select Filial,Sum_plan, Period from Plann ) pl on pl.Filial = Motivation.kurator_filial AND pl.Period like " + Period + ";";

            cmd.ExecuteNonQuery();
            cmd.CommandText = "update Motivation SET vnutr_zakr = z.summa from Motivation inner join (select Filial,SUM(Summa_trud) summa from Zakr where Filial like '%Сервис%' and Klient  = 'Техстройконтракт' and ZO_zakr_date >= '" + textBox3.Text.Trim() + "' and ZO_zakr_date < dateadd(month, +1, '" + textBox3.Text.Trim() + "') group by Filial) z on z.Filial like '%' + Motivation.kurator_filial + '%';";
            cmd.ExecuteNonQuery();
            cmd.CommandText = "delete from Virabotka where Virabotka.FIO LIKE '%?%' OR Virabotka.Data_priema > '" + textBox3.Text.Trim() + "' AND Period like " + Period + "; ";
            cmd.ExecuteNonQuery();
            cmd.CommandText = "update Motivation SET virabotka = vir.summa, mehan_count = m_count from Motivation inner join (select Filial,SUM(time_vsego) summa, COUNT(Virabotka.FIO) m_count, Virabotka.Period from Virabotka inner join (select * from Shtat where Period like " + Period + " ) t on t.tab = Virabotka.Tab_num AND t.Period = Virabotka.Period group by Filial, Virabotka.Period) vir on vir.Filial like '%' + Motivation.kurator_filial + '%' AND vir.Period like " + Period + ";";
            cmd.ExecuteNonQuery();
            cmd.CommandText = "delete from Virabotka where Virabotka.FIO LIKE '%?%' OR Virabotka.Data_priema > '" + textBox3.Text.Trim() + "' AND Period like " + Period + "; ";
            cmd.ExecuteNonQuery();


            // Зарплата

        
            // Коммит 1

            // Комит 2

            //Создаем workbook
            var workbook1 = new XLWorkbook();
            //Название страницы
            var worksheet1 = workbook1.Worksheets.Add("Премия за ЗнР");
            //Заполняем ячейки

            var ZP_po_FIO = (from c in db.Test_1 where c.Period == comboBox1.SelectedItem.ToString() select c).ToList();
            ZP_po_FIO.Add(new Test_1 { Truck = " " });
            int row_num;

            worksheet1.Row(1).Style.Font.Bold = true;
            worksheet1.Row(1).Style.Alignment.WrapText = true;
            worksheet1.Row(1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            worksheet1.Range(1, 1, 1, worksheet1.ColumnsUsed().Count()).Style.Fill.BackgroundColor = XLColor.LightGray;


            worksheet1.Cell(1, 1).Value = "Номер ЗнР";
            worksheet1.Cell(1, 2).Value = "Клиент";
            worksheet1.Cell(1, 3).Value = "Машина";
            worksheet1.Cell(1, 4).Value = "Дата закр. ЗО";
            worksheet1.Cell(1, 5).Value = "Дата закр. ЗнР";
            worksheet1.Cell(1, 6).Value = "Давн";
            worksheet1.Cell(1, 7).Value = "Причина";
            worksheet1.Cell(1, 8).Value = "Труд";
            worksheet1.Cell(1, 9).Value = "Расходы";
            worksheet1.Cell(1, 10).Value = "Материал";
            worksheet1.Cell(1, 11).Value = "ЗП прод труд";
            worksheet1.Cell(1, 12).Value = "ЗП прод мат";
            worksheet1.Cell(1, 13).Value = "ЗП бриг труд";
            worksheet1.Cell(1, 14).Value = "ЗП бриг мат";
            worksheet1.Cell(1, 15).Value = "ЗП мех труд закр";
            worksheet1.Cell(1, 16).Value = "ЗП мех расход закр";
            worksheet1.Cell(1, 17).Value = "ЗП мех труд док";
            worksheet1.Cell(1, 18).Value = "ЗП мех расход док";
            worksheet1.Cell(1, 19).Value = "ЗП оформ мат";
            worksheet1.Cell(1, 20).Value = "ЗП оформ труд";
            worksheet1.Cell(1, 21).Value = "% прод мат";
            worksheet1.Cell(1, 22).Value = "% прод труд";
            worksheet1.Cell(1, 23).Value = "% бригад труд";
            worksheet1.Cell(1, 24).Value = "% бригад мат";
            worksheet1.Cell(1, 25).Value = "% мех труд закр";
            worksheet1.Cell(1, 26).Value = "% мех труд док";
            worksheet1.Cell(1, 27).Value = "% оформ мат";
            worksheet1.Cell(1, 28).Value = "% оформ труд";
            worksheet1.Cell(1, 29).Value = "Период";

            row_num = 2;

            foreach (var c in ZP_po_FIO)
            {
                            
               
                    worksheet1.Cell(row_num, 1).Value = c.Remont_num;
                worksheet1.Cell(row_num, 2).Value = c.Klient;
                worksheet1.Cell(row_num, 3).Value = c.Truck;

                worksheet1.Cell(row_num, 4).Value = c.Data_zakr_ZO;
                worksheet1.Cell(row_num, 4).Style.NumberFormat.Format = "mmm-yy";

                worksheet1.Cell(row_num, 5).Value = c.Data_zakr_ZNR;
                worksheet1.Cell(row_num, 5).Style.NumberFormat.Format = "mmm-yy";

                worksheet1.Cell(row_num, 6).Value = c.Davnost;
                worksheet1.Cell(row_num, 7).Value = c.Prichina;

                worksheet1.Cell(row_num, 8).Value = c.Summa_trud;
                worksheet1.Cell(row_num, 8).Style.NumberFormat.Format = "#";

                worksheet1.Cell(row_num, 9).Value = c.Summa_rashod;
                worksheet1.Cell(row_num, 9).Style.NumberFormat.Format = "#";

                worksheet1.Cell(row_num, 10).Value = c.Summa_mat;
                worksheet1.Cell(row_num, 10).Style.NumberFormat.Format = "#";

                worksheet1.Cell(row_num, 11).Value = c.ZP_prod_trud;
                worksheet1.Cell(row_num, 11).Style.NumberFormat.Format = "#";

                worksheet1.Cell(row_num, 12).Value = c.ZP_prod_mat;
                worksheet1.Cell(row_num, 12).Style.NumberFormat.Format = "#";

                worksheet1.Cell(row_num, 13).Value = c.ZP_brigad_trud;
                worksheet1.Cell(row_num, 13).Style.NumberFormat.Format = "#";

                worksheet1.Cell(row_num, 14).Value = c.ZP_brigad_mat;
                worksheet1.Cell(row_num, 14).Style.NumberFormat.Format = "#";

                worksheet1.Cell(row_num, 15).Value = c.ZP_meh_trud_zakr;
                worksheet1.Cell(row_num, 15).Style.NumberFormat.Format = "#";

                worksheet1.Cell(row_num, 16).Value = c.ZP_meh_rashod_zakr;
                worksheet1.Cell(row_num, 16).Style.NumberFormat.Format = "#";

                worksheet1.Cell(row_num, 17).Value = c.ZP_meh_trud_dok;
                worksheet1.Cell(row_num, 17).Style.NumberFormat.Format = "#";

                worksheet1.Cell(row_num, 18).Value = c.ZP_meh_rashod_doc;
                worksheet1.Cell(row_num, 18).Style.NumberFormat.Format = "#";

                worksheet1.Cell(row_num, 19).Value = c.ZP_oform_mat;
                worksheet1.Cell(row_num, 19).Style.NumberFormat.Format = "#";

                worksheet1.Cell(row_num, 20).Value = c.ZP_oform_trud;
                worksheet1.Cell(row_num, 20).Style.NumberFormat.Format = "#";

                worksheet1.Cell(row_num, 21).Value = c.Procent_prod_mat;
                worksheet1.Cell(row_num, 22).Value = c.Procent_prod_trud;
                worksheet1.Cell(row_num, 23).Value = c.Procent_brigad_trud;
                worksheet1.Cell(row_num, 24).Value = c.Procent_brigad_mat;
                worksheet1.Cell(row_num, 25).Value = c.Procent_meh_trud_zakr;
                worksheet1.Cell(row_num, 26).Value = c.Procent_meh_trud_doc;
                worksheet1.Cell(row_num, 27).Value = c.Procent_oform_mat;
                worksheet1.Cell(row_num, 28).Value = c.Procent_oform_trud;
                worksheet1.Cell(row_num, 29).Value = c.Period;

                row_num++;
            }

             workbook1.Worksheets.Add("Анализ ЗнР");

            worksheet1.Cell(1, 1).Value = "Выбираем ЗнР, которые есть в закрывашках, но нет в ЗП ";

            /*

            */

            workbook1.SaveAs(textBox1.Text + "\\zp_fio_svod.xlsx");

         
               // MessageBox.Show(«Документ создан!», «Внимание!», MessageBoxButtons.OK, MessageBoxIcon.Information);
            }

        private void button4_Click(object sender, EventArgs e)
        {
            int m = 2, n = 2, row_num;
            //Создаем workbook
            var workbook = new XLWorkbook();
            //Название страницы
            var worksheet = workbook.Worksheets.Add("Мотивация куратора");
            //Заполняем ячейки

            //  var CurrentPeriod = (from c in db.Bonus_za_ZNR where (c.Period == comboBox1.Text.ToString()) select c).ToList();

            var Kurator_Itog = (from c in db.Motivation where c.Period == comboBox1.SelectedItem.ToString() orderby c.kurator_fio select c).ToList();
            Kurator_Itog.Add(new Motivation { kurator_fio = " " });
            m = 2; n = 2;

            worksheet.Row(1).Style.Font.Bold = true;
            worksheet.Row(1).Style.Alignment.WrapText = true;
            worksheet.Row(1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            worksheet.Range(1, 1, 1, worksheet.ColumnsUsed().Count()).Style.Fill.BackgroundColor = XLColor.LightGray;


            worksheet.Cell(1, 1).Value = "Ответственный";
            worksheet.Cell(1, 2).Value = "Таб";
            worksheet.Cell(1, 3).Value = "Филиал";
            worksheet.Cell(1, 4).Value = "Период";
            worksheet.Cell(1, 5).Value = "Кол-во продавцов";
            worksheet.Cell(1, 6).Value = "План (выезд)";
            worksheet.Cell(1, 7).Value = "Факт (выезд)";
            worksheet.Cell(1, 8).Value = "Факт (выезд) макс";
            worksheet.Cell(1, 9).Value = "План (звонок)";
            worksheet.Cell(1, 10).Value = "Факт (звонок)";
            worksheet.Cell(1, 11).Value = "План (смета)";
            worksheet.Cell(1, 12).Value = "Факт (смета)";
            worksheet.Cell(1, 13).Value = "K (crm)";
            worksheet.Cell(1, 14).Value = "ЗП (crm)";
            worksheet.Cell(1, 15).Value = "План закрывашки";
            worksheet.Cell(1, 16).Value = "Факт закрывашки";
            worksheet.Cell(1, 17).Value = "K (закр)";
            worksheet.Cell(1, 18).Value = "ЗП (закр)";
            worksheet.Cell(1, 19).Value = "План ОД";
            worksheet.Cell(1, 20).Value = "Факт ОД";
            worksheet.Cell(1, 21).Value = "ЗП (ОД)";
            worksheet.Cell(1, 22).Value = "План деб";
            worksheet.Cell(1, 23).Value = "Факт деб";
            worksheet.Cell(1, 24).Value = "ЗП (деб)";
            worksheet.Cell(1, 25).Value = "Кол-во механиков";
            worksheet.Cell(1, 26).Value = "Выработка";
            worksheet.Cell(1, 27).Value = "К (выработка)";
            worksheet.Cell(1, 28).Value = "ЗП (выработка)";
            worksheet.Cell(1, 29).Value = "Внутр закр";
            worksheet.Cell(1, 30).Value = "K внутр";
            worksheet.Cell(1, 31).Value = "ЗП внутр";
            worksheet.Cell(1, 32).Value = "ЗП Итого";
            worksheet.Cell(1, 33).Value = "Корректировочный коэффициент";
            worksheet.Cell(1, 34).Value = "Корректировка";
            worksheet.Cell(1, 35).Value = "Итого мотивация";
            worksheet.Cell(1, 36).Value = "Санкции ККД";
            workbook.CalculateMode = XLCalculateMode.Auto;

            var oklad = (from okl in db.Stavki select okl).ToList();
            var pers_kurator = (from crm in db.crm_max where crm.Period == comboBox1.SelectedItem.ToString() select crm).ToList();


            foreach (var c in Kurator_Itog)
            {
                row_num = 5;

                if (m > 2 && (c.kurator_fio.Equals(worksheet.Cell(m - 1, 1).Value.ToString()) == false || c.kurator_fio.Equals(" ") == true) && !c.kurator_fio.Contains("итого") )
                {
                    worksheet.Cell(m, 1).Value = "Итого " + worksheet.Cell(m - 1, 1).Value.ToString();

                    worksheet.Cell(m, 2).Value = "ставка";



                    foreach (var d in oklad)
                    {
                        if (d.Tab_num.Value.ToString().Equals(worksheet.Cell(m - 1, 2).Value.ToString()))
                            worksheet.Cell(m, 3).Value = d.Oklad.Value;
                        worksheet.Cell(m, 3).Style.NumberFormat.Format = "#";
                    }
                    //  worksheet.Cell(m, 5).FormulaA1 = "=SUM(E" + n + ":E" + (m - 1) + ")";

                    // CRM

                    worksheet.Cell(m, row_num).FormulaA1 = "=SUM(" + worksheet.Range(worksheet.Cell(n, row_num).Address, worksheet.Cell(m - 1, row_num).Address) + ")"; worksheet.Cell(m, row_num).Style.Fill.BackgroundColor = XLColor.LightGray;
                    string prod_count_Address = worksheet.Cell(m, row_num).Address.ToString();
                    row_num++;
                    worksheet.Cell(m, row_num).FormulaA1 = "=SUM(" + worksheet.Range(worksheet.Cell(n, row_num).Address, worksheet.Cell(m - 1, row_num).Address) + ")+10"; worksheet.Cell(m, row_num).Style.Fill.BackgroundColor = XLColor.LightGray;
                    string plan_viezd_Address = worksheet.Cell(m, row_num).Address.ToString();
                    row_num++;
                    worksheet.Cell(m, row_num).FormulaA1 = "=SUM(" + worksheet.Range(worksheet.Cell(n, row_num).Address, worksheet.Cell(m - 1, row_num).Address) + ")"; worksheet.Cell(m, row_num).Style.Fill.BackgroundColor = XLColor.LightGray;
                    string fact_viezd_Address = worksheet.Cell(m, row_num).Address.ToString();
                    row_num++;

                    int kurator_viezd = 0;

                    foreach (var f in pers_kurator)
                    {
                        if (f.Tab_num.Value.ToString().Equals(worksheet.Cell(m - 1, 2).Value.ToString()))
                            kurator_viezd = f.viezd_pers.Value;
                    }


                    worksheet.Cell(m, row_num).FormulaA1 = "=SUM(" + worksheet.Range(worksheet.Cell(n, row_num).Address, worksheet.Cell(m - 1, row_num).Address) + ")" + "+" + kurator_viezd; worksheet.Cell(m, row_num).Style.Fill.BackgroundColor = XLColor.LightGray;
                    string fact_viezd_maks_Address = worksheet.Cell(m, row_num).Address.ToString();
                    row_num++;
                    worksheet.Cell(m, row_num).FormulaA1 = "=SUM(" + worksheet.Range(worksheet.Cell(n, row_num).Address, worksheet.Cell(m - 1, row_num).Address) + ")"; worksheet.Cell(m, row_num).Style.Fill.BackgroundColor = XLColor.LightGray;
                    string plan_zvonok_Address = worksheet.Cell(m, row_num).Address.ToString();
                    row_num++;
                    worksheet.Cell(m, row_num).FormulaA1 = "=SUM(" + worksheet.Range(worksheet.Cell(n, row_num).Address, worksheet.Cell(m - 1, row_num).Address) + ")"; worksheet.Cell(m, row_num).Style.Fill.BackgroundColor = XLColor.LightGray;
                    string fact_zvonok_Address = worksheet.Cell(m, row_num).Address.ToString();
                    row_num++;
                    worksheet.Cell(m, row_num).FormulaA1 = "=SUM(" + worksheet.Range(worksheet.Cell(n, row_num).Address, worksheet.Cell(m - 1, row_num).Address) + ")"; worksheet.Cell(m, row_num).Style.Fill.BackgroundColor = XLColor.LightGray;
                    string plan_smeta_Address = worksheet.Cell(m, row_num).Address.ToString();
                    row_num++;
                    worksheet.Cell(m, row_num).FormulaA1 = "=SUM(" + worksheet.Range(worksheet.Cell(n, row_num).Address, worksheet.Cell(m - 1, row_num).Address) + ")"; worksheet.Cell(m, row_num).Style.Fill.BackgroundColor = XLColor.LightGray;
                    string fact_smeta_Address = worksheet.Cell(m, row_num).Address.ToString();
                    row_num++; //12

                    //=(ЕСЛИ((K18/J18)>1;1;K18/J18)*50+ЕСЛИ((M18/L18)>1;1;M18/L18)*30+ЕСЛИ((O18/N18)>1;1;O18/N18)*20)/100
                    //K_crm
                    worksheet.Cell(m, row_num).FormulaA1 = "=(IF((" + fact_viezd_maks_Address + "/" + plan_viezd_Address + ")>1,1," + fact_viezd_maks_Address + "/" + plan_viezd_Address + ")*50" + "+IF((" + fact_zvonok_Address + "/" + plan_zvonok_Address + ")>1,1," + fact_zvonok_Address + "/" + plan_zvonok_Address + ")*30" + "+IF((" + fact_smeta_Address + "/" + plan_smeta_Address + ")>1,1," + fact_smeta_Address + "/" + plan_smeta_Address + ")*20" + ")/100";
                    worksheet.Cell(m, row_num).Style.NumberFormat.Format = "#.##";

                    row_num++;
                    //ЗП crm
                    worksheet.Cell(m, row_num).Style.NumberFormat.Format = "#,#";
                    worksheet.Cell(m, row_num).FormulaA1 = "=" + worksheet.Cell(m, 3).Address.ToString() + "*0.15*" + worksheet.Cell(m, row_num - 1).Address.ToString();
                    string zp_crm = worksheet.Cell(m, row_num).Address.ToString();
                    row_num++;
                    // закрывашки план-факт

                    worksheet.Cell(m, row_num).Style.NumberFormat.Format = "#,#";
                    worksheet.Cell(m, row_num).FormulaA1 = "=SUM(" + worksheet.Range(worksheet.Cell(n, row_num).Address, worksheet.Cell(m - 1, row_num).Address) + ")"; worksheet.Cell(m, row_num).Style.Fill.BackgroundColor = XLColor.LightGray;
                    row_num++; //16
                    worksheet.Cell(m, row_num).Style.NumberFormat.Format = "#,#";
                    worksheet.Cell(m, row_num).FormulaA1 = "=SUM(" + worksheet.Range(worksheet.Cell(n, row_num).Address, worksheet.Cell(m - 1, row_num).Address) + ")"; worksheet.Cell(m, row_num).Style.Fill.BackgroundColor = XLColor.LightGray;
                    row_num++; //17

                    // закрывашки расчет ЗП

                    worksheet.Cell(m, row_num).FormulaA1 = "=(P" + m + "/O" + m + ")";
                    worksheet.Cell(m, row_num).Style.NumberFormat.Format = "#.##";

                    worksheet.Range(m, 1, m, worksheet.ColumnsUsed().Count()).Style.Fill.BackgroundColor = XLColor.LightGray;

                    row_num++; // k закрывашки
                    worksheet.Cell(m, row_num).Style.NumberFormat.Format = "#.##";
                    worksheet.Cell(m, row_num).FormulaA1 = "=" + worksheet.Cell(m, 3).Address.ToString() + "*0.5*" + worksheet.Cell(m, row_num - 1).Address.ToString();
                    string zp_zakr = worksheet.Cell(m, row_num).Address.ToString();

                    row_num++; // k непоступившие ОД план

                    worksheet.Cell(m, row_num).Style.NumberFormat.Format = "#,#";
                    worksheet.Cell(m, row_num).FormulaA1 = "=SUM(" + worksheet.Range(worksheet.Cell(n, row_num).Address, worksheet.Cell(m - 1, row_num).Address) + ")"; worksheet.Cell(m, row_num).Style.Fill.BackgroundColor = XLColor.LightGray;
                    string od_plan_Adreess = worksheet.Cell(m, row_num).Address.ToString();
                    row_num++; // k непоступившие ОД факт

                    worksheet.Cell(m, row_num).Style.NumberFormat.Format = "#,#";
                    worksheet.Cell(m, row_num).FormulaA1 = "=SUM(" + worksheet.Range(worksheet.Cell(n, row_num).Address, worksheet.Cell(m - 1, row_num).Address) + ")"; worksheet.Cell(m, row_num).Style.Fill.BackgroundColor = XLColor.LightGray;
                    string od_fact_Adreess = worksheet.Cell(m, row_num).Address.ToString();
                    row_num++; // ЗП за ОД

                    worksheet.Cell(m, row_num).FormulaA1 = "=IF(" + od_plan_Adreess + "<" + od_fact_Adreess + ",0," + worksheet.Cell(m, 3).Address.ToString() + "*0.15)";
                    string zp_od = worksheet.Cell(m, row_num).Address.ToString();

                    row_num++; // дебиторка план

                    worksheet.Cell(m, row_num).Style.NumberFormat.Format = "#,#";
                    worksheet.Cell(m, row_num).FormulaA1 = "=SUM(" + worksheet.Range(worksheet.Cell(n, row_num - 7).Address, worksheet.Cell(m - 1, row_num - 7).Address) + ")"; worksheet.Cell(m, row_num).Style.Fill.BackgroundColor = XLColor.LightGray;
                    string deb_plan_Adreess = worksheet.Cell(m, row_num).Address.ToString();

                    row_num++; // дебиторка факт

                    worksheet.Cell(m, row_num).Style.NumberFormat.Format = "#,#";
                    worksheet.Cell(m, row_num).FormulaA1 = "=SUM(" + worksheet.Range(worksheet.Cell(n, row_num).Address, worksheet.Cell(m - 1, row_num).Address) + ")"; worksheet.Cell(m, row_num).Style.Fill.BackgroundColor = XLColor.LightGray;
                    string deb_fact_Adreess = worksheet.Cell(m, row_num).Address.ToString();
                    row_num++; // K дебиторка

                    worksheet.Cell(m, row_num).FormulaA1 = "=IF(" + deb_plan_Adreess + "<" + deb_fact_Adreess + ",0," + worksheet.Cell(m, 3).Address.ToString() + "*0.05)";
                    string zp_debitor = worksheet.Cell(m, row_num).Address.ToString();

                    row_num++; // кол-во механиков

                    worksheet.Cell(m, row_num).Style.NumberFormat.Format = "#,#";
                    worksheet.Cell(m, row_num).FormulaA1 = "=SUM(" + worksheet.Range(worksheet.Cell(n, row_num).Address, worksheet.Cell(m - 1, row_num).Address) + ")"; worksheet.Cell(m, row_num).Style.Fill.BackgroundColor = XLColor.LightGray;
                    string meh_cont_Adreess = worksheet.Cell(m, row_num).Address.ToString();

                    row_num++; // выработка

                    worksheet.Cell(m, row_num).Style.NumberFormat.Format = "#,#";
                    worksheet.Cell(m, row_num).FormulaA1 = "=SUM(" + worksheet.Range(worksheet.Cell(n, row_num).Address, worksheet.Cell(m - 1, row_num).Address) + ")"; worksheet.Cell(m, row_num).Style.Fill.BackgroundColor = XLColor.LightGray;
                    string meh_vir_Adreess = worksheet.Cell(m, row_num).Address.ToString();
                    row_num++; // K мех

                    worksheet.Cell(m, row_num).FormulaA1 = "=" + meh_vir_Adreess + "/" + meh_cont_Adreess + "/135";
                    string k_meh_Adreess = worksheet.Cell(m, row_num).Address.ToString();
                    worksheet.Cell(m, row_num).Style.NumberFormat.Format = "#.##";

                    row_num++; // ЗП выработка

                    worksheet.Cell(m, row_num).Style.NumberFormat.Format = "#,#";
                    worksheet.Cell(m, row_num).FormulaA1 = "=" + k_meh_Adreess + "*" + worksheet.Cell(m, 3).Address.ToString() + "*0.15";
                    string zp_meh = worksheet.Cell(m, row_num).Address.ToString();
                    row_num++; // внутренние

                    worksheet.Cell(m, row_num).Style.NumberFormat.Format = "#,#";
                    worksheet.Cell(m, row_num).FormulaA1 = "=SUM(" + worksheet.Range(worksheet.Cell(n, row_num).Address, worksheet.Cell(m - 1, row_num).Address) + ")"; worksheet.Cell(m, row_num).Style.Fill.BackgroundColor = XLColor.LightGray;
                    string vnutr_Adreess = worksheet.Cell(m, row_num).Address.ToString();

                    row_num++; // K внутр

                    worksheet.Cell(m, row_num).Style.NumberFormat.Format = "#,#";
                    worksheet.Cell(m, row_num).FormulaA1 = "=" + "IF(AND(" + vnutr_Adreess + ">0," + vnutr_Adreess + "<200000),0.03," + "IF(AND(" + vnutr_Adreess + ">200000," + vnutr_Adreess + "<500000),0.02," + "IF(AND(" + vnutr_Adreess + ">500000," + vnutr_Adreess + "<1000000),0.018," + "IF(AND(" + vnutr_Adreess + ">1000000," + vnutr_Adreess + "<3500000),0.008," + "IF(AND(" + vnutr_Adreess + ">3500000," + vnutr_Adreess + "<10000000000),0.006,0)" + ")" + ")" + ")" + ")";
                    string k_vnutr_Adreess = worksheet.Cell(m, row_num).Address.ToString();

                    row_num++; // ЗП внутр

                    worksheet.Cell(m, row_num).Style.NumberFormat.Format = "#,#";
                    worksheet.Cell(m, row_num).FormulaA1 = "=" + vnutr_Adreess + "*" + k_vnutr_Adreess;
                    string zp_vnutr = worksheet.Cell(m, row_num).Address.ToString();

                    row_num++;  // Итого

                    worksheet.Cell(m, row_num).Style.NumberFormat.Format = "#,#";
                    worksheet.Cell(m, row_num).FormulaA1 = "=" + zp_crm + "+" + zp_zakr + "+" + zp_od + "+" + zp_debitor + "+" + zp_meh + "+" + zp_vnutr;
                    string zp_itog = worksheet.Cell(m, row_num).Address.ToString();

                    row_num = row_num + 3;

                    worksheet.Cell(m, row_num).Style.NumberFormat.Format = "#,#";
                    worksheet.Cell(m, row_num).FormulaA1 = "=" + zp_itog + "*" + worksheet.Cell(m, 33).Address.ToString() + "-" + worksheet.Cell(m, 34).Address.ToString() + "-" + worksheet.Cell(m, 36).Address.ToString();
                    worksheet.Cell(m, 33).Value = 1;



                    n = m + 1;
                    m++;


                }

                if (!c.kurator_fio.Equals(" ") && !c.kurator_fio.Contains("итого"))
                {

                    worksheet.Cell(m, 1).Value = c.kurator_fio.ToString();
                    worksheet.Cell(m, 2).Value = c.kurator_id.ToString();
                    if (c.kurator_filial != null)
                    worksheet.Cell(m, 3).Value = c.kurator_filial.ToString();
                    worksheet.Cell(m, 4).Value = c.Period.ToString();
                    worksheet.Cell(m, 4).Style.NumberFormat.Format = "mmm-yy";
                    worksheet.Cell(m, 5).Value = c.prod_count.ToString();
                    worksheet.Cell(m, 6).Value = c.plan_viezd.ToString();
                    worksheet.Cell(m, 7).Value = c.fact_viezd.ToString();
                    worksheet.Cell(m, 8).Value = c.fact_viezd_max.ToString();
                    worksheet.Cell(m, 9).Value = c.plan_zvonok.ToString();
                    worksheet.Cell(m, 10).Value = c.fact_zvonok.ToString();
                    worksheet.Cell(m, 11).Value = c.plan_smeta.ToString();
                    worksheet.Cell(m, 12).Value = c.fact_smeta.ToString();
                    worksheet.Cell(m, 15).Value = c.plan_zakr.ToString();
                    worksheet.Cell(m, 15).Style.NumberFormat.Format = "#,#";

                    worksheet.Cell(m, 16).Value = c.fact_zakr.ToString();
                    worksheet.Cell(m, 16).Style.NumberFormat.Format = "#,#";


                    worksheet.Cell(m, 20).Value = c.net_od.ToString();
                    worksheet.Cell(m, 20).Style.NumberFormat.Format = "#,#";

                    if (!c.plan_zakr.ToString().Equals(""))
                    {
                        worksheet.Cell(m, 19).Value = Double.Parse(c.plan_zakr.ToString()) * 0.2; // непоступившие ОД норма
                        worksheet.Cell(m, 19).Style.NumberFormat.Format = "#,#";
                    }

                    worksheet.Cell(m, 23).Value = c.debitora.ToString(); // дебиторка
                    worksheet.Cell(m, 23).Style.NumberFormat.Format = "#,#";


                    worksheet.Cell(m, 25).Value = c.mehan_count.ToString();
                    worksheet.Cell(m, 26).Value = c.virabotka.ToString();
                    worksheet.Cell(m, 29).Value = c.vnutr_zakr.ToString();
                    worksheet.Cell(m, 15).Style.NumberFormat.Format = "#,#";
                    m++;
                }
            }

            workbook.SaveAs(textBox1.Text + "\\motiv.xlsx");
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }
       
        private static Excel.Workbook MyBook = null;
        private static Excel.Application MyApp = null;
        private static Excel.Worksheet MySheet = null;

        class Kurator_Debitor
        {
            public String Kurator { get; set; }
            public Double Deb_doc { get; set; }
            public Double Deb_no_doc { get; set; }
            public int Period { get; set; }
        }

        private void button1_Click_1(object sender, EventArgs e)
        {
             string format = "dd.MM.yyyy h:mm:ss";
                         
             CultureInfo provider = CultureInfo.CreateSpecificCulture("ru-RU");
            style = NumberStyles.Number;
            culture = CultureInfo.CreateSpecificCulture("en-GB");

            List<DateTime> last_date = new List<DateTime>();

            

            double deb_current_no_doc, deb_current_doc;

            List<int> list1 = new List<int>();
            bool debitor = false;

           
            List<String> list3 = new List<string>();

            var Deb_current = new List<Kurator_Debitor>();

            var Kurator = new List<Kurator_Filial>();

           
                // Only get files that begin with the letter "c."
                
                List<String> Files = new List<string>();
                Files = Directory.GetFiles(textBox5.Text, "*.xlsx").ToList();


                //Формируем список куратор-филиал

                var wb1 = new XLWorkbook(textBox6.Text + "\\Кураторы-филиалы.xlsx");
                var ws1 = wb1.Worksheet(1);

              

                int excelRow = ws1.RowsUsed().Count();

                int excelColumn = ws1.ColumnsUsed().Count();

                for (int i = 1; i <= excelRow; i++)
                {
                  Kurator.Add(new Kurator_Filial() { RegDir = ws1.Cell(i, 4).Value.ToString(), Kurator = ws1.Cell(i, 1).Value.ToString(), Filial = ws1.Cell(i, 3).Value.ToString() });
                }

                wb1.Dispose();
                
               listBox1.Items.Add("Обнаружено " + Files.Count + " файлов в каталоге.");

                int Files_count = 1;

           

                foreach (string dir in Files)
                {

                 last_date.Add( DateTime.Today.AddMonths(-2));

                    var wb = new XLWorkbook(dir);
                    var ws = wb.Worksheet(1);

                    debitor = false;

                                  
                     excelRow = ws.RowsUsed().Count();

                     excelColumn = ws.ColumnsUsed().Count();

                    for (int i = 1; i <= 3; i++)
                    {
                        for (int j = 1; j <= excelColumn; j++)
                        {
                           
                            if (ws.Cell(i, j).Value.ToString().Equals("Остаток - Просроченная задолженность на конец периода"))
                            {
                                listBox1.Items.Add("Обнаружен отчет дебеторка. Имя " + dir);
                                listBox1.Update();
                                debitor = true;


                            }

                            if (debitor == true && ws.Cell(i, j).Value.ToString().Equals("Филиал продажи"))
                            {

                                ws.Column(j).InsertColumnsBefore(2);
                                ws.Cell(i, j).Value = "Ответственный";
                                ws.Cell(i, j+1).Value = "Региональный директор";

                                for (int k = i + 1; k <= excelRow; k++)
                                {

                                    ws.Cell(k, j).Value = from p in Kurator
                                                 where p.Filial.Equals(ws.Cell(k, j-1).Value)
                                                 select p.Kurator;
                                     ws.Cell(k, j+1).Value = from p in Kurator
                                                 where p.Filial.Equals(ws.Cell(k, j-1).Value)
                                                 select p.RegDir;

                                }
                                
                                
                                j=j+2;
                            }

                            if (debitor == true && ws.Cell(i, j).Value.ToString().Equals("Дата прихода в архив"))
                            {

                                ws.Column(j).InsertColumnsBefore(1);
                                ws.Cell(i, j).Value = "Документы";
                                for (int k=i+1; k <= excelRow; k++)
                                {

                                    if (ws.Cell(k, j + 1).Value.Equals("") == true && ws.Cell(k, j + 2).Value.Equals("") == true)
                                        ws.Cell(k, j).Value = "без документов";
                                    else
                                        ws.Cell(k, j).Value = "с документами";

                                }
                                    j++;
                            }

                             if (debitor == true && ws.Cell(i, j).Value.ToString().Equals("Срок"))
                            {

                             DateTime cur_date =  DateTime.Today.AddMonths(-2);
                                
                              

                             for (int k = i + 1; k <= excelRow; k++)
                                {

                                cur_date = DateTime.ParseExact(ws.Cell(k, j).Value.ToString(), format, provider);

                                if (cur_date.CompareTo(last_date[Files_count-1]) >1 )
                                last_date[Files_count-1] = cur_date;

                                  
                                }

                            
                                
                            }

                        if (debitor == true && ws.Cell(i, j).Value.ToString().Equals("Задачи по ИНН"))
                        {

                            ws.Column(j).Delete();
                         
                            j++;
                        }

                    }
                    }


                    wb.Save();

                    var source = ws.Range(3, 1,excelRow-2, excelColumn).AsTable();
                    var range = source.DataRange;
                    var header = ws.Range(3,1, 3, excelColumn);
                    var dataRange = ws.Range(header.FirstCell(), range.LastCell());
                    var name = "Сводная по дебиторке";
                    var wsPT = wb.Worksheets.Add(name);
                //    pt  = ws.Range(3, 1, excelRow - 2, excelColumn).CreatePivotTable(wsPT.Cell("A1"),name).;


                    var pt = wsPT.PivotTables.AddNew(name, wsPT.Cell(1, 1), source);
                  //  wsPT.Column(2).Style.NumberFormat.Format = "#,##.0";

                    pt.AutofitColumns = true;
           
                
                   
                    pt.RowLabels.Add("Региональный директор");
                    pt.RowLabels.Add("Ответственный");
                    pt.RowLabels.Add("Филиал отгрузки");             
                    pt.ColumnLabels.Add("Документы");
                    pt.Values.Add("Сальдо в рублях");

                //   wb.SaveAs("D:\\YandexDisk\\Скрипты\\Дебитора\\"+ Files_count+".xlsx");

                wb.Save();
                wb.Dispose();


                       
                try
                {
                    MyApp = new Excel.Application();
                    MyApp.Visible = true;


                    MyBook = MyApp.Workbooks.Open(dir);

                    MyBook.Save();
                    MyApp.Quit();
                }  
                finally
                {
                    Marshal.ReleaseComObject(MyBook);
                    Marshal.FinalReleaseComObject(MyApp);
                    MyApp = null;
                    MyBook = null;

               

                GC.Collect();
                }


                wb = new XLWorkbook(dir);
                   ws = wb.Worksheet(2);

                ws.Columns().AdjustToContents();

                int excelRow2 = ws.RowsUsed().Count();

                    for (int i = 1; i <= excelRow2; i++)
                    {
                    if (Double.TryParse(ws.Cell(i, 2).Value.ToString().Replace(",", "."), style, culture, out deb_current_no_doc) == true && Double.TryParse(ws.Cell(i, 3).Value.ToString().Replace(",", "."), style, culture, out deb_current_doc) == true)

                        Deb_current.Add(new Kurator_Debitor() { Kurator = ws.Cell(i, 1).Value.ToString(), Deb_no_doc = deb_current_no_doc, Deb_doc = deb_current_doc, Period = Files_count });
                    }

    

                

                    wb.Dispose();

                    Files_count++;
                }

                var wb_1 = new XLWorkbook(Files[0]);
                var ws_comp = wb_1.Worksheets.Add("Сравнение");

            int []PeriodFirst ={1,2};
                if (last_date[0] > last_date[1])
                    {
                 ws_comp.Cell("B1").Value = "Срез первый ("+last_date[0].ToString()+")";
                ws_comp.Cell("E1").Value = "Срез второй ("+last_date[1].ToString()+")";
                    PeriodFirst[0] = 1;
                 PeriodFirst[1] = 2;
                    }
                else
                    {
                  ws_comp.Cell("B1").Value = "Срез первый ("+last_date[1].ToString()+")";
                ws_comp.Cell("E1").Value = "Срез второй ("+last_date[0].ToString()+")";
                      PeriodFirst[0] = 2;
                 PeriodFirst[1] = 1;
                    }


                ws_comp.Cell(1, 1).Value = "Куратор";
                ws_comp.Cell(2, 1).Value = (from p in Kurator                                
                                      select p.Kurator).Distinct();

                ws_comp.Cell(2, 3).Value = "без документов";
                ws_comp.Cell(2, 2).Value = "с документами";
                ws_comp.Cell(2, 4).Value = "Итого";
          
               
                ws_comp.Cell("H1").Value = "Изменение";
                ws_comp.Range("B1:D1").Row(1).Merge();
                ws_comp.Range("E1:G1").Row(1).Merge();
                ws_comp.Range("H1:J1").Row(1).Merge();

                ws_comp.Cell(2, 6).Value = "без документов";
                ws_comp.Cell(2, 5).Value = "с документами";
                ws_comp.Cell(2, 7).Value = "Итого";

                ws_comp.Cell(2, 9).Value = "без документов";
                ws_comp.Cell(2, 8).Value = "с документами";
                ws_comp.Cell(2, 10).Value = "Итого";

                excelRow = ws_comp.RowsUsed().Count();

            ws_comp.Columns().AdjustToContents();

            

            for (int k = 3; k <= excelRow; k++)
                {

                
                    // заполнение первого периода

                    ws_comp.Cell(k, 2).Value = from p in Deb_current
                                               where (p.Kurator.Equals(ws_comp.Cell(k, 1).Value) && p.Period == PeriodFirst[0])
                                          select p.Deb_no_doc;

                    ws_comp.Cell(k, 3).Value = from p in Deb_current
                                               where (p.Kurator.Equals(ws_comp.Cell(k, 1).Value) && p.Period == PeriodFirst[0])
                                               select p.Deb_doc;
                if (Double.TryParse(ws_comp.Cell(k, 3).Value.ToString().Replace(",", "."), style, culture, out deb_current_no_doc) == true && Double.TryParse(ws_comp.Cell(k, 2).Value.ToString().Replace(",", "."), style, culture, out deb_current_doc) == true)

                    ws_comp.Cell(k, 4).Value = deb_current_no_doc + deb_current_doc;


                    // заполнение второго периода

                    ws_comp.Cell(k, 5).Value = from p in Deb_current
                                               where (p.Kurator.Equals(ws_comp.Cell(k, 1).Value) && p.Period == PeriodFirst[1])
                                               select p.Deb_no_doc;

                    ws_comp.Cell(k, 6).Value = from p in Deb_current
                                               where (p.Kurator.Equals(ws_comp.Cell(k, 1).Value) && p.Period == PeriodFirst[1])
                                               select p.Deb_doc;

                if (Double.TryParse(ws_comp.Cell(k, 5).Value.ToString().Replace(",", "."), style, culture, out deb_current_no_doc) == true && Double.TryParse(ws_comp.Cell(k, 6).Value.ToString().Replace(",", "."), style, culture, out deb_current_doc) == true)

                    ws_comp.Cell(k, 7).Value = deb_current_no_doc + deb_current_doc;

                // заполнение сравнения

               ws_comp.Cell(k, 8).FormulaA1 = "=B"+k+"-"+"E"+k; 

                    ws_comp.Cell(k, 9).FormulaA1 = "=C" + k + "-" + "F" + k;

                ws_comp.Cell(k, 10).FormulaA1 = "=D" + k + "-" + "G" + k;


            }

             
                ws_comp = wb_1.Worksheets.Add("Сравнение 2");

                ws_comp.Cell(1, 1).Value = "Региональный Директор";
                ws_comp.Cell(2, 1).Value = (from p in Kurator                                
                                      select p.RegDir).Distinct();

                ws_comp.Cell(2, 3).Value = "без документов";
                ws_comp.Cell(2, 2).Value = "с документами";
                ws_comp.Cell(2, 4).Value = "Итого";
          
                ws_comp.Cell("B1").Value = "Срез первый";
                ws_comp.Cell("E1").Value = "Срез второй";
                ws_comp.Cell("H1").Value = "Изменение";
                ws_comp.Range("B1:D1").Row(1).Merge();
                ws_comp.Range("E1:G1").Row(1).Merge();
                ws_comp.Range("H1:J1").Row(1).Merge();

                ws_comp.Cell(2, 6).Value = "без документов";
                ws_comp.Cell(2, 5).Value = "с документами";
                ws_comp.Cell(2, 7).Value = "Итого";

                ws_comp.Cell(2, 9).Value = "без документов";
                ws_comp.Cell(2, 8).Value = "с документами";
                ws_comp.Cell(2, 10).Value = "Итого";

                excelRow = ws_comp.RowsUsed().Count();

            ws_comp.Columns().AdjustToContents();

            for (int k = 3; k <= excelRow; k++)
                {

                    // заполнение первого периода

                    ws_comp.Cell(k, 2).Value = from p in Deb_current
                                               where (p.Kurator.Equals(ws_comp.Cell(k, 1).Value) && p.Period == PeriodFirst[0])
                                          select p.Deb_no_doc;

                    ws_comp.Cell(k, 3).Value = from p in Deb_current
                                               where (p.Kurator.Equals(ws_comp.Cell(k, 1).Value) && p.Period == PeriodFirst[0])
                                               select p.Deb_doc;
                if (Double.TryParse(ws_comp.Cell(k, 3).Value.ToString().Replace(",", "."), style, culture, out deb_current_no_doc) == true && Double.TryParse(ws_comp.Cell(k, 2).Value.ToString().Replace(",", "."), style, culture, out deb_current_doc) == true)

                    ws_comp.Cell(k, 4).Value = deb_current_no_doc + deb_current_doc;


                    // заполнение второго периода

                    ws_comp.Cell(k, 5).Value = from p in Deb_current
                                               where (p.Kurator.Equals(ws_comp.Cell(k, 1).Value) && p.Period == PeriodFirst[1])
                                               select p.Deb_no_doc;

                    ws_comp.Cell(k, 6).Value = from p in Deb_current
                                               where (p.Kurator.Equals(ws_comp.Cell(k, 1).Value) && p.Period == PeriodFirst[1])
                                               select p.Deb_doc;

                if (Double.TryParse(ws_comp.Cell(k, 5).Value.ToString().Replace(",", "."), style, culture, out deb_current_no_doc) == true && Double.TryParse(ws_comp.Cell(k, 6).Value.ToString().Replace(",", "."), style, culture, out deb_current_doc) == true)

                    ws_comp.Cell(k, 7).Value = deb_current_no_doc + deb_current_doc;

                // заполнение сравнения

               ws_comp.Cell(k, 8).FormulaA1 = "=B"+k+"-"+"E"+k; 

                    ws_comp.Cell(k, 9).FormulaA1 = "=C" + k + "-" + "F" + k;

                ws_comp.Cell(k, 10).FormulaA1 = "=D" + k + "-" + "G" + k;


            }




                wb_1.Save();

            /* }
           catch (Exception f)
            {
                string s = String.Concat("The process failed: ", f.ToString());
                listBox1.Items.Add(s); listBox1.Update();
            }
            */

        }

        public void UnmergeColumn(IXLWorksheet ws, int Column)
        {

            int excelRow = ws.RowsUsed().Count();

            int StartRow = 1;

            progressBar1.Value = 0;

            progressBar1.Maximum = excelRow;

           
            for (int k = 1; k <= excelRow ; k++)
            {

                progressBar1.Value++;

                if (ws.Range(StartRow, Column, k, Column).IsMerged() == true && ws.Cell(k, Column).Value.Equals("") == false)
                {
                    ws.Range(StartRow, Column, k - 1, Column).Unmerge();
                    StartRow = k;

                }
                if (ws.Range(StartRow, Column, k, Column).IsMerged() == true && k == excelRow)
                {
                    ws.Range(StartRow, Column, excelRow+1, Column).Unmerge();
                   
                }

                if (ws.Range(StartRow, Column, k, Column).IsMerged() == true && ws.Cell(k, Column).Value.Equals("") == false)
                    StartRow = k;

                if (ws.Cell(k, Column).Value.ToString().Equals("") == true && StartRow >1)
                    ws.Cell(k, Column).Value = ws.Cell(k - 1, Column).Value;
            }

        }
        private void button5_Click(object sender, EventArgs e)
        {


            style = NumberStyles.Number;
            culture = CultureInfo.CreateSpecificCulture("en-GB");

            
            List<int> list1 = new List<int>();
            bool debitor = false;

            List<String> Files = new List<string>();
            List<String> list3 = new List<string>();

            var Deb_current = new List<Kurator_Debitor>();

            var Shtat = new List<Shtat>();


            // Only get files that begin with the letter "c."
            Files = Directory.GetFiles(textBox7.Text, "*.xlsx").ToList();


            //Формируем список сотрудников с опорными должностями

            var wb1 = new XLWorkbook(textBox6.Text + "\\Штат.xlsx");
            var ws1 = wb1.Worksheet(1);



            int excelRow = ws1.RowsUsed().Count();

            int excelColumn = ws1.ColumnsUsed().Count();

            for (int i = 1; i <= excelRow; i++)
            {


                Shtat.Add(new Shtat() { FIO = ws1.Cell(i, 5).Value.ToString(), Filial = ws1.Cell(i, 3).Value.ToString(), Tab = ws1.Cell(i, 4).Value.ToString(), OD = ws1.Cell(i, 10).Value.ToString() });
            }

            wb1.Dispose();

            /////////////////////////////////

            listBox1.Items.Add("Обнаружено " + Files.Count + " файлов в каталоге.");

            int Files_count = 1;

            foreach (string dir in Files)
            {
                var wb = new XLWorkbook(dir);
                var ws = wb.Worksheet(1);

                debitor = false;


                excelRow = ws.RowsUsed().Count();

                excelColumn = ws.ColumnsUsed().Count();

              

                for (int i = 1; i <= 3; i++)
                {
                    for (int j = 1; j <= excelColumn; j++)
                    {

                       

                        if (ws.Cell(i, j).Value.ToString().Equals("№ ЗО"))
                        {
                            ws.Column(j).Delete();
                        }
                        if (ws.Cell(i, j).Value.ToString().Equals("Статус ЗО"))
                        {
                            ws.Column(j).Delete();
                        }
                        if (ws.Cell(i, j).Value.ToString().Equals("Тип ЗО"))
                        {
                            ws.Column(j).Delete();
                        }
                        if (ws.Cell(i, j).Value.ToString().Equals("Дата создания WIP"))
                        {
                            string  format;
                            format = "dd.MM.yyyy h:mm:ss";
                            CultureInfo provider = CultureInfo.CreateSpecificCulture("ru-RU");

                            for (int k = 4; k <= excelRow; k++)
                            {
                                
                                DateTime oldDate = DateTime.ParseExact(ws.Cell(k, 8).Value.ToString(), format, provider);
                                DateTime newDate = DateTime.Now;

                             
                                TimeSpan ts = newDate - oldDate;

                           
                                ws.Cell(k, 9).Value = ts.Days;

                            }
                        }

                        if (ws.Cell(i, j).Value.ToString().Equals("Сумма з/ч"))
                        {
                          
                            for (int k = 4; k <= excelRow; k++)
                            {

                             if (ws.Cell(k, j).Value.ToString().Equals("0"))
                                    ws.Row(k).Delete();

                            }
                        }

                        if (ws.Cell(i, j).Value.ToString().Equals("представление: Repairs: Контроль незакрытых WIP заданий (drill6)"))
                        {
                            listBox1.Items.Add("Обнаружен отчет Открытые WIP. Имя " + dir);
                            listBox1.Update();
                            debitor = true;

                          
                            UnmergeColumn(ws, 6);
                            listBox1.Items.Add("Разъединение строки 1 " + dir); listBox1.Update(); 
                            UnmergeColumn(ws, 2);
                            listBox1.Items.Add("Разъединение строки 2 " + dir); listBox1.Update();
                            UnmergeColumn(ws, 7);
                            listBox1.Items.Add("Разъединение строки 3 " + dir); listBox1.Update();

                            listBox1.Items.Add(ws.Cell(3, 1).Value); listBox1.Update();
                            listBox1.Items.Add(ws.Cell(3, 2).Value); listBox1.Update();
                            listBox1.Items.Add(ws.Cell(3, 3).Value); listBox1.Update();
                            listBox1.Items.Add(ws.Cell(3, 4).Value); listBox1.Update();
                            listBox1.Items.Add(ws.Cell(3, 5).Value); listBox1.Update();
                            listBox1.Items.Add(ws.Cell(3, 6).Value); listBox1.Update();
                            listBox1.Items.Add(ws.Cell(3, 7).Value); listBox1.Update();



                            ws.Column(1).InsertColumnsAfter(1);
                            ws.Cell(3, 2).Value = "Ответственный";

                          

                            for (int k = 4; k <= excelRow; k++)
                            {

                                ws.Cell(k, 2).Value = from p in Shtat
                                                      where (p.Filial.Equals(ws.Cell(k, 3).Value)&&(p.OD.Equals("Заместитель директора филиала.")|| p.OD.Equals("Менеджер по продажам.")))
                                                      select p.FIO;
                                
                            }

                            wb.Save();

                        }

                      

                    }
                }

                wb.Save();
                wb.Dispose();
                /*

                var source = ws.Range(3, 1, excelRow - 2, excelColumn).AsTable();
                var range = source.DataRange;
                var header = ws.Range(3, 1, 3, excelColumn);
                var dataRange = ws.Range(header.FirstCell(), range.LastCell());
                var name = "Сводная по дебиторке";
                var wsPT = wb.Worksheets.Add(name);
                //    pt  = ws.Range(3, 1, excelRow - 2, excelColumn).CreatePivotTable(wsPT.Cell("A1"),name).;


                var pt = wsPT.PivotTables.AddNew(name, wsPT.Cell(1, 1), source);
                //  wsPT.Column(2).Style.NumberFormat.Format = "#,##.0";

                pt.AutofitColumns = true;




                pt.RowLabels.Add("Ответственный");
                pt.RowLabels.Add("Филиал отгрузки");
                pt.ColumnLabels.Add("Документы");
                pt.Values.Add("Сальдо в рублях");

                //   wb.SaveAs("D:\\YandexDisk\\Скрипты\\Дебитора\\"+ Files_count+".xlsx");
                */



                /*

                try
                {
                    MyApp = new Excel.Application();
                    MyApp.Visible = true;


                    MyBook = MyApp.Workbooks.Open(dir);

                    MyBook.Save();
                    MyApp.Quit();
                }
                finally
                {
                    Marshal.ReleaseComObject(MyBook);
                    Marshal.FinalReleaseComObject(MyApp);
                    MyApp = null;
                    MyBook = null;



                    GC.Collect();
                }*/

            }

            }

        private void button2_Click_1(object sender, EventArgs e)
        {

        }

        private List<int> GetColumnByName(List<string> file_dir, string list, string name)
        {
            int list_num = 1;
            string list_name = "";
            IXLWorksheet ws;
            
                        try
                        {
                            list_num = Int32.Parse(list);

                        }
                        catch (FormatException)
                        {
                            list_name = list;
                        }

             var wb1 = new XLWorkbook(file_dir[0]+file_dir[1]);
            
                    if (list_name != "")
                        ws = wb1.Worksheet(list_name);
                    else
                        ws = wb1.Worksheet(list_num);
               

           
           
            List<int> col = new List<int>();

            int f_cell = ws.FirstCellUsed().Address.RowNumber;
            int last_column = ws.LastColumnUsed().ColumnNumber();

            for (int i = f_cell; i <= f_cell + 5; i++)
                for (int j = 1; j <= last_column; j++)
                {
                    if (ws.Cell(i, j).Value.ToString().Equals(name))
                    {
                        col.Add(ws.Cell(i, j).Address.RowNumber); col.Add(ws.Cell(i, j).Address.ColumnNumber);

                    }

                }

            return col;

            wb1.Dispose();
            

        }

        private List<int> GetColumnByName(string filename, string list, string name)
        {
            int list_num = 1;
            string list_name = "";
            IXLWorksheet ws;
            
                        try
                        {
                            list_num = Int32.Parse(list);

                        }
                        catch (FormatException)
                        {
                            list_name = list;
                        }

             var wb1 = new XLWorkbook(filename);
            
                    if (list_name != "")
                        ws = wb1.Worksheet(list_name);
                    else
                        ws = wb1.Worksheet(list_num);
               

           
           
            List<int> col = new List<int>();

            int f_cell = ws.FirstCellUsed().Address.RowNumber;
            int last_column = ws.LastColumnUsed().ColumnNumber();

            for (int i = f_cell; i <= f_cell+5; i++)
                for (int j = 1; j <= last_column; j++)
                    {
                    if(ws.Cell(i,j).Value.ToString().Equals(name))
                        {
                        col.Add(ws.Cell(i,j).Address.RowNumber);col.Add(ws.Cell(i,j).Address.ColumnNumber);
                     
                        }
               
                    }

          return col;

            wb1.Dispose();
            

        }

        private string GetFilename(string file_directory, string str)
            {
             List<String> Files = new List<string>();
             Files = Directory.GetFiles(file_directory, "*.xlsx").ToList();
            string ret = "";

            foreach (var file in Files)
                {
               
                if (GetColumnByName(file, "1", str).Count > 0)
                ret=  file.ToString().Split("\\".ToCharArray(), StringSplitOptions.RemoveEmptyEntries).Last();
               
                   

                }


            return ret;
            }


        private string GetList(List<string> file_dir, int list_num)
        {
            var wb1 = new XLWorkbook(file_dir[0]+file_dir[1]);
            var ws = wb1.Worksheet(list_num);

            string name = ws.Name;

            wb1.Dispose();

            return "'" + file_dir[0]+ "[" + file_dir[1] + "]" + name + "'!";
            

        }

        private void ParceXL(List<string> commands)
        {


            Dictionary<string, List<string>> files = new Dictionary<string, List<string>>();
            Dictionary<string, string> ranges = new Dictionary<string, string>();
            
           
            foreach (var command in commands)
            {
                string com = "";
                string[] prm = new string[10];


                com = command.Split("(".ToCharArray(), StringSplitOptions.RemoveEmptyEntries)[0];

                prm = (command.Split("(".ToCharArray(), StringSplitOptions.RemoveEmptyEntries)[1].Split(")".ToCharArray(), StringSplitOptions.RemoveEmptyEntries)[0]).Split(",".ToCharArray(), StringSplitOptions.RemoveEmptyEntries);
                
                if (com.Equals("Файл", StringComparison.InvariantCultureIgnoreCase) && prm[0]!="" && prm[0] != null && prm[1] != "" && prm[1] != null) // добавляем строку до столбца по номеру или до слолбца по названию
                 {
                    List<string> file_dir = new List<string>();

                    if (prm[2].Contains("xlsx"))
                        {                         
                         file_dir.Add(prm[1].TrimEnd().TrimStart());
                         file_dir.Add(prm[2].TrimEnd().TrimStart());
                        }
                   else
                    {
                        file_dir.Add(prm[1].TrimEnd().TrimStart());
                        file_dir.Add(GetFilename(prm[1].TrimEnd().TrimStart(),prm[2].TrimEnd().TrimStart()));

                    }



                    files.Add(prm[0],file_dir);
                 }

                // '[Дебиторка.xlsx]представление_ trx_III_ контрол'!$I:$O

                if (com.Equals("Диапазон", StringComparison.InvariantCultureIgnoreCase) && prm[0] != "" && prm[0] != null && prm[1] != "" && prm[1] != null) // добавляем строку до столбца по номеру или до слолбца по названию
                {
                  
                    List<string> file_dir = new List<string>();

                    if(files.TryGetValue(prm[1].TrimStart().TrimEnd(), out file_dir))
                    {
                       
                       
                        ranges.Add(prm[0], GetList(file_dir, Convert.ToInt32(prm[2])) + prm[3].TrimEnd().TrimStart());
                    }

                }
                  


            }
                foreach (var command in commands)
            {

                string com = "";
                string [] prm = new string[10];

               
                com = command.Split("(".ToCharArray(), StringSplitOptions.RemoveEmptyEntries)[0];
                
                prm = (command.Split("(".ToCharArray(), StringSplitOptions.RemoveEmptyEntries)[1].Split(")".ToCharArray(), StringSplitOptions.RemoveEmptyEntries)[0]).Split(",".ToCharArray(), StringSplitOptions.RemoveEmptyEntries);

                if (com.Equals("ВПР", StringComparison.InvariantCultureIgnoreCase)) // добавляем строку до столбца по номеру или до слолбца по названию
                {

                    string worksheet_name = "";
                    int worksheet_num = 0;
                    int column_num = 1;
                    int row_num = 1;
                    string column_name = "";
                    
                    string filename = prm[0];
                    List<string> file_dir = new List<string>();
                    files.TryGetValue(prm[0], out file_dir);
                    filename = file_dir[0]+file_dir[1];
                    
                    string range = "";
                   
                    
                    var wb1 = new XLWorkbook(filename);
                    IXLWorksheet ws;

                    if (prm[2] != "" && prm[2] != null)
                    {

                        try
                        {
                            column_num = Int32.Parse(prm[2]);

                        }
                        catch (FormatException)
                        {
                            column_name = prm[2];
                        }


                    }
                    if (prm[1] != "" && prm[1] != null)
                    {

                        try
                        {
                            worksheet_num = Int32.Parse(prm[1]);

                        }
                        catch (FormatException)
                        {
                            worksheet_name = prm[1];
                        }


                    }
                    
                    if (worksheet_name != "")
                        ws = wb1.Worksheet(worksheet_name);
                    else
                        ws = wb1.Worksheet(worksheet_num);


                    int strok_vniz = 0;

                   
                    if (prm.Count() == 7)
                    {

                        range = prm[5];
                        ranges.TryGetValue(prm[5], out range);
                        strok_vniz = Convert.ToInt32(prm[4]);

                        //"'C:\\Отчеты\\Константы\\[Кураторы-филиалы.xlsx]Лист1'!$A:$E"
                        
                        for (int i = 1; i < ws.CellsUsed().Count(); i++)
                        {
                            if (ws.Cell(i, column_num).Value.Equals("") == false)
                            {
                                row_num = i;
                                break;
                            }
                        }

                    }      

                    if(prm.Count() == 6)
                       {
                         range = prm[4];
                         ranges.TryGetValue(prm[4], out range);
                         strok_vniz = Convert.ToInt32(prm[3]);
                        
                        row_num = GetColumnByName(file_dir,prm[1],prm[2])[0];
                        column_num =  GetColumnByName(file_dir,prm[1],prm[2])[1];


                        }
                    
                    var wb2 = new XLWorkbook(range.Split("'".ToCharArray(), StringSplitOptions.RemoveEmptyEntries)[0].Split("[]".ToCharArray(), StringSplitOptions.RemoveEmptyEntries)[0] + range.Split("'".ToCharArray(), StringSplitOptions.RemoveEmptyEntries)[0].Split("[]".ToCharArray(), StringSplitOptions.RemoveEmptyEntries)[1]);
                    var ws2 = wb2.Worksheet(range.Split("'".ToCharArray(), StringSplitOptions.RemoveEmptyEntries)[0].Split("[]".ToCharArray(), StringSplitOptions.RemoveEmptyEntries)[2]);

                    IXLRange range_2 = ws2.Range(range.Split("!".ToCharArray(), StringSplitOptions.RemoveEmptyEntries)[1]);

                    int cells_in_table = ws.RowsUsed().Count();

                    //var range_2 = ws2.Range(range.Split("!".ToCharArray(), StringSplitOptions.RemoveEmptyEntries)[1]);

                    if (strok_vniz > 0)
                    {
                        for(int i = row_num+1;i<= strok_vniz+ row_num + 1;i++ )
                        { 
                            for(int j = range_2.FirstRowUsed().RowNumber(); j <= range_2.LastRowUsed().RowNumber();j++ )
                            {
                                if (ws.Cell(i, column_num + 1).Value.Equals(ws2.Cell(j, range_2.FirstColumnUsed().ColumnNumber()).Value))
                                    ws.Cell(i, column_num).Value = ws2.Cell(j, Convert.ToInt32(prm[6])).Value;


                            }
                        }
                    }
                    
                    if (strok_vniz == 0)
                    {
                        for(int i = row_num+1;i <= cells_in_table - row_num + 1;i++ )
                        {
                            for (int j = range_2.FirstRowUsed().RowNumber(); j <= range_2.LastRowUsed().RowNumber(); j++)
                            {
                                if (ws.Cell(i, column_num + 1).Value.Equals(ws2.Cell(j, range_2.FirstColumnUsed().ColumnNumber()).Value))
                                    ws.Cell(i, column_num).Value = ws2.Cell(j, Convert.ToInt32(prm[5])).Value;


                            }
                        }

                    }


                    wb2.Dispose();

                    wb1.Save();

                    wb1.Dispose();
                }


                #region Добавить строку
                if (com.Equals("УдалитьСтолбец", StringComparison.InvariantCultureIgnoreCase)) // добавляем строку до столбца по номеру или до слолбца по названию
                {

                    string worksheet_name = "";
                    int worksheet_num = 0;

                    int column_num = 1;
                    int row_num = 1;
                    string column_name = "";

                    string filename = prm[0];
                    List<string> file_dir = new List<string>();
                    files.TryGetValue(prm[0], out file_dir);
                    filename = file_dir[0] + file_dir[1];

                    if (prm[2] != "" && prm[2] != null)
                    {

                        try
                        {
                            column_num = Int32.Parse(prm[2]);

                        }
                        catch (FormatException)
                        {
                            column_name = prm[2];
                        }


                    }
                    if (prm[1] != "" && prm[1] != null)
                    {

                        try
                        {
                            worksheet_num = Int32.Parse(prm[1]);

                        }
                        catch (FormatException)
                        {
                            worksheet_name = prm[1];
                        }


                    }

                    var wb1 = new XLWorkbook(filename);



                    IXLWorksheet ws;



                    if (worksheet_name != "")
                        ws = wb1.Worksheet(worksheet_name);
                    else
                        ws = wb1.Worksheet(worksheet_num);

                    if (column_name != "")
                        column_num = GetColumnByName(file_dir, prm[1], prm[2])[1];

                    ws.Column(column_num).Delete();
                   
                    wb1.Save();

                    wb1.Dispose();

                }
                if (com.Equals("ДобавитьСтроку",StringComparison.InvariantCultureIgnoreCase)) // добавляем строку до столбца по номеру или до слолбца по названию
                {
                   
                    string worksheet_name = "";
                    int worksheet_num = 0;

                    int column_num = 1;
                    int row_num = 1;
                    string column_name = "";

                    string filename = prm[0];
                    List<string> file_dir = new List<string>();
                    files.TryGetValue(prm[0], out file_dir);
                    filename = file_dir[0]+file_dir[1];
                  
                    if (prm[2]!="" && prm[2] != null)
                    {

                        try
                        {
                            column_num = Int32.Parse(prm[2]);
                           
                        }
                        catch (FormatException)
                        {
                             column_name = prm[2];
                        }
                       

                    }
                    if (prm[1] != "" && prm[1] != null)
                    {

                        try
                        {
                            worksheet_num = Int32.Parse(prm[1]);

                        }
                        catch (FormatException)
                        {
                            worksheet_name = prm[1];
                        }


                    }

                    var wb1 = new XLWorkbook(filename);

                    

                    IXLWorksheet ws;

                    

                    if (worksheet_name != "")                    
                        ws = wb1.Worksheet(worksheet_name);                   
                    else                   
                        ws = wb1.Worksheet(worksheet_num);

                    if (column_name != "")
                       column_num =  GetColumnByName(file_dir,prm[1],prm[2])[1];

                    for(int i = 1; i < ws.CellsUsed().Count(); i++)
                    {
                        if (ws.Cell(i, column_num).Value.Equals("") == false)
                        {
                            row_num = i;
                            break;
                        }
                    }

                   
                    ws.Column(column_num).InsertColumnsBefore(1);
                    ws.Cell(row_num, column_num).Value = prm[3].ToString();

                    string str = ws.Range(1,1,4,4).RangeAddress.ToString();

                    wb1.Save();

                    wb1.Dispose();

                }
                #endregion

                if (com.Equals("ДобавитьДокументы", StringComparison.InvariantCultureIgnoreCase)) // добавляем строку до столбца по номеру или до слолбца по названию
                {

                    string worksheet_name = "";
                    int worksheet_num = 0;

                    int column_num = 1;
                    int row_num = 1;
                    string column_name = "";

                    string filename = prm[0];
                    List<string> file_dir = new List<string>();
                    files.TryGetValue(prm[0], out file_dir);
                    filename = file_dir[0] + file_dir[1];

                    if (prm[2] != "" && prm[2] != null)
                    {

                        try
                        {
                            column_num = Int32.Parse(prm[2]);

                        }
                        catch (FormatException)
                        {
                            column_name = prm[2];
                        }


                    }
                    if (prm[1] != "" && prm[1] != null)
                    {

                        try
                        {
                            worksheet_num = Int32.Parse(prm[1]);

                        }
                        catch (FormatException)
                        {
                            worksheet_name = prm[1];
                        }


                    }

                    var wb1 = new XLWorkbook(filename);



                    IXLWorksheet ws;



                    if (worksheet_name != "")
                        ws = wb1.Worksheet(worksheet_name);
                    else
                        ws = wb1.Worksheet(worksheet_num);

                    if (column_name != "")
                        column_num = GetColumnByName(file_dir, prm[1], prm[2])[1];

                    int total_rows = ws.RowsUsed().Count();

                    for (int i = 1; i < total_rows; i++)
                    {
                        if (ws.Cell(i, column_num).Value.Equals("") == false)
                        {
                            row_num = i;
                            break;
                        }
                    }


                    ws.Column(column_num).InsertColumnsBefore(1);
                    ws.Cell(row_num, column_num).Value = prm[3].ToString();


                    for (int i = row_num + 1; i < total_rows; i++)
                    {
                        if (ws.Cell(i, column_num + 1).Value.Equals("") == true && ws.Cell(i, column_num + 2).Value.Equals("") == true)
                        {

                            ws.Cell(i, column_num).Value = "без документов";

                        }
                        if (ws.Cell(i, column_num + 1).Value.Equals("")==false || ws.Cell(i, column_num + 2).Value.Equals("") == false)
                        
                         {
                            ws.Cell(i, column_num).Value = "с документами";

                        }

                    }


                    string str = ws.Range(1, 1, 4, 4).RangeAddress.ToString();

                    wb1.Save();

                    wb1.Dispose();

                }

                if (com.Equals("ПоКлиенту", StringComparison.InvariantCultureIgnoreCase)) // добавляем строку до столбца по номеру или до слолбца по названию
                {

                    string worksheet_name = "";
                    int worksheet_num = 0;

                    int column_num = 1;
                    int row_num = 1;
                   

                    string filename = prm[0];
                    List<string> file_dir = new List<string>();
                    files.TryGetValue(prm[0], out file_dir);
                    filename = file_dir[0] + file_dir[1];

                   
                    if (prm[1] != "" && prm[1] != null)
                    {

                        try
                        {
                            worksheet_num = Int32.Parse(prm[1]);

                        }
                        catch (FormatException)
                        {
                            worksheet_name = prm[1];
                        }


                    }

                    var wb1 = new XLWorkbook(filename);

                    var wb2 = new XLWorkbook();
                    var ws2 = wb2.AddWorksheet("ДЗ_по_клиенту");

                    IXLWorksheet ws;

                    List<string> clients = new List<string>();

                    if (prm[2].Contains(";"))
                    { 
                    
                    clients.AddRange(prm[2].Split(";".ToCharArray(), StringSplitOptions.RemoveEmptyEntries));
                    }
                    else
                    {
                        clients.Add(prm[2]);
                    }

                    if (worksheet_name != "")
                        ws = wb1.Worksheet(worksheet_name);
                    else
                        ws = wb1.Worksheet(worksheet_num);

                    int total_rows = ws.RowsUsed().Count();
                    int total_columns = ws.ColumnsUsed().Count();
                    int first = ws.FirstRowUsed().RowNumber();

                    int search_column = GetColumnByName(file_dir, prm[1], prm[3])[1];
                    row_num = GetColumnByName(file_dir, prm[1], prm[3])[0];

                    int new_row = 1;

                    ws.Row(row_num).CopyTo(ws2.Row(new_row));
                                      

                    for (int i = 1; i < total_rows; i++)
                    {
                       
                        for (int j = 1; j < total_columns; j++)
                        {
                          
                            foreach(var client in clients)
                            {
                                if (i == 1346 && j == 10)
                                {
                                    ;
                                }
                                if (ws.Cell(i, j).Value.ToString().TrimEnd().TrimStart().ToUpper().Equals(client.ToString().TrimEnd().TrimStart().ToUpper()) == true)
                                {
                                    new_row++;
                                    ws.Row(i).CopyTo(ws2.Row(new_row));
                                }

                            }
                        }
                    }

                    wb2.SaveAs(file_dir[0] + prm[4] + ".xlsx");

                    wb2.Dispose();

                    wb1.Save();

                    wb1.Dispose();

                }

                if (com.Equals("Сводная", StringComparison.InvariantCultureIgnoreCase)) // добавляем строку до столбца по номеру или до слолбца по названию
                {

                    string worksheet_name = "";
                    int worksheet_num = 0;

                    int column_num = 1;
                    int row_num = 1;


                    string filename = prm[0];
                    List<string> file_dir = new List<string>();
                    files.TryGetValue(prm[0], out file_dir);
                    filename = file_dir[0] + file_dir[1];


                    if (prm[1] != "" && prm[1] != null)
                    {

                        try
                        {
                            worksheet_num = Int32.Parse(prm[1]);

                        }
                        catch (FormatException)
                        {
                            worksheet_name = prm[1];
                        }


                    }

                    var wb1 = new XLWorkbook(filename);                              

                    IXLWorksheet ws;
                    

                    List<string> RowLabels = new List<string>(), ColumnLabels = new List<string>();

                    if (prm[3].Contains(";"))
                    {

                        RowLabels.AddRange(prm[3].Split(";".ToCharArray(), StringSplitOptions.RemoveEmptyEntries));
                    }
                    else
                    {
                        RowLabels.Add(prm[3]);
                    }
                    if (prm[4].Contains(";"))
                    {

                        ColumnLabels.AddRange(prm[4].Split(";".ToCharArray(), StringSplitOptions.RemoveEmptyEntries));
                    }
                    else
                    {
                        ColumnLabels.Add(prm[4]);
                    }

                    if (worksheet_name != "")
                        ws = wb1.Worksheet(worksheet_name);
                    else
                        ws = wb1.Worksheet(worksheet_num);

                    int first_table_row = GetColumnByName(file_dir, prm[1], RowLabels[0])[0]; //вычисляем порядковый номер первой строки таблицы

                    var source = ws.Range(first_table_row,1, ws.LastRowUsed().RowNumber(), ws.LastColumnUsed().ColumnNumber()).AsTable(); // формируем диапазон сводной талицы

                    var range = source.DataRange;

                    var header = ws.Range(first_table_row, 1, first_table_row, ws.LastColumnUsed().ColumnNumber()).AsTable();

                    var dataRange = ws.Range(header.FirstCell(), range.LastCell());

                    var name = prm[2];

                    var wsPT = wb1.Worksheets.Add(name);
                   
                    var pt = wsPT.PivotTables.AddNew(name, wsPT.Cell(1, 1), source);
                   
                    pt.AutofitColumns = true;


                    foreach(var rowlabel in RowLabels)
                    pt.RowLabels.Add(rowlabel);

                    foreach (var columlabel in ColumnLabels)
                        pt.ColumnLabels.Add(columlabel);

                    pt.Values.Add(prm[5]);



                    wb1.Save();

                    wb1.Dispose();

                }

            }

            foreach (var range in ranges)
            listBox1.Items.Add(range);

        }
        private void button6_Click(object sender, EventArgs e)
        {
            
                List<string> lst= new List<string>();
                string line;

                // Read the file and display it line by line.  
                System.IO.StreamReader file =
                    new System.IO.StreamReader(listBox2.SelectedItem.ToString(), Encoding.GetEncoding(1251));

                while ((line = file.ReadLine()) != null)
                lst.Add(line);
                          

                file.Close();

                    ParceXL(lst);

                    lst.Clear();

                }
            }
        }
  
