using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace SergeysApp
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            
        }

        private void LoadFromDefaultPath()
        {
            Excel.LoadWorkBook(Environment.CurrentDirectory + @"\ExcelList\MainList.xlsx");
            FillDictionaryHWell();
            FillDictionaryVolDiameter();
            FillComboBoxList();
            ChangeCollectionVolume();
            FillDictionaryVolDiameter(1);

        }

       

        private void Form1_Load(object sender, EventArgs e)
        {
            LoadFromDefaultPath();
            Processing.Text = " ";
        }

        private void Form_Closing(object sender, EventArgs e)
        {
            Excel.CloseExcel();
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            
        }

        private void ExportExcel()
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.DefaultExt = "*.xls;*.xlsx";
            ofd.Filter = "файл Excel (Spisok.xlsx)|*.xlsx";
            ofd.Title = "Выберите файл базы данных";
            if (!(ofd.ShowDialog() == DialogResult.OK))
                return;
            Excel.LoadWorkBook(ofd.FileName);
        }

        private void FillDictionaryVolDiameter()
        {
                Excel.LoadVolDiameter();
        }

        private void FillDictionaryVolDiameter(int a) //перегрузил метод
        {
            Excel.LoadVolDiameter2(Excel.GetSheetNumberByName(comboboxChooseList.Text)); //поменял лист на калькулятор2Лист
            ChangeCollectionVolume(2);
        }
        private void FillDictionaryHWell()
        {
            Excel.LoadHwell(Excel.HWellNumber);
        }


        private void ChangeCollectionDiameter()
        {
            SelectDiameter.Items.Clear();
            //SelectDiameter.Text = "";
            foreach (KeyValuePair<VolumeDiameter, string> valuePair in Excel.VolDiamDictionaryList[Excel.GetSheetNumberByName(comboboxList1.Text)])
            {
                if (valuePair.Key.Volume == SelectVolume.Text)
                {
                    SelectDiameter.Items.Add(valuePair.Key.Diameter);
                }
            }           
        }

        private void ChangeCollectionDiameter(int a)
        {
            SelectDiameter2.Items.Clear();
            SelectDiameter2.Text = "";
            foreach (KeyValuePair<VolumeDiameter, string> valuePair in Excel.VolDiamDictionary)
            {
                if (valuePair.Key.Volume == SelectVolume2.Text)
                {
                    SelectDiameter2.Items.Add(valuePair.Key.Diameter);
                }
            }
        }        

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            TextBoxLenght.Clear();
            TextBoxWeight.Clear();
            SetVolumeDiameter();
            ChangeCollectionDiameter();
        }



        private void SelectDiameter_SelectedIndexChanged(object sender, EventArgs e)
        {
            TextBoxLenght.Clear();
            TextBoxWeight.Clear();
            SetVolumeDiameter();
            SetLenghtWeight();
        }

        private void comboBox1_SelectedIndexChanged_1(object sender, EventArgs e)
        {
            SetPriceHWell();
        }

        private void SetPriceHWell()
        {
            string current;
            if (Excel.HwellDictionary.TryGetValue(SelectH.Text, out current))
            {
                HwellPrice.Text = current;
            }
            Calculate();
        }

        private void SetPriceHWell(int a)
        {
            string current;
            if (Excel.HwellDictionary.TryGetValue(SelectH2.Text, out current))
            {
                HwellPrice2.Text = current;
            }
            Calculate(2);
        }

        private void SetLenghtWeight()
        {
            foreach (KeyValuePair<VolumeDiameter, string> valuePair in Excel.VolDiamDictionaryList[Excel.GetSheetNumberByName(comboboxList1.Text)])
            {
                if (valuePair.Key.Volume == SelectVolume.Text && valuePair.Key.Diameter == SelectDiameter.Text)
                {
                    TextBoxLenght.Text = valuePair.Key.Length;
                    TextBoxWeight.Text = valuePair.Key.Weight;
                }
            }
        }

        private void SetHeightWeight()
        {
            foreach (KeyValuePair<VolumeDiameter, string> valuePair in Excel.VolDiamDictionary)
            {
                if (valuePair.Key.Volume == SelectVolume2.Text && valuePair.Key.Diameter == SelectDiameter2.Text)
                {
                    TextBoxHeight.Text = valuePair.Key.Height;
                    TextBoxWeight2.Text = valuePair.Key.Weight;
                }
            }
        }

        private void SetVolumeDiameter()
        {
            string current;
            VolumeDiameter vCurrent = new VolumeDiameter(SelectVolume.Text, SelectDiameter.Text);
            if (Excel.VolDiamDictionaryList[Excel.GetSheetNumberByName(comboboxList1.Text)].TryGetValue(vCurrent, out current))
            {
                VolDiamSum.Text = current;
            }
            else
            {
                VolDiamSum.Text = "Значение отсутсвует";
            }
            Calculate();
        }

        private void SetVolumeDiameter(int a) //перегрузил метод
        {
            string current;
            VolumeDiameter vCurrent = new VolumeDiameter(SelectVolume2.Text, SelectDiameter2.Text);
            if (Excel.VolDiamDictionary.TryGetValue(vCurrent, out current))
            {
                VolDiamSum2.Text = current;
            }
            else
            {
                VolDiamSum2.Text = "Значение отсутсвует";
            }
            Calculate(2);
        }

        private void Calculate()
        {
            double volDiamSum, hWellPrice, summ;
            try
            {
                volDiamSum = Double.Parse(VolDiamSum.Text);
                hWellPrice = Double.Parse(HwellPrice.Text);
                summ = volDiamSum + hWellPrice;
                LabelTotal.Text = summ.ToString();
            }
            catch
            {

            }
        }

        private void Calculate(int a)
        {
            double volDiamSum2, hWellPrice2, summ2;
            try
            {
                volDiamSum2 = Double.Parse(VolDiamSum2.Text);
                hWellPrice2 = Double.Parse(HwellPrice2.Text);
                summ2 = volDiamSum2 + hWellPrice2;
                LabelTotal2.Text = summ2.ToString();
            }
            catch
            {

            }
        }

        private void VolDiamSum_Click(object sender, EventArgs e)
        {

        }

        private void LabelTotal_Click(object sender, EventArgs e)
        {

        }

        private void LabelTotal_TextChanged(object sender, EventArgs e)
        {

        }

        private void TextBoxLenght_TextChanged(object sender, EventArgs e)
        {

        }

        private void comboBox1_SelectedIndexChanged_2(object sender, EventArgs e)
        {
            CalculateDiscount();
        }

        public void CalculateDiscount()
        {
            if (LabelTotal.Text == "")
            {
                AfterDiscount.Clear();
                TotalAndDiscount.Clear();
            }
            else
            {
                try
                {
                    double _labelTotal = Double.Parse(LabelTotal.Text);
                    double _comboBoxDiscount = Double.Parse(comboBoxDiscount.Text);
                    double _afterDiscount;
                    _afterDiscount = (_labelTotal * _comboBoxDiscount) * 0.01;
                    AfterDiscount.Text = _afterDiscount.ToString();
                    TotalAndDiscount.Text = (_labelTotal - _afterDiscount).ToString();
                }
                catch
                {
                    AfterDiscount.Text = "";
                    TotalAndDiscount.Text = "";
                }
            }
        }
        
        public void CalculateDiscount(int a) //перегружаем
        {
            if (LabelTotal2.Text == "")
            {
                AfterDiscount2.Clear();
                TotalAndDiscount2.Clear();
            }
            else
            {
                try
                {
                    double _labelTotal = Double.Parse(LabelTotal2.Text);
                    double _comboBoxDiscount = Double.Parse(comboBoxDiscount2.Text);
                    double _afterDiscount;
                    _afterDiscount = (_labelTotal * _comboBoxDiscount) * 0.01;
                    AfterDiscount2.Text = _afterDiscount.ToString();
                    TotalAndDiscount2.Text = (_labelTotal - _afterDiscount).ToString();
                }
                catch
                {
                    AfterDiscount2.Text = "";
                    TotalAndDiscount2.Text = "";
                }
            }
        }

        private void comboBox1_SelectedIndexChanged_3(object sender, EventArgs e)
        {
            TextBoxHeight.Clear();
            TextBoxWeight2.Clear();
            SetVolumeDiameter(2);
            ChangeCollectionDiameter(2);
        }

        private void Calc2LoadFile_Click(object sender, EventArgs e)
        {
            RefreshWindows();
            Excel.HwellDictionary.Clear();
            ExportExcel();
            FillDictionaryHWell();
            FillComboBoxList();
        }

        private void comboboxList1_SelectedIndexChanged(object sender, EventArgs e)
        {
            ChangeCollectionVolume();
            usePictureByName();
        }

        private void TabControlIndexChange(object sender, EventArgs e)
        {
            pictureBox1.Image = null;
        }

        public void usePictureByName()
        {

            //string _thePath = @"ExcelList\ЖИР ВЕРТ.png";
            string _thePath = " ";
            if (comboboxList1.Text == "ЕМКОСТИ ГОР")
            {
                pictureBox1.Image = null;
            }
            else
            {
                var replacement = _thePath.Replace(" ", @"ExcelList\" + comboboxList1.Text + ".png");
                Image image = Image.FromFile(replacement);
                pictureBox1.Image = image;

                //pictureBox1.ImageLocation = Environment.CurrentDirectory + @"ExcelList\girGor.png";
                pictureBox1.SizeMode = PictureBoxSizeMode.AutoSize;
            }
           
        }
        private void usePictureByName(int a)
        {
            string _thePath = " ";
            var replacement = _thePath.Replace(" ", @"ExcelList\" + comboboxChooseList.Text + ".png");
            Image image = Image.FromFile(replacement);
            pictureBox1.Image = image;

            //pictureBox1.ImageLocation = Environment.CurrentDirectory + @"ExcelList\girGor.png";
            pictureBox1.SizeMode = PictureBoxSizeMode.AutoSize;
            
        }


        private void Refresh_Click(object sender, EventArgs e)
        {
            RefreshWindows();
        }

        private void comboBoxList2_SelectedIndexChanged(object sender, EventArgs e)
        {
            Excel.HwellDictionary.Clear();
            //FillDictionary2();
        }

        private void RefreshWindows() //перегружаем
        {
            RefreshExceptLists();
            comboboxList1.Text = "";
        }
        private void RefreshWindows(int a)
        {
            RefreshExceptLists(2);
            comboboxChooseList.Text = "";
        }

        private void RefreshExceptLists()
        {
            SelectVolume.Text = "";
            SelectDiameter.Text = "";
            SelectH.Text = "";
            TextBoxLenght.Text = "";
            TextBoxWeight.Text = "";
            LabelTotal.Text = "";
            comboBoxDiscount.Text = "";
            AfterDiscount.Text = "";
            TotalAndDiscount.Text = "";
            VolDiamSum.Text = "";
            HwellPrice.Text = "";
            SelectVolume.Items.Clear();
            SelectDiameter.Items.Clear();

        }

        private void RefreshExceptLists(int a)
        {
            SelectVolume2.Text = "";
            SelectDiameter2.Text = "";
            SelectH2.Text = "";
            TextBoxHeight.Text = "";
            TextBoxWeight2.Text = "";
            LabelTotal2.Text = "";
            comboBoxDiscount2.Text = "";
            AfterDiscount2.Text = "";
            TotalAndDiscount2.Text = "";
            VolDiamSum2.Text = "";
            HwellPrice2.Text = "";
            SelectVolume2.Items.Clear();
            SelectDiameter2.Items.Clear();
            TextBoxH2.Text = "";
            TextBoxH3.Text = "";
            TextBoxH4.Text = "";
            TextBoxD2.Text = "";
        }


        private void FillComboBoxList()
        {
            comboboxList1.Items.Clear();
            comboboxChooseList.Items.Clear();
            for (int i = 1; i <= Excel.SheetsCount; i++)
            {
                if (Excel.GetLastRow(i) > 30 && Excel.OpenWorkSheet(i).Name != "КНС корпуса" && Excel.OpenWorkSheet(i).Name != "ЖИР ВЕРТ")
                {
                    comboboxList1.Items.Add(Excel.OpenWorkSheet(i).Name);
                }
            }
            comboboxChooseList.Items.Add(Excel.OpenWorkSheet(Excel.GetMainCell()).Name);
            comboboxChooseList.SelectedIndex = 0;
            comboboxList1.SelectedIndex = 0;
        }

        
        
        
        
        private void FillComboBoxList(int a) //перегружаю метод, чтобы заполнял список листов исключительно для калькулятора 2
                                            //если хоть одна строка содержит h2 значение, значит берем этот номер листа. Тем самым находим ЖИРВЕРТ
        {
            comboboxChooseList.Items.Clear();
            comboboxList1.Items.Clear();
            comboboxChooseList.Items.Add(Excel.OpenWorkSheet(Excel.GetMainCell()).Name);
            FillComboBoxList();
        }

        private void ChangeCollectionVolume()
        {
            SelectVolume.Items.Clear();
            //SelectVolume.Text = "";           
            foreach (KeyValuePair<VolumeDiameter, string> valuePair in Excel.VolDiamDictionaryList[Excel.GetSheetNumberByName(comboboxList1.Text)])
            {
                if (!SelectVolume.Items.Contains(valuePair.Key.Volume))
                {
                    SelectVolume.Items.Add(valuePair.Key.Volume);
                }
            }
        }
        private void ChangeCollectionVolume(int a) //перегрузил метод
        {
            SelectVolume2.Items.Clear();
            //SelectVolume.Text = "";           
            foreach (KeyValuePair<VolumeDiameter, string> valuePair in Excel.VolDiamDictionary)
            {
                if (!SelectVolume2.Items.Contains(valuePair.Key.Volume))
                {
                    SelectVolume2.Items.Add(valuePair.Key.Volume);
                }
            }
        }

       

        private void label21_Click(object sender, EventArgs e)
        {

        }

        private void comboboxChooseList_SelectedIndexChanged(object sender, EventArgs e)
        {
            ChangeCollectionVolume(2);
            usePictureByName(2);
        }

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            TextBoxHeight.Clear();
            TextBoxWeight2.Clear();
            SetVolumeDiameter(2);
            SetHeightWeight();
            SetOtherValues();
        }

        private void comboBox3_SelectedIndexChanged(object sender, EventArgs e)
        {
            SetPriceHWell(2);
        }

        private void TextBoxLenght2_TextChanged(object sender, EventArgs e)
        {

        }

        private void comboBoxDiscount2_SelectedIndexChanged(object sender, EventArgs e)
        {
            CalculateDiscount(2);
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void SetOtherValues()
        {
            foreach (KeyValuePair<VolumeDiameter, string> valuePair in Excel.VolDiamDictionary)
            {
                if (valuePair.Key.Volume == SelectVolume2.Text && valuePair.Key.Diameter == SelectDiameter2.Text)
                {
                    TextBoxH2.Text = valuePair.Key.H2;
                    TextBoxH3.Text = valuePair.Key.H3;
                    TextBoxH4.Text = valuePair.Key.H4;
                    TextBoxD2.Text = valuePair.Key.D2;
                }
            }
        }

        private void tabPage2_Click(object sender, EventArgs e)
        {

        }

        private void button2_Click(object sender, EventArgs e)
        {
            RefreshWindows(2);
        }

        private void button1_Click_1(object sender, EventArgs e)
        {
            RefreshWindows();
            Excel.HwellDictionary.Clear();
            ExportExcel();
            FillDictionaryHWell();
            FillComboBoxList(2);
        }
    }
}

