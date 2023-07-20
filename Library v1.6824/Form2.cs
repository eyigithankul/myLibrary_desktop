using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;
namespace my_library
{

    public partial class Form2 : Form
    {
        public string lastname;
        private int formKoor, formKoorX, formKoorY;
        public Form2()
        {
            InitializeComponent();
        }

        private void label1_Click(object sender, EventArgs e)
        {
            System.Environment.Exit(0);
        }

        private void panel1_MouseDown(object sender, MouseEventArgs e)
        {
            formKoor = 1;
            formKoorX = e.X;
            formKoorY = e.Y;
        }

        private void Form2_Load(object sender, EventArgs e)
        {
            KayitGetir();

        }

        private void KayitGetir()
        {
            DataSet dset = new DataSet();
            XmlReader reader = XmlReader.Create(@"kutuphane.xml", new XmlReaderSettings());
            dset.ReadXml(reader);
            dataGridView1.DataSource = dset.Tables[0];
            reader.Close();
            label6.Text = (dataGridView1.Rows.Count - 1).ToString();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            XDocument xdosya = XDocument.Load(@"kutuphane.xml");
            XElement rootelement = xdosya.Root;
            XElement element = new XElement("Kitap");
            XElement adi = new XElement("KitapAdi", textBox1.Text);
            XElement yazar = new XElement("Yazar", textBox2.Text);
            XElement raf = new XElement("Raf", textBox3.Text);
            XElement not = new XElement("Not", textBox4.Text);
            element.Add(adi, yazar, raf, not);
            rootelement.Add(element);
            xdosya.Save(@"kutuphane.xml");
            MessageBox.Show("Kayıt Eklendi. ");
            textBox1.Text = "";
            textBox2.Text = "";
            textBox3.Text = "";
            textBox4.Text = "";
            KayitGetir();


        }

        private void button2_Click(object sender, EventArgs e)
        {

            XDocument xdosya = XDocument.Load(@"kutuphane.xml");
            xdosya.Root.Elements().Where(x => x.Element("KitapAdi").Value == dataGridView1.CurrentRow.Cells[0].Value.ToString()).Remove();

            xdosya.Save(@"kutuphane.xml");
            MessageBox.Show("Kayıt Silindi.");
            KayitGetir();


        }

        private void button3_Click(object sender, EventArgs e)
        {
            XDocument xdosya = XDocument.Load(@"kutuphane.xml");
            XElement element = xdosya.Element("Kitaplar").Elements("Kitap").FirstOrDefault(x => x.Element("KitapAdi").Value == lastname);
            if (element != null)
            {
                element.SetElementValue("KitapAdi", textBox1.Text);
                element.SetElementValue("Yazar", textBox2.Text);
                element.SetElementValue("Raf", textBox3.Text);
                element.SetElementValue("Not", textBox4.Text);
                xdosya.Save(@"kutuphane.xml");
                MessageBox.Show("Kayıt Güncellendi.");
                KayitGetir();
                textBox1.Text = "";
            textBox2.Text = "";
            textBox3.Text = "";
            textBox4.Text = "";
            }


        }

        private void dataGridView1_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            textBox1.Text = dataGridView1.CurrentRow.Cells[0].Value.ToString();
            textBox2.Text = dataGridView1.CurrentRow.Cells[1].Value.ToString();
            textBox3.Text = dataGridView1.CurrentRow.Cells[2].Value.ToString();
            textBox4.Text = dataGridView1.CurrentRow.Cells[3].Value.ToString();
            lastname = textBox1.Text;
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void listBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void button4_Click(object sender, EventArgs e)
        {
            try
            {
                var secili = comboBox1.SelectedItem;
                string kolonadi = "";
                if (secili != null)
                {
                    if (secili == "Kitap Adı ile Arama")
                    {
                        kolonadi = "KitapAdi";
                    }
                    else if (secili == "Yazar Adı ile Arama")
                    {
                        kolonadi = "Yazar";
                    }
                    else
                    {
                        kolonadi = "Raf";
                    }
                ((System.Data.DataTable)dataGridView1.DataSource).DefaultView.RowFilter = string.Format("" + kolonadi + " like '%{0}%'", textBox5.Text.Trim().Replace("'", "''"));
                    label6.Text = (dataGridView1.Rows.Count-1).ToString();
                }
                else
                {
                    MessageBox.Show("Arama türünü seçiniz...");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error Code:6824");
            }
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void button5_Click(object sender, EventArgs e)
        {
            KayitGetir();
            textBox5.Text = "";
        }

        private void label6_Click(object sender, EventArgs e)
        {

        }

        private void button6_Click(object sender, EventArgs e)
        {
            //Microsoft.Office.Interop.Excel._Application app = new Microsoft.Office.Interop.Excel.Application();
            //Microsoft.Office.Interop.Excel._Workbook workbook = app.Workbooks.Add(Type.Missing);
            //Microsoft.Office.Interop.Excel._Worksheet worksheet = null;
            //app.Visible = true;
            //worksheet = workbook.Sheets["Sheet1"];
            //worksheet = workbook.ActiveSheet;
            //for (int i = 1; i < dataGridView1.Columns.Count + 1; i++)
            //{
            //    worksheet.Cells[1, i] = dataGridView1.Columns[i - 1].HeaderText;
            //}
            //for (int i = 0; i < dataGridView1.Rows.Count - 1; i++)
            //{
            //    for (int j = 0; j < dataGridView1.Columns.Count; j++)
            //    {
            //        if (dataGridView1.Rows[i].Cells[j].Value != null)
            //        {
            //            worksheet.Cells[i + 2, j + 1] = dataGridView1.Rows[i].Cells[j].Value.ToString();
            //        }
            //        else
            //        {
            //            worksheet.Cells[i + 2, j + 1] = "";
            //        }
            //    }
            //}

            Excel.Application excel = new Excel.Application();
            excel.Visible = true;
            object Missing = Type.Missing;
            Workbook workbook = excel.Workbooks.Add(Missing);
            Worksheet sheet1 = (Worksheet)workbook.Sheets[1];
            int StartCol = 1;
            int StartRow = 1;
            for (int j = 0; j < dataGridView1.Columns.Count; j++)
            {
                Range myRange = (Range)sheet1.Cells[StartRow, StartCol + j];
                myRange.Value2 = dataGridView1.Columns[j].HeaderText;
            }
            StartRow++;
            for (int i = 0; i < dataGridView1.Rows.Count; i++)
            {
                for (int j = 0; j < dataGridView1.Columns.Count; j++)
                {

                    Range myRange = (Range)sheet1.Cells[StartRow + i, StartCol + j];
                    myRange.Value2 = dataGridView1[j, i].Value == null ? "" : dataGridView1[j, i].Value;
                    myRange.Select();
                }
            }

        }

        private void releaseObject(Excel.Workbook xlWorkBook)
        {
            throw new NotImplementedException();
        }

        private void panel1_MouseMove(object sender, MouseEventArgs e)
        {
            if (formKoor == 1)
            {
                this.SetDesktopLocation(MousePosition.X - formKoorX, MousePosition.Y - formKoorY);
            }
        }

        private void panel1_MouseUp(object sender, MouseEventArgs e)
        {
            formKoor = 0;
        }
    }
}
