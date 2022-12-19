using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net.NetworkInformation;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.ProgressBar;
using Excel = Microsoft.Office.Interop.Excel;
using Word = Microsoft.Office.Interop.Word;
using ExcelDataReader;


namespace WindowsFormsApp4
{
    public partial class Excelator : Form
    {
       
        private List<string> fullFileName;
        private List<string> fullFileName2;

        int rozmiar = 0;
        string check;
        string[] checker;

        string[,] magazyn = new string[20000, 20000];
        public Excelator()
        {
            InitializeComponent();
        }
        static bool ConvertExcelToCsv(string excelFilePath, string csvOutputFile, int worksheetNumber = 1)
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            using (var stream = new FileStream(excelFilePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                IExcelDataReader reader = null;
                if (excelFilePath.EndsWith(".xls"))
                {
                    reader = ExcelReaderFactory.CreateBinaryReader(stream);
                }
                else if (excelFilePath.EndsWith(".xlsx"))
                {
                    reader = ExcelReaderFactory.CreateOpenXmlReader(stream);
                }

                if (reader == null)
                    return false;

                var ds = reader.AsDataSet(new ExcelDataSetConfiguration()
                {
                    ConfigureDataTable = (tableReader) => new ExcelDataTableConfiguration()
                    {
                        UseHeaderRow = false
                    }
                });

                var csvContent = string.Empty;
                int row_no = 0;
                while (row_no < ds.Tables[0].Rows.Count)
                {
                    var arr = new List<string>();
                    for (int i = 0; i < ds.Tables[0].Columns.Count; i++)
                    {
                        arr.Add(ds.Tables[0].Rows[row_no][i].ToString());
                    }
                    row_no++;
                    csvContent += string.Join(";", arr) + "\n";
                }
                StreamWriter csv = new StreamWriter(csvOutputFile, false, Encoding.GetEncoding("ISO-8859-1"));
                csv.Write(csvContent);
                csv.Close();
                return true;
            }
        
        }


        void odczyt()
        {       
            checker = textBox2.Text.Split('\t');
            check = checker[0];
            rozmiar = checker.Length;
                                    
            int c = 0;

            foreach (string fileName in fullFileName)
            {
                int x = 0;
                listBox1.Items.Add(fileName.Substring(fileName.LastIndexOf(@"\") + 1));
                bool warunek = false;
                int starter = 0;
                using (var reader = new StreamReader(fileName, Encoding.GetEncoding("ISO-8859-1")))
                {
                    while (!reader.EndOfStream)
                    {

                        var line = reader.ReadLine();
                        var values = line.Split(';');
                                        
                        while (warunek == false && x < values.Length)
                        {
                            if (values[x] == check)
                            {
                                warunek = true;
                                starter = x;
                            }
                            else
                            {
                                x++;
                            }
                        }
                        int start = 0;
                        if (warunek)
                        {
                            start = starter;
                            for (int j = 0; j < checker.Length; ++j)
                            {
                                
                                magazyn[c, j] = values[start];
                                start++;
                            }
                        }
                        else
                        {
                            x = 0;
                        }
                        if (warunek)
                        {
                            c++;
                        }
                    }
                        warunek = false;
                }
            }
        }
        void concat()
        {

            
            var box = textBox1.Text;
            string fullPath = ".\\" + box + ".csv";

            StreamWriter sw = new StreamWriter(fullPath, true, Encoding.GetEncoding("ISO-8859-1"));
            for (int w=0;w<checker.Length;w++) {
               sw.Write(checker[w]);
                if (w != checker.Length-1)
                {
                    sw.Write(';');
                }
            }
            for (int i = 0; i < 199; i++)
            {
                if (magazyn[i, 0] == check)
                {

                }
                else
                {
                    for (int j = 0; j < rozmiar + 1; j++)
                    {
                        sw.Write(magazyn[i, j]+';');
                   
                    }
                    sw.Write('\n');
                }

               
            }
            sw.Close();
        }

        void cleaner()
        {
            Array.Clear(magazyn,0,magazyn.Length);           
        }
        private void button1_Click(object sender, EventArgs e)
        {
            cleaner();
            listBox1.Items.Clear();
            OpenFileDialog OpenFileDialog1 = new OpenFileDialog();
            OpenFileDialog1.Multiselect = true;
            OpenFileDialog1.Filter = "csv Files|*.csv";
            OpenFileDialog1.Title = "Seclect a csv File";
            if (OpenFileDialog1.ShowDialog() == DialogResult.OK)
            {
                fullFileName = new List<String>(OpenFileDialog1.FileNames);

                string errormessage = "Niepoprawna zawartość pliku";
                string title = "Wybór";
                try
                {
                odczyt();
                label1.Visible = true;
                textBox1.Visible = true;
                button3.Visible = true;
                button2.Visible = true;
                }
                catch (Exception)
                {
                    MessageBox.Show(errormessage, title);
                    listBox1.Items.Clear();
                    cleaner();

                }

            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void button2_Click(object sender, EventArgs e)
        {
            string message = "Pomyślnie wyczyszczono";
            string errormessage = "Lista jest pusta, nie wymaga czyszczenia";

            string title = "Czyszczenie";
            if (listBox1.Items.Count == 0)
            {
                cleaner();
                MessageBox.Show(errormessage, title);
            }
            else
            {
                listBox1.Items.Clear();
                cleaner();
                label1.Visible = false;
                textBox1.Visible = false;
                button3.Visible = false;
                button2.Visible = false;
                MessageBox.Show(message, title);
            }


        }

        private void button3_Click(object sender, EventArgs e)
        {      
            string title = "Połączenie";
   
            try
            {

                if (listBox1.Items.Count == 0)
                {
                    MessageBox.Show("Najpierw wybierz pliki do połączenia", title);
                }
                else if (string.IsNullOrEmpty(textBox2.Text))
                {
                    MessageBox.Show("Najpierw wklej nazwy kolumn", title);
                }
                else
                {
                    concat();
                    MessageBox.Show("Pomyślnie połączono", title);
                    label1.Visible = false;
                    textBox1.Visible = false;
                    button3.Visible = false;
                    button2.Visible = false;
                    textBox1.Clear();
                    textBox2.Clear();
                    listBox1.Items.Clear();
                    cleaner();
                }

            }
            catch (Exception)
            {
                MessageBox.Show("Coś poszło nie tak", title);
            }


           
        }



        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void textBox1_TextChanged_1(object sender, EventArgs e)
        {

        }


        private void Form1_Load_1(object sender, EventArgs e)
        {

        }

        private void label2_Click(object sender, EventArgs e)
        {

        }

        private void label3_Click(object sender, EventArgs e)
        {

        }

        private void label4_Click(object sender, EventArgs e)
        {

        }

        private void button4_Click(object sender, EventArgs e)
        {
             listBox2.Items.Clear();
                       
            OpenFileDialog OpenFileDialog2 = new OpenFileDialog();
            OpenFileDialog2.Multiselect = true;
            OpenFileDialog2.Filter = "xlsx Files|*.xlsx| xls Files|*.xls|All files (*.*)|*.*";
            OpenFileDialog2.Title = "Seclect a xlsx or xls File";
            if (OpenFileDialog2.ShowDialog() == DialogResult.OK)
            {
                fullFileName2 = new List<String>(OpenFileDialog2.FileNames);
                foreach (string fileName in fullFileName2)
                {

                    listBox2.Items.Add(fileName.Substring(fileName.LastIndexOf(@"\") + 1));

                    string errormessage = "Niepoprawna zawartość pliku";
                    string title = "Wybór";
                    try
                    {

                        button5.Visible = true;
                        button6.Visible = true;
                    }
                    catch (Exception)
                    {
                        MessageBox.Show(errormessage, title);
                        listBox2.Items.Clear();
                        cleaner();

                    }
                }

            }
        }
        private void button5_Click(object sender, EventArgs e)
        {
            string message = "Pomyślnie wyczyszczono";
            string errormessage = "Lista jest pusta, nie wymaga czyszczenia";

            string title = "Czyszczenie";
            if (listBox2.Items.Count == 0)
            {
                cleaner();
                MessageBox.Show(errormessage, title);
            }
            else
            {
                listBox2.Items.Clear();

                button5.Visible = false;
                button6.Visible = false;
                MessageBox.Show(message, title);
            }


        }

        private void button6_Click(object sender, EventArgs e)
        {

            try
            {
                foreach (string fileName in fullFileName2)
                {

                    string kp = fileName.Remove(fileName.LastIndexOf(@"."));
                    var k2 = kp + ".csv";

                    ConvertExcelToCsv(fileName, k2);
                    listBox2.Items.Clear();
                    button5.Visible = false;
                    button6.Visible = false;
                }
                MessageBox.Show("Pomyślnie przekonwertowano");
            }
            catch ( Exception e2)
            {
                throw e2;
            }
            listBox2.Items.Clear();
            button5.Visible = false;
            button6.Visible = false;
            
        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {

        }

        private void label6_Click(object sender, EventArgs e)
        {

        }

        private void label7_Click(object sender, EventArgs e)
        {

        }
    }
}
