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

namespace WindowsFormsApp4
{
    public partial class Excelator : Form
    {
        List<string> DATA_ODCZYTU = new List<string>();
        List<string> ZUZYCIE_BIEZACE = new List<string>();
        List<string> JEDNOSTKA = new List<string>();
        List<string> DATA_POPRZEDNIEGO_ODCZYTU = new List<string>();
        List<string> ODCZYT_BIEZACY = new List<string>();
        List<string> ODCZYT_POPRZEDNI = new List<string>();
        List<string> STREFA_EC = new List<string>();
        List<string> ADRES_PPE = new List<string>();
        List<string> TYP_ODCZYTU = new List<string>();
        List<string> NUMER_LICZNIKA = new List<string>();
        List<string> PUNKT_POBORU = new List<string>();
        List<string> SKLADNIK = new List<string>();
        private List<string> fullFileName;
        private List<string> fullFileName2;
        int czynnik = 0;
        int nowyczynnik = 0;
        public Excelator()
        {
            InitializeComponent();
        }
        static void ConvertExcelToCsv(string excelFilePath, string csvOutputFile, int worksheetNumber = 1)
        {
            if (!File.Exists(excelFilePath)) throw new FileNotFoundException(excelFilePath);
            if (File.Exists(csvOutputFile)) throw new ArgumentException("File exists: " + csvOutputFile);

            // connection string
            var cnnStr = String.Format("Provider=Microsoft.Jet.OLEDB.4.0;Data Source={0};Extended Properties=\"Excel 8.0;IMEX=1;HDR=NO\"", excelFilePath);
            var cnn = new OleDbConnection(cnnStr);

            // get schema, then data
            var dt = new DataTable();
            try
            {
                cnn.Open();
                var schemaTable = cnn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                if (schemaTable.Rows.Count < worksheetNumber) throw new ArgumentException("The worksheet number provided cannot be found in the spreadsheet");
                string worksheet = schemaTable.Rows[worksheetNumber - 1]["table_name"].ToString().Replace("'", "");
                string sql = String.Format("select * from [{0}]", worksheet);
                var da = new OleDbDataAdapter(sql, cnn);
                da.Fill(dt);
            }
            catch (Exception e)
            {
                // ???
                throw e;
            }
            finally
            {
                // free resources
                cnn.Close();
            }

            // write out CSV data
            using (var wtr = new StreamWriter(csvOutputFile, true, Encoding.GetEncoding("UTF-8")))
            {
                foreach (DataRow row in dt.Rows)
                {
                    bool firstLine = true;
                    foreach (DataColumn col in dt.Columns)
                    {
                        if (!firstLine) { wtr.Write(";"); } else { firstLine = false; }
                        var data = row[col.ColumnName].ToString().Replace("\"", "\"\"");
                        wtr.Write(String.Format("\"{0}\"", data));
                    }
                    wtr.WriteLine();
                }
            }
        }
        void odczyt()
        {


            string check = "DATA_ODCZYTU";

            foreach (string fileName in fullFileName)
            {
                listBox1.Items.Add(fileName.Substring(fileName.LastIndexOf(@"\") + 1));

                using (var reader = new StreamReader(fileName, Encoding.GetEncoding("ISO-8859-2")))
                {
                    int z = 0;
                    while (!reader.EndOfStream)
                    {

                        var line = reader.ReadLine();
                        var values = line.Split(';');
                        ;
                        // wpisanie danych do list

                        
                        void przypis()
                        {
                            DATA_ODCZYTU.Add(values[z]);
                            ZUZYCIE_BIEZACE.Add(values[z + 1]);
                            JEDNOSTKA.Add(values[z + 2]);
                            DATA_POPRZEDNIEGO_ODCZYTU.Add(values[z + 3]);
                            ODCZYT_BIEZACY.Add(values[z + 4]);
                            ODCZYT_POPRZEDNI.Add(values[z + 5]);
                            STREFA_EC.Add(values[z + 6]);
                            ADRES_PPE.Add(values[z + 7]);
                            TYP_ODCZYTU.Add(values[z + 8]);
                            NUMER_LICZNIKA.Add(values[z + 9]);
                            PUNKT_POBORU.Add(values[z + 10]);
                            SKLADNIK.Add(values[z + 11]);
                        }

                       
                        bool result;// = check.Equals(values[0]);
                        //   xer = values.Length;
                        int i = 0;
                        do
                        {
                            if (check.Equals(values[i]))
                            {
                                result = true;
                                z = i;
                            }
                            else
                            {
                                result = false;
                            }
                            //result = check.Equals(values[i]);
                            i++;

                        } while (i < values.Length && result == false);

                        bool b1 = false;
                        for (int j = z; j < z+12; j++)
                        {
                            if (string.IsNullOrEmpty(values[j]) == true)
                            {
                                b1 = true;
                            }

                        }

                        if (result == true && czynnik == 0) // sprawdza czy wiersz z nazwami tabeli został już wczytany, jeżeli czynnik = 0 to znaczy, że pętla wykonuje się pierwszy raz 
                        {
                            przypis();
                            czynnik = 1;
                        }
                        else if (result && czynnik == 1) // gdy ponownie pojawi się nazwa kolumny a czynnik jest już 1, pominie cały wiersz excela
                        {

                        }
                        else if (b1 == true)
                        {

                        } 
                        else // w każdym innym przypadku po prostu dodaje elementy do listy
                        {
                            przypis();
                        }
                    }
                }
            }
        }

        void concat()
        {

            var dl_listy = DATA_ODCZYTU.Count;
            var box = textBox1.Text;
            string fullPath = ".\\" + box + ".csv";

            StreamWriter sw = new StreamWriter(fullPath, true, Encoding.GetEncoding("ISO-8859-2"));
            for (int i = 0; i < dl_listy; i++)
            {
                sw.WriteLine(DATA_ODCZYTU[i] + ';' + ZUZYCIE_BIEZACE[i] + ';' + JEDNOSTKA[i] + ';' + DATA_POPRZEDNIEGO_ODCZYTU[i] + ';' + ODCZYT_BIEZACY[i] + ';' + ODCZYT_POPRZEDNI[i] + ';' + STREFA_EC[i] + ';' + ADRES_PPE[i] + ';' + TYP_ODCZYTU[i] + ';' + NUMER_LICZNIKA[i] + ';' + PUNKT_POBORU[i] + ';' + SKLADNIK[i]);

            }
            sw.Close();
        }

        void cleaner()
        {
            DATA_ODCZYTU.Clear();
            ZUZYCIE_BIEZACE.Clear();
            JEDNOSTKA.Clear();
            DATA_POPRZEDNIEGO_ODCZYTU.Clear();
            ODCZYT_BIEZACY.Clear();
            ODCZYT_POPRZEDNI.Clear();
            STREFA_EC.Clear();
            ADRES_PPE.Clear();
            TYP_ODCZYTU.Clear();
            NUMER_LICZNIKA.Clear();
            PUNKT_POBORU.Clear();
            SKLADNIK.Clear();
            czynnik = 0;
        }
        private void button1_Click(object sender, EventArgs e)
        {
            cleaner();
            listBox1.Items.Clear();
            OpenFileDialog OpenFileDialog1 = new OpenFileDialog();
            OpenFileDialog1.Multiselect = true;
            OpenFileDialog1.Filter = "csv Files|*.csv";
            OpenFileDialog1.Title = "Seclect a csv File";
            if (OpenFileDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK)
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



            string errormessage = "Nie udało się połączyć plików";
            string errormessage2 = "Najpierw wybierz pliki do połączenia";
            string title = "Połączenie";
            string message = "Pomyślnie połączono";


            try
            {

                if (listBox1.Items.Count == 0)
                {
                    MessageBox.Show(errormessage2, title);
                }
                else
                {
                    concat();
                    MessageBox.Show(message, title);
                    textBox1.Clear();
                    label1.Visible = false;
                    textBox1.Visible = false;
                    button3.Visible = false;
                    button2.Visible = false;
                }

            }
            catch (Exception)
            {
                MessageBox.Show(errormessage, title);
            }


            listBox1.Items.Clear();
            cleaner();
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
    }
}
