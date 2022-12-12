using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net.NetworkInformation;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.ProgressBar;

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
        int czynnik = 0;
        public Excelator()
        {
            InitializeComponent();
        }

        void odczyt()
        {
            
            int liczbaplikow = 0;
            string check = "DATA_ODCZYTU";

            foreach (string fileName in fullFileName)
            {
                listBox1.Items.Add(fileName.Substring(fileName.LastIndexOf(@"\") + 1));
                liczbaplikow++;
                using (var reader = new StreamReader(fileName, Encoding.GetEncoding("ISO-8859-2")))
                {
                    while (!reader.EndOfStream)
                    {

                        var line = reader.ReadLine();
                        var values = line.Split(';');
                        bool result = check.Equals(values[0]);
                        // wpisanie danych do list
                        if (result && czynnik == 0) // sprawdza czy wiersz z nazwami tabeli został już wczytany, jeżeli czynnik = 0 to znaczy, że pętla wykonuje się pierwszy raz 
                        {
                            DATA_ODCZYTU.Add(values[0]);
                            ZUZYCIE_BIEZACE.Add(values[1]);
                            JEDNOSTKA.Add(values[2]);
                            DATA_POPRZEDNIEGO_ODCZYTU.Add(values[3]);
                            ODCZYT_BIEZACY.Add(values[4]);
                            ODCZYT_POPRZEDNI.Add(values[5]);
                            STREFA_EC.Add(values[6]);
                            ADRES_PPE.Add(values[7]);
                            TYP_ODCZYTU.Add(values[8]);
                            NUMER_LICZNIKA.Add(values[9]);
                            PUNKT_POBORU.Add(values[10]);
                            SKLADNIK.Add(values[11]);
                            czynnik = 1;
                        }
                        else if (result && czynnik == 1) // gdy ponownie pojawi się nazwa kolumny a czynnik jest już 1, pominie cały wiersz excela
                        {

                        }
                        else // w każdym innym przypadku po prostu dodaje elementy do listy
                        {

                            DATA_ODCZYTU.Add(values[0]);
                            ZUZYCIE_BIEZACE.Add(values[1]);
                            JEDNOSTKA.Add(values[2]);
                            DATA_POPRZEDNIEGO_ODCZYTU.Add(values[3]);
                            ODCZYT_BIEZACY.Add(values[4]);
                            ODCZYT_POPRZEDNI.Add(values[5]);
                            STREFA_EC.Add(values[6]);
                            ADRES_PPE.Add(values[7]);
                            TYP_ODCZYTU.Add(values[8]);
                            NUMER_LICZNIKA.Add(values[9]);
                            PUNKT_POBORU.Add(values[10]);
                            SKLADNIK.Add(values[11]);
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

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void textBox1_TextChanged_1(object sender, EventArgs e)
        {

        }

        private void button4_Click(object sender, EventArgs e)
        {
            var box = textBox1.Text;
            string fullPath = ".\\"+ box + ".csv";
            MessageBox.Show(fullPath, "nic");
        }

        private void Form1_Load_1(object sender, EventArgs e)
        {

        }
    }
}
