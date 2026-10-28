
using CDNBase;
using System;
using ADODB;
using System.Collections.Generic;
using static ImportFS.CSV;
//using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using ExcelDataReader;
using System.Runtime.Remoting.Contexts;

namespace ImportFS
{
    public partial class Form1 : Form
    {
        // public static string sKnt_Kod = "OSTREWKA";
       /* public static string Operator = "ADMIN";
        public static string Haslo = "Platforma2020!";
        public static string Firma = "PPBW_PPBW";
        public static string Baza = "[CDN_PPBW_PPBW].[CDN].";*/

         public static string sKnt_Kod = "!NIEOKREŚLONY!";
         public static string Operator = "ADMIN";
        public static string Haslo = "NE!w0rk12";
        public static string Firma = "Handen 19062024";//"Handen IRIS 112023 jpk testy";//"Handen";//"Handen IRIS 112023 jpk testy";//"Prezentacja_KH";//"TESTOWA_Handen";
                                              // public static string Baza = "[CDN_Testowa_HANDEN].[CDN].";
        protected static CDNBase.Application Application = null;
       protected static ILogin Login = null;
        public Form1()
        {

            InitializeComponent();
        }

        static protected void LogowanieAutomatyczne()
        {
            object[] hPar = new object[] {
                         0,  0,   0,  0,  0,   0,  0,    0,   0,   0,   0,   0,   0,   0,  0,   1, 1 };    // do jakich modułów się logujemy
            /* Kolejno: KP, KH, KHP, ST, FA, MAG, PK, PKXL, CRM, ANL, DET, BIU, SRW, ODB, KB, KBP, HAP*/
            // katalog, gdzie jest zainstalowana Optima (bez ustawienia tej zmiennej nie zadziała, chyba że program odpalimy z katalogu O!)
            System.Environment.CurrentDirectory = @"C:\Program Files (x86)\Comarch ERP Optima";
            // tworzymy nowy obiekt apliakcji
            Application = new CDNBase.Application();
            //Application.HASPServerCommandLine = "OPTIMA1:5000247229";
            // blokujemy
            Application.LockApp(1);//(513, 5000, null, null, null, null);
                                   // logujemy się do podanej Firmy, na danego operatora, do podanych modułów
                                   Login = Application.Login(vUser: Operator, vChk: Haslo, vFirm: Firma, _ModulKP: hPar[0], _ModulKH: hPar[1], _ModulKHP: hPar[2], _ModulST: hPar[3], _ModulFA: hPar[4], _ModulMAG: hPar[5], _ModulPK: hPar[6], _ModulPKXL: hPar[7], _ModulCRM: hPar[8], _ModulANL: hPar[9], _ModulDET: hPar[10], _ModulBIU: hPar[11], _ModulSRW: hPar[12], _ModulOBD: hPar[13], _ModulKB: hPar[14], _ModulKBP: hPar[15], _ModulHAP: hPar[16]);
           // Login = Application.LoginOpe(Operator, Haslo, Firma);
           // MessageBox.Show("Server: " + Application.HASPKeyServer.ToString() + "\nKlucz: " + Application.AppWorkMode + "\nConnection: " + Application.KonfigConnectStr.ToString());
        }

        protected static void Wylogowanie()
        {
            // niszczymy Login
            Login = null;
            // odblokowanie (wylogowanie) O!
            Application.UnlockApp();
            // niszczymy obiekt Aplikacji
            Application = null;
        }


        void WyjscieZApki(string a)
        {
            Cursor.Current = Cursors.Default;
            Wylogowanie();
            if (a == "done")
            {
                MessageBox.Show("Zakończono przetwarzanie pliku.", "Komunikat",
                MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            {
                MessageBox.Show("Generowanie dokumentów zakończyło się niepowodzeniem.", "Komunikat",
                MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }

            return;
        }


        private void Button1_Click(object sender, EventArgs e)
        {


            #region excel5
            Cursor.Current = Cursors.WaitCursor;
            if (tB1.Text == "nie wybrano pliku")
            {
                MessageBox.Show("Nie wybrano pliku arkusza.", "Komunikat",
                MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            string filePath = tB1.Text;
            DataTable dataTable;
            using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
            {
                IExcelDataReader reader;
                reader = ExcelDataReader.ExcelReaderFactory.CreateReader(stream);

                var conf = new ExcelDataSetConfiguration
                {
                    ConfigureDataTable = _ => new ExcelDataTableConfiguration
                    {
                        UseHeaderRow = false
                    }
                };

                var dataSet = reader.AsDataSet(conf);
                dataTable = dataSet.Tables[0];
            }

            //Przypisywanie danych z datasetu do listy przyszłych pozycji.
            //List<ONE032Data> listaPozycjiPz = new List<ONE032Data>();
            if (dataTable.Rows.Count == 0)
            {
                MessageBox.Show("Pusty arkusz", "Komunikat",
                MessageBoxButtons.OK, MessageBoxIcon.Error);//błąd
                return;
            }


            bool czyWszytskieSa = true;

            Indeksy indeksy = WyszukajIndeksow(dataTable, ref czyWszytskieSa);

            if (czyWszytskieSa == false)
                return;

            

            #endregion

            #region dzialanie na optimie
            #region logowanie
            //Logowanie do O! bez wyświetlania okienka logowania
            Cursor.Current = Cursors.WaitCursor;

            

                try
                {
                    LogowanieAutomatyczne();
                }
                catch (COMException eo)
                {
                    if (eo.InnerException != null)
                        MessageBox.Show("Błąd logowania: " + eo.InnerException.Message, "Komunikat",
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                    else
                        MessageBox.Show("Błąd logowania: " + eo.Message, "Komunikat",
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                catch (Exception ei)
                {
                    if (ei.InnerException != null)
                        MessageBox.Show("Błąd logowania: " + ei.InnerException.Message, "Komunikat",
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                    else
                        MessageBox.Show("Błąd logowania: " + ei.Message, "Komunikat",
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            //MessageBox.Show("Zalogowano ", "Komunikat",
            // MessageBoxButtons.OK, MessageBoxIcon.Error);
            #endregion



            // CDNBase.AdoSession Sesja = Login.CreateSession();
            #endregion

            string nr_pliku = "";
            if (filePath.Contains("sap_GAZ 2"))
                nr_pliku = "01";
            else if (filePath.Contains("sap_EE 2"))
                nr_pliku = "02";
            else if (filePath.Contains("sap_PD EE 2"))
                nr_pliku = "03";
            else if (filePath.Contains("sap_PD GAZ 2"))
                nr_pliku = "04";
            else if (filePath.Contains("sap_PD GAZ EUR 2"))
                nr_pliku = "05";
            else if (filePath.Contains("sap_PD PD 2"))
                nr_pliku = "06";
            else if (filePath.Contains("sap_PD PD EUR 2"))
                nr_pliku = "07";
            

              dodawanieFaktury(dataTable,  indeksy, nr_pliku);



                //zapisujemy
                //Sesja.Save();

                WyjscieZApki("done");
            
            //ActiveForm.Dispose();
        }

        
        void WybraniePliku()
        {
            OpenFileDialog openFileDialog1 = new OpenFileDialog()
            {
                InitialDirectory = Directory.GetCurrentDirectory(),
                DefaultExt = "xlsx",
                Filter = "Pliki xlsx (*.xlsx)|*.xlsx|Pliki xls (*.xls)|*.xls|Wszystkie pliki (*.*)|*.*",
            };
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                tB1.Text = openFileDialog1.FileName;
                //MessageBox.Show(path);
            }
            return;
        }

        private void Button2_Click(object sender, EventArgs e)
        {
            WybraniePliku();
        }


        private void dodawanieFaktury(DataTable dataTable,  Indeksy indeksy, string nr_pliku)
        {

            try
            {
                
                    //MessageBox.Show(records[0].Opis);
                    
                    int k = 0;
                    pBar1.Visible = true;
                    // Set Minimum to 1 to represent the first file being copied.
                    pBar1.Minimum = 1;
                    // Set Maximum to the total number of files to copy.
                    pBar1.Maximum = dataTable.Rows.Count;
                    // Set the initial value of the ProgressBar.
                    pBar1.Value = 1;
                    // Set the Step property to a value of 1 to represent each file being copied.
                    pBar1.Step = 1;


                List<string> faktury = new List<string>();
                CDNBase.AdoSession Sesja = Login.CreateSession();


                for (int i = 1; i < dataTable.Rows.Count; i++)
                {
                    pBar1.PerformStep();
                    try
                    {
                        object dummy = Type.Missing;
                        string czyIstniejeQuery = "if exists(SELECT * FROM cdn.VatNag where VaN_Dokument='" + dataTable.Rows[i][indeksy.nr_faktury].ToString().Trim() + "') SELECT 'TAK' else SELECT 'NIE'";
                        var czyIstniejeWynik = Sesja.Connection.Execute(czyIstniejeQuery, out dummy, 0);
                        string czyIstnieje = "TAK";
                        if (!czyIstniejeWynik.EOF)
                           czyIstnieje = czyIstniejeWynik.GetString().Trim();

                       // MessageBox.Show("czyIstnieje: "+czyIstnieje);
                        if (!faktury.Contains(dataTable.Rows[i][indeksy.nr_faktury].ToString()) && dataTable.Rows[i][indeksy.nr_faktury].ToString()!="" && czyIstnieje=="NIE")
                            {
                            
                            string Knt_Kod = "";
                           /* if (dataTable.Rows[i][indeksy.nip].ToString().Length == 11)
                            {

                                string KntQuery = "SELECt Knt_Kod from cdn.Kontrahenci where Knt_Pesel='" + dataTable.Rows[i][indeksy.nip].ToString() + "'";
                                var wynikQuery = Sesja.Connection.Execute(KntQuery, out dummy, 0);
                                if (!wynikQuery.EOF)
                                    Knt_Kod = wynikQuery.GetString();

                            }
                            else*/ if (dataTable.Rows[i][indeksy.nip].ToString().Length == 10)
                            {
                                string KntQuery = "SELECt TOP(1) Knt_Kod from cdn.Kontrahenci where Knt_Nip='" + dataTable.Rows[i][indeksy.nip].ToString() + "'";
                                var wynikQuery = Sesja.Connection.Execute(KntQuery, out dummy, 0);
                                if (!wynikQuery.EOF)
                                    Knt_Kod = wynikQuery.GetString();
                            }
                            if(Knt_Kod=="")
                            {
                                string KntQuery = " SELECT TOP(1) KNT_Kod from cdn.Kontrahenci join cdn.KntAtrybuty on KnA_PodmiotId=Knt_KntId join cdn.DefAtrybuty on DeA_DeAId=KnA_DeAId "
                                                    + " WHERE DeA_Kod = 'NUMER PLATNIKA' and KnA_WartoscTxt = '" + dataTable.Rows[i][indeksy.nr_plat].ToString() + "'";
                                var wynikQuery = Sesja.Connection.Execute(KntQuery, out dummy, 0);
                                if (!wynikQuery.EOF)
                                    Knt_Kod = wynikQuery.GetString();
                            }



                            if (Knt_Kod == "")
                            {
                                Knt_Kod = dodajKontrahenta(dataTable.Rows[i], indeksy);
                            }

                           // MessageBox.Show("Kontrahent kod  : " + Knt_Kod, "Komunikat",
                             //                       MessageBoxButtons.OK, MessageBoxIcon.Information);

                            string numerFaktury = dataTable.Rows[i][indeksy.nr_faktury].ToString();
                           

                            dodajDokument(dataTable.Rows[i], indeksy, Knt_Kod, i, dataTable, nr_pliku);

                            faktury.Add(numerFaktury);




                        }
                    }

                    catch (Exception es)
                    {
                        MessageBox.Show("Błąd w dodawaniu faktury "+ dataTable.Rows[i][indeksy.nr_faktury].ToString() +" : " + Convert.ToString(es), "Komunikat",
                                                MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }

                pBar1.Visible = false;
                Sesja.Save();
                Sesja.Dispose();
                faktury.Clear();

            }
            catch (Exception e)
            {
                /* if(e.ToString().Contains("current financial period"))
                 {
                     MessageBox.Show("Zmień okres obrachunkowy na: " + Convert.ToDateTime(dataTable.Rows[dataTable.Rows.Count - 1][0]).Year, "Komunikat", MessageBoxButtons.OK, MessageBoxIcon.Error);
                 }*/
                MessageBox.Show("Błąd w dodawaniu faktury: " + Convert.ToString(e), "Komunikat",
                                                MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

        }

        private Indeksy WyszukajIndeksow(DataTable dataTable, ref bool czyWszystkieSa)
        {
            Indeksy indeksy = new Indeksy();

            for (int i =0; i<dataTable.Columns.Count; i++)
            {
                switch (dataTable.Rows[0][i].ToString())
                {
                    case "NIP/PESEL(płatnika)":
                        {
                            indeksy.nip = i;
                            break;
                        }
                    case "Numer Płatnika":
                        {
                            indeksy.nr_plat = i;
                            break;
                        }
                    case "Nazwa Płatnika":
                        {
                            indeksy.nazwa_plat = i;
                            break;
                        }
                    case "Kod kraju płatnika":
                        {
                            indeksy.kraj = i;
                            break;
                        }
                    case "Miejscowość (płatnik)":
                        {
                            indeksy.miejscowosc = i;
                            break;
                        }
                    case "Kod pocztowy (płatnik)":
                        {
                            indeksy.kod_poczt = i;
                            break;
                        }
                    case "Ulica (płatnik)":
                        {
                            indeksy.ulica = i;
                            break;
                        }
                    case "Numer domu (płatnik)":
                        {
                            indeksy.nr_dom = i;
                            break;
                        }
                    case "Numer lokalu(płatnik)":
                        {
                            indeksy.nr_lok = i;
                            break;
                        }
                    case "Numer Faktury":
                        {
                            indeksy.nr_faktury = i;
                            break;
                        }
                    case "Numer dokumentu korygowanego":
                        {
                            indeksy.nr_kor = i;
                            break;
                        }
                    case "Id faktury":
                        {
                            indeksy.id_faktury = i;
                            break;
                        }
                    case "Data Wystawienia":
                        {
                            indeksy.data_wystawienia = i;
                            break;
                        }
                    case "Data sprzedaży":
                        {
                            indeksy.data_sprzedazy = i;
                            break;
                        }
                    case "Kod podatkowy":
                        {
                            indeksy.kod_podatkowy = i;
                            break;
                        }
                    case "Jednostka powiązana (1-Tak, 0-Nie)":
                        {
                            indeksy.powiazana = i;
                            break;
                        }
                    case "Termin płatności":
                        {
                            indeksy.termin_plat = i;
                            break;
                        }
                    case "Energia elektryczna":
                        {
                            indeksy.energia = i;
                            break;
                        }
                    case "Akcyza - Pozycja":
                        {
                            indeksy.akcyza = i;
                            break;
                        }
                    case "Stawka Vat":
                        {
                            indeksy.vat = i;
                            break;
                        }
                    case "Stawka VAT":
                        {
                            indeksy.vat = i;
                            break;
                        }
                    case "Opłata handlowa":
                        {
                            indeksy.oplata = i;
                            break;
                        }
                    case "Paliwo gazowe":
                        {
                            indeksy.paliwo = i;
                            break;
                        }
                    case "Opłata handlowa - Gaz":
                        {
                            indeksy.oplata = i;
                            break;
                        }
                    case "Opłata sieciowa zmienna":
                        {
                            indeksy.oplata_zmienna = i;
                            break;
                        }
                    case "Opłata sieciowa stała (kWh/h*h)":
                        {
                            indeksy.oplata_stala = i;
                            break;
                        }
                    case "Opłata sieciowa - stała (pp/m-c)":
                        {
                            indeksy.oplata_stala2 = i;
                            break;
                        }
                    case "Opłata za przekroczenie mocy":
                        {
                            indeksy.przekroczenie_mocy = i;
                            break;
                        }
                    case "Wartość netto pozycji":
                        {
                            indeksy.netto = i;
                            break;
                        }
                    case "Wartość netto pozycji zł":
                        {
                            indeksy.netto = i;
                            break;
                        }
                    case "Nazwa składnika bilingowego":
                        {
                            indeksy.nazwa = i;
                            break;
                        }
                    case "Waluta":
                        {
                            indeksy.waluta = i;
                            break;
                        }
                    case "Kurs przeliczeniowy":
                        {
                            indeksy.kurs = i;
                            break;
                        }
                    case "Data przeliczenia":
                        {
                            indeksy.data_kursu = i;
                            break;
                        }
                    case "Wartość netto pozycji EURO":
                        {
                            indeksy.netto_euro = i;
                            break;
                        }
                    case "Wartość Vat":
                        {
                            indeksy.wartosc_vat = i;
                            break;
                        }
                    case "Wartość Vat faktury":
                        {
                            indeksy.wartosc_vat = i;
                            break;
                        }
                    case "Wartość Vat faktury EURO":
                        {
                            indeksy.wartosc_vat = i;
                            break;
                        }
                    case "Wartość Brutto":
                        {
                            indeksy.brutto = i;
                            break;
                        }
                    case "Wartość brutto faktury":
                        {
                            indeksy.brutto = i;
                            break;
                        }
                    case "Wartość brutto faktury zł":
                        {
                            indeksy.brutto = i;
                            break;
                        }
                    case "Wartość brutto faktury EURO":
                        {
                            indeksy.brutto_eur = i;
                            break;
                        }


                }


            }


            if (indeksy.nip==0)
            {
                MessageBox.Show("Brak kolumny 'NIP/PESEL(płatnika)'", "Komunikat",
                                               MessageBoxButtons.OK, MessageBoxIcon.Error);
                czyWszystkieSa = false;
                return indeksy;
            }


            if (indeksy.nr_plat == 0)
            {
                MessageBox.Show("Brak kolumny 'Numer Płatnika'", "Komunikat",
                                               MessageBoxButtons.OK, MessageBoxIcon.Error);
                czyWszystkieSa = false;
                return indeksy;
            }


            if (indeksy.nr_faktury == 0)
            {
                MessageBox.Show("Brak kolumny 'Numer Faktury'", "Komunikat",
                                               MessageBoxButtons.OK, MessageBoxIcon.Error);
                czyWszystkieSa = false;
                return indeksy;
            }

            if (indeksy.data_wystawienia == 0)
            {
                MessageBox.Show("Brak kolumny 'Data wystawienia'", "Komunikat",
                                               MessageBoxButtons.OK, MessageBoxIcon.Error);
                czyWszystkieSa = false;
                return indeksy;
            }

            if (indeksy.data_sprzedazy == 0)
            {
                MessageBox.Show("Brak kolumny 'Data sprzedaży'", "Komunikat",
                                               MessageBoxButtons.OK, MessageBoxIcon.Error);
                czyWszystkieSa = false;
                return indeksy;
            }


            if (indeksy.termin_plat == 0)
            {
                MessageBox.Show("Brak kolumny 'Termin płatności'", "Komunikat",
                                               MessageBoxButtons.OK, MessageBoxIcon.Error);
                czyWszystkieSa = false;
                return indeksy;
            }

            if(indeksy.energia==0 && indeksy.paliwo==0 && indeksy.nazwa==0)
            {
                MessageBox.Show("Zły format pliku", "Komunikat",
                                               MessageBoxButtons.OK, MessageBoxIcon.Error);
                czyWszystkieSa = false;
                return indeksy;
            }


            return indeksy;
        }

        private string dodajKontrahenta(System.Data.DataRow dataTable, Indeksy indeksy)
        {
            try
            {
                CDNBase.AdoSession Sesja = Login.CreateSession();


                CDNHeal.Kontrahenci Kontrahenci = (CDNHeal.Kontrahenci)Sesja.CreateObject("CDN.Kontrahenci", null);
                CDNHeal.IKontrahent Kontrahent = (CDNHeal.IKontrahent)Kontrahenci.AddNew(null);

                CDNHeal.IAdres Adres = Kontrahent.Adres;
                CDNHeal.INumerNIP NumerNIP = Kontrahent.NumerNIP;

                object dummy = Type.Missing;

                if (dataTable[indeksy.kraj].ToString() != "" && indeksy.kraj != 0)
                {
                    string krajQuery = "SELECT Kra_Nazwa FROM CDN_KNF_Konfiguracja.cdn.Kraje where Kra_Kod = '" + dataTable[indeksy.kraj].ToString() + "'";
                    var wynikQuery = Sesja.Connection.Execute(krajQuery, out dummy, 0);
                    if (!wynikQuery.EOF)
                        Adres.Kraj = wynikQuery.GetString();
                    string kodKraj = dataTable[indeksy.kraj].ToString();
                  /*  if (kodKraj == "AT" || kodKraj == "BE" || kodKraj == "BG" || kodKraj == "HR" || kodKraj == "CY" || kodKraj == "CZ" || kodKraj == "DK" || kodKraj == "EE" || kodKraj == "FI" || kodKraj == "FR" || kodKraj == "GR" || kodKraj == "ES"
                        || kodKraj == "IE" || kodKraj == "LU" || kodKraj == "LT" || kodKraj == "LV" || kodKraj == "MT" || kodKraj == "NL" || kodKraj == "DE" || kodKraj == "PT" || kodKraj == "RO" || kodKraj == "SK" || kodKraj == "SI" || kodKraj == "SE"
                        || kodKraj == "HU" || kodKraj == "IT")
                        Kontrahent.Export = 3;
                    else*/
                        Kontrahent.Export = 7;
                }
                else
                
                    Adres.Kraj = "Polska";

                Adres.KodPocztowy = (dataTable[indeksy.kod_poczt].ToString() != "" && indeksy.kod_poczt != 0) ? dataTable[indeksy.kod_poczt].ToString() : "";
                Adres.NrDomu = (dataTable[indeksy.nr_dom].ToString() != "" && indeksy.nr_dom != 0) ? dataTable[indeksy.nr_dom].ToString() : "";
                Adres.Miasto = (dataTable[indeksy.miejscowosc].ToString() != "" && indeksy.miejscowosc != 0) ? dataTable[indeksy.miejscowosc].ToString() : "";
                Adres.Ulica = (dataTable[indeksy.ulica].ToString() != "" && indeksy.ulica != 0) ? dataTable[indeksy.ulica].ToString() : "";
                Adres.NrLokalu = (dataTable[indeksy.nr_lok].ToString() != "" && indeksy.nr_lok != 0) ? dataTable[indeksy.nr_lok].ToString() : "";
                Kontrahent.KrajISO = (dataTable[indeksy.kraj].ToString() != "" && indeksy.kraj != 0) ? dataTable[indeksy.kraj].ToString() : "PL";

                if (dataTable[indeksy.nip].ToString().Length == 10)
                {
                    NumerNIP.UstawNIP((dataTable[indeksy.kraj].ToString() != "" && indeksy.kraj != 0) ? dataTable[indeksy.kraj].ToString() : "PL", dataTable[indeksy.nip].ToString(), 1);
                }
                if (dataTable[indeksy.nip].ToString().Length == 11)
                {
                    try
                    {
                        Kontrahent.Pesel = dataTable[indeksy.nip].ToString();
                    }
                    catch
                    {
                        string nip = dataTable[indeksy.nip].ToString().Trim();
                        if (dataTable[indeksy.kraj].ToString().Length != 0 && indeksy.kraj != 0)
                            nip = nip.Substring(dataTable[indeksy.kraj].ToString().Length);
                        NumerNIP.UstawNIP((dataTable[indeksy.kraj].ToString() != "" && indeksy.kraj != 0) ? dataTable[indeksy.kraj].ToString() : "PL", nip, 1);
                    }
                }
                else
                {
                    string nip = dataTable[indeksy.nip].ToString().Trim();
                    if (dataTable[indeksy.kraj].ToString().Length != 0 && indeksy.kraj != 0)
                        nip = nip.Substring(dataTable[indeksy.kraj].ToString().Length);
                    NumerNIP.UstawNIP((dataTable[indeksy.kraj].ToString() != "" && indeksy.kraj != 0) ? dataTable[indeksy.kraj].ToString() : "PL", nip, 1);
                }

                Kontrahent.Nazwa1 = dataTable[indeksy.nazwa_plat].ToString();



                var formyPlatnosci = (CDNBase.ICollection)(Sesja.CreateObject("CDN.FormyPlatnosci", null));
                var fPl = (OP_KASBOLib.FormaPlatnosci)formyPlatnosci[2];
                Kontrahent.FormaPlatnosci = fPl;




                Sesja.Save();


                if (dataTable[indeksy.powiazana].ToString().Trim() == "1")
                {
                    CDNBase.AdoSession sesja2 = Login.CreateSession();
                    object dummy2 = Type.Missing;
                    //var rejestrVat = (CDNRVAT.VAT)Sesja.CreateObject("CDN.RejestryVAT", $"VaN_VaNID = {dokId}");
                    var kodJPKWynik = sesja2.Connection.Execute("SELECT KV7_Kv7ID from CDN.KodyJPKV7 where KV7_Kod='TP'", out dummy2, 0);
                    int kodJPK = 0;
                    if (!kodJPKWynik.EOF)
                    {
                        //MessageBox.Show(kodJPKWynik.GetString());
                        kodJPK = Convert.ToInt32(kodJPKWynik.GetString().Trim());
                    }
                    string insert = "INSERT INTO [CDN].[DokKodyJPKV7] ([DKV7_ParentTyp] ,[DKV7_ParentID] ,[DKV7_KV7ID]  ,[DKV7_Status]) VALUES (0," + Kontrahent.ID + "," + kodJPK + ",0)";
                    // MessageBox.Show(insert);
                    sesja2.Connection.Execute(insert, out dummy2, 0);



                    sesja2.Save();

                }


                var rAtr = Sesja.CreateObject("CDN.DefAtrybuty").Item("DeA_Kod='NUMER PLATNIKA'");

                var rAtrybut = Kontrahent.Atrybuty.AddNew();
                rAtrybut.DefAtrybut = rAtr;
                rAtrybut.Wartosc = dataTable[indeksy.nr_plat].ToString();
                Sesja.Save();
                Sesja.Dispose();


                return Kontrahent.Akronim;
            }
            catch(Exception an)
            {
                MessageBox.Show("Błąd funkcji dodajKontrahenta dla faktury " + dataTable[indeksy.nr_faktury].ToString() + " : " + an.Message);
                return "";
            }
        
        }
        
        private void dodajDokument(DataRow dataRow, Indeksy indeksy, string knt_Kod, int i, DataTable dataTable, string nr_pliku)
        {
            try
            {
                CDNBase.AdoSession Sesja = Login.CreateSession();
                var formyPlatnosci = (CDNBase.ICollection)(Sesja.CreateObject("CDN.FormyPlatnosci", null));
                //var waluty = (CDNBase.ICollection)(Sesja.CreateObject("CDN.Waluty", null));
                var rejestryVAT = (CDNBase.ICollection)(Sesja.CreateObject("CDN.RejestryVAT", null));
                var kontrahenci = (CDNBase.ICollection)(Sesja.CreateObject("CDN.Kontrahenci", null));
                // pobieranie kontrahenta, formy platnosci  i waluty
                var kontrahent = (CDNHeal.IKontrahent)kontrahenci["Knt_Kod='" + knt_Kod.Trim() + "'"];
                // w konfiguracji jest tylko jedna waluta (PLN)
                //			CDNHeal.Waluta Waluta			= (CDNHeal.Waluta)Waluty[ "WNa_Symbol='EUR'" ];
                var fPl = (OP_KASBOLib.FormaPlatnosci)formyPlatnosci[2];
                // utworzenie nowego obiektu rejestru VAT
                var rejestrVAT = (CDNRVAT.VAT)rejestryVAT.AddNew(null);
                //ustawianie parametrów rejestru
                rejestrVAT.Typ = 2;//	1 - zakupu; 2 - sprzedaży
                                   //Console.WriteLine("Typ ustawiony");
                rejestrVAT.Podmiot = (CDNHeal.IPodmiot)kontrahent;


                rejestrVAT.Rejestr = "SPRZEDAŻ IRIS";    // nazwa rejestru
                                                         //Console.WriteLine("Rejestr ustawiony");

                rejestrVAT.Dokument = dataRow[indeksy.nr_faktury].ToString();
                //Console.WriteLine("Dokument ustawiony");

                //rejestrVAT.IdentKsieg = "2007/05/28-oop";
                //Console.WriteLine("IdentKsieg ustawiony");


                rejestrVAT.DataWys = Convert.ToDateTime(dataRow[indeksy.data_wystawienia].ToString());
                rejestrVAT.DataZap = Convert.ToDateTime(dataRow[indeksy.data_wystawienia].ToString());
                rejestrVAT.DataOpe = Convert.ToDateTime(dataRow[indeksy.data_sprzedazy].ToString());
                //Console.WriteLine("DataZap ustawiona");

                rejestrVAT.FormaPlatnosci = fPl;


                //DATA DEKLARACJI
                //Console.WriteLine("Forma platnosci ustawiona");
                //  if (dataRow[indeksy.kod_podatkowy].ToString().Trim()=="A9" || dataRow[indeksy.kod_podatkowy].ToString().Trim() == "A3" || dataRow[indeksy.kod_podatkowy].ToString().Trim() == "EE" )
                // {
                rejestrVAT.RozliczacVAT = 1;

                if(kontrahent.KrajISO!="PL")
                {
                    rejestrVAT.Export = 7;
                }

                rejestrVAT.RozliczacVAT27 = 0;
                /*   }
                   else
                   {
                       DateTime data = Convert.ToDateTime(dataRow[indeksy.data_wystawienia].ToString());
                       DateTime date4;
                       if (data.Month != 12)
                       {

                           date4 = Convert.ToDateTime(data.Year + "-" + (data.Month+1) + "-01") ;
                       }
                       else
                       {

                           date4 = Convert.ToDateTime((data.Year+1) + "-01-01");
                       }
                       rejestrVAT.DeklMiesiacVAT = date4.Month;
                       rejestrVAT.DeklRokVAT = date4.Year;
                   }*/



                if (indeksy.waluta != 0)
                {

                    // var typyKursowOptima = (CDNBase.ICollection)(Sesja.CreateObject("CDN.WalKursy", null));
                    var typyKursowOptima = (CDNHeal.TypyKursowWalut)(Sesja.CreateObject("CDN.TypyKursowWalut", null));
                    CDNHeal.TypKursuWaluty typKursu2 = (CDNHeal.TypKursuWaluty)typyKursowOptima["WKu_WKuID = 1"];

                    rejestrVAT.TypKursuWalutyDoVAT = typKursu2;
                    rejestrVAT.TypKursuWaluty = typKursu2;
                    var Waluty = (CDNBase.ICollection)(Sesja.CreateObject("CDN.Waluty", null));
                    CDNHeal.Waluta Waluta = (CDNHeal.Waluta)Waluty["WNa_Symbol='" + dataRow[indeksy.waluta].ToString() + "'"];
                    rejestrVAT.Waluta = Waluta;
                    rejestrVAT.DataKur = Convert.ToDateTime(dataRow[indeksy.data_kursu].ToString());
                    rejestrVAT.WalutaDoVAT = Waluta;
                    rejestrVAT.DataKurDoVAT = Convert.ToDateTime(dataRow[indeksy.data_kursu].ToString());
                    //sprobować przypisać kurs
                    rejestrVAT.KursL = Convert.ToDouble(dataRow[indeksy.kurs].ToString());
                    rejestrVAT.KursLDoVAT = Convert.ToDouble(dataRow[indeksy.kurs].ToString());
                   


                }

                rejestrVAT.Termin = Convert.ToDateTime(dataRow[indeksy.termin_plat].ToString());

                rejestrVAT.DeklMiesiacVAT = Convert.ToDateTime(dataRow[indeksy.data_wystawienia].ToString()).Month;
                rejestrVAT.DeklRokVAT = Convert.ToDateTime(dataRow[indeksy.data_wystawienia].ToString()).Year;

                // if (kontrahent.Export == 3 || kontrahent.Export == 1)
                //   rejestrVAT.RozliczacVAT = 0;

                if (indeksy.nr_kor != 0 && dataRow[indeksy.nr_kor].ToString() != "")
                {
                    rejestrVAT.Korekta = 1;
                    rejestrVAT.KorektaDo = dataRow[indeksy.nr_kor].ToString();
                }

                //Console.WriteLine("Podmiot ustawiony");
                // kategoria ustawia się sama, gdy ustawiany jest kontrahent

                // waluty nie ustawiam, bo w konf. jest na razie tylko jedna waluta (PLN)
                //			RejestrVAT.WalutaDoVAT = Waluta;	
                //			Console.WriteLine( "Waluta ustawiona " + RejestrVAT.Waluta.Symbol );

                // dodanie elementów rejestru VAT
                for (int k = i; k < dataTable.Rows.Count; k++)
                {
                    if (dataTable.Rows[k][indeksy.nr_faktury].ToString() == dataRow[indeksy.nr_faktury].ToString())
                    {
                        if (indeksy.energia != 0)
                            DodajElementyDoRejestru_typ1(rejestrVAT, dataTable.Rows[k], indeksy, Sesja);
                        else if (indeksy.paliwo != 0)
                            DodajElementyDoRejestru_typ2(rejestrVAT, dataTable.Rows[k], indeksy, Sesja);
                        else if (indeksy.nazwa != 0)
                            DodajElementyDoRejestru_typ3(rejestrVAT, dataTable.Rows[k], indeksy, Sesja);
                    }
                }


                // zapisanie zmian
                Sesja.Save();
                var rAtr = Sesja.CreateObject("CDN.DefAtrybuty").Item("DeA_Kod='NR_PLIKU'");
                var rAtrybut = rejestrVAT.Atrybuty.AddNew();
                rAtrybut.DefAtrybut = rAtr;
                rAtrybut.Wartosc = nr_pliku;
                Sesja.Save();

                if (indeksy.wartosc_vat != 0)
                {
                    object dummy2 = Type.Missing;
                    //var rejestrVat = (CDNRVAT.VAT)Sesja.CreateObject("CDN.RejestryVAT", $"VaN_VaNID = {dokId}");
                    var vatWynik = Sesja.Connection.Execute("SELECT SUM(VAT_VAT) FROM cdn.VatTab WHERE VAT_VaNID = " + rejestrVAT.ID, out dummy2, 0);
                    double vat = 0;
                    if (!vatWynik.EOF)
                    {
                        //MessageBox.Show("Suma vat w pliku: "+vatWynik.GetString());
                        vat = Math.Round(Convert.ToDouble(dataRow[indeksy.wartosc_vat].ToString()) - Convert.ToDouble(vatWynik.GetString()), 2, MidpointRounding.AwayFromZero);
                        //MessageBox.Show("Różnica vat: " + vat.ToString());

                        if (vat != 0)
                        {
                            //MessageBox.Show("Różnica vat: " + vat.ToString());
                            vatWynik = Sesja.Connection.Execute("SELECT VAT_VAT FROM cdn.VATTab WHERE Vat_VANID = " + rejestrVAT.ID + " and VAT_Lp = 1", out dummy2, 0);
                            if (!vatWynik.EOF)
                            {

                                vat = Convert.ToDouble(vatWynik.GetString()) + vat;
                                //MessageBox.Show("Przypisywany vat: " + vat.ToString()+" w fakturze "+dataRow[indeksy.nr_faktury].ToString());
                                Sesja.Connection.Execute("UPDATE [CDN].[VatTab] SET [VaT_VAT] = " + vat.ToString().Replace(",", ".") + " ,[VaT_VATDoVAT] =" + vat.ToString().Replace(",", ".") + " ,[VaT_VATWal] = " + vat.ToString().Replace(",", ".") + " ,[VaT_VATOSS] = " + vat.ToString().Replace(",", ".") + " WHERE Vat_VANID = " + rejestrVAT.ID + " and VAT_Lp = 1", out dummy2, 0);
                                if (indeksy.waluta == 0 || dataRow[indeksy.waluta].ToString() == "PLN")
                                {

                                    Sesja.Connection.Execute("UPDATE [CDN].[VatNag] SET [VaN_RazemVAT] = " + dataRow[indeksy.wartosc_vat].ToString().Replace(",", ".") + " ,[VaN_RazemVATDoVAT] =" + dataRow[indeksy.wartosc_vat].ToString().Replace(",", ".") +
                                        ",[VaN_RazemBrutto] =" + dataRow[indeksy.brutto].ToString().Replace(",", ".") + ",[VaN_RazemBruttoDoVAT] = " + dataRow[indeksy.brutto].ToString().Replace(",", ".") + " ,[VaN_RazemBruttoWal] = " + dataRow[indeksy.brutto].ToString().Replace(",", ".") +
                                        ",[VaN_RazemBruttoOSS] = " + dataRow[indeksy.brutto].ToString().Replace(",", ".") + " ,[VaN_RazemVATOSS] = " + dataRow[indeksy.wartosc_vat].ToString().Replace(",", ".") + " WHERE VaN_VANID = " + rejestrVAT.ID, out dummy2, 0);


                                }
                                else
                                {
                                    Sesja.Connection.Execute("UPDATE [CDN].[VatNag] SET [VaN_RazemVAT] = " + dataRow[indeksy.wartosc_vat].ToString().Replace(",", ".") + " ,[VaN_RazemVATDoVAT] =" + dataRow[indeksy.wartosc_vat].ToString().Replace(",", ".") +
                                        ",[VaN_RazemBrutto] =" + dataRow[indeksy.brutto].ToString().Replace(",", ".") + ",[VaN_RazemBruttoDoVAT] = " + dataRow[indeksy.brutto].ToString().Replace(",", ".") + " ,[VaN_RazemBruttoWal] = " + dataRow[indeksy.brutto_eur].ToString().Replace(",", ".") +
                                        " ,[VaN_RazemBruttoOSS] = " + dataRow[indeksy.brutto_eur].ToString().Replace(",", ".") + " ,[VaN_RazemVATOSS] = " + dataRow[indeksy.wartosc_vat].ToString().Replace(",", ".") + " WHERE VaN_VANID = " + rejestrVAT.ID, out dummy2, 0);
                                }

                            }

                        }
                    }
                }
                try
                {
                    Sesja.Save();
                    Sesja.Resync();
                }
                catch (Exception on)
                {

                }

                try
                {
                    if (indeksy.waluta == 0 || dataRow[indeksy.waluta].ToString() == "PLN")
                    {
                        object dummy = Type.Missing;

                        Sesja.Connection.Execute("update cdn.BnkZdarzenia set Bzd_Rozliczono=0, Bzd_Rozliczono2=0, Bzd_KwotaSys=" + dataRow[indeksy.brutto].ToString().Replace(",", ".").Replace("-", "") + ", Bzd_Kwota=" + dataRow[indeksy.brutto].ToString().Replace(",", ".").Replace("-", "") + ",  BZd_SplitPayKwotaVAT=0 WHERE Bzd_Numer='" + dataRow[indeksy.nr_faktury].ToString() + "'", out dummy, 0);
                        //, Bzd_KwotaSys=" + dataRow[indeksy.brutto].ToString().Replace(",", ".").Replace("-", "") + ", Bzd_Kwota=" + dataRow[indeksy.brutto].ToString().Replace(",", ".").Replace("-", "") +
                    }
                    else
                    {
                        object dummy = Type.Missing;

                        Sesja.Connection.Execute("update cdn.BnkZdarzenia set  Bzd_KwotaSys=" + dataRow[indeksy.brutto].ToString().Replace(",", ".").Replace("-", "") + ", Bzd_Kwota=" + dataRow[indeksy.brutto_eur].ToString().Replace(",", ".").Replace("-", "") + ",  BZd_SplitPayKwotaVAT=0 WHERE Bzd_Numer='" + dataRow[indeksy.nr_faktury].ToString() + "'", out dummy, 0);
                    }
                    Sesja.Save();
                }
                catch (Exception e)
                {
                    MessageBox.Show("Błąd zmiany Bank Zdarzenia dla faktury " + dataRow[indeksy.nr_faktury].ToString() + " : " + e.Message);
                }


                try
                {
                    Sesja.Save();
                    Sesja.Dispose();
                }
                catch (Exception)
                {

                }
            }
            catch(Exception an)
            {
                MessageBox.Show("Błąd funkcji dodajDokument dla faktury " + dataRow[indeksy.nr_faktury].ToString() + " : " + an.Message);
            }



        }

        private void DodajElementyDoRejestru_typ1(CDNRVAT.VAT rejestrVAT, DataRow dataRow, Indeksy indeksy, CDNBase.AdoSession Sesja)
        {
            try
            {
                CDNBase.ICollection Elementy = rejestrVAT.Elementy;
                #region element1
                var Element = (CDNRVAT.VATElement)Elementy.AddNew(null);
                int stawka = 0;
                if (dataRow[indeksy.vat].ToString() != "NP." && indeksy.vat != 0)
                {
                    stawka = Convert.ToInt32(dataRow[indeksy.vat].ToString());
                    Element.Stawka = Convert.ToDouble(stawka);
                }
                else if (dataRow[indeksy.vat].ToString() == "NP.")
                {
                    Element.Flaga = 4;
                    Element.Stawka = 0;

                }
                else
                {
                    // dodanie elementu o stawce zwolnionej z VAT
                    Element.Flaga = 1;
                    Element.Stawka = 0;
                }
                // Element.Flaga = 4;	// Flaga=4 oznacza stwkę NP.
                //Console.WriteLine("\tStawka ustawiona");

                /*
                 * Rodzaje zakupów:
                 * Towary		 = 1;
                 * Inne			 = 2;
                 * Trwale		 = 3;
                 * Uslugi		 = 4;
                 * NoweSrTran	 = 5;
                 * Nieruchomosci = 6;
                 * Paliwo		 = 7;
                 */
                Element.RodzajZakupu = 4;
                //Console.WriteLine("\tRodzaj zakupu ustawiony");

                Element.Netto = Convert.ToDecimal(dataRow[indeksy.energia].ToString()) - Convert.ToDecimal(dataRow[indeksy.akcyza].ToString()); // m na końcu oznacza typ decimal
                                                                                                                                           //Console.WriteLine("\tNetto ustawione");
                                                                                                                                         

                /* Porperty Odliczenia:
                 *  dla rej zakupu(odliczenia):		0-nie, 1-tak, 2-warunkowo
                 *	dla rej sprzed(uwz. w prop):	0-nie uwzgledniaj, 1-Uwzgledniaj w proporcji, 2- tylko w mianowniku
                 */
                Element.Odliczenia = 1;

                if (dataRow[indeksy.powiazana].ToString() == "1")
                {
                    var kategorie = (CDNBase.ICollection)(Sesja.CreateObject("CDN.Kategorie", null));
                    // pobieranie kontrahenta, formy platnosci  i waluty
                    var kategoria = (CDNHeal.Kategoria)kategorie["Kat_KodSzczegol='733-01-02-01'"];
                    Element.Kategoria = kategoria;
                }

                else
                {
                    var kategorie = (CDNBase.ICollection)(Sesja.CreateObject("CDN.Kategorie", null));
                    // pobieranie kontrahenta, formy platnosci  i waluty
                    var kategoria = (CDNHeal.Kategoria)kategorie["Kat_KodSzczegol='733-02-02-01'"];
                    Element.Kategoria = kategoria;
                }

                var kategorie2 = (CDNBase.ICollection)(Sesja.CreateObject("CDN.Kategorie", null));
                // pobieranie kontrahenta, formy platnosci  i waluty
                var kategoria2 = (CDNHeal.Kategoria)kategorie2["Kat_KodSzczegol='P.01-02-02'"];
                Element.Kategoria2 = kategoria2;
                

                Element.KatOpis = "Energia elektryczna";
                #endregion


                #region element2
                var Element2 = (CDNRVAT.VATElement)Elementy.AddNew(null);

                Element2.Flaga = Element.Flaga;
                Element2.Stawka = Element.Stawka;


                Element2.RodzajZakupu = Element.RodzajZakupu;
                //Console.WriteLine("\tRodzaj zakupu ustawiony");

                if (dataRow[indeksy.oplata].ToString() != "")
                    Element2.Netto = Convert.ToDecimal(dataRow[indeksy.oplata].ToString()); // m na końcu oznacza typ decimal
                                                                                            //Console.WriteLine("\tNetto ustawione");
                else
                    Element2.Netto = 0m;

                /* Porperty Odliczenia:
                 *  dla rej zakupu(odliczenia):		0-nie, 1-tak, 2-warunkowo
                 *	dla rej sprzed(uwz. w prop):	0-nie uwzgledniaj, 1-Uwzgledniaj w proporcji, 2- tylko w mianowniku
                 */
                Element2.Odliczenia = 1;
                if (dataRow[indeksy.powiazana].ToString() == "1")
                {
                    var kategorie = (CDNBase.ICollection)(Sesja.CreateObject("CDN.Kategorie", null));
                    // pobieranie kontrahenta, formy platnosci  i waluty
                    var kategoria = (CDNHeal.Kategoria)kategorie["Kat_KodSzczegol='733-01-02-02'"];
                    Element2.Kategoria = kategoria;
                }

                else
                {
                    var kategorie = (CDNBase.ICollection)(Sesja.CreateObject("CDN.Kategorie", null));
                    // pobieranie kontrahenta, formy platnosci  i waluty
                    var kategoria = (CDNHeal.Kategoria)kategorie["Kat_KodSzczegol='733-02-02-02'"];
                    Element2.Kategoria = kategoria;
                }
                Element2.Kategoria2 = kategoria2;
                Element2.KatOpis = "Energia elektryczna";
                #endregion


                #region element3
                var Element3 = (CDNRVAT.VATElement)Elementy.AddNew(null);

                Element3.Flaga = Element.Flaga;
                Element3.Stawka = Element.Stawka;

                Element3.RodzajZakupu = Element.RodzajZakupu;
                //Console.WriteLine("\tRodzaj zakupu ustawiony");

                Element3.Netto = Convert.ToDecimal(dataRow[indeksy.akcyza].ToString()); // m na końcu oznacza typ decimal
                                                                                        //Console.WriteLine("\tNetto ustawione");

                /* Porperty Odliczenia:
                 *  dla rej zakupu(odliczenia):		0-nie, 1-tak, 2-warunkowo
                 *	dla rej sprzed(uwz. w prop):	0-nie uwzgledniaj, 1-Uwzgledniaj w proporcji, 2- tylko w mianowniku
                 */
                Element3.Odliczenia = 1;
                if (dataRow[indeksy.powiazana].ToString() == "1")
                {
                    var kategorie = (CDNBase.ICollection)(Sesja.CreateObject("CDN.Kategorie", null));
                    // pobieranie kontrahenta, formy platnosci  i waluty
                    var kategoria = (CDNHeal.Kategoria)kategorie["Kat_KodSzczegol='733-01-02-03'"];
                    Element3.Kategoria = kategoria;
                }

                else
                {
                    var kategorie = (CDNBase.ICollection)(Sesja.CreateObject("CDN.Kategorie", null));
                    // pobieranie kontrahenta, formy platnosci  i waluty
                    var kategoria = (CDNHeal.Kategoria)kategorie["Kat_KodSzczegol='733-02-02-03'"];
                    Element3.Kategoria = kategoria;
                }

                Element3.Kategoria2 = kategoria2;
                Element3.KatOpis = "Energia elektryczna";
                #endregion

                /* #region element4
                 var Element4 = (CDNRVAT.VATElement)Elementy.AddNew(null);

                 Element4.Flaga = Element.Flaga;
                 Element4.Stawka = Element.Stawka;

                 Element4.RodzajZakupu = Element.RodzajZakupu;
                 //Console.WriteLine("\tRodzaj zakupu ustawiony");

                 Element4.Netto = -Convert.ToDecimal(dataRow[indeksy.akcyza].ToString()); // m na końcu oznacza typ decimal
                                                                                         //Console.WriteLine("\tNetto ustawione");

                 /* Porperty Odliczenia:
                  *  dla rej zakupu(odliczenia):		0-nie, 1-tak, 2-warunkowo
                  *	dla rej sprzed(uwz. w prop):	0-nie uwzgledniaj, 1-Uwzgledniaj w proporcji, 2- tylko w mianowniku
                  */
                /* Element4.Odliczenia = 1;
                 Element4.KatOpis = "Energia elektryczna";

                 #endregion


                 #region element5
                 var Element5 = (CDNRVAT.VATElement)Elementy.AddNew(null);

                 Element5.Flaga = Element.Flaga;
                 Element5.Stawka = Element.Stawka;

                 Element5.RodzajZakupu = Element.RodzajZakupu;
                 //Console.WriteLine("\tRodzaj zakupu ustawiony");

                 Element5.Netto = Convert.ToDecimal(dataRow[indeksy.akcyza].ToString()); // m na końcu oznacza typ decimal
                                                                                         //Console.WriteLine("\tNetto ustawione");

                 /* Porperty Odliczenia:
                  *  dla rej zakupu(odliczenia):		0-nie, 1-tak, 2-warunkowo
                  *	dla rej sprzed(uwz. w prop):	0-nie uwzgledniaj, 1-Uwzgledniaj w proporcji, 2- tylko w mianowniku
                  */
                /* Element5.Odliczenia = 1;
                 Element5.KatOpis = "Energia elektryczna";
                 #endregion*/




            }
            catch (Exception e)
            {
                MessageBox.Show("Błąd w dodawaniu pozycji: " + Convert.ToString(e), "Komunikat",
                                                MessageBoxButtons.OK, MessageBoxIcon.Error);


            }
           // Console.WriteLine("\tOdliczenia ustawione\n");
           
        }
        private void DodajElementyDoRejestru_typ2(CDNRVAT.VAT rejestrVAT, DataRow dataRow, Indeksy indeksy, CDNBase.AdoSession Sesja)
        {
            try
            {
                CDNBase.ICollection Elementy = rejestrVAT.Elementy;
                #region element1
                var Element = (CDNRVAT.VATElement)Elementy.AddNew(null);
                int stawka = 0;
                if (dataRow[indeksy.vat].ToString() != "NP." && indeksy.vat != 0)
                {
                    stawka = Convert.ToInt32(dataRow[indeksy.vat].ToString());
                    Element.Stawka = Convert.ToDouble(stawka);
                }
                else if (dataRow[indeksy.vat].ToString() == "NP.")
                {
                    Element.Flaga = 4;
                    Element.Stawka = 0;

                }
                else
                {
                    // dodanie elementu o stawce zwolnionej z VAT
                    Element.Flaga = 1;
                    Element.Stawka = 0;
                }
                // Element.Flaga = 4;	// Flaga=4 oznacza stwkę NP.
                //Console.WriteLine("\tStawka ustawiona");

                /*
                 * Rodzaje zakupów:
                 * Towary		 = 1;
                 * Inne			 = 2;
                 * Trwale		 = 3;
                 * Uslugi		 = 4;
                 * NoweSrTran	 = 5;
                 * Nieruchomosci = 6;
                 * Paliwo		 = 7;
                 */
                Element.RodzajZakupu = 4;
                //Console.WriteLine("\tRodzaj zakupu ustawiony");

                Element.Netto = Convert.ToDecimal(dataRow[indeksy.paliwo].ToString()) - Convert.ToDecimal(dataRow[indeksy.akcyza].ToString()); // m na końcu oznacza typ decimal
                                                                                                                                                //Console.WriteLine("\tNetto ustawione");

                /* Porperty Odliczenia:
                 *  dla rej zakupu(odliczenia):		0-nie, 1-tak, 2-warunkowo
                 *	dla rej sprzed(uwz. w prop):	0-nie uwzgledniaj, 1-Uwzgledniaj w proporcji, 2- tylko w mianowniku
                 */
                Element.Odliczenia = 1;
                
                if (dataRow[indeksy.powiazana].ToString() == "1")
                {
                    var kategorie = (CDNBase.ICollection)(Sesja.CreateObject("CDN.Kategorie", null));
                    // pobieranie kontrahenta, formy platnosci  i waluty
                    var kategoria = (CDNHeal.Kategoria)kategorie["Kat_KodSzczegol='733-01-01-01'"];
                    Element.Kategoria = kategoria;
                }

                else
                {
                    var kategorie = (CDNBase.ICollection)(Sesja.CreateObject("CDN.Kategorie", null));
                    // pobieranie kontrahenta, formy platnosci  i waluty
                    var kategoria = (CDNHeal.Kategoria)kategorie["Kat_KodSzczegol='733-02-01-01'"];
                    Element.Kategoria = kategoria;
                }

                var kategorie2 = (CDNBase.ICollection)(Sesja.CreateObject("CDN.Kategorie", null));
                // pobieranie kontrahenta, formy platnosci  i waluty
                var kategoria2 = (CDNHeal.Kategoria)kategorie2["Kat_KodSzczegol='P.01-01-02'"];
                Element.Kategoria2 = kategoria2;

                Element.KatOpis = "Paliwo gazowe";
                #endregion


                #region element2
                var Element2 = (CDNRVAT.VATElement)Elementy.AddNew(null);

                Element2.Flaga = Element.Flaga;
                Element2.Stawka = Element.Stawka;


                Element2.RodzajZakupu = Element.RodzajZakupu;
                //Console.WriteLine("\tRodzaj zakupu ustawiony");

                if (dataRow[indeksy.oplata].ToString() != "")
                    Element2.Netto = Convert.ToDecimal(dataRow[indeksy.oplata].ToString()); // m na końcu oznacza typ decimal
                                                                                            //Console.WriteLine("\tNetto ustawione");
                else
                    Element2.Netto = 0m;

                /* Porperty Odliczenia:
                 *  dla rej zakupu(odliczenia):		0-nie, 1-tak, 2-warunkowo
                 *	dla rej sprzed(uwz. w prop):	0-nie uwzgledniaj, 1-Uwzgledniaj w proporcji, 2- tylko w mianowniku
                 */
                Element2.Odliczenia = 1;
                if (dataRow[indeksy.powiazana].ToString() == "1")
                {
                    var kategorie = (CDNBase.ICollection)(Sesja.CreateObject("CDN.Kategorie", null));
                    // pobieranie kontrahenta, formy platnosci  i waluty
                    var kategoria = (CDNHeal.Kategoria)kategorie["Kat_KodSzczegol='733-01-01-02'"];
                    Element2.Kategoria = kategoria;
                }

                else
                {
                    var kategorie = (CDNBase.ICollection)(Sesja.CreateObject("CDN.Kategorie", null));
                    // pobieranie kontrahenta, formy platnosci  i waluty
                    var kategoria = (CDNHeal.Kategoria)kategorie["Kat_KodSzczegol='733-02-01-02'"];
                    Element2.Kategoria = kategoria;
                }
                Element2.Kategoria2 = kategoria2;

                Element2.KatOpis = "Paliwo gazowe";
                #endregion


                #region element3
                var Element3 = (CDNRVAT.VATElement)Elementy.AddNew(null);

                Element3.Flaga = Element.Flaga;
                Element3.Stawka = Element.Stawka;

                Element3.RodzajZakupu = Element.RodzajZakupu;
                //Console.WriteLine("\tRodzaj zakupu ustawiony");

                Element3.Netto = Convert.ToDecimal(dataRow[indeksy.oplata_zmienna].ToString()); // m na końcu oznacza typ decimal
                                                                                        //Console.WriteLine("\tNetto ustawione");

                /* Porperty Odliczenia:
                 *  dla rej zakupu(odliczenia):		0-nie, 1-tak, 2-warunkowo
                 *	dla rej sprzed(uwz. w prop):	0-nie uwzgledniaj, 1-Uwzgledniaj w proporcji, 2- tylko w mianowniku
                 */
                Element3.Odliczenia = 1;
                if (dataRow[indeksy.powiazana].ToString() == "1")
                {
                    var kategorie = (CDNBase.ICollection)(Sesja.CreateObject("CDN.Kategorie", null));
                    // pobieranie kontrahenta, formy platnosci  i waluty
                    var kategoria = (CDNHeal.Kategoria)kategorie["Kat_KodSzczegol='733-01-01-03'"];
                    Element3.Kategoria = kategoria;
                }

                else
                {
                    var kategorie = (CDNBase.ICollection)(Sesja.CreateObject("CDN.Kategorie", null));
                    // pobieranie kontrahenta, formy platnosci  i waluty
                    var kategoria = (CDNHeal.Kategoria)kategorie["Kat_KodSzczegol='733-02-01-03'"];
                    Element3.Kategoria = kategoria;
                }

                Element3.Kategoria2 = kategoria2;
                Element3.KatOpis = "Paliwo gazowe";
                #endregion

                #region element4
                var Element4 = (CDNRVAT.VATElement)Elementy.AddNew(null);

                Element4.Flaga = Element.Flaga;
                Element4.Stawka = Element.Stawka;

                Element4.RodzajZakupu = Element.RodzajZakupu;
                //Console.WriteLine("\tRodzaj zakupu ustawiony");

                decimal netto = 0.0m;
                if(indeksy.oplata_stala!=0 && dataRow[indeksy.oplata_stala].ToString()!="")
                    netto = Convert.ToDecimal(dataRow[indeksy.oplata_stala].ToString()); // m na końcu oznacza typ decimal
                                                                                         //Console.WriteLine("\tNetto ustawione");
                if (indeksy.oplata_stala2!=0 && dataRow[indeksy.oplata_stala2].ToString() != "")
                    netto += Convert.ToDecimal(dataRow[indeksy.oplata_stala2].ToString());

                Element4.Netto = netto;

                /* Porperty Odliczenia:
                 *  dla rej zakupu(odliczenia):		0-nie, 1-tak, 2-warunkowo
                 *	dla rej sprzed(uwz. w prop):	0-nie uwzgledniaj, 1-Uwzgledniaj w proporcji, 2- tylko w mianowniku
                 */
                Element4.Odliczenia = 1;
                if (dataRow[indeksy.powiazana].ToString() == "1")
                {
                    var kategorie = (CDNBase.ICollection)(Sesja.CreateObject("CDN.Kategorie", null));
                    // pobieranie kontrahenta, formy platnosci  i waluty
                    var kategoria = (CDNHeal.Kategoria)kategorie["Kat_KodSzczegol='733-01-01-04'"];
                    Element4.Kategoria = kategoria;
                }

                else
                {
                    var kategorie = (CDNBase.ICollection)(Sesja.CreateObject("CDN.Kategorie", null));
                    // pobieranie kontrahenta, formy platnosci  i waluty
                    var kategoria = (CDNHeal.Kategoria)kategorie["Kat_KodSzczegol='733-02-01-04'"];
                    Element4.Kategoria = kategoria;
                }
                Element4.Kategoria2 = kategoria2;
                Element4.KatOpis = "Paliwo gazowe";
                #endregion


                #region element5
                var Element5 = (CDNRVAT.VATElement)Elementy.AddNew(null);

                Element5.Flaga = Element.Flaga;
                Element5.Stawka = Element.Stawka;

                Element5.RodzajZakupu = Element.RodzajZakupu;
                //Console.WriteLine("\tRodzaj zakupu ustawiony");
                if (dataRow[indeksy.przekroczenie_mocy].ToString() != "")
                    Element5.Netto = Convert.ToDecimal(dataRow[indeksy.przekroczenie_mocy].ToString()); // m na końcu oznacza typ decimal
                                                                                                        //Console.WriteLine("\tNetto ustawione");
                else
                    Element5.Netto = 0m;

                /* Porperty Odliczenia:
                 *  dla rej zakupu(odliczenia):		0-nie, 1-tak, 2-warunkowo
                 *	dla rej sprzed(uwz. w prop):	0-nie uwzgledniaj, 1-Uwzgledniaj w proporcji, 2- tylko w mianowniku
                 */
                Element5.Odliczenia = 1;
                if (dataRow[indeksy.powiazana].ToString() == "1")
                {
                    var kategorie = (CDNBase.ICollection)(Sesja.CreateObject("CDN.Kategorie", null));
                    // pobieranie kontrahenta, formy platnosci  i waluty
                    var kategoria = (CDNHeal.Kategoria)kategorie["Kat_KodSzczegol='733-01-01-06'"];
                    Element5.Kategoria = kategoria;
                }

                else
                {
                    var kategorie = (CDNBase.ICollection)(Sesja.CreateObject("CDN.Kategorie", null));
                    // pobieranie kontrahenta, formy platnosci  i waluty
                    var kategoria = (CDNHeal.Kategoria)kategorie["Kat_KodSzczegol='733-02-01-06'"];
                    Element5.Kategoria = kategoria;
                }
                Element5.Kategoria2 = kategoria2;
                Element5.KatOpis = "Paliwo gazowe";
                #endregion

                #region element6
                var Element6 = (CDNRVAT.VATElement)Elementy.AddNew(null);

                Element6.Flaga = Element.Flaga;
                Element6.Stawka = Element.Stawka;

                Element6.RodzajZakupu = Element.RodzajZakupu;
                //Console.WriteLine("\tRodzaj zakupu ustawiony");
                if (dataRow[indeksy.akcyza].ToString() != "")
                    Element6.Netto = Convert.ToDecimal(dataRow[indeksy.akcyza].ToString()); // m na końcu oznacza typ decimal
                                                                                                        //Console.WriteLine("\tNetto ustawione");
                else
                    Element6.Netto = 0m;

                /* Porperty Odliczenia:
                 *  dla rej zakupu(odliczenia):		0-nie, 1-tak, 2-warunkowo
                 *	dla rej sprzed(uwz. w prop):	0-nie uwzgledniaj, 1-Uwzgledniaj w proporcji, 2- tylko w mianowniku
                 */
                Element6.Odliczenia = 1;
                if (dataRow[indeksy.powiazana].ToString() == "1")
                {
                    var kategorie = (CDNBase.ICollection)(Sesja.CreateObject("CDN.Kategorie", null));
                    // pobieranie kontrahenta, formy platnosci  i waluty
                    var kategoria = (CDNHeal.Kategoria)kategorie["Kat_KodSzczegol='733-01-01-08'"];
                    Element6.Kategoria = kategoria;
                }

                else
                {
                    var kategorie = (CDNBase.ICollection)(Sesja.CreateObject("CDN.Kategorie", null));
                    // pobieranie kontrahenta, formy platnosci  i waluty
                    var kategoria = (CDNHeal.Kategoria)kategorie["Kat_KodSzczegol='733-02-01-08'"];
                    Element6.Kategoria = kategoria;
                }

                Element6.Kategoria2 = kategoria2;
                Element6.KatOpis = "Paliwo gazowe";
                #endregion




            }
            catch (Exception e)
            {
                MessageBox.Show("Błąd w dodawaniu pozycji: " + Convert.ToString(e), "Komunikat",
                                                MessageBoxButtons.OK, MessageBoxIcon.Error);


            }
            // Console.WriteLine("\tOdliczenia ustawione\n");

        }

        private void DodajElementyDoRejestru_typ3(CDNRVAT.VAT rejestrVAT, DataRow dataRow, Indeksy indeksy, CDNBase.AdoSession Sesja)
        {
            try
            {
                CDNBase.ICollection Elementy = rejestrVAT.Elementy;
                #region element1
                var Element = (CDNRVAT.VATElement)Elementy.AddNew(null);
                int stawka = 0;
               // if(dataRow[indeksy.nr_faktury].ToString()== "F/PD/00008002/23/23")
                 //   MessageBox.Show("Vat: " + dataRow[indeksy.vat].ToString());
                if (!dataRow[indeksy.vat].ToString().Contains("NP") && !dataRow[indeksy.vat].ToString().Contains("ZW") && indeksy.vat != 0)
                {
                    stawka = Convert.ToInt32(dataRow[indeksy.vat].ToString());
                    Element.Stawka = Convert.ToDouble(stawka);
                }
                else if (dataRow[indeksy.vat].ToString().Contains("NP"))
                {
                    Element.Flaga = 4;
                    Element.Stawka = 0;

                }
                else
                {
                    // dodanie elementu o stawce zwolnionej z VAT
                    Element.Flaga = 1;
                    Element.Stawka = 0;
                }
                // Element.Flaga = 4;	// Flaga=4 oznacza stwkę NP.
                //Console.WriteLine("\tStawka ustawiona");

                /*
                 * Rodzaje zakupów:
                 * Towary		 = 1;
                 * Inne			 = 2;
                 * Trwale		 = 3;
                 * Uslugi		 = 4;
                 * NoweSrTran	 = 5;
                 * Nieruchomosci = 6;
                 * Paliwo		 = 7;
                 */
                Element.RodzajZakupu = 4;
                //Console.WriteLine("\tRodzaj zakupu ustawiony");

                if(indeksy.waluta!=0)
                {
                    Element.Netto = Convert.ToDecimal(dataRow[indeksy.netto_euro].ToString());
                }
                else
                    Element.Netto = Convert.ToDecimal(dataRow[indeksy.netto].ToString()); // m na końcu oznacza typ decimal
                                                                                                                                               //Console.WriteLine("\tNetto ustawione");

                /* Porperty Odliczenia:
                 *  dla rej zakupu(odliczenia):		0-nie, 1-tak, 2-warunkowo
                 *	dla rej sprzed(uwz. w prop):	0-nie uwzgledniaj, 1-Uwzgledniaj w proporcji, 2- tylko w mianowniku
                 */
                Element.Odliczenia = 1;
                if (dataRow[indeksy.powiazana].ToString() == "1")
                {
                    if (dataRow[indeksy.nazwa].ToString().Contains("wysokometanowy"))
                    {
                        var kategorie = (CDNBase.ICollection)(Sesja.CreateObject("CDN.Kategorie", null));
                        // pobieranie kontrahenta, formy platnosci  i waluty
                        var kategoria = (CDNHeal.Kategoria)kategorie["Kat_KodSzczegol='733-01-01-01'"];
                        Element.Kategoria = kategoria;
                    }
                    else
                    {
                        switch(dataRow[indeksy.nazwa].ToString())
                        {
                            case "Oplata za energię elektryczną bilansującą":
                                {
                                    var kategorie = (CDNBase.ICollection)(Sesja.CreateObject("CDN.Kategorie", null));
                                    // pobieranie kontrahenta, formy platnosci  i waluty
                                    var kategoria = (CDNHeal.Kategoria)kategorie["Kat_KodSzczegol='733-01-02-01'"];
                                    Element.Kategoria = kategoria;

                                    var kategorie2 = (CDNBase.ICollection)(Sesja.CreateObject("CDN.Kategorie", null));
                                    // pobieranie kontrahenta, formy platnosci  i waluty
                                    var kategoria2 = (CDNHeal.Kategoria)kategorie2["Kat_KodSzczegol='P.01-02-01'"];
                                    Element.Kategoria2 = kategoria2;
                                    break;
                                }
                            case "Opłata za energię elektryczną (Mwh)":
                                {
                                    var kategorie = (CDNBase.ICollection)(Sesja.CreateObject("CDN.Kategorie", null));
                                    // pobieranie kontrahenta, formy platnosci  i waluty
                                    var kategoria = (CDNHeal.Kategoria)kategorie["Kat_KodSzczegol='733-01-02-01'"];
                                    Element.Kategoria = kategoria;

                                    var kategorie2 = (CDNBase.ICollection)(Sesja.CreateObject("CDN.Kategorie", null));
                                    // pobieranie kontrahenta, formy platnosci  i waluty
                                    var kategoria2 = (CDNHeal.Kategoria)kategorie2["Kat_KodSzczegol='P.01-02-01'"];
                                    Element.Kategoria2 = kategoria2;
                                    break;
                                }
                            case "Neutrality charge":
                                {
                                    var kategorie = (CDNBase.ICollection)(Sesja.CreateObject("CDN.Kategorie", null));
                                    // pobieranie kontrahenta, formy platnosci  i waluty
                                    var kategoria = (CDNHeal.Kategoria)kategorie["Kat_KodSzczegol='733-01-01-01'"];
                                    Element.Kategoria = kategoria;
                                    break;
                                }
                            case "Imbalance":
                                {
                                    var kategorie = (CDNBase.ICollection)(Sesja.CreateObject("CDN.Kategorie", null));
                                    // pobieranie kontrahenta, formy platnosci  i waluty
                                    var kategoria = (CDNHeal.Kategoria)kategorie["Kat_KodSzczegol='733-01-01-05'"];
                                    Element.Kategoria = kategoria;
                                    break;
                                }
                            case "Opłata za Gotowość (kWh)":
                                {
                                    var kategorie = (CDNBase.ICollection)(Sesja.CreateObject("CDN.Kategorie", null));
                                    // pobieranie kontrahenta, formy platnosci  i waluty
                                    var kategoria = (CDNHeal.Kategoria)kategorie["Kat_KodSzczegol='733-01-01-02'"];
                                    Element.Kategoria = kategoria;
                                    break;
                                }
                        }
                    }
                }

                else
                {
                    if (dataRow[indeksy.nazwa].ToString().Contains("wysokometanowy"))
                    {
                        var kategorie = (CDNBase.ICollection)(Sesja.CreateObject("CDN.Kategorie", null));
                        // pobieranie kontrahenta, formy platnosci  i waluty
                        var kategoria = (CDNHeal.Kategoria)kategorie["Kat_KodSzczegol='733-02-01-01'"];
                        Element.Kategoria = kategoria;
                    }
                    else
                    {
                        switch (dataRow[indeksy.nazwa].ToString())
                        {
                            case "Oplata za energię elektryczną bilansującą":
                                {
                                    var kategorie = (CDNBase.ICollection)(Sesja.CreateObject("CDN.Kategorie", null));
                                    // pobieranie kontrahenta, formy platnosci  i waluty
                                    var kategoria = (CDNHeal.Kategoria)kategorie["Kat_KodSzczegol='733-02-02-01'"];
                                    Element.Kategoria = kategoria;

                                    var kategorie2 = (CDNBase.ICollection)(Sesja.CreateObject("CDN.Kategorie", null));
                                    // pobieranie kontrahenta, formy platnosci  i waluty
                                    var kategoria2 = (CDNHeal.Kategoria)kategorie2["Kat_KodSzczegol='P.01-02-01'"];
                                    Element.Kategoria2 = kategoria2;
                                    break;
                                }
                            case "Opłata za energię elektryczną (Mwh)":
                                {
                                    var kategorie = (CDNBase.ICollection)(Sesja.CreateObject("CDN.Kategorie", null));
                                    // pobieranie kontrahenta, formy platnosci  i waluty
                                    var kategoria = (CDNHeal.Kategoria)kategorie["Kat_KodSzczegol='733-02-02-01'"];
                                    Element.Kategoria = kategoria;

                                    var kategorie2 = (CDNBase.ICollection)(Sesja.CreateObject("CDN.Kategorie", null));
                                    // pobieranie kontrahenta, formy platnosci  i waluty
                                    var kategoria2 = (CDNHeal.Kategoria)kategorie2["Kat_KodSzczegol='P.01-02-01'"];
                                    Element.Kategoria2 = kategoria2;
                                    break;
                                }
                            case "Neutrality charge":
                                {
                                    var kategorie = (CDNBase.ICollection)(Sesja.CreateObject("CDN.Kategorie", null));
                                    // pobieranie kontrahenta, formy platnosci  i waluty
                                    var kategoria = (CDNHeal.Kategoria)kategorie["Kat_KodSzczegol='733-02-01-01'"];
                                    Element.Kategoria = kategoria;
                                    break;
                                }
                            case "Imbalance":
                                {
                                    var kategorie = (CDNBase.ICollection)(Sesja.CreateObject("CDN.Kategorie", null));
                                    // pobieranie kontrahenta, formy platnosci  i waluty
                                    var kategoria = (CDNHeal.Kategoria)kategorie["Kat_KodSzczegol='733-02-01-05'"];
                                    Element.Kategoria = kategoria;
                                    break;
                                }
                            case "Opłata za Gotowość (kWh)":
                                {
                                    var kategorie = (CDNBase.ICollection)(Sesja.CreateObject("CDN.Kategorie", null));
                                    // pobieranie kontrahenta, formy platnosci  i waluty
                                    var kategoria = (CDNHeal.Kategoria)kategorie["Kat_KodSzczegol='733-02-01-07'"];
                                    Element.Kategoria = kategoria;
                                    break;
                                }
                            case "Opłata związana z neutralnością bilansowania":
                                {
                                    var kategorie = (CDNBase.ICollection)(Sesja.CreateObject("CDN.Kategorie", null));
                                    // pobieranie kontrahenta, formy platnosci  i waluty
                                    var kategoria = (CDNHeal.Kategoria)kategorie["Kat_KodSzczegol='733-02-01-01'"];
                                    Element.Kategoria = kategoria;
                                    break;
                                }
                        }
                    }

                }

                if(Element.Kat2Opis=="")
                {
                    var kategorie2 = (CDNBase.ICollection)(Sesja.CreateObject("CDN.Kategorie", null));
                    // pobieranie kontrahenta, formy platnosci  i waluty
                    var kategoria2 = (CDNHeal.Kategoria)kategorie2["Kat_KodSzczegol='P.01-01-04'"];
                    Element.Kategoria2 = kategoria2;
                }

                Element.KatOpis = dataRow[indeksy.nazwa].ToString();
                #endregion






            }
            catch (Exception e)
            {
                MessageBox.Show("Błąd w dodawaniu pozycji: " + Convert.ToString(e), "Komunikat",
                                                MessageBoxButtons.OK, MessageBoxIcon.Error);


            }
            // Console.WriteLine("\tOdliczenia ustawione\n");

        }
    }
}
