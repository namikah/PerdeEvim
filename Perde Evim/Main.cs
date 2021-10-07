using Nsoft;
using System;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Windows.Forms;

namespace Perde_Evim
{
    public partial class Main : Form
    {
        public Main()
        {

            InitializeComponent();
        }


        Excel.Application oXL;
        Excel._Workbook oWB;
        Excel._Worksheet oSheet;

        private void KreditNomresiRefresh()
        {
            cbKreditNomresi.Items.Clear();
            cbKreditNomresi.Items.Add("");
            cbKreditNomresi.Text = "";

            MyData.selectCommand("Arxiv\\baza.accdb", "SELECT * FROM Kredit where a10 <> '0' and a2 Like '%" + txtAxtarKredit.Text + "%'");
            MyData.dtmainKredit=new DataTable();
            MyData.oledbadapter1.Fill(MyData.dtmainKredit);

            progressBar2.Value = 0;
            progressBar2.Maximum = MyData.dtmainKredit.Rows.Count;
            progressBar2.Step = 1;


            try
            {
                for (int i = 0; i < MyData.dtmainKredit.Rows.Count; i++)
                {
                    progressBar2.PerformStep();
                    base.Text = progressBar2.Value * 100 / MyData.dtmainKredit.Rows.Count + "% Kreditlər yenilənir #" + (MyData.dtmainKredit.Rows.Count - i) + " Baxılmamış sənəd qalıb";
                    cbKreditNomresi.Items.Add(MyData.dtmainKredit.Rows[i]["a5"] + " " + MyData.dtmainKredit.Rows[i]["a2"]);
                }

            }
            catch { }

            base.Text = "Pərdə Evim";

        }

        private void KreditRefresh()
        {
            if (cbBaglanmisKreditler.Checked == true)
            {
                string commandText = "SELECT * FROM Kredit where 1=1";
                commandText += " and a2 Like '%" + txtAxtarKredit.Text + "%'";
                commandText += " or a3 Like '%" + txtAxtarKredit.Text + "%'";
                commandText += " or a5 Like '%" + txtAxtarKredit.Text + "%' order by id desc";
                MyData.selectCommand("Arxiv\\baza.accdb", commandText);
                MyData.dtmainKredit = new DataTable();
                MyData.oledbadapter1.Fill(MyData.dtmainKredit);
            }
            else
            {
                string commandText = "SELECT * FROM Kredit where 1=1";
                commandText += " and a10 <> '0'";
                commandText += " and a2 Like '%" + txtAxtarKredit.Text + "%' order by id desc";
                MyData.selectCommand("Arxiv\\baza.accdb", commandText);
                MyData.dtmainKredit = new DataTable();
                MyData.oledbadapter1.Fill(MyData.dtmainKredit);
            }

            try { DataKredit.DataSource = MyData.dtmainKredit; } catch { }

            //this.DataKredit.Sort(this.DataKredit.Columns["id2"], ListSortDirection.Descending);

        }

        private void QrafikAvtoYarat()
        {
            try { dtTarix.Text = DataKredit.Rows[DataKredit.CurrentCell.RowIndex].Cells["a1"].Value.ToString(); }
            catch { }
            try { dtBirinciOdenis.Text = dtTarix.Value.AddMonths(1).ToShortDateString(); }
            catch { }
            try { txtadi.Text = DataKredit.Rows[DataKredit.CurrentCell.RowIndex].Cells["a2"].Value.ToString(); }
            catch { }
            try { txtunvan.Text = DataKredit.Rows[DataKredit.CurrentCell.RowIndex].Cells["a3"].Value.ToString(); }
            catch { }
            try { txtTel.Text = DataKredit.Rows[DataKredit.CurrentCell.RowIndex].Cells["a4"].Value.ToString(); }
            catch { }
            try { txtKreditNomresi.Text = DataKredit.Rows[DataKredit.CurrentCell.RowIndex].Cells["a5"].Value.ToString(); }
            catch { }
            try { txtMebleg.Text = DataKredit.Rows[DataKredit.CurrentCell.RowIndex].Cells["a7"].Value.ToString(); }
            catch { }
            try { txtİlkinOdenis.Text = DataKredit.Rows[DataKredit.CurrentCell.RowIndex].Cells["a8"].Value.ToString(); }
            catch { }
            try { txtMuddet.Text = DataKredit.Rows[DataKredit.CurrentCell.RowIndex].Cells["a12"].Value.ToString(); }
            catch { txtMuddet.Text = "8"; }
            try
            {
                if (DataKredit.Rows[DataKredit.CurrentCell.RowIndex].Cells["a12"].Value.ToString() == "") txtMuddet.Text = "8";
                if (DataKredit.Rows[DataKredit.CurrentCell.RowIndex].Cells["a12"].Value.ToString() == "0") txtMuddet.Text = "1";
            }
            catch { }


            try
            {
                MyData.updateCommand("Arxiv\\baza.accdb", "UPDATE Kredit SET "
                                                                                     + "a1 ='" + DataKredit.Rows[DataKredit.CurrentCell.RowIndex].Cells["a1"].Value.ToString() + "',"
                                                                                     + "a2 ='" + DataKredit.Rows[DataKredit.CurrentCell.RowIndex].Cells["a2"].Value.ToString() + "',"
                                                                                     + "a3 ='" + DataKredit.Rows[DataKredit.CurrentCell.RowIndex].Cells["a3"].Value.ToString() + "',"
                                                                                     + "a4 ='" + DataKredit.Rows[DataKredit.CurrentCell.RowIndex].Cells["a4"].Value.ToString() + "',"
                                                                                     + "a5 ='" + DataKredit.Rows[DataKredit.CurrentCell.RowIndex].Cells["a5"].Value.ToString() + "',"
                                                                                     + "a6 ='" + DataKredit.Rows[DataKredit.CurrentCell.RowIndex].Cells["a6"].Value.ToString() + "',"
                                                                                     + "a7 ='" + DataKredit.Rows[DataKredit.CurrentCell.RowIndex].Cells["a7"].Value.ToString() + "',"
                                                                                     + "a8 ='" + DataKredit.Rows[DataKredit.CurrentCell.RowIndex].Cells["a8"].Value.ToString() + "',"
                                                                                     + "a9 ='" + txtAyliqOdenis.Text + "',"
                                                                                     + "a10 ='" + DataKredit.Rows[DataKredit.CurrentCell.RowIndex].Cells["a10"].Value.ToString() + "',"
                                                                                     + "a11 ='" + DataKredit.Rows[DataKredit.CurrentCell.RowIndex].Cells["a11"].Value.ToString() + "',"
                                                                                     + "a12 ='" + txtMuddet.Text + "'"
                                                                                     + " WHERE id Like '" + DataKredit.Rows[DataKredit.CurrentCell.RowIndex].Cells["id2"].Value.ToString() + "'");
            }
            catch { }

            FreeAdd();
        }

        private void QrafikRefresh()
        {
            txtGecikmeQrafik.Text = "0 AZN";

            try
            {
                MyData.selectCommand("Qrafik\\" + lbKreditNomresiQrafik.Text + ".accdb", "SELECT * FROM Qrafik");
                //oledbadapter1.SelectCommand.CommandText += " or e2 Like '%" + txtAxtarHesab.Text + "%'";
                MyData.dtmainQrafik=new DataTable();
                MyData.oledbadapter1.Fill(MyData.dtmainQrafik);
                DataQrafik.DataSource = MyData.dtmainQrafik;
            }
            catch (System.Exception excep)
            {
                MessageBox.Show(excep.Message + Environment.NewLine + "Qrafik Tapılmadığı üçün yeni Qrafik yaradıldı.");
                QrafikAvtoYarat(); 
                return;
            }


            progressBar2.Value = 0;
            progressBar2.Maximum = MyData.dtmainQrafik.Rows.Count;
            progressBar2.Step = 1;

            int i = 0;
            string gecikmeGunleri = "";
            double qaliqQrafikUzre = 0, SsudaQrafik = 0, odenisler = 0;// ?????????????
            DateTime dt = DateTime.Today;

            qaliqQrafikUzre = Convert.ToDouble(txtQaliqQrafik.Text.Substring(0, txtQaliqQrafik.Text.Length - 4));
            lbQrafikUzreOdenilib.Text = "Kredit üzrə ödənilib (0 ay)";
            lbUmumiQalibQrafik.Text = "Ümumi borc (" + DataQrafik.Rows.Count.ToString() + " ay)";
            lbUmumiQalibQrafik.ForeColor = Color.Red;
            txtQaliqQrafik.ForeColor = Color.Red;
            if (Convert.ToDouble(txtQaliqQrafik.Text.Substring(0, txtQaliqQrafik.Text.Length - 4)) <= 0)
            {
                lbUmumiQalibQrafik.ForeColor = Color.Green;
                txtQaliqQrafik.ForeColor = Color.Green;
            }

            lbAydanQalanOdenis.Text  = "1 - ci aydan qalan";
            try
            {
                txtAydanQalanOdenis.Text = (Convert.ToDouble(DataQrafik.Rows[0].Cells["q2"].Value) - Convert.ToDouble(txtOdenilibQrafik.Text.Substring(0, txtOdenilibQrafik.Text.Length - 4))).ToString() + " AZN";
            }
            catch { }

            try
            {
                for (i = 0; i < DataQrafik.Rows.Count; i++)
                {
                    progressBar2.PerformStep();
                    base.Text = progressBar2.Value * 100 / DataQrafik.Rows.Count + "% Qrafik yenilənir #" + (DataQrafik.Rows.Count - i) + " Baxılmamış sənəd qalıb";

                    SsudaQrafik = Convert.ToDouble(DataQrafik.Rows[i].Cells["q5"].Value);

                    if (qaliqQrafikUzre <= SsudaQrafik)
                    {
                        odenisler += Convert.ToDouble(DataQrafik.Rows[i].Cells["q4"].Value);
                        //if (i < DataQrafik.Rows.Count - 1) odenisler += Convert.ToDouble(DataQrafik.Rows[i + 1].Cells["q2"].Value);//???????????????
                        DataQrafik.Rows[i].Cells["q2"].Style.BackColor = Color.GreenYellow;
                        DataQrafik.Rows[i].Cells["q3"].Style.BackColor = Color.GreenYellow;
                        DataQrafik.Rows[i].Cells["q4"].Style.BackColor = Color.GreenYellow;
                        DataQrafik.Rows[i].Cells["q5"].Style.BackColor = Color.GreenYellow;

                        //DataQrafik.Rows[i].Cells["q2"].Style.Font = new Font(DataQrafik.DefaultCellStyle.Font, FontStyle.Italic);
                        //DataQrafik.Rows[i].Cells["q3"].Style.Font = new Font(DataQrafik.DefaultCellStyle.Font, FontStyle.Italic);
                        //DataQrafik.Rows[i].Cells["q4"].Style.Font = new Font(DataQrafik.DefaultCellStyle.Font, FontStyle.Italic);
                        //DataQrafik.Rows[i].Cells["q5"].Style.Font = new Font(DataQrafik.DefaultCellStyle.Font, FontStyle.Italic);

                        lbQrafikUzreOdenilib.Text = "Kredit üzrə ödənilib (" + (i + 1) + " ay)";
                        lbUmumiQalibQrafik.Text = "Ümumi borc (" + (DataQrafik.Rows.Count - (i + 1)).ToString() + " ay)";

                        if (i < DataQrafik.Rows.Count - 1)
                        {
                            //Gecikme Gunlerinin hesablanmasi/////////////////////////////////////////////////
                            DateTime dt2 = DateTime.Parse(MyData.dtmainQrafik.Rows[i + 1]["Tarix"].ToString()).Date; //Gecikme gunleri ucun
                            gecikmeGunleri = (dt - dt2).TotalDays.ToString();
                            if ((dt - dt2).TotalDays <= 0 ||Convert.ToDouble(txtQaliqQrafik.Text.Substring(0,txtQaliqQrafik.Text.Length-4)) <= 0) { gecikmeGunleri = "0";}
                            lbGecikmeQrafik.Text = "Gecikmə (" + gecikmeGunleri + " gün)";
                            
                            //////////////////////////////////////////////////////////////////////////////////

                            txtAydanQalanOdenis.Text = (Math.Round(odenisler + Convert.ToDouble(DataQrafik.Rows[i + 1].Cells["q4"].Value) - Convert.ToDouble(txtOdenilibQrafik.Text.Substring(0, txtOdenilibQrafik.Text.Length - 4)), 2)).ToString() + " AZN";   //??????????
                            lbAydanQalanOdenis.Text = i + 2 + " - ci aydan qalan ";
                            lbAydanQalanOdenis.ForeColor = Color.Red;
                            txtAydanQalanOdenis.ForeColor = Color.Red;
                            if (Convert.ToDouble(txtAydanQalanOdenis.Text.Substring(0, txtAydanQalanOdenis.Text.Length - 4)) <= 0)
                            {
                                lbAydanQalanOdenis.ForeColor = Color.Green;
                                txtAydanQalanOdenis.ForeColor = Color.Green;
                            }
                        }

                    }

                    if (i == 0 && qaliqQrafikUzre > SsudaQrafik)
                    {
                        //Gecikme Gunlerinin hesablanmasi///////////////////////////////////////////////////////////////////
                        DateTime dt2 = DateTime.Parse(MyData.dtmainQrafik.Rows[0]["Tarix"].ToString()).Date; //Gecikme gunleri ucun
                        gecikmeGunleri = (dt - dt2).TotalDays.ToString();
                        if ((dt - dt2).TotalDays <= 0) gecikmeGunleri = "0";
                        lbGecikmeQrafik.Text = "Gecikmə (" + gecikmeGunleri + " gün)";
                        ////////////////////////////////////////////////////////////////////////////////////////////////////
                    }

                    ////////////////////////Ayin tutusdurulmasi//////////////////////////



                    int k = 0, kk = 0;
                    k = Convert.ToInt32(DataQrafik.Rows[i].Cells["q1"].Value.ToString().Substring(6, 4) + DataQrafik.Rows[i].Cells["q1"].Value.ToString().Substring(3, 2) + DataQrafik.Rows[i].Cells["q1"].Value.ToString().Substring(0, 2));
                    kk = Convert.ToInt32(dt.ToShortDateString().Substring(6, 4) + dt.ToShortDateString().Substring(3, 2) + dt.ToShortDateString().Substring(0, 2));

                    try
                    {
                        if (kk >= k)
                        {
                            //DataQrafik.Rows[i].Cells["id3"].Style.Font = new Font(DataQrafik.DefaultCellStyle.Font, FontStyle.Italic);
                            //DataQrafik.Rows[i].Cells["q1"].Style.Font = new Font(DataQrafik.DefaultCellStyle.Font, FontStyle.Italic);
                            DataQrafik.Rows[i].Cells["id3"].Style.BackColor = Color.GreenYellow;
                            DataQrafik.Rows[i].Cells["q1"].Style.BackColor = Color.GreenYellow;
                            //DataQrafik.Rows[i].Cells["id3"].Style.BackColor = Color.FromArgb(192, 255, 192);
                            //DataQrafik.Rows[i].Cells["q1"].Style.BackColor = Color.FromArgb(192,255,192);
                            //DataQrafik.Rows[ss].Cells["q5"].Style.BackColor = Color.LightGreen;
                            //DataQrafik.Rows[ss].DefaultCellStyle.Font = new Font(DataQrafik.DefaultCellStyle.Font, FontStyle.Bold);
                        }

                        if (kk > k)
                        {
                            txtGecikmeQrafik.Text = Math.Round(Convert.ToDouble(txtQaliqQrafik.Text.Substring(0, txtQaliqQrafik.Text.Length - 4)) - Convert.ToDouble(DataQrafik.Rows[i].Cells["q5"].Value), 2).ToString() + " AZN";
                        }

                    }
                    catch { }

                } //End for (i)
            } //End try (ii)
            catch (System.Exception excep)
            {
                MessageBox.Show(excep.Message + Environment.NewLine + "Tapılmadı.");
            }

            base.Text = "Pərdə Evim";
        }

        private void QrafikOdenislerRefresh()
        {
            double CemiGiris = 0, CemiCixis = 0; ;
            dataOdenislerQrafik.Rows.Clear();

            string commandText = "SELECT kassa1,kassa3,kassa4,kassa5 FROM Kassa where 1=1";
            commandText += " and kassa3 Like '%" + lbKreditNomresiQrafik.Text + "%'";
            MyData.selectCommand("Arxiv\\baza.accdb", commandText);
            MyData.dtmainKassa = new DataTable();
            MyData.oledbadapter1.Fill(MyData.dtmainKassa);

            for (int i = 0; i < MyData.dtmainKassa.Rows.Count; i++)
            {
                dataOdenislerQrafik.Rows.Add(MyData.dtmainKassa.Rows[i]["kassa1"], MyData.dtmainKassa.Rows[i]["kassa3"], MyData.dtmainKassa.Rows[i]["kassa4"], MyData.dtmainKassa.Rows[i]["kassa5"]);
                CemiGiris += Convert.ToDouble(MyData.dtmainKassa.Rows[i]["kassa4"]);
                CemiCixis += Convert.ToDouble(MyData.dtmainKassa.Rows[i]["kassa5"]);
            }

            dataOdenislerQrafik.Rows.Add("YEKUN", "", CemiGiris, CemiCixis);
            dataOdenislerQrafik.Rows[dataOdenislerQrafik.Rows.Count - 1].DefaultCellStyle.ForeColor = Color.Maroon;
            dataOdenislerQrafik.Rows[dataOdenislerQrafik.Rows.Count - 1].DefaultCellStyle.BackColor = Color.FromArgb(192, 192, 255);
        }

        private void KassaRefresh()
        {
            MyData.selectCommand("Arxiv\\baza.accdb", "SELECT * FROM Kassa where 1=1");
            MyData.dtmainKassa = new DataTable();
            MyData.oledbadapter1.Fill(MyData.dtmainKassa);

            try
            {
                txtQaliqKassa.Text = MyData.dtmainKassa.Rows[MyData.dtmainKassa.Rows.Count - 1]["kassa6"].ToString();

            }
            catch { }

            string commandText = "SELECT * FROM Kassa where 1=1";
            commandText += " and kassa3 Like '%" + txtAxtarKassa.Text + "%'";
            commandText += " or kassa2 Like '%" + txtAxtarKassa.Text + "%'";
            commandText += " or kassa1 Like '%" + txtAxtarKassa.Text + "%'";
            MyData.selectCommand("Arxiv\\baza.accdb", commandText);
            MyData.dtmainKassa = new DataTable();
            MyData.oledbadapter1.Fill(MyData.dtmainKassa);

            progressBar2.Value = 0;
            progressBar2.Maximum = MyData.dtmainKassa.Rows.Count;
            progressBar2.Step = 1;

            try
            {
                dataKassa.DataSource = MyData.dtmainKassa;
            }
            catch { }

            //this.dataKassa.Sort(this.dataKassa.Columns["idKassa"], ListSortDirection.Descending);

            double cemCixis = 0, cemGiris = 0, cemiQaliq = 0;

            for (int i = 0; i < dataKassa.Rows.Count; i++)
            {
                progressBar2.PerformStep();
                base.Text = progressBar2.Value * 100 / dataKassa.Rows.Count + "% Kassa yenilənir #" + (dataKassa.Rows.Count - i) + " Baxılmamış sənəd qalıb";

                cemCixis += Convert.ToDouble(dataKassa.Rows[i].Cells["kassa5"].Value);
                cemGiris += Convert.ToDouble(dataKassa.Rows[i].Cells["kassa4"].Value);
                cemiQaliq += Convert.ToDouble(dataKassa.Rows[i].Cells["kassa4"].Value) - Convert.ToDouble(dataKassa.Rows[i].Cells["kassa5"].Value);

            }

            try
            {
                txtCemiCixisKassa.Text = Math.Round(cemCixis, 2, MidpointRounding.AwayFromZero).ToString() + " AZN";
                txtCemiGirisKassa.Text = Math.Round(cemGiris, 2, MidpointRounding.AwayFromZero).ToString() + " AZN";
                txtCemiQaliqKassa.Text = Math.Round(cemiQaliq, 2, MidpointRounding.AwayFromZero).ToString() + " AZN";
            }
            catch { }

            base.Text = "Pərdə Evim";
        }

        private void SifarislerRefresh()
        {
            progressBar2.Value = 0;
            progressBar2.Maximum = 100;

            string commandText = "SELECT * FROM Sifarisler where 1=1";
            commandText += " and s1 Like '%" + txtAxtarSifarisler.Text + "%'";
            commandText += " or s13 Like '%" + txtAxtarSifarisler.Text + "%'";
            commandText += " or s10 Like '%" + txtAxtarSifarisler.Text + "%'";
            MyData.selectCommand("Arxiv\\baza.accdb", commandText);
            MyData.dtmainSifarisler = new DataTable();
            MyData.oledbadapter1.Fill(MyData.dtmainSifarisler);

            try { dataSifarisler.DataSource = MyData.dtmainSifarisler; }
            catch { }

            progressBar2.Value = 100;
            //this.dataSifarisler.Sort(this.dataSifarisler.Columns["idSifarisler"], ListSortDirection.Descending);

            base.Text = "Pərdə Evim";
        }

        private void QaliqlarRefresh()
        {
            double giris = 0, cixis = 0, qaliq = 0, Say = 0;

            MyData.selectCommand("Arxiv\\baza.accdb", "SELECT * FROM Kassa");
            MyData.dtmainKassa=new DataTable();
            MyData.oledbadapter1.Fill(MyData.dtmainKassa);
            dataKassa.DataSource = MyData.dtmainKassa;

            progressBar2.Value = 0;
            progressBar2.Maximum = MyData.dtmainKassa.Rows.Count;
            progressBar2.Step = 1;

            for (int i = 0; i < MyData.dtmainKassa.Rows.Count; i++)
            {
                progressBar2.PerformStep();
                base.Text = progressBar2.Value * 100 / MyData.dtmainKassa.Rows.Count + "% Qalıq yenilənir #" + (MyData.dtmainKassa.Rows.Count - i) + " Baxılmamış sənəd qalıb";

                try
                {
                    Say = Convert.ToDouble(MyData.dtmainKassa.Rows[i]["id"]);
                    giris = Convert.ToDouble(MyData.dtmainKassa.Rows[i]["kassa4"]);
                    cixis = Convert.ToDouble(MyData.dtmainKassa.Rows[i]["kassa5"]);
                    qaliq += giris - cixis;

                    MyData.updateCommand("Arxiv\\baza.accdb", "UPDATE Kassa SET kassa6 ='" + Math.Round(qaliq, 2, MidpointRounding.AwayFromZero).ToString() + "' WHERE id=" + Say);
                }
                catch (System.Exception excep)
                {
                    MessageBox.Show(excep.Message + Environment.NewLine + " Əməliyyat baş tutmadı ");
                }

            }

            base.Text = "Pərdə Evim";
        }

        private void QaliqlarQrafikRefresh()
        {
            try
            {
                MyData.selectCommand("Qrafik\\" + lbKreditNomresiQrafik.Text + ".accdb", "SELECT * FROM Qrafik");
                MyData.dtmainQrafik=new DataTable();
                MyData.oledbadapter1.Fill(MyData.dtmainQrafik);
            }
            catch (System.Exception excep)
            {
                MessageBox.Show(excep.Message + Environment.NewLine + "Qrafik Tapılmadığı üçün yeni Qrafik yaradıldı.");
                QrafikAvtoYarat();
                return;
            }

            progressBar2.Value = 0;
            progressBar2.Maximum = MyData.dtmainQrafik.Rows.Count;
            progressBar2.Step = 1;

            double Esas = 0, Faiz = 0, qaliq = Convert.ToDouble(txtMeblegQrafik.Text.Substring(0, txtMeblegQrafik.Text.Length - 4)) - Convert.ToDouble(txtAvansQrafik.Text.Substring(0, txtAvansQrafik.Text.Length - 4)), id = 0;

            for (int i = 0; i < MyData.dtmainQrafik.Rows.Count; i++)
            {
                progressBar2.PerformStep();
                base.Text = progressBar2.Value * 100 / MyData.dtmainQrafik.Rows.Count + "% Qalıq yenilənir #" + (MyData.dtmainQrafik.Rows.Count - i) + " Baxılmamış sənəd qalıb";

                try
                {
                    id = Convert.ToDouble(MyData.dtmainQrafik.Rows[i]["id"]);
                    Esas = Convert.ToDouble(MyData.dtmainQrafik.Rows[i]["Əsas"]);
                    //Son sudanin/qaligin son odenisle eynilesdirilmesi
                    if (i == MyData.dtmainQrafik.Rows.Count - 1) Esas = qaliq;
                    ////////////////////////////////////////////////////
                    Faiz = Convert.ToDouble(MyData.dtmainQrafik.Rows[i]["Faiz"]);
                    qaliq -= Esas;

                    MyData.updateCommand("Qrafik\\" + lbKreditNomresiQrafik.Text + ".accdb", "UPDATE Qrafik SET Ssuda ='" + Math.Round(qaliq, 2, MidpointRounding.AwayFromZero).ToString() + "', Ödəniş ='" + Math.Round(Esas + Faiz, 2, MidpointRounding.AwayFromZero).ToString() + "', Əsas ='" + Math.Round(Esas, 2, MidpointRounding.AwayFromZero).ToString() + "' WHERE id=" + id);
                }
                catch (System.Exception excep)
                {
                    MessageBox.Show(excep.Message + Environment.NewLine + " Əməliyyat baş tutmadı ");
                }
            }

            base.Text = "Pərdə Evim";
        }

        /* ANNUITET ADD
            private void AnnuitetAdd() //Debet ve Kredit ucun
        {
            try
            {
                File.Copy("Qrafik\\Qrafik.accdb", "Qrafik\\" + txtKreditNomresi.Text + ".accdb", true);
            }
            catch { MessageBox.Show("Qrafik.accdb tapılmadı."); }

            try
            {
                DateTime dt = dtTarix.Value;
                string Time = "";
                double Odenis = 0, f = 0, t = 0, m = 0, ssuda = 0, Faiz = 0, Esas = 0, Faiz2 = 0, Esas2 = 0;
                f = Convert.ToDouble(txtFaiz.Text);
                if (txtFaiz.Text == "0") f = 0.00000001;
                m = Convert.ToDouble(txtMebleg.Text) - Convert.ToDouble(txtİlkinOdenis.Text);
                if (txtMebleg.Text == "0") m = 0.00000001;
                t = Convert.ToDouble(txtMuddet.Text);
                if (txtMuddet.Text == "0") t = 0.00000001;
                Odenis = m * (f / 100 / 12) / (1 - 1 / Math.Pow(1 + f / 100 / 12, t));
                ssuda = Convert.ToDouble(txtMebleg.Text) - Convert.ToDouble(txtİlkinOdenis.Text);
                Faiz = (Convert.ToDouble(txtMebleg.Text) - Convert.ToDouble(txtİlkinOdenis.Text)) * Convert.ToDouble(txtFaiz.Text) / 1200;
                Esas = ssuda - Faiz;

                for (int a = 0; a < Convert.ToDouble(txtMuddet.Text); a++)
                {
                    Faiz2 = ssuda * Convert.ToDouble(txtFaiz.Text) / 1200;
                    Esas2 = Odenis - Faiz2;
                    ssuda = ssuda - Esas2;
                    Time = dt.AddMonths(a + 1).ToShortDateString();

                    CreateSqlConnectionQrafik();
                    oledbadapter1.InsertCommand = new OleDbCommand();
                    oledbadapter1.InsertCommand.Connection = oledbconnection1;
                    oledbconnection1.Open();
                    oledbadapter1.InsertCommand.CommandText = "INSERT INTO Qrafik (Tarix, Ödəniş, Faiz, Əsas, Ssuda)values("


                                                                                                        + "'" + Time + "',"
                                                                                                        + "'" + Math.Round(Odenis, 2).ToString() + "',"
                                                                                                        + "'" + Math.Round(Faiz2, 2).ToString() + "',"
                                                                                                        + "'" + Math.Round(Esas2, 2).ToString() + "',"
                                                                                                        + "'" + Math.Round(ssuda, 2).ToString() + "')";

                    oledbadapter1.InsertCommand.ExecuteNonQuery();
                    oledbconnection1.Close();


                }
            }
            catch { return; }
        }
        */

        private void QebzPrint()
        {
            try
            {
                File.Copy("Excel\\Qebz.xlsx", "C:\\Users\\" + Environment.UserName + "\\Desktop\\Qebz.xlsx", true);
            }
            catch { MessageBox.Show("Qebz.xlsx tapılmadı."); }


            //Get a new workbook.
            oXL = new Excel.Application();
            oWB = (Excel._Workbook)(oXL.Workbooks.Open("C:\\Users\\" + Environment.UserName + "\\Desktop\\Qebz.xlsx"));
            oSheet = (Excel._Worksheet)oWB.Sheets[1];
            oSheet.Visible = Excel.XlSheetVisibility.xlSheetVisible;

            oXL.Visible = true;
            oSheet.Activate();
            oSheet.Range["A1"].Select();
            oSheet.PageSetup.Orientation = Excel.XlPageOrientation.xlPortrait;

            try
            {
              

            oSheet.Cells[2, "O"] = "Tarix: " + dtTarixKassa.Text.ToString();
            oSheet.Cells[5, "W"] = txtGiris.Text;
            oSheet.Cells[6, "E"] = MyChange.ReqemToMetn(Convert.ToDouble(txtGiris.Text));
            oSheet.Cells[7, "E"] = txtAciqlama.Text;
            oSheet.Cells[8, "E"] = txtOdeyen.Text;

            oSheet.Cells[15, "O"] = "Tarix: " + dtTarixKassa.Text.ToString();
            oSheet.Cells[18, "W"] = txtGiris.Text;
            oSheet.Cells[19, "E"] = MyChange.ReqemToMetn(Convert.ToDouble(txtGiris.Text));
            oSheet.Cells[20, "E"] = txtAciqlama.Text;
            oSheet.Cells[21, "E"] = txtOdeyen.Text;
            }
            catch { }
        }

        private void FreeAdd()
        {

             try
            {
                File.Copy("Qrafik\\Qrafik.accdb", "Qrafik\\" + txtKreditNomresi.Text + ".accdb", true);
            }
             catch (System.Exception excep)
             {
                 MessageBox.Show(excep.Message + Environment.NewLine + " Qrafik.accdb tapılmadı ");
             }

            try
            {
                DateTime dt = dtBirinciOdenis.Value;
                string Time = dt.ToShortDateString();
                double Odenis = 0, t = 0, m = 0, ssuda = 0, Faiz = 0, Esas = 0, Faiz2 = 0, Esas2 = 0;
                m = Convert.ToDouble(txtMebleg.Text) - Convert.ToDouble(txtİlkinOdenis.Text);
                if (txtMebleg.Text == "0") m = 0.00000001;
                t = Convert.ToDouble(txtMuddet.Text);
                if (txtMuddet.Text == "0") t = 1;
                if (txtMuddet.Text == "") t = 8;
                Odenis = Math.Round(Convert.ToDouble(txtAyliqOdenis.Text), 2, MidpointRounding.AwayFromZero);
                ssuda = Convert.ToDouble(txtMebleg.Text) - Convert.ToDouble(txtİlkinOdenis.Text);
                Faiz = 0;
                Esas = ssuda - Faiz;

                progressBar2.Value = 0;
                try { progressBar2.Maximum = Convert.ToInt32(t); }
                catch { }
                progressBar2.Step = 1;

                for (int a = 0; a < t; a++)
                {
                    progressBar2.PerformStep();
                    base.Text = progressBar2.Value * 100 / t + "% Qrafik yaradılır #" + a.ToString();

                    Faiz2 = ssuda * 0 / 1200;
                    Esas2 = Odenis - Faiz2;
                    //Son ssudanin/qaligin ayliq odenisle beraberlesdirilmesi ucun
                    if (a == t - 1) 
                    {
                        Esas2 = ssuda;
                        Odenis = Esas2 + Faiz2;
                    } ///////////////////////////////////////////////////////////////
                    ssuda = ssuda - Esas2;

                    MyData.insertCommand("Qrafik\\" + txtKreditNomresi.Text + ".accdb'", "INSERT INTO Qrafik (Tarix, Ödəniş, Faiz, Əsas, Ssuda)values("


                                                                                                        + "'" + Time + "',"
                                                                                                        + "'" + Math.Round(Odenis, 2, MidpointRounding.AwayFromZero).ToString() + "',"
                                                                                                        + "'" + Math.Round(Faiz2, 2, MidpointRounding.AwayFromZero).ToString() + "',"
                                                                                                        + "'" + Math.Round(Esas2, 2, MidpointRounding.AwayFromZero).ToString() + "',"
                                                                                                        + "'" + Math.Round(ssuda, 2, MidpointRounding.AwayFromZero).ToString() + "')");

                    Time = dt.AddMonths(a + 1).ToShortDateString();
                }
            }
            catch (System.Exception excep)
            {
                MessageBox.Show(excep.Message + Environment.NewLine + " Əməliyyat baş tutmadı ");
                return;
            }
            base.Text = "Pərdə Evim";

        } 

        private void infoRefresh()
        {
            MyData.selectCommand("Arxiv\\baza.accdb", "Select * From Kredit where a10 <> '0'");
            MyData.dtmainKredit=new DataTable();
            MyData.oledbadapter1.Fill(MyData.dtmainKredit);

            progressBar2.Value = 0;
            progressBar2.Maximum = MyData.dtmainKredit.Rows.Count;
            progressBar2.Step = 1;

            dataGecikme.Rows.Clear();
            dataBugun.Rows.Clear();
            dataSabah.Rows.Clear();

            int k = 0, kk = 0, kkk = 0, i = 0, ii = 0;
            double gecikmeCemi = 0, YekunUmumiBorc = 0;
            try
            {
                for (ii = 0; ii < MyData.dtmainKredit.Rows.Count; ii++)
                {
                    progressBar2.PerformStep();
                    base.Text = progressBar2.Value * 100 / MyData.dtmainKredit.Rows.Count + "% info yenilənir #" + (MyData.dtmainKredit.Rows.Count - ii) + " Baxılmamış sənəd qalıb";

                    try
                    {
                        MyData.selectCommand("Qrafik\\" + MyData.dtmainKredit.Rows[ii]["a5"].ToString() + ".accdb", "Select * From Qrafik");
                        MyData.dtmainQrafik=new DataTable();
                        MyData.oledbadapter1.Fill(MyData.dtmainQrafik);
                    }

                    catch (System.Exception excep)
                    {
                        MessageBox.Show(excep.Message + Environment.NewLine + "Qrafik Tapılmadı."); return;
                    }

                    DateTime dt = DateTime.Today;
                    double GecikmeAZN = 0, OdenmeliMebleğ = 0, umumiBorc = Convert.ToDouble(MyData.dtmainKredit.Rows[ii]["a10"]);

                    for (i = 0; i < MyData.dtmainQrafik.Rows.Count; i++)
                    {
                        OdenmeliMebleğ = Convert.ToDouble(MyData.dtmainQrafik.Rows[i]["Ödəniş"]);

                        k = Convert.ToInt32(MyData.dtmainQrafik.Rows[i]["Tarix"].ToString().Substring(6, 4) + MyData.dtmainQrafik.Rows[i]["Tarix"].ToString().Substring(3, 2) + MyData.dtmainQrafik.Rows[i]["Tarix"].ToString().Substring(0, 2));
                        kk = Convert.ToInt32(dt.ToShortDateString().Substring(6, 4) + dt.ToShortDateString().Substring(3, 2) + dt.ToShortDateString().Substring(0, 2));
                        kkk = Convert.ToInt32(dt.AddDays(1).ToShortDateString().Substring(6, 4) + dt.AddDays(1).ToShortDateString().Substring(3, 2) + dt.AddDays(1).ToShortDateString().Substring(0, 2));


                        if (kk > k)  // GECIKMEDE OLANLAR
                        {
                            GecikmeAZN = Math.Round(Convert.ToDouble(MyData.dtmainKredit.Rows[ii]["a10"]) - Convert.ToDouble(MyData.dtmainQrafik.Rows[i]["Ssuda"]), 2);

                        }


                        if (kk == k && Convert.ToDouble(MyData.dtmainKredit.Rows[ii]["a10"]) > 0) // BU GUNE ODENISLER 
                        {
                            if (OdenmeliMebleğ > umumiBorc) OdenmeliMebleğ = umumiBorc;
                            if (Convert.ToDouble(MyData.dtmainKredit.Rows[ii]["a10"]) > Convert.ToDouble(MyData.dtmainQrafik.Rows[i]["Ssuda"])) dataBugun.Rows.Add(MyData.dtmainKredit.Rows[ii]["a5"].ToString(), MyData.dtmainKredit.Rows[ii]["a2"].ToString(), OdenmeliMebleğ + GecikmeAZN, umumiBorc, "Ödənilməyib", MyData.dtmainKredit.Rows[ii]["a4"].ToString(), MyData.dtmainKredit.Rows[ii]["a11"].ToString());
                            else dataBugun.Rows.Add(MyData.dtmainKredit.Rows[ii]["a5"].ToString(), MyData.dtmainKredit.Rows[ii]["a2"].ToString(), 0, umumiBorc, "Ödənilib", MyData.dtmainKredit.Rows[ii]["a4"].ToString(), MyData.dtmainKredit.Rows[ii]["a11"].ToString());
                        }

                        if (kkk == k && Convert.ToDouble(MyData.dtmainKredit.Rows[ii]["a10"]) > 0) // SABAHA ODENISLER 
                        {
                            if (OdenmeliMebleğ > umumiBorc) OdenmeliMebleğ = umumiBorc;
                            if (Convert.ToDouble(MyData.dtmainKredit.Rows[ii]["a10"]) > Convert.ToDouble(MyData.dtmainQrafik.Rows[i]["Ssuda"])) dataSabah.Rows.Add(MyData.dtmainKredit.Rows[ii]["a5"].ToString(), MyData.dtmainKredit.Rows[ii]["a2"].ToString(), OdenmeliMebleğ + GecikmeAZN, umumiBorc, "Ödənilməyib", MyData.dtmainKredit.Rows[ii]["a4"].ToString(), MyData.dtmainKredit.Rows[ii]["a11"].ToString());
                            else dataSabah.Rows.Add(MyData.dtmainKredit.Rows[ii]["a5"].ToString(), MyData.dtmainKredit.Rows[ii]["a2"].ToString(), 0, umumiBorc, "Ödənilib", MyData.dtmainKredit.Rows[ii]["a4"].ToString(), MyData.dtmainKredit.Rows[ii]["a11"].ToString());
                        }


                    }//End for

                    if (GecikmeAZN > 0) // GECIKMEDE OLANLAR
                    {
                        gecikmeCemi += GecikmeAZN;
                        dataGecikme.Rows.Add(MyData.dtmainKredit.Rows[ii]["a5"].ToString(), MyData.dtmainKredit.Rows[ii]["a2"].ToString(), GecikmeAZN.ToString(), umumiBorc, MyData.dtmainKredit.Rows[ii]["a6"].ToString(), MyData.dtmainKredit.Rows[ii]["a4"].ToString(), MyData.dtmainKredit.Rows[ii]["a11"].ToString());
                        YekunUmumiBorc += Convert.ToDouble(MyData.dtmainKredit.Rows[ii]["a10"]);
                    }



                }//End for

                dataGecikme.Rows.Add("YEKUN", "", gecikmeCemi, YekunUmumiBorc, "", "", ""); //Yekun meblegin yazilmasi
                lbGecikmedeOlanlar.Text = "GECİKMƏDƏ OLANLAR (" + gecikmeCemi + " AZN)";

                double bugun = 0, bugunumumi = 0, sabah = 0, sabahumumi = 0;
                try
                {
                    for (int cem = 0; cem < dataBugun.Rows.Count; cem++)
                    {
                        bugun += Convert.ToDouble(dataBugun.Rows[cem].Cells["odenilmeli"].Value);
                        bugunumumi += Convert.ToDouble(dataBugun.Rows[cem].Cells["umumi"].Value);
                    }
                }
                catch { }

                try
                {
                    for (int cem = 0; cem < dataSabah.Rows.Count; cem++)
                    {
                        sabah += Convert.ToDouble(dataSabah.Rows[cem].Cells["odenilmeliSabah"].Value);
                        sabahumumi += Convert.ToDouble(dataSabah.Rows[cem].Cells["umumiSabah"].Value);
                    }
                }
                catch { }

                lbBuguneCemi.Text = "BUGÜN ÖDƏNMƏLİDİR (" + bugun.ToString() + " AZN) ";
                lbSabahCemi.Text = "SABAH ÖDƏNMƏLİDİR (" + sabah.ToString() + " AZN)";

                dataBugun.Rows.Add("YEKUN", "", bugun, bugunumumi, "", "", "");
                dataSabah.Rows.Add("YEKUN", "", sabah, sabahumumi, "", "", "");

                dataBugun.Rows[dataBugun.Rows.Count - 1].DefaultCellStyle.ForeColor = Color.Maroon;
                dataSabah.Rows[dataSabah.Rows.Count - 1].DefaultCellStyle.ForeColor = Color.Maroon;
                dataGecikme.Rows[dataGecikme.Rows.Count - 1].DefaultCellStyle.ForeColor = Color.Maroon;
                dataGecikme.Rows[dataGecikme.Rows.Count - 1].DefaultCellStyle.BackColor = Color.FromArgb(192, 192, 255);
                dataBugun.Rows[dataBugun.Rows.Count - 1].DefaultCellStyle.BackColor = Color.FromArgb(192, 192, 255);
                dataSabah.Rows[dataSabah.Rows.Count - 1].DefaultCellStyle.BackColor = Color.FromArgb(192, 192, 255);
            }
            catch (System.Exception excep)
            {
                MessageBox.Show(excep.Message + Environment.NewLine + "Tapılmadı.");
            }

            base.Text = "Pərdə Evim";
        }

        private void KassaEmeliyyati()
        {
            MyCheck.ParolYoxla();
            if (!MyCheck.Parolicaze) return;

            if(!MyCheck.davamYesNo()) return;

            try
            {
                MyData.insertCommand("Arxiv\\baza.accdb", "INSERT INTO Kassa (kassa1,kassa2,kassa3,kassa4,kassa5,kassa6)values("



                                                                                                    + "'" + dtTarixKassa.Text + "',"
                                                                                                    + "'" + txtStatus.Text + "',"
                                                                                                    + "'" + txtAciqlama.Text + "',"
                                                                                                    + "'" + Math.Round(Convert.ToDouble(txtGiris.Text), 2, MidpointRounding.AwayFromZero).ToString() + "',"
                                                                                                    + "'" + Math.Round(Convert.ToDouble(txtCixis.Text), 2, MidpointRounding.AwayFromZero).ToString() + "',"
                                                                                                    + "'" + Math.Round((Convert.ToDouble(txtQaliqKassa.Text) + Convert.ToDouble(txtGiris.Text) - Convert.ToDouble(txtCixis.Text)), 2, MidpointRounding.AwayFromZero).ToString() + "')");
            }
            catch (System.Exception excep)
            {
                MessageBox.Show(excep.Message + Environment.NewLine + Environment.NewLine + "Əməliyyat baş tutmadı. " + Environment.NewLine + Environment.NewLine + "Giriş/Çıxış məbləğləri reqem olmalıdır.");
                return;
            }

            //QaliqlarRefresh();
            KassaRefresh();
            KreditRefresh();
            QrafikRefresh();

        }

        private void KreditOdenisi()
        {
            MyCheck.ParolYoxla();
            if (!MyCheck.Parolicaze) return;

            if (!MyCheck.davamYesNo()) return;
            ////////////////////////////////////////////////Odenisin Borcdan Silinmesi ucun////////////////////////////////

            int i = 0;
            string um = "";
            try
            {
                for (i = 0; i < cbKreditNomresi.Text.Length; i++)
                {
                    if (cbKreditNomresi.Text.Substring(i, 1) == " ")
                    {
                        um = cbKreditNomresi.Text.Substring(0, i);
                        break;
                    }
                }
            } 
            catch (System.Exception ex) { MessageBox.Show( ex.Message + Environment.NewLine + "ComboBox siyahıda Kredit nömrəsi tapılmadı."); return; }


            try
            {
                string commandText = "SELECT * FROM Kredit where 1=1";
                commandText += " and a5 = '" + um + "'";
                MyData.selectCommand("Arxiv\\baza.accdb", commandText);
                MyData.dtmainKredit=new DataTable();
                MyData.oledbadapter1.Fill(MyData.dtmainKredit);
                string qaliq = MyData.dtmainKredit.Rows[0]["a10"].ToString();

                MyData.updateCommand("Arxiv\\baza.accdb", "UPDATE Kredit SET a10 ='" + Math.Round((Convert.ToDouble(qaliq) - Convert.ToDouble(txtGiris.Text) + Convert.ToDouble(txtCixis.Text)), 2, MidpointRounding.AwayFromZero).ToString() + "', a6 = '" + dtTarixKassa.Value.ToShortDateString() + "' WHERE a5='" + um + "'");
            }
            catch (System.Exception ex) { MessageBox.Show(ex.Message + Environment.NewLine + "Kreditlərdə Kredit nömrəsi tapılmadı."); return; }

            try
            {
                MyData.insertCommand("Arxiv\\baza.accdb", "INSERT INTO Kassa (kassa1,kassa2,kassa3,kassa4,kassa5,kassa6)values("



                                                                                                    + "'" + dtTarixKassa.Text + "',"
                                                                                                    + "'" + txtStatus.Text + "',"
                                                                                                    + "'" + txtAciqlama.Text + "',"
                                                                                                    + "'" + Math.Round(Convert.ToDouble(txtGiris.Text), 2, MidpointRounding.AwayFromZero).ToString() + "',"
                                                                                                    + "'" + Math.Round(Convert.ToDouble(txtCixis.Text), 2, MidpointRounding.AwayFromZero).ToString() + "',"
                                                                                                    + "'" + Math.Round((Convert.ToDouble(txtQaliqKassa.Text) + Convert.ToDouble(txtGiris.Text) - Convert.ToDouble(txtCixis.Text)), 2, MidpointRounding.AwayFromZero).ToString() + "')");
            }
            catch (System.Exception excep)
            {
                MessageBox.Show(excep.Message + Environment.NewLine + " Əməliyyat baş tutmadı ");
                return;
            }


            QebzPrint();

            //QaliqlarRefresh();
            KassaRefresh();
            KreditRefresh();
            QrafikRefresh();
            qaliqmusterikassa();
        }

        private void AvansOdenisi()
        {
            MyCheck.ParolYoxla();
            if (!MyCheck.Parolicaze) return;

            if (!MyCheck.davamYesNo()) return;

            ////////////////////////////////////////////////Odenisin Borcdan Silinmesi ucun////////////////////////////////

            int i = 0;
            string um = "";
            try
            {
                for (i = 0; i < cbKreditNomresi.Text.Length; i++)
                {
                    if (cbKreditNomresi.Text.Substring(i, 1) == " ")
                    {
                        um = cbKreditNomresi.Text.Substring(0, i);
                        break;
                    }
                }
            }
            catch { MessageBox.Show("Kredit nömrəsi tapılmadı."); return; }


            try
            {
                MyData.selectCommand("Arxiv\\baza.accdb", "SELECT * FROM Kredit where a5 = '" + um + "'");
                MyData.dtmainKredit=new DataTable();
                MyData.oledbadapter1.Fill(MyData.dtmainKredit);
                string qaliq = MyData.dtmainKredit.Rows[0]["a10"].ToString(),
                       Avans = MyData.dtmainKredit.Rows[0]["a8"].ToString();

                MyData.updateCommand("Arxiv\\baza.accdb", "UPDATE Kredit SET a6 = '" + dtTarixKassa.Value.ToShortDateString() + "' WHERE a5='" + um + "'");
            }
            catch { MessageBox.Show("Kredit nömrəsi tapılmadı."); return; }

            try
            {
                MyData.insertCommand("Arxiv\\baza.accdb", "INSERT INTO Kassa (kassa1,kassa2,kassa3,kassa4,kassa5,kassa6)values("



                                                                                                    + "'" + dtTarixKassa.Text + "',"
                                                                                                    + "'" + txtStatus.Text + "',"
                                                                                                    + "'" + txtAciqlama.Text + "',"
                                                                                                    + "'" + Math.Round(Convert.ToDouble(txtGiris.Text), 2, MidpointRounding.AwayFromZero).ToString() + "',"
                                                                                                    + "'" + Math.Round(Convert.ToDouble(txtCixis.Text), 2, MidpointRounding.AwayFromZero).ToString() + "',"
                                                                                                    + "'" + Math.Round((Convert.ToDouble(txtQaliqKassa.Text) + Convert.ToDouble(txtGiris.Text) - Convert.ToDouble(txtCixis.Text)), 2, MidpointRounding.AwayFromZero).ToString() + "')");
            }
            catch (System.Exception excep)
            {
                MessageBox.Show(excep.Message + Environment.NewLine + " Əməliyyat baş tutmadı ");
                return;
            }


            QebzPrint();

            //QaliqlarRefresh();
            KassaRefresh();
            KreditRefresh();
            QrafikRefresh();
            qaliqmusterikassa();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            MyCheck.ParolYoxla();
            if (!MyCheck.Parolicaze) return;

                try { KreditRefresh(); }
            catch (System.Exception excep)
            {
                MessageBox.Show(excep.Message + Environment.NewLine + "Kredit Yenilənmədi.");
            }

            try { QrafikRefresh(); }
            catch (System.Exception excep)
            {
                MessageBox.Show(excep.Message + Environment.NewLine + "Qrafik Yenilənmədi.");
            }

            try { SifarislerRefresh(); }
            catch (System.Exception excep)
            {
                MessageBox.Show(excep.Message + Environment.NewLine + "Sifarişlər Yenilənmədi.");
            }

            try { KassaRefresh(); }
            catch (System.Exception excep)
            {
                MessageBox.Show(excep.Message + Environment.NewLine + "Kassa Yenilənmədi.");
            }

            try { infoRefresh(); }
            catch (System.Exception excep)
            {
                MessageBox.Show(excep.Message + Environment.NewLine + "İnfo Yenilənmədi.");
            }

             try {  KreditNomresiRefresh();}
             catch (System.Exception excep)
             {
                 MessageBox.Show(excep.Message + Environment.NewLine + "Kredit Nomreleri Yenilənmədi.");
             }
           
        }

        private void btYaddaSaxla_Click(object sender, EventArgs e)
        {
            MyCheck.ParolYoxla();
            if (!MyCheck.Parolicaze) return;

            if (!MyCheck.davamYesNo()) return;

            txtKreditNomresi.BackColor = Color.White;
                txtAyliqOdenis.BackColor = Color.White;
                txtMuddet.BackColor = Color.White;
                txtQaliq.BackColor = Color.White;
                txtİlkinOdenis.BackColor = Color.White;
                txtMebleg.BackColor = Color.White;

                if (txtKreditNomresi.Text == "") { MessageBox.Show("Kredit nömrəsi Qeyd olunmayıb"); txtKreditNomresi.BackColor = Color.Red; return; }
                if (txtMebleg.Text == "") { MessageBox.Show("Ümumi məbləğ qeyd olunmayıb."); txtMebleg.BackColor = Color.Red; return; }
                if (txtİlkinOdenis.Text == "") { MessageBox.Show("İlkin ödəniş qeyd olunmayıb."); txtİlkinOdenis.BackColor = Color.Red; return; }
                if (txtQaliq.Text == "") { MessageBox.Show("Qalıq qeyd olunmayıb."); txtQaliq.BackColor = Color.Red; return; }
                if (txtMuddet.Text == "") { MessageBox.Show("Kreditin Müddəti qeyd olunmayıb."); txtMuddet.BackColor = Color.Red; return; }
                if (txtAyliqOdenis.Text == "") { MessageBox.Show("Aylıq ödəniş qeyd olunmayıb."); txtAyliqOdenis.BackColor = Color.Red; return; }

            MyData.selectCommand("Arxiv\\baza.accdb", "Select a5 From Kredit where a5='" + txtKreditNomresi.Text + "'");
                MyData.dtmainKredit=new DataTable();
                MyData.oledbadapter1.Fill(MyData.dtmainKredit);

                if (MyData.dtmainKredit.Rows.Count > 0) { MessageBox.Show("Kredit nömrəsi artıq mövcuddur."); return; }

                //if (txtFaiz.Text == "0") { txtFaiz.ForeColor = Color.Red; return; } else { txtFaiz.ForeColor = Color.Black; }

           try
           {
                MyData.insertCommand("Arxiv\\baza.accdb", "INSERT INTO Kredit (a1,a2,a3,a4,a5,a6,a7,a8,a9,a10,a11,a12)values("



                                                                                                        + "'" + dtTarix.Value.ToShortDateString() + "',"
                                                                                                    + "'" + txtadi.Text + "',"
                                                                                                    + "'" + txtunvan.Text + "',"
                                                                                                    + "'" + txtTel.Text + "',"
                                                                                                    + "'" + txtKreditNomresi.Text + "',"
                                                                                                    + "'',"
                                                                                                    + "'" + txtMebleg.Text + "',"
                                                                                                    + "'" + txtİlkinOdenis.Text + "',"
                                                                                                    + "'" + txtAyliqOdenis.Text + "',"
                                                                                                    + "'" + txtQaliq.Text + "',"
                                                                                                    + "'" + txtQeyd.Text + "',"
                                                                                                    + "'" + txtMuddet.Text + "')");


                FreeAdd();
                KreditRefresh();
                KreditNomresiRefresh();
                QrafikRefresh();
            }
           catch (System.Exception excep)
            {
               MessageBox.Show(excep.Message + Environment.NewLine + " Əməliyyat baş tutmadı ");
            }
        }

        private void DataKredit_DoubleClick(object sender, EventArgs e)
        {
            tabControl1.SelectedTab = tabPage3;
            QrafikRefresh();
        }

        private void DataQrafik_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            DataQrafik.EditMode = DataGridViewEditMode.EditProgrammatically;

            try
            {
                MyData.updateCommand("Qrafik\\" + lbKreditNomresiQrafik.Text + ".accdb", "UPDATE Qrafik SET "
                                                                                     + "Tarix ='" + DataQrafik.Rows[DataQrafik.CurrentCell.RowIndex].Cells["q1"].Value.ToString() + "',"
                                                                                     + "Ödəniş ='" + DataQrafik.Rows[DataQrafik.CurrentCell.RowIndex].Cells["q2"].Value.ToString() + "',"
                                                                                     + "Faiz ='" + DataQrafik.Rows[DataQrafik.CurrentCell.RowIndex].Cells["q3"].Value.ToString() + "',"
                                                                                     + "Əsas ='" + DataQrafik.Rows[DataQrafik.CurrentCell.RowIndex].Cells["q4"].Value.ToString() + "',"
                                                                                     + "SSuda ='" + DataQrafik.Rows[DataQrafik.CurrentCell.RowIndex].Cells["q5"].Value.ToString() + "'"
                                                                                     + " WHERE id Like  '" + DataQrafik.Rows[DataQrafik.CurrentCell.RowIndex].Cells["id3"].Value.ToString() + "'");
            }
            catch { MessageBox.Show("Əməliyyat baş tutmadı."); }
        }

        private void toolStripMenuItem1_Click(object sender, EventArgs e)
        {
            MyCheck.ParolYoxla();
            if (!MyCheck.Parolicaze) return;
            if (MyCheck.ParolAdminYesNo()) return;

            DataQrafik.EditMode = DataGridViewEditMode.EditOnEnter;
        }

        private void DataKredit_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            DataKredit.EditMode = DataGridViewEditMode.EditProgrammatically;
            try
            {
                MyData.updateCommand("Arxiv\\baza.accdb", "UPDATE Kredit SET "
                                                                                       + "a1 ='" + DataKredit.Rows[DataKredit.CurrentCell.RowIndex].Cells["a1"].Value.ToString() + "',"
                                                                                     + "a2 ='" + DataKredit.Rows[DataKredit.CurrentCell.RowIndex].Cells["a2"].Value.ToString() + "',"
                                                                                     + "a3 ='" + DataKredit.Rows[DataKredit.CurrentCell.RowIndex].Cells["a3"].Value.ToString() + "',"
                                                                                     + "a4 ='" + DataKredit.Rows[DataKredit.CurrentCell.RowIndex].Cells["a4"].Value.ToString() + "',"
                                                                                     + "a5 ='" + DataKredit.Rows[DataKredit.CurrentCell.RowIndex].Cells["a5"].Value.ToString() + "',"
                                                                                     + "a6 ='" + DataKredit.Rows[DataKredit.CurrentCell.RowIndex].Cells["a6"].Value.ToString() + "',"
                                                                                     + "a7 ='" + DataKredit.Rows[DataKredit.CurrentCell.RowIndex].Cells["a7"].Value.ToString() + "',"
                                                                                     + "a8 ='" + DataKredit.Rows[DataKredit.CurrentCell.RowIndex].Cells["a8"].Value.ToString() + "',"
                                                                                     + "a9 ='" + DataKredit.Rows[DataKredit.CurrentCell.RowIndex].Cells["a9"].Value.ToString() + "',"
                                                                                     + "a10 ='" + DataKredit.Rows[DataKredit.CurrentCell.RowIndex].Cells["a10"].Value.ToString() + "',"
                                                                                     + "a11 ='" + DataKredit.Rows[DataKredit.CurrentCell.RowIndex].Cells["a11"].Value.ToString() + "',"
                                                                                     + "a12 ='" + DataKredit.Rows[DataKredit.CurrentCell.RowIndex].Cells["a12"].Value.ToString() + "'"
                                                                                     + " WHERE id Like '" + DataKredit.Rows[DataKredit.CurrentCell.RowIndex].Cells["id2"].Value.ToString() + "'");
            }
            catch { MessageBox.Show("Əməliyyat baş tutmadı."); }

            if (DataKredit.Columns[DataKredit.CurrentCell.ColumnIndex].Name == "a7" || DataKredit.Columns[DataKredit.CurrentCell.ColumnIndex].Name == "a8" || DataKredit.Columns[DataKredit.CurrentCell.ColumnIndex].Name == "a9" || DataKredit.Columns[DataKredit.CurrentCell.ColumnIndex].Name == "a12")
            {
                QrafikAvtoYarat(); 
                KreditRefresh();

                MessageBox.Show("Seçilmiş müştərinin qrafiki yeniləndi.");
            }
        }

        private void toolStripMenuItem2_Click(object sender, EventArgs e)
        {
            MyCheck.ParolYoxla();
            if (!MyCheck.Parolicaze) return;
            if (!MyCheck.ParolAdminYesNo()) return;

            DataKredit.EditMode = DataGridViewEditMode.EditOnEnter;
        }

        private void toolStripMenuItem11_Click(object sender, EventArgs e)
        {
            MyCheck.ParolYoxla();
            if (!MyCheck.Parolicaze) return;
            if (!MyCheck.ParolAdminYesNo()) return;
            if (!MyCheck.davamYesNo()) return;

            try
            {
                MyData.deleteCommand($"Qrafik\\{lbKreditNomresiQrafik.Text}.accdb", $"DELETE FROM Qrafik WHERE id Like  '{DataQrafik.Rows[DataQrafik.CurrentCell.RowIndex].Cells["id3"].Value.ToString()}'");

                QrafikRefresh();
            }
            catch { };
        }

        private void toolStripMenuItem3_Click(object sender, EventArgs e)
        {
            MyCheck.ParolYoxla();
            if (!MyCheck.Parolicaze) return;
            if (!MyCheck.ParolAdminYesNo()) return;
            if (!MyCheck.davamYesNo()) return;

            try
            {
                MyData.deleteCommand("Arxiv\\baza.accdb", $"DELETE FROM Kredit WHERE id Like  '{DataKredit.Rows[DataKredit.CurrentCell.RowIndex].Cells["id2"].Value.ToString()}'");

                KreditRefresh();
                KreditNomresiRefresh();
            }
            catch { };


          
        }

        private void txtMebleg_TextChanged(object sender, EventArgs e)
        {

            try
            {
                txtQaliq.Text = (Convert.ToDouble(txtMebleg.Text) - Convert.ToDouble(txtİlkinOdenis.Text)).ToString();
            }
            catch { }
          
        }

        private void txtİlkinOdenis_TextChanged(object sender, EventArgs e)
        {
            try
            {
                txtQaliq.Text = (Convert.ToDouble(txtMebleg.Text) - Convert.ToDouble(txtİlkinOdenis.Text)).ToString();
            }
            catch { }
        }

        private void label7_Click(object sender, EventArgs e)
        {
            if (txtQaliq.Enabled == false) { txtQaliq.Enabled = true; return; }
            if (txtQaliq.Enabled == true) { txtQaliq.Enabled = false; return; }
        }

        private void label16_Click(object sender, EventArgs e)
        {
            if (txtQaliqKassa.Enabled == false) { txtQaliqKassa.Enabled = true; return; }
            if (txtQaliqKassa.Enabled == true) { txtQaliqKassa.Enabled = false; return; }
        }

        private void txtAxtarKredit_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            { KreditRefresh(); }
        }

        private void txtAxtarKassa_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            { KassaRefresh(); }
        }

        private void toolStripMenuItem4_Click(object sender, EventArgs e)
        {
            MyCheck.ParolYoxla();
            if (!MyCheck.Parolicaze) return;
            if (!MyCheck.ParolAdminYesNo()) return;

            dataKassa.EditMode = DataGridViewEditMode.EditOnEnter;
        }

        private void toolStripMenuItem5_Click(object sender, EventArgs e)
        {
            MyCheck.ParolYoxla();
            if (!MyCheck.Parolicaze) return;
            if (!MyCheck.ParolAdminYesNo()) return;
            if (!MyCheck.davamYesNo()) return;

            try
            {
                MyData.insertCommand("Arxiv\\baza.accdb", "DELETE FROM Kassa WHERE id Like  '" + dataKassa.Rows[dataKassa.CurrentCell.RowIndex].Cells["idKassa"].Value.ToString() + "'");

                KassaRefresh();
            }
            catch { };
        }

        private void dataKassa_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            dataKassa.EditMode = DataGridViewEditMode.EditProgrammatically;

            try
            {
                MyData.updateCommand("Arxiv\\baza.accdb", "UPDATE Kassa SET "
                                                                                     + "kassa1 ='" + dataKassa.Rows[dataKassa.CurrentCell.RowIndex].Cells["kassa1"].Value.ToString() + "',"
                                                                                     + "kassa2 ='" + dataKassa.Rows[dataKassa.CurrentCell.RowIndex].Cells["kassa2"].Value.ToString() + "',"
                                                                                     + "kassa3 ='" + dataKassa.Rows[dataKassa.CurrentCell.RowIndex].Cells["kassa3"].Value.ToString() + "',"
                                                                                     + "kassa4 ='" + Math.Round(Convert.ToDouble(dataKassa.Rows[dataKassa.CurrentCell.RowIndex].Cells["kassa4"].Value), 2, MidpointRounding.AwayFromZero).ToString() + "',"
                                                                                     + "kassa5 ='" + Math.Round(Convert.ToDouble(dataKassa.Rows[dataKassa.CurrentCell.RowIndex].Cells["kassa5"].Value), 2, MidpointRounding.AwayFromZero).ToString() + "',"
                                                                                     + "kassa6 ='" + Math.Round(Convert.ToDouble(dataKassa.Rows[dataKassa.CurrentCell.RowIndex].Cells["kassa6"].Value), 2, MidpointRounding.AwayFromZero).ToString() + "'"
                                                                                     + " WHERE id Like  '" + dataKassa.Rows[dataKassa.CurrentCell.RowIndex].Cells["idKassa"].Value.ToString() + "'");
            }
            catch { MessageBox.Show("Əməliyyat baş tutmadı."); }
        }

        private void cbKreditNomresi_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                string commandText = "SELECT * FROM Kredit where 1=1";
                commandText += $" and a5 Like '%{cbKreditNomresi.Text}%'";
                commandText += $" or a2 Like '%{cbKreditNomresi.Text}%'";
                MyData.selectCommand("Arxiv\\baza.accdb", commandText);
                MyData.dtmainKredit = new DataTable();
                MyData.oledbadapter1.Fill(MyData.dtmainKredit);
                try
                {
                    cbKreditNomresi.Text = $"{MyData.dtmainKredit.Rows[0]["a5"].ToString()} {MyData.dtmainKredit.Rows[0]["a2"].ToString()}";
                }
                catch { }

                qaliqmusterikassa();
            }
        }

        private void txtStatus_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                MyData.selectCommand("Arxiv\\baza.accdb", "Select * From Kassa WHERE kassa2 Like '%" + txtStatus.Text + "%'");
                MyData.dtmainKassa=new DataTable();
                MyData.oledbadapter1.Fill(MyData.dtmainKassa);

                try
                {
                    txtStatus.Text = MyData.dtmainKassa.Rows[0]["kassa2"].ToString();

                }
                catch { };
            }
        }

        private void txtAciqlama_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                MyData.selectCommand("Arxiv\\baza.accdb", "Select * From Kassa WHERE kassa3 Like '%" + txtAciqlama.Text + "%'");
                MyData.dtmainKassa=new DataTable();
                MyData.oledbadapter1.Fill(MyData.dtmainKassa);

                try
                {
                    txtAciqlama.Text = MyData.dtmainKassa.Rows[0]["kassa3"].ToString();

                }
                catch { };
            }
        }

        private void ödənişEtToolStripMenuItem_Click(object sender, EventArgs e)
        {
            tabControl1.SelectedTab = tabPage6;
            cbEmeliyyatNovu.Text = "Kredit ödənişi";
            cbKreditNomresi.Text = DataKredit.Rows[DataKredit.CurrentCell.RowIndex].Cells["a5"].Value.ToString() + " " + DataKredit.Rows[DataKredit.CurrentCell.RowIndex].Cells["a2"].Value.ToString();
            KassaRefresh();
        }

        private void txtQaliq_TextChanged(object sender, EventArgs e)
        {
            try 
            {
                txtAyliqOdenis.Text = Math.Round((Convert.ToDouble(txtQaliq.Text) / Convert.ToDouble(txtMuddet.Text)), 2).ToString();
            }
            catch { }
        }

        private void txtMuddet_TextChanged(object sender, EventArgs e)
        {
            try 
            {
                txtAyliqOdenis.Text = Math.Round((Convert.ToDouble(txtQaliq.Text) / Convert.ToDouble(txtMuddet.Text)),2).ToString();
            }
            catch { }
        }

        private void button11_Click(object sender, EventArgs e)
        {
            tabControl1.SelectedTab = tabPage6;
            cbEmeliyyatNovu.Text = "Kredit ödənişi";
            cbKreditNomresi.Text = DataKredit.Rows[DataKredit.CurrentCell.RowIndex].Cells["a5"].Value.ToString() + " " + DataKredit.Rows[DataKredit.CurrentCell.RowIndex].Cells["a2"].Value.ToString();
            KassaRefresh();
        }

        private void button12_Click(object sender, EventArgs e)
        {
            MyCheck.ParolYoxla();
            if (!MyCheck.Parolicaze) return;
            if (!MyCheck.davamYesNo()) return;
            try
            {
                MyData.insertCommand("Arxiv\\baza.accdb", "INSERT INTO Sifarisler (s1,s2,s3,s4,s5,s6,s7,s8,s9,s10,s11,s12,s13)values("



                                                                                                    + "'" + txtAdiSifaris.Text + "',"
                                                                                                    + "'" + txtUnvanSifaris.Text + "',"
                                                                                                    + "'" + txtTelefonSifaris.Text + "',"
                                                                                                    + "'" + txtSatisSifaris.Text + "',"
                                                                                                    + "'" + txtOlcuSifaris.Text + "',"
                                                                                                    + "'" + txtTikisSifaris.Text + "',"
                                                                                                    + "'" + txtTehTarSidaris.Text + "',"
                                                                                                    + "'" + txtAZNSifaris.Text + "',"
                                                                                                    + "'" + txtAcotSifaris.Text + "',"
                                                                                                    + "'" + txtNagdKreditSifaris.Text + "',"
                                                                                                    + "'" + txtProSifaris.Text + "',"
                                                                                                    + "'" + txtQeydSifaris.Text + "',"
                                                                                                    + "'" + txtQeydlerSifaris.Text + "')");

                SifarislerRefresh();
            }
            catch (System.Exception excep)
            {
                MessageBox.Show(excep.Message + Environment.NewLine + " Əməliyyat baş tutmadı ");
            }
        }

        private void txtAxtarSifarisler_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                SifarislerRefresh();
            }
        }

        private void toolStripMenuItem6_Click(object sender, EventArgs e)
        {
            MyCheck.ParolYoxla();
            if (!MyCheck.Parolicaze) return;
            if (!MyCheck.ParolAdminYesNo()) return;

            dataSifarisler.EditMode = DataGridViewEditMode.EditOnEnter;
        }

        private void dataSifarisler_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            dataSifarisler.EditMode = DataGridViewEditMode.EditProgrammatically;

            try
            {
                MyData.updateCommand("Arxiv\\baza.accdb", "UPDATE Sifarisler SET "
                                                                                 + "s1 ='" + dataSifarisler.Rows[dataSifarisler.CurrentCell.RowIndex].Cells["ss1"].Value.ToString() + "',"
                                                                                 + "s2 ='" + dataSifarisler.Rows[dataSifarisler.CurrentCell.RowIndex].Cells["ss2"].Value.ToString() + "',"
                                                                                 + "s3 ='" + dataSifarisler.Rows[dataSifarisler.CurrentCell.RowIndex].Cells["ss3"].Value.ToString() + "',"
                                                                                 + "s4 ='" + dataSifarisler.Rows[dataSifarisler.CurrentCell.RowIndex].Cells["ss4"].Value.ToString() + "',"
                                                                                 + "s5 ='" + dataSifarisler.Rows[dataSifarisler.CurrentCell.RowIndex].Cells["ss5"].Value.ToString() + "',"
                                                                                 + "s6 ='" + dataSifarisler.Rows[dataSifarisler.CurrentCell.RowIndex].Cells["ss6"].Value.ToString() + "',"
                                                                                 + "s7 ='" + dataSifarisler.Rows[dataSifarisler.CurrentCell.RowIndex].Cells["ss7"].Value.ToString() + "',"
                                                                                 + "s8 ='" + dataSifarisler.Rows[dataSifarisler.CurrentCell.RowIndex].Cells["ss8"].Value.ToString() + "',"
                                                                                 + "s9 ='" + dataSifarisler.Rows[dataSifarisler.CurrentCell.RowIndex].Cells["ss9"].Value.ToString() + "',"
                                                                                 + "s10 ='" + dataSifarisler.Rows[dataSifarisler.CurrentCell.RowIndex].Cells["ss10"].Value.ToString() + "',"
                                                                                 + "s11 ='" + dataSifarisler.Rows[dataSifarisler.CurrentCell.RowIndex].Cells["ss11"].Value.ToString() + "',"
                                                                                 + "s12 ='" + dataSifarisler.Rows[dataSifarisler.CurrentCell.RowIndex].Cells["ss12"].Value.ToString() + "',"
                                                                                 + "s13 ='" + dataSifarisler.Rows[dataSifarisler.CurrentCell.RowIndex].Cells["ss13"].Value.ToString() + "'"
                                                                                 + " WHERE id Like  '" + dataSifarisler.Rows[dataSifarisler.CurrentCell.RowIndex].Cells["idSifarisler"].Value.ToString() + "'");
            }
            catch { MessageBox.Show("Əməliyyat baş tutmadı."); }
        }

        private void toolStripMenuItem7_Click(object sender, EventArgs e)
        {
            MyCheck.ParolYoxla();
            if (!MyCheck.Parolicaze) return;
            if (!MyCheck.ParolAdminYesNo()) return;
            if (!MyCheck.davamYesNo()) return;

            try
            {
                MyData.selectCommand("Arxiv\\baza.accdb", "DELETE FROM Sifarisler WHERE id Like  '" + dataSifarisler.Rows[dataSifarisler.CurrentCell.RowIndex].Cells["idSifarisler"].Value.ToString() + "'");

                SifarislerRefresh();
            }
            catch { };
        }

        private void PrintSifarisler()
        {
            int s = 0, k = 0;

            try { File.Copy("Excel\\Empty.xlsx", "C:\\Users\\" + Environment.UserName + "\\Desktop\\Sifarisler.xlsx", true); }
            catch { MessageBox.Show("'Sifarisler.xlsx' tapılmadı."); }

            //Get a new workbook.
            oXL = new Excel.Application();
            oWB = (Excel._Workbook)(oXL.Workbooks.Open("C:\\Users\\" + Environment.UserName + "\\Desktop\\Sifarisler.xlsx"));
            oSheet = (Excel._Worksheet)oWB.Sheets["Sheet1"];

            oSheet.Visible = Excel.XlSheetVisibility.xlSheetVisible;

            oXL.Visible = true;
            oSheet.Activate();
            oSheet.Range["A1"].Select();
            oSheet.PageSetup.Orientation = Excel.XlPageOrientation.xlLandscape;

            try
            {
            for (k = 0; k < dataSifarisler.Rows.Count; k++)
            {

                for (s = 0; s < dataSifarisler.Columns.Count; s++)
                {
                    oSheet.Cells[1, s + 1] = dataSifarisler.Columns[s].HeaderText;
                    oSheet.Cells[k + 2, s + 1] = dataSifarisler.Rows[k].Cells[s].Value.ToString();
                    oSheet.Cells[k + 2, 1] = k + 1;
                }

                oSheet.Range["A" + (k + 2)].Borders.LineStyle = Excel.Constants.xlSolid;
                oSheet.Range["B" + (k + 2)].Borders.LineStyle = Excel.Constants.xlSolid;
                oSheet.Range["C" + (k + 2)].Borders.LineStyle = Excel.Constants.xlSolid;
                oSheet.Range["D" + (k + 2)].Borders.LineStyle = Excel.Constants.xlSolid;
                oSheet.Range["E" + (k + 2)].Borders.LineStyle = Excel.Constants.xlSolid;
                oSheet.Range["F" + (k + 2)].Borders.LineStyle = Excel.Constants.xlSolid;
                oSheet.Range["G" + (k + 2)].Borders.LineStyle = Excel.Constants.xlSolid;
                oSheet.Range["H" + (k + 2)].Borders.LineStyle = Excel.Constants.xlSolid;
                oSheet.Range["I" + (k + 2)].Borders.LineStyle = Excel.Constants.xlSolid;
                oSheet.Range["J" + (k + 2)].Borders.LineStyle = Excel.Constants.xlSolid;
                oSheet.Range["K" + (k + 2)].Borders.LineStyle = Excel.Constants.xlSolid;
                oSheet.Range["L" + (k + 2)].Borders.LineStyle = Excel.Constants.xlSolid;
                oSheet.Range["M" + (k + 2)].Borders.LineStyle = Excel.Constants.xlSolid;
                oSheet.Range["N" + (k + 2)].Borders.LineStyle = Excel.Constants.xlSolid;

            }
            oSheet.Range["A" + 1].Borders.LineStyle = Excel.Constants.xlSolid;
            oSheet.Range["B" + 1].Borders.LineStyle = Excel.Constants.xlSolid;
            oSheet.Range["C" + 1].Borders.LineStyle = Excel.Constants.xlSolid;
            oSheet.Range["D" + 1].Borders.LineStyle = Excel.Constants.xlSolid;
            oSheet.Range["E" + 1].Borders.LineStyle = Excel.Constants.xlSolid;
            oSheet.Range["F" + 1].Borders.LineStyle = Excel.Constants.xlSolid;
            oSheet.Range["G" + 1].Borders.LineStyle = Excel.Constants.xlSolid;
            oSheet.Range["H" + 1].Borders.LineStyle = Excel.Constants.xlSolid;
            oSheet.Range["I" + 1].Borders.LineStyle = Excel.Constants.xlSolid;
            oSheet.Range["J" + 1].Borders.LineStyle = Excel.Constants.xlSolid;
            oSheet.Range["K" + 1].Borders.LineStyle = Excel.Constants.xlSolid;
            oSheet.Range["L" + 1].Borders.LineStyle = Excel.Constants.xlSolid;
            oSheet.Range["M" + 1].Borders.LineStyle = Excel.Constants.xlSolid;
            oSheet.Range["N" + 1].Borders.LineStyle = Excel.Constants.xlSolid;
            }
            catch (System.Exception excep)
            {
            MessageBox.Show(excep.Message + " Məlumatlar Excel-ə doldurularkən, Excel-ə toxunmaq olmaz !!! ");
            }

            oSheet.Columns.AutoFit();
            oSheet.Rows.AutoFit();

            oXL.Visible = true;
            oSheet.Activate();
            oSheet.Range["A1"].Select();
            oSheet.PageSetup.Orientation = Excel.XlPageOrientation.xlPortrait;
            //oSheet.PrintOut();
            //oWB.Close(SaveChanges: true);
            //oXL.Application.Quit();
    
           
        }

        private void PrintQrafik()
        {
            int s = 0, k = 0;
            Excel.Application oXL;
            Excel._Workbook oWB;
            Excel._Worksheet oSheet;

            try { File.Copy("Excel\\Empty.xlsx", "C:\\Users\\" + Environment.UserName + "\\Desktop\\Qrafik.xlsx", true); }
            catch { MessageBox.Show("'Qrafik.xlsx' tapılmadı."); }

            //Get a new workbook.
            oXL = new Excel.Application();
            oWB = (Excel._Workbook)(oXL.Workbooks.Open("C:\\Users\\" + Environment.UserName + "\\Desktop\\Qrafik.xlsx"));
            oSheet = (Excel._Worksheet)oWB.Sheets["Sheet1"];

            oSheet.Visible = Excel.XlSheetVisibility.xlSheetVisible;

            oXL.Visible = true;
            oSheet.Activate();
            oSheet.Range["A1"].Select();
            oSheet.PageSetup.Orientation = Excel.XlPageOrientation.xlLandscape;

            try
            {
                oSheet.Cells[1, 1] = "Adı: " + lbAd.Text + " (" + lbKreditNomresiQrafik.Text + ")";
                oSheet.Cells[2, 1] = "";
                oSheet.Cells[3, 1] = "Ümumi məbləğ: " + txtMeblegQrafik.Text;
                oSheet.Cells[4, 1] = "Avans: " + txtAvansQrafik.Text;
                oSheet.Cells[5, 1] = "Kredit məbləği: " + txtAvansQrafik.Text;
                oSheet.Cells[6, 1] = "Kreditin müddəti: " + DataQrafik.Rows.Count + " ay";
                oSheet.Cells[7, 1] = "";
                oSheet.Cells[DataQrafik.Rows.Count + 10, 1] = "Kredit üzrə ödənilib: " + txtOdenilibQrafik.Text;
                oSheet.Cells[DataQrafik.Rows.Count + 11, 1] = "Ümumi borc: " + txtQaliqQrafik.Text;

            oSheet.Range[oSheet.Cells[1, 1], oSheet.Cells[1, 6]].Merge();
            oSheet.Range[oSheet.Cells[2, 1], oSheet.Cells[2, 6]].Merge();
            oSheet.Range[oSheet.Cells[3, 1], oSheet.Cells[3, 6]].Merge();
            oSheet.Range[oSheet.Cells[4, 1], oSheet.Cells[4, 6]].Merge();
            oSheet.Range[oSheet.Cells[5, 1], oSheet.Cells[5, 6]].Merge();
            oSheet.Range[oSheet.Cells[6, 1], oSheet.Cells[6, 6]].Merge();
            oSheet.Range[oSheet.Cells[7, 1], oSheet.Cells[7, 6]].Merge();
            oSheet.Range[oSheet.Cells[DataQrafik.Rows.Count + 10, 1], oSheet.Cells[DataQrafik.Rows.Count + 10, 6]].Merge();
            oSheet.Range[oSheet.Cells[DataQrafik.Rows.Count + 11, 1], oSheet.Cells[DataQrafik.Rows.Count + 11, 6]].Merge();

            for (k = 0; k < DataQrafik.Rows.Count; k++)
            {

                for (s = 0; s < DataQrafik.Columns.Count; s++)
                {
                    oSheet.Cells[8, s + 1] = DataQrafik.Columns[s].HeaderText;
                    oSheet.Cells[k + 9, s + 1] = DataQrafik.Rows[k].Cells[s].Value.ToString();
                    oSheet.Cells[k + 9, 1] = k + 1;
                }

                oSheet.Range["A" + (k + 9)].Borders.LineStyle = Excel.Constants.xlSolid;
                oSheet.Range["B" + (k + 9)].Borders.LineStyle = Excel.Constants.xlSolid;
                oSheet.Range["C" + (k + 9)].Borders.LineStyle = Excel.Constants.xlSolid;
                oSheet.Range["D" + (k + 9)].Borders.LineStyle = Excel.Constants.xlSolid;
                oSheet.Range["E" + (k + 9)].Borders.LineStyle = Excel.Constants.xlSolid;
                oSheet.Range["F" + (k + 9)].Borders.LineStyle = Excel.Constants.xlSolid;

            }
            oSheet.Range["A" + 8].Borders.LineStyle = Excel.Constants.xlSolid;
            oSheet.Range["B" + 8].Borders.LineStyle = Excel.Constants.xlSolid;
            oSheet.Range["C" + 8].Borders.LineStyle = Excel.Constants.xlSolid;
            oSheet.Range["D" + 8].Borders.LineStyle = Excel.Constants.xlSolid;
            oSheet.Range["E" + 8].Borders.LineStyle = Excel.Constants.xlSolid;
            oSheet.Range["F" + 8].Borders.LineStyle = Excel.Constants.xlSolid;

            }
            catch (System.Exception excep)
            {
            MessageBox.Show(excep.Message + " Məlumatlar Excel-ə doldurularkən, Excel-ə toxunmaq olmaz !!! ");
            }

            oSheet.Columns.AutoFit();
            oSheet.Rows.AutoFit();

            oXL.Visible = true;
            oSheet.Activate();
            oSheet.Range["A1"].Select();
            oSheet.PageSetup.Orientation = Excel.XlPageOrientation.xlPortrait;
            //oSheet.PrintOut();
            //oWB.Close(SaveChanges: true);
            //oXL.Application.Quit();
    
        }

        private void PrintKredit()
        {
            int s = 0, k = 0;
            Excel.Application oXL;
            Excel._Workbook oWB;
            Excel._Worksheet oSheet;

            try { File.Copy("Excel\\Empty.xlsx", "C:\\Users\\" + Environment.UserName + "\\Desktop\\Kredit.xlsx", true); }
            catch { MessageBox.Show("'Kredit.xlsx' tapılmadı."); }

            //Get a new workbook.
            oXL = new Excel.Application();
            oWB = (Excel._Workbook)(oXL.Workbooks.Open("C:\\Users\\" + Environment.UserName + "\\Desktop\\Kredit.xlsx"));
            oSheet = (Excel._Worksheet)oWB.Sheets["Sheet1"];

            oSheet.Visible = Excel.XlSheetVisibility.xlSheetVisible;

            oXL.Visible = true;
            oSheet.Activate();
            oSheet.Range["A1"].Select();
            oSheet.PageSetup.Orientation = Excel.XlPageOrientation.xlLandscape;

            try
            { 

            //oSheet.Range[oSheet.Cells[DataKredit.Rows.Count + 2, 10], oSheet.Cells[DataKredit.Rows.Count + 2, 12]].Merge();

            for (k = 0; k < DataKredit.Rows.Count; k++)
            {

                for (s = 0; s < DataKredit.Columns.Count; s++)
                {
                    oSheet.Cells[1, s + 1] = DataKredit.Columns[s].HeaderText;
                    oSheet.Cells[k + 2, s + 1] = DataKredit.Rows[k].Cells[s].Value.ToString();
                    oSheet.Cells[k + 2, 1] = k + 1;
                }

                oSheet.Range["A" + (k + 2)].Borders.LineStyle = Excel.Constants.xlSolid;
                oSheet.Range["B" + (k + 2)].Borders.LineStyle = Excel.Constants.xlSolid;
                oSheet.Range["C" + (k + 2)].Borders.LineStyle = Excel.Constants.xlSolid;
                oSheet.Range["D" + (k + 2)].Borders.LineStyle = Excel.Constants.xlSolid;
                oSheet.Range["E" + (k + 2)].Borders.LineStyle = Excel.Constants.xlSolid;
                oSheet.Range["F" + (k + 2)].Borders.LineStyle = Excel.Constants.xlSolid;
                oSheet.Range["G" + (k + 2)].Borders.LineStyle = Excel.Constants.xlSolid;
                oSheet.Range["H" + (k + 2)].Borders.LineStyle = Excel.Constants.xlSolid;
                oSheet.Range["I" + (k + 2)].Borders.LineStyle = Excel.Constants.xlSolid;
                oSheet.Range["J" + (k + 2)].Borders.LineStyle = Excel.Constants.xlSolid;
                oSheet.Range["K" + (k + 2)].Borders.LineStyle = Excel.Constants.xlSolid;
                oSheet.Range["L" + (k + 2)].Borders.LineStyle = Excel.Constants.xlSolid;
                oSheet.Range["M" + (k + 2)].Borders.LineStyle = Excel.Constants.xlSolid;

            }
            oSheet.Range["A" + 1].Borders.LineStyle = Excel.Constants.xlSolid;
            oSheet.Range["B" + 1].Borders.LineStyle = Excel.Constants.xlSolid;
            oSheet.Range["C" + 1].Borders.LineStyle = Excel.Constants.xlSolid;
            oSheet.Range["D" + 1].Borders.LineStyle = Excel.Constants.xlSolid;
            oSheet.Range["E" + 1].Borders.LineStyle = Excel.Constants.xlSolid;
            oSheet.Range["F" + 1].Borders.LineStyle = Excel.Constants.xlSolid;
            oSheet.Range["G" + 1].Borders.LineStyle = Excel.Constants.xlSolid;
            oSheet.Range["H" + 1].Borders.LineStyle = Excel.Constants.xlSolid;
            oSheet.Range["I" + 1].Borders.LineStyle = Excel.Constants.xlSolid;
            oSheet.Range["J" + 1].Borders.LineStyle = Excel.Constants.xlSolid;
            oSheet.Range["K" + 1].Borders.LineStyle = Excel.Constants.xlSolid;
            oSheet.Range["L" + 1].Borders.LineStyle = Excel.Constants.xlSolid;
            oSheet.Range["M" + 1].Borders.LineStyle = Excel.Constants.xlSolid;

            }
            catch (System.Exception excep)
            {
            MessageBox.Show(excep.Message + " Məlumatlar Excel-ə doldurularkən, Excel-ə toxunmaq olmaz !!! ");
            }

            oSheet.Columns.AutoFit();
            oSheet.Rows.AutoFit();

            oXL.Visible = true;
            oSheet.Activate();
            oSheet.Range["A1"].Select();
            oSheet.PageSetup.Orientation = Excel.XlPageOrientation.xlPortrait;
            //oSheet.PrintOut();
            //oWB.Close(SaveChanges: true);
            //oXL.Application.Quit();
        }

        private void PrintKassa()
        {
            int s = 0, k = 0;
            Excel.Application oXL;
            Excel._Workbook oWB;
            Excel._Worksheet oSheet;

            try { File.Copy("Excel\\Empty.xlsx", "C:\\Users\\" + Environment.UserName + "\\Desktop\\Kassa.xlsx", true); }
            catch { MessageBox.Show("'Kassa.xlsx' tapılmadı."); }

            //Get a new workbook.
            oXL = new Excel.Application();
            oWB = (Excel._Workbook)(oXL.Workbooks.Open("C:\\Users\\" + Environment.UserName + "\\Desktop\\Kassa.xlsx"));
            oSheet = (Excel._Worksheet)oWB.Sheets["Sheet1"];

            oSheet.Visible = Excel.XlSheetVisibility.xlSheetVisible;

            oXL.Visible = true;
            oSheet.Activate();
            oSheet.Range["A1"].Select();
            oSheet.PageSetup.Orientation = Excel.XlPageOrientation.xlLandscape;

            try
            { 

            //oSheet.Range[oSheet.Cells[DataKredit.Rows.Count + 2, 10], oSheet.Cells[DataKredit.Rows.Count + 2, 12]].Merge();

            for (k = 0; k < dataKassa.Rows.Count; k++)
            {

                for (s = 0; s < dataKassa.Columns.Count; s++)
                {
                    oSheet.Cells[1, s + 1] = dataKassa.Columns[s].HeaderText;
                    oSheet.Cells[k + 2, s + 1] = dataKassa.Rows[k].Cells[s].Value.ToString();
                    oSheet.Cells[k + 2, 1] = k + 1;
                }

                oSheet.Range["A" + (k + 2)].Borders.LineStyle = Excel.Constants.xlSolid;
                oSheet.Range["B" + (k + 2)].Borders.LineStyle = Excel.Constants.xlSolid;
                oSheet.Range["C" + (k + 2)].Borders.LineStyle = Excel.Constants.xlSolid;
                oSheet.Range["D" + (k + 2)].Borders.LineStyle = Excel.Constants.xlSolid;
                oSheet.Range["E" + (k + 2)].Borders.LineStyle = Excel.Constants.xlSolid;
                oSheet.Range["F" + (k + 2)].Borders.LineStyle = Excel.Constants.xlSolid;
                oSheet.Range["G" + (k + 2)].Borders.LineStyle = Excel.Constants.xlSolid;

            }
            oSheet.Range["A" + 1].Borders.LineStyle = Excel.Constants.xlSolid;
            oSheet.Range["B" + 1].Borders.LineStyle = Excel.Constants.xlSolid;
            oSheet.Range["C" + 1].Borders.LineStyle = Excel.Constants.xlSolid;
            oSheet.Range["D" + 1].Borders.LineStyle = Excel.Constants.xlSolid;
            oSheet.Range["E" + 1].Borders.LineStyle = Excel.Constants.xlSolid;
            oSheet.Range["F" + 1].Borders.LineStyle = Excel.Constants.xlSolid;
            oSheet.Range["G" + 1].Borders.LineStyle = Excel.Constants.xlSolid;

            }
            catch (System.Exception excep)
            {
            MessageBox.Show(excep.Message + " Məlumatlar Excel-ə doldurularkən, Excel-ə toxunmaq olmaz !!! ");
            }

            oSheet.Columns.AutoFit();
            oSheet.Rows.AutoFit();

            oXL.Visible = true;
            oSheet.Activate();
            oSheet.Range["A1"].Select();
            oSheet.PageSetup.Orientation = Excel.XlPageOrientation.xlPortrait;
            //oSheet.PrintOut();
            //oWB.Close(SaveChanges: true);
            //oXL.Application.Quit();
        }

        private void qaliqmusterikassa()
        {
            if (cbKreditNomresi.Text == "")
            {
                txtQaliqMusteriKassa.Text = "0"; 
                txtAciqlama.Text = "";
                return;
            }



            string um = "";
            try
            {
                for (int i = 0; i < cbKreditNomresi.Text.Length; i++)
                {
                    if (cbKreditNomresi.Text[i] == ' ')
                    {
                        um = cbKreditNomresi.Text.Substring(0, i);
                        break;
                    }
                }
            }
            catch { }

            try
            {
                MyData.selectCommand("Arxiv\\baza.accdb", $"Select * From Kredit WHERE a5 Like '%{um}%'");
                MyData.dtmainKredit=new DataTable();
                MyData.oledbadapter1.Fill(MyData.dtmainKredit);

                txtQaliqMusteriKassa.Text = MyData.dtmainKredit.Rows[0]["a10"].ToString();
                txtAciqlama.Text = cbKreditNomresi.Text + " " + cbEmeliyyatNovu.Text;
            }
            catch { txtQaliqMusteriKassa.Text = "0"; txtAciqlama.Text = ""; }
        }

        private void cbKreditNomresi_TextChanged(object sender, EventArgs e)
        {
            qaliqmusterikassa();
        }

        private void txtGecikmeQrafik_TextChanged(object sender, EventArgs e)
        {

            try
            {
                if (Convert.ToDouble(txtGecikmeQrafik.Text.Substring(0, txtGecikmeQrafik.Text.Length - 4)) <= 0)
                {
                    txtGecikmeQrafik.Text = "0 AZN";
                    txtGecikmeQrafik.ForeColor = Color.Green;
                    lbGecikmeQrafik.ForeColor = Color.Green;
                }
                else
                {
                    txtGecikmeQrafik.ForeColor = Color.Red;
                    lbGecikmeQrafik.ForeColor = Color.Red;
                }
            }
            catch { }
        }

        private void button19_Click(object sender, EventArgs e)
        {
            if (cbEmeliyyatNovu.Text == "Kredit ödənişi")
            {
                KreditOdenisi();
            }

            if (cbEmeliyyatNovu.Text == "Avans ödənişi")
            {
                AvansOdenisi();
            }

            if (cbEmeliyyatNovu.Text == "Kassa əməliyyatı")
            {
                KassaEmeliyyati();
            }
        }

        private void button20_Click(object sender, EventArgs e)
        {
            if (tabControl1.SelectedTab == tabPage1) infoRefresh();

            if (tabControl1.SelectedTab == tabPage2) { KreditRefresh(); KreditNomresiRefresh(); }
            
            if (tabControl1.SelectedTab == tabPage5) SifarislerRefresh();
            
            if (tabControl1.SelectedTab == tabPage3)
            {
                QaliqlarQrafikRefresh();
                QrafikRefresh();
                QrafikOdenislerRefresh();
                
            }

            if (tabControl1.SelectedTab == tabPage6)
            {
                txtAxtarKassa.Text = "";
                KassaRefresh();
                KreditNomresiRefresh();
            }
        }

        private void btAxtarKredit_Click(object sender, EventArgs e)
        {
            KreditRefresh();
        }

        private void btAxtarSifarisler_Click(object sender, EventArgs e)
        {
            SifarislerRefresh();
        }

        private void btAxtarKassa_Click(object sender, EventArgs e)
        {
            KassaRefresh();
        }

        private void btAxtarStatus_Click(object sender, EventArgs e)
        {
            MyData.selectCommand("Arxiv\\baza.accdb", "Select * From Kassa WHERE kassa2 Like '%" + txtStatus.Text + "%'");
            MyData.dtmainKassa=new DataTable();
            MyData.oledbadapter1.Fill(MyData.dtmainKassa);

            try
            {
                txtStatus.Text = MyData.dtmainKassa.Rows[0]["kassa2"].ToString();

            }
            catch { };
        }

        private void btAxtarAciqlama_Click(object sender, EventArgs e)
        {
            MyData.selectCommand("Arxiv\\baza.accdb", "Select * From Kassa WHERE kassa3 Like '%" + txtAciqlama.Text + "%'");
            MyData.dtmainKassa=new DataTable();
            MyData.oledbadapter1.Fill(MyData.dtmainKassa);

            try
            {
                txtAciqlama.Text = MyData.dtmainKassa.Rows[0]["kassa3"].ToString();

            }
            catch { };
        }

        private void button1_Click(object sender, EventArgs e)
        {
            MyData.selectCommand("Arxiv\\baza.accdb", "SELECT * FROM Kredit where a5 Like '%" + cbKreditNomresi.Text + "%' or a2 Like '%" + cbKreditNomresi.Text + "%'");
            MyData.dtmainKredit=new DataTable();
            MyData.oledbadapter1.Fill(MyData.dtmainKredit);
            try { cbKreditNomresi.Text = MyData.dtmainKredit.Rows[0]["a5"].ToString() + " " + MyData.dtmainKredit.Rows[0]["a2"].ToString(); }
            catch { }
        }

        private void köçürToolStripMenuItem_Click(object sender, EventArgs e)
        {
            tabControl1.SelectedTab = tabPage4;

            try { dtTarix.Text = DataKredit.Rows[DataKredit.CurrentCell.RowIndex].Cells["a1"].Value.ToString(); }
            catch { }
            try { dtBirinciOdenis.Text = dtTarix.Value.AddMonths(1).ToShortDateString(); }
            catch { }
            try { txtadi.Text = DataKredit.Rows[DataKredit.CurrentCell.RowIndex].Cells["a2"].Value.ToString(); }
            catch { }
            try { txtunvan.Text = DataKredit.Rows[DataKredit.CurrentCell.RowIndex].Cells["a3"].Value.ToString(); }
            catch { }
            try { txtTel.Text = DataKredit.Rows[DataKredit.CurrentCell.RowIndex].Cells["a4"].Value.ToString(); }
            catch { }
            try { txtKreditNomresi.Text = DataKredit.Rows[DataKredit.CurrentCell.RowIndex].Cells["a5"].Value.ToString(); }
            catch { }
            try { txtMebleg.Text = DataKredit.Rows[DataKredit.CurrentCell.RowIndex].Cells["a7"].Value.ToString(); }
            catch { }
            try { txtİlkinOdenis.Text = DataKredit.Rows[DataKredit.CurrentCell.RowIndex].Cells["a8"].Value.ToString(); }
            catch { }
            try { txtMuddet.Text = DataKredit.Rows[DataKredit.CurrentCell.RowIndex].Cells["a12"].Value.ToString(); }
            catch { txtMuddet.Text = "8"; }
        }

        private void cbBaglanmisKreditler_CheckedChanged(object sender, EventArgs e)
        {
            KreditRefresh();
        }

        private void haqqındaToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            info info = new info();
            info.ShowDialog();
        }

        private void parolToolStripMenuItem_Click(object sender, EventArgs e)
        {
            YeniParol parol = new YeniParol();
            parol.ShowDialog();
        }

        private void resetQrafikToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            MyCheck.ParolYoxla();
            if (!MyCheck.Parolicaze) return;
            if (!MyCheck.ParolAdminYesNo()) return;
            if (!MyCheck.davamYesNo()) return;

            MyData.selectCommand("Arxiv\\baza.accdb", "Select * From Kredit");
            MyData.dtmainKredit=new DataTable();
            MyData.oledbadapter1.Fill(MyData.dtmainKredit);
            DataKredit.DataSource = MyData.dtmainKredit;

            progressBar2.Value = 0;
            progressBar2.Maximum = MyData.dtmainKredit.Rows.Count;
            progressBar2.Step = 1;

            for (int i = 0; i < DataKredit.Rows.Count; i++)
            {
                progressBar2.PerformStep();
                base.Text = progressBar2.Value * 100 / DataKredit.Rows.Count + "% Qrafik yaradılır #" + (DataKredit.Rows.Count - i) + " Qrafik qalır";

                try { dtTarix.Text = DataKredit.Rows[i].Cells["a1"].Value.ToString(); }
                catch { }
                try { dtBirinciOdenis.Text = dtTarix.Value.AddMonths(1).ToShortDateString(); }
                catch { }
                try { txtadi.Text = DataKredit.Rows[i].Cells["a2"].Value.ToString(); }
                catch { }
                try { txtunvan.Text = DataKredit.Rows[i].Cells["a3"].Value.ToString(); }
                catch { }
                try { txtTel.Text = DataKredit.Rows[i].Cells["a4"].Value.ToString(); }
                catch { }
                try { txtKreditNomresi.Text = DataKredit.Rows[i].Cells["a5"].Value.ToString(); }
                catch { }
                try { txtMebleg.Text = DataKredit.Rows[i].Cells["a7"].Value.ToString(); }
                catch { }
                try { txtİlkinOdenis.Text = DataKredit.Rows[i].Cells["a8"].Value.ToString(); }
                catch { }
                try { txtMuddet.Text = DataKredit.Rows[i].Cells["a12"].Value.ToString(); }
                catch { txtMuddet.Text = "8"; }
                try
                {
                    if (DataKredit.Rows[i].Cells["a12"].Value.ToString() == "") txtMuddet.Text = "8";
                    if (DataKredit.Rows[i].Cells["a12"].Value.ToString() == "0") txtMuddet.Text = "1";
                }
                catch { }


                try
                {
                    MyData.updateCommand("Arxiv\\baza.accdb", "UPDATE Kredit SET "
                                                                                         + "a1 ='" + DataKredit.Rows[i].Cells["a1"].Value.ToString() + "',"
                                                                                         + "a2 ='" + DataKredit.Rows[i].Cells["a2"].Value.ToString() + "',"
                                                                                         + "a3 ='" + DataKredit.Rows[i].Cells["a3"].Value.ToString() + "',"
                                                                                         + "a4 ='" + DataKredit.Rows[i].Cells["a4"].Value.ToString() + "',"
                                                                                         + "a5 ='" + DataKredit.Rows[i].Cells["a5"].Value.ToString() + "',"
                                                                                         + "a6 ='" + DataKredit.Rows[i].Cells["a6"].Value.ToString() + "',"
                                                                                         + "a7 ='" + DataKredit.Rows[i].Cells["a7"].Value.ToString() + "',"
                                                                                         + "a8 ='" + DataKredit.Rows[i].Cells["a8"].Value.ToString() + "',"
                                                                                         + "a9 ='" + txtAyliqOdenis.Text + "',"
                                                                                         + "a10 ='" + DataKredit.Rows[i].Cells["a10"].Value.ToString() + "',"
                                                                                         + "a11 ='" + DataKredit.Rows[i].Cells["a11"].Value.ToString() + "',"
                                                                                         + "a12 ='" + txtMuddet.Text + "'"
                                                                                         + " WHERE id Like '" + DataKredit.Rows[i].Cells["id2"].Value.ToString() + "'");
                }
                catch { }

                ///////////////////////////freeADD////////////////////////////////
                try
                {
                    File.Copy("Qrafik\\Qrafik.accdb", "Qrafik\\" + txtKreditNomresi.Text + ".accdb", true);
                }
                catch (System.Exception excep)
                {
                    MessageBox.Show(excep.Message + Environment.NewLine + " Qrafik.accdb tapılmadı ");
                }

                try
                {
                    DateTime dt = dtBirinciOdenis.Value;
                    string Time = dt.ToShortDateString();
                    double Odenis = 0, t = 0, m = 0, ssuda = 0, Faiz = 0, Esas = 0, Faiz2 = 0, Esas2 = 0;
                    m = Convert.ToDouble(txtMebleg.Text) - Convert.ToDouble(txtİlkinOdenis.Text);
                    if (txtMebleg.Text == "0") m = 0.00000001;
                    t = Convert.ToDouble(txtMuddet.Text);
                    if (txtMuddet.Text == "0") t = 1;
                    if (txtMuddet.Text == "") t = 8;
                    Odenis = Math.Round(Convert.ToDouble(txtAyliqOdenis.Text), 2, MidpointRounding.AwayFromZero);
                    ssuda = Convert.ToDouble(txtMebleg.Text) - Convert.ToDouble(txtİlkinOdenis.Text);
                    Faiz = 0;
                    Esas = ssuda - Faiz;

                    for (int a = 0; a < t; a++)
                    {
                        Faiz2 = ssuda * 0 / 1200;
                        Esas2 = Odenis - Faiz2;
                        //Son ssudanin/qaligin ayliq odenisle beraberlesdirilmesi ucun
                        if (a == t - 1)
                        {
                            Esas2 = ssuda;
                            Odenis = Esas2 + Faiz2;
                        } ///////////////////////////////////////////////////////////////
                        ssuda = ssuda - Esas2;

                        MyData.insertCommand("Qrafik\\" + txtKreditNomresi.Text + ".accdb'", "INSERT INTO Qrafik (Tarix, Ödəniş, Faiz, Əsas, Ssuda)values("


                                                                                                            + "'" + Time + "',"
                                                                                                            + "'" + Math.Round(Odenis, 2, MidpointRounding.AwayFromZero).ToString() + "',"
                                                                                                            + "'" + Math.Round(Faiz2, 2, MidpointRounding.AwayFromZero).ToString() + "',"
                                                                                                            + "'" + Math.Round(Esas2, 2, MidpointRounding.AwayFromZero).ToString() + "',"
                                                                                                            + "'" + Math.Round(ssuda, 2, MidpointRounding.AwayFromZero).ToString() + "')");

                        Time = dt.AddMonths(a + 1).ToShortDateString();
                    }
                }
                catch (System.Exception excep)
                {
                    MessageBox.Show(excep.Message + Environment.NewLine + " Əməliyyat baş tutmadı ");
                    return;
                }


            }

            base.Text = "Pərdə Evim";
        }

        private void infoToolStripMenuItem_Click(object sender, EventArgs e)
        {
            tabControl1.SelectedTab = tabPage1;
        }

        private void kreditToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            tabControl1.SelectedTab = tabPage2;
        }

        private void qrafikToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            tabControl1.SelectedTab = tabPage3;
        }

        private void yeniMüştəriToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            tabControl1.SelectedTab = tabPage4;
        }

        private void sifarişlərToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            tabControl1.SelectedTab = tabPage5;
        }

        private void kassaToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            tabControl1.SelectedTab = tabPage6;
        }

        private void yeniləToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            if (tabControl1.SelectedTab == tabPage1) infoRefresh();

            if (tabControl1.SelectedTab == tabPage2) KreditRefresh();

            if (tabControl1.SelectedTab == tabPage5) SifarislerRefresh();

            if (tabControl1.SelectedTab == tabPage3)
            {
                QaliqlarQrafikRefresh();
                QrafikRefresh();

            }

            if (tabControl1.SelectedTab == tabPage6)
            {
                MyCheck.ParolYoxla();
                if (!MyCheck.Parolicaze) return;
                if (!MyCheck.ParolAdminYesNo()) return;
                if (MyCheck.davamYesNo("Qalıqlar yenidən hesablansın?" + Environment.NewLine + "Qeyd: Bu Əməliyyat biraz vaxt aparacaq." + Environment.NewLine + "Davam etmək istəyirsiniz?"))
                {
                    QaliqlarRefresh();
                    txtAxtarKassa.Text = "";
                    KassaRefresh();
                }
                else
                {
                    txtAxtarKassa.Text = "";
                    KassaRefresh();
                }
            }
        }

        private void çıxışToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            base.Close();
        }

        private void btPrint_Click(object sender, EventArgs e)
        {
            DialogResult result2 = MessageBox.Show("Davam etmək istəyirsiniz?", "?", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2);
            if (result2 == DialogResult.No) return;

            if (tabControl1.SelectedTab == tabPage2) PrintKredit();

            if (tabControl1.SelectedTab == tabPage3) PrintQrafik();

            if (tabControl1.SelectedTab == tabPage5) PrintSifarisler();

            if (tabControl1.SelectedTab == tabPage6) PrintKassa();
        }

        private void çapEtToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (tabControl1.SelectedTab == tabPage2) PrintKredit();

            if (tabControl1.SelectedTab == tabPage3) PrintQrafik();

            if (tabControl1.SelectedTab == tabPage5) PrintSifarisler();

            if (tabControl1.SelectedTab == tabPage6) PrintKassa();
        }

        private void DataKredit_SelectionChanged(object sender, EventArgs e)
        {
            try { lbKreditNomresiQrafik.Text = DataKredit.Rows[DataKredit.CurrentCell.RowIndex].Cells["a5"].Value.ToString(); }
            catch { }
            try { lbAd.Text = DataKredit.Rows[DataKredit.CurrentCell.RowIndex].Cells["a2"].Value.ToString(); }
            catch { }
            try { txtMeblegQrafik.Text = DataKredit.Rows[DataKredit.CurrentCell.RowIndex].Cells["a7"].Value.ToString() + " AZN"; }
            catch { }
            try { txtAvansQrafik.Text = DataKredit.Rows[DataKredit.CurrentCell.RowIndex].Cells["a8"].Value.ToString() + " AZN"; }
            catch { }
            try { txtKreditQrafik.Text = Math.Round(Convert.ToDouble(DataKredit.Rows[DataKredit.CurrentCell.RowIndex].Cells["a7"].Value) - Convert.ToDouble(DataKredit.Rows[DataKredit.CurrentCell.RowIndex].Cells["a8"].Value), 2, MidpointRounding.AwayFromZero).ToString() + " AZN"; }
            catch { }
            try { lbAvansOdenilibQrafik.Text = "Avans (" + Math.Round(Convert.ToDouble(DataKredit.Rows[DataKredit.CurrentCell.RowIndex].Cells["a8"].Value) / Convert.ToDouble(DataKredit.Rows[DataKredit.CurrentCell.RowIndex].Cells["a7"].Value) * 100, 2, MidpointRounding.AwayFromZero).ToString() + " %)"; }
            catch { }
            try { lbKreditQrafik.Text = "Kredit məbləği (" + Math.Round(100 - Convert.ToDouble(DataKredit.Rows[DataKredit.CurrentCell.RowIndex].Cells["a8"].Value) / Convert.ToDouble(DataKredit.Rows[DataKredit.CurrentCell.RowIndex].Cells["a7"].Value) * 100, 2, MidpointRounding.AwayFromZero).ToString() + " %)"; }
            catch { }
            try { txtQaliqQrafik.Text = DataKredit.Rows[DataKredit.CurrentCell.RowIndex].Cells["a10"].Value.ToString() + " AZN"; }
            catch { }
            try { txtOdenilibQrafik.Text = Math.Round(Convert.ToDouble(DataKredit.Rows[DataKredit.CurrentCell.RowIndex].Cells["a7"].Value) - Convert.ToDouble(DataKredit.Rows[DataKredit.CurrentCell.RowIndex].Cells["a8"].Value) - Convert.ToDouble(DataKredit.Rows[DataKredit.CurrentCell.RowIndex].Cells["a10"].Value), 2).ToString() + " AZN"; }
            catch { }

            QrafikRefresh();
        }

        private void resetCurrentQrafikToolStripMenuItem_Click(object sender, EventArgs e)
        {
            MyCheck.ParolYoxla();
            if (!MyCheck.Parolicaze) return;
            if (!MyCheck.ParolAdminYesNo()) return;
            if (!MyCheck.davamYesNo()) return;

            tabControl1.SelectedTab = tabPage2;
            int i = DataKredit.CurrentCell.RowIndex;

            QrafikAvtoYarat();
            KreditRefresh();
            DataKredit.Rows[i].Selected = true;
        }

        private void lbKreditNomresiQrafik_TextChanged(object sender, EventArgs e)
        {
            QrafikOdenislerRefresh();
        }

        private void cbEmeliyyatNovu_TextChanged(object sender, EventArgs e)
        {
            if (cbEmeliyyatNovu.Text == "Kredit ödənişi" || cbEmeliyyatNovu.Text == "Avans ödənişi")
            {
                cbKreditNomresi.Enabled = true;
                label17.Enabled = true;
                label45.Enabled = true;
                button1.Enabled = true;
                txtOdeyen.Enabled = true;
            }
            else
            {
                cbKreditNomresi.Enabled = false;
                label17.Enabled = false;
                label45.Enabled = false;
                button1.Enabled = false;
                txtOdeyen.Enabled = false;
            }
        }

        private void toolStripMenuItem8_Click(object sender, EventArgs e)
        {
            try
            {
                tabControl1.SelectedTab = tabPage6;
                cbEmeliyyatNovu.Text = "Kredit ödənişi";
                cbKreditNomresi.Text = dataBugun.Rows[dataBugun.CurrentCell.RowIndex].Cells["kreditNomresiBugun"].Value.ToString() + " " + dataBugun.Rows[dataBugun.CurrentCell.RowIndex].Cells["adiBugun"].Value.ToString();
                txtGiris.Text = dataBugun.Rows[dataBugun.CurrentCell.RowIndex].Cells["odenilmeli"].Value.ToString();
                KassaRefresh();
            }
            catch { }
        }

        private void toolStripMenuItem9_Click(object sender, EventArgs e)
        {
            try
            {
                tabControl1.SelectedTab = tabPage6;
                cbEmeliyyatNovu.Text = "Kredit ödənişi";
                cbKreditNomresi.Text = dataSabah.Rows[dataSabah.CurrentCell.RowIndex].Cells["kreditNomresiSabah"].Value.ToString() + " " + dataSabah.Rows[dataSabah.CurrentCell.RowIndex].Cells["adiSabah"].Value.ToString();
                txtGiris.Text = dataSabah.Rows[dataSabah.CurrentCell.RowIndex].Cells["odenilmeliSabah"].Value.ToString();
                KassaRefresh();
            }
            catch { }
        }

        private void toolStripMenuItem10_Click(object sender, EventArgs e)
        {
            try
            {
                tabControl1.SelectedTab = tabPage6;
                cbEmeliyyatNovu.Text = "Kredit ödənişi";
                cbKreditNomresi.Text = dataGecikme.Rows[dataGecikme.CurrentCell.RowIndex].Cells["kreditNomresiGecikme"].Value.ToString() + " " + dataGecikme.Rows[dataGecikme.CurrentCell.RowIndex].Cells["adiGecikme"].Value.ToString();
                txtGiris.Text = dataGecikme.Rows[dataGecikme.CurrentCell.RowIndex].Cells["gecikmisBorc"].Value.ToString();
                KassaRefresh();
            }
            catch { }
        }

        private void AdminPanelToolStripMenuItem_Click(object sender, EventArgs e)
        {
            MyCheck.ParolYoxla();
            if (!MyCheck.Parolicaze) return;
            if (!MyCheck.ParolAdminYesNo()) return;

            AdminPanel adminPanel = new AdminPanel();
            adminPanel.Show();
        }

        private void Main_TextChanged(object sender, EventArgs e)
        {
            
        }

        private void dataGecikme_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }
    }
}
