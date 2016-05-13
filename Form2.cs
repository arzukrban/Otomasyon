using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Otomasyon
{
    public partial class Form2 : Form
    {
        public Form2()
        {
            InitializeComponent();
        }

        SqlConnection conn = new SqlConnection("Data Source=MUTRF3240VYQ; Initial Catalog=VDKOtomasyon; User Id=.; password=Marmara23");

        private void BtnKaydet_Click(object sender, EventArgs e)
        {
            try
            {
                //myCommand.CommandText = "SELECT MAX(FormID) FROM tbl_Form";
                //int maxId = Convert.ToInt32(myCommand.ExecuteScalar());
                //SqlCommand kmt = new SqlCommand("Insert into  TblSorun(Tarih,UnvanId,AcilisSifresi,OdaNumarasi,Dahili,Sorun,Cozum,AdSoyad) values ('" + TxtDate.Text + "','" + CmbUnvan.SelectedValue + "','" + TxtAcilisSifresi.Text + "','" + TxtOdaNo.Text + "','" + TxtDahili.Text + "','" + TxtSorun.Text + "','" + TxtSorunCozum.Text + "','" + TxtAdSoyad.Text + "' )", conn);
                //kmt.ExecuteNonQuery();
                //conn.Close();
                //DatagridYenile();
                //Temizle();
                conn.Open();
                string strUpdate = "insert into  TblSorun (Tarih,UnvanId,AcilisSifresi,OdaNumarasi,Dahili,Sorun,Cozum,AdSoyad) values (@Tarih,@UnvanId,@AcilisSifresi,@OdaNumarasi,@Dahili,@Sorun,@Cozum,@AdSoyad)";
                SqlCommand cmdekle = new SqlCommand(strUpdate, conn);
                cmdekle.Parameters.AddWithValue("@Tarih", TxtDate.Text);
                cmdekle.Parameters.AddWithValue("@UnvanId", CmbUnvan.SelectedValue);
                cmdekle.Parameters.AddWithValue("@AcilisSifresi", TxtAcilisSifresi.Text);
                cmdekle.Parameters.AddWithValue("@OdaNumarasi", TxtOdaNo.Text);
                cmdekle.Parameters.AddWithValue("@Dahili", TxtDahili.Text);
                cmdekle.Parameters.AddWithValue("@Sorun", TxtSorun.Text);
                cmdekle.Parameters.AddWithValue("@Cozum", TxtSorunCozum.Text);
                cmdekle.Parameters.AddWithValue("@AdSoyad", TxtAdSoyad.Text);
                cmdekle.ExecuteNonQuery();
                conn.Close();
                DatagridYenile();
                Temizle();
            }

            catch (Exception)
            {
                conn.Close();
                throw;

            }
        }
        private void Temizle()
        {
            TxtAcilisSifresi.Text = "";
            TxtSorun.Text = "";
            TxtOdaNo.Text = "";
            TxtDahili.Text = "";
            TxtDate.Text = DateTime.Now.ToString();
            TxtSorunCozum.Text = "";
            TxtAdSoyad.Text = "";
        }
        protected void DatagridYenile()
        {
            try
            {
                conn.Open();
                System.Data.DataTable tbl = new System.Data.DataTable();
                SqlDataAdapter adptr = new SqlDataAdapter("select TblSorun.Id,UnvanAdi,Tarih,Sorun,Dahili,Cozum,AdSoyad from TblSorun inner join TblUnvan on TblSorun.UnvanId=TblUnvan.ID ", conn);
                adptr.Fill(tbl);
                conn.Close();
                dataGridView1.DataSource = tbl;
                dataGridView1styleOlustur();
            }
            catch (Exception)
            {
                throw;
            }
        }

        private void ComboDoldur()
        {
            try
            {
                conn.Open();
                System.Data.DataTable tblUnvan = new System.Data.DataTable();
                SqlDataAdapter adap = new SqlDataAdapter("Select ID,UnvanAdi from TblUnvan ", conn);
                adap.Fill(tblUnvan);
                conn.Close();
                CmbUnvan.DataSource = tblUnvan;
                CmbUnvan.ValueMember = "ID";
                CmbUnvan.DisplayMember = "UnvanAdi";
            }
            catch (Exception)
            {
                throw;
            }
        }

        private void TxtDahili_TextChanged(object sender, EventArgs e)
        {

        }

        private void dataGridView1_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            //foreach (DataGridViewRow row in dgv.Rows)
            //{
            //    if ((int)row.Tag == ma.ID)//ma.ID is the selected combo box value
            //    {
            //        row.Selected = true;
            //        dgv.CurrentCell = row.Cells[0];
            //    }
            //}
            BtnKaydet.Enabled = false;
            int row = dataGridView1.CurrentRow.Index;
            if (row >= 0)
            {
                TxtDate.Text = dataGridView1[2, row].Value.ToString();
                CmbUnvan.Text = dataGridView1[1, row].Value.ToString();
                TxtSorun.Text = dataGridView1[3, row].Value.ToString();
                TxtDahili.Text = dataGridView1[4, row].Value.ToString();
                TxtSorunCozum.Text = dataGridView1[5, row].Value.ToString();
                TxtAdSoyad.Text = dataGridView1[6, row].Value.ToString();
                TxtId.Text = dataGridView1[0, row].Value.ToString();
            }
        }

        private void BtnGuncelle_Click(object sender, EventArgs e)
        {
            try
            {
                conn.Open();
                System.Data.DataTable tbl = new System.Data.DataTable();
                SqlDataAdapter adp = new SqlDataAdapter("Update TblSorun set Cozum='" + TxtSorunCozum.Text + "' ,Tarih='" + TxtDate.Text + "' ,OdaNumarasi='" + TxtOdaNo.Text + "',Dahili='" + TxtDahili.Text + "',Sorun='" + TxtSorun.Text + "',UnvanId='" + CmbUnvan.SelectedValue + "',AdSoyad='" + TxtAdSoyad.Text + "' where Id='" + TxtId.Text + "'", conn);
                adp.Fill(tbl);
                conn.Close();
                dataGridView1.DataSource = tbl;
                DatagridYenile();
                Temizle();
                BtnKaydet.Enabled = false;
            }
            catch (Exception)
            {
                throw;
            }
        }

        private void btnSil_Click(object sender, EventArgs e)
        {
            try
            {
                conn.Open();
                System.Data.DataTable tbl = new System.Data.DataTable();
                SqlDataAdapter adp = new SqlDataAdapter("Delete from TblSorun  where Id='" + TxtId.Text + "'", conn);
                adp.Fill(tbl);
                conn.Close();
                dataGridView1.DataSource = tbl;
                DatagridYenile();
                Temizle();
                BtnKaydet.Enabled = false;
            }
            catch (Exception)
            {
                throw;
            }
        }

        private void btnYeni_Click(object sender, EventArgs e)
        {
            Temizle();
            BtnKaydet.Enabled = true;
        }

        private void BtnFiltre_Click(object sender, EventArgs e)
        {
            try
            {
                conn.Open();
                System.Data.DataTable tbl = new System.Data.DataTable();
                SqlDataAdapter adp = new SqlDataAdapter("select TblSorun.Id,UnvanAdi,Tarih,Sorun,Dahili,Cozum from TblSorun inner join TblUnvan on TblSorun.UnvanId=TblUnvan.ID  where  Tarih='" + TxtDate.Text + "'", conn);
                adp.Fill(tbl);
                conn.Close();
                dataGridView1.DataSource = tbl;
                //DatagridYenile();
                BtnKaydet.Enabled = false;
            }
            catch (Exception)
            {

                throw;
            }
        }

        private void label8_Click(object sender, EventArgs e)
        {

        }

        //This code Displays row number header:
        private void dataGridView1_DataBindingComplete(object sender, DataGridViewBindingCompleteEventArgs e)
        {
            DataGridView gridView = sender as DataGridView;
            if (null != gridView)
            {

                foreach (DataGridViewRow r in gridView.Rows)
                {
                    gridView.Rows[r.Index].HeaderCell.Value = (r.Index + 1).ToString();
                    gridView.Rows[r.Index].ReadOnly = true;
                }
            }
        }

        private void pictureBox7_Click(object sender, EventArgs e)
        {

        }

        private void Form2_Load(object sender, EventArgs e)
        {
            this.FormBorderStyle = FormBorderStyle.FixedSingle;
            ComboDoldur();
            DatagridYenile();
            PersonelGridYenile();
            //DataGridViewCellStyle cell_style = new DataGridViewCellStyle();
            //cell_style.BackColor = Color.Lavender;
            //cell_style.Alignment = DataGridViewContentAlignment.MiddleRight;
            //cell_style.Format = "C2";    
            //dataGridView1.Columns[0].DefaultCellStyle = cell_style;
            dataGridView1styleOlustur();
            GrdPersonelStyleOlustur();
        }

        private void GrdPersonelStyleOlustur()
        {
            GrdPersonel.Columns[0].Width = 40;
            GrdPersonel.Columns[1].Width = 350;
            GrdPersonel.Columns[2].Width = 350;
            GrdPersonel.Columns[3].Width = 350;
        }

        private void dataGridView1styleOlustur()
        {
            dataGridView1.Columns[0].Width = 40;
            dataGridView1.Columns[1].Width = 80;
            dataGridView1.Columns[2].Width = 100;
            dataGridView1.Columns[3].Width = 300;
            dataGridView1.Columns[4].Width = 100;
            dataGridView1.Columns[5].Width = 300;
            dataGridView1.Columns[6].Width = 150;

        }

        private void BtnDialogAc_Click(object sender, EventArgs e)
        {
            openFileDialog1.Title = "Dosya Seçiniz !";
            openFileDialog1.Filter = "Excel Dosyası |*.xlsx| Excel Dosyası|*.xls";
            openFileDialog1.InitialDirectory = "C:\\Users\\arzu.kurban\\Desktop\\TURNIKE PERSONEL GIRIS-CIKIS";
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                label11.Text = openFileDialog1.FileName;
                exceldata(openFileDialog1.FileName);
                PersonelGridYenile();
                MessageBox.Show("Aktarma işlemi tamamlanmıştır.", "Bilgi", MessageBoxButtons.OK);
            }
            else
                label11.Text = "";
        }
        public static System.Data.DataTable exceldata(string filePath)
        {
            try
            {
                SqlConnection conn = new SqlConnection("Data Source=MUTRF3240VYQ; Initial Catalog=VDKOtomasyon; User Id=.; password=Marmara23");
                System.Data.DataTable dtexcel = new System.Data.DataTable();
                bool hasHeaders = false;
                string HDR = hasHeaders ? "Yes" : "No";
                string strConn;
                if (filePath.Substring(filePath.LastIndexOf('.')).ToLower() == ".xlsx")
                    strConn = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + filePath + ";Extended Properties=\"Excel 12.0;HDR=" + HDR + ";IMEX=0\"";
                else
                    strConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + filePath + ";Extended Properties=\"Excel 8.0;HDR=" + HDR + ";IMEX=0\"";
                OleDbConnection connect = new OleDbConnection(strConn);
                connect.Open();
                System.Data.DataTable schemaTable = connect.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, new object[] { null, null, null, "TABLE" });
                DataRow schemaRow = schemaTable.Rows[0];
                string sheet = schemaRow["TABLE_NAME"].ToString();
                if (!sheet.EndsWith("_"))
                {
                    string query = "SELECT  * FROM [Sayfa1$]";
                    OleDbDataAdapter daexcel = new OleDbDataAdapter(query, connect);
                    dtexcel.Locale = CultureInfo.CurrentCulture;
                    daexcel.Fill(dtexcel);
                    conn.Open();
                    for (int i = 0; i < dtexcel.Rows.Count; i++)
                    {
                        String AdSoyad = dtexcel.Rows[i]["ADI SOYADI"].ToString();
                        String Unvani = dtexcel.Rows[i]["UNVANI"].ToString();
                        String GorevYeri = dtexcel.Rows[i]["GÖREV YERİ"].ToString();

                        SqlCommand kmtselect = new SqlCommand("Select COUNT(*) FROM  TblCalisan where  AdSoyad= '" + AdSoyad + "' and Unvani = '" + Unvani + "' and GorevYeri = '" + GorevYeri + "'", conn);
                        //if records exists
                        string rowsAffected = kmtselect.ExecuteScalar().ToString();

                        if (Int32.Parse(rowsAffected) >= 1)
                        {
                            MessageBox.Show(AdSoyad + "isimli kullanıcı mevcuttur.", "Bilgi", MessageBoxButtons.OK);
                        }
                        else
                        {
                            SqlCommand kmtekle = new SqlCommand("Insert into  TblCalisan(AdSoyad,Unvani,GorevYeri) values ('" + AdSoyad + "','" + Unvani + "','" + GorevYeri + "' )", conn);
                            kmtekle.ExecuteNonQuery();
                        }

                    }
                    conn.Close();
                }
                connect.Close();
                return dtexcel;
            }
            catch (Exception)
            {
                throw;
            }
        }
        protected void PersonelGridYenile()
        {
            try
            {
                conn.Open();
                System.Data.DataTable tbl = new System.Data.DataTable();
                SqlDataAdapter adptr = new SqlDataAdapter("select ID,AdSoyad,Unvani,GorevYeri from TblCalisan", conn);
                adptr.Fill(tbl);
                conn.Close();
                GrdPersonel.DataSource = tbl;
                GrdPersonelStyleOlustur();
            }
            catch (Exception)
            {

                throw;
            }

        }

        private void PrsYeni_Click(object sender, EventArgs e)
        {
            PersonelEkranTemizle();
            PrsKaydet.Enabled = true;
        }

        private void PersonelEkranTemizle()
        {
            TxtPrsAdSoyadi.Text = "";
            TxtPrsGorevYeri.Text = "";

        }

        private void PrsKaydet_Click(object sender, EventArgs e)
        {
            try
            {
                conn.Open();
                SqlCommand kmtekle = new SqlCommand("Insert into  TblCalisan(AdSoyad,Unvani,GorevYeri) values ('" + TxtPrsAdSoyadi.Text + "','" + cmbPrsUnvan.SelectedText + "','" + TxtPrsGorevYeri.Text + "' )", conn);
                kmtekle.ExecuteNonQuery();
                conn.Close();
                PersonelGridYenile();
                PersonelEkranTemizle();
            }
            catch (Exception)
            {

                throw;
            }

        }

        private void PrsGuncelle_Click(object sender, EventArgs e)
        {
            try
            {
                conn.Open();
                System.Data.DataTable tbl = new System.Data.DataTable();
                SqlDataAdapter adp = new SqlDataAdapter("Update TblCalisan set AdSoyad='" + TxtPrsAdSoyadi.Text + "' ,Unvani='" + cmbPrsUnvan.Text + "',GorevYeri='" + TxtPrsGorevYeri.Text + "' where Id='" + TxtPrsId.Text + "'", conn);
                adp.Fill(tbl);
                conn.Close();
                GrdPersonel.DataSource = tbl;
                PersonelGridYenile();
                PersonelEkranTemizle();
                BtnKaydet.Enabled = false;
            }
            catch (Exception)
            {

                throw;
            }

        }

        private void PrsSil_Click(object sender, EventArgs e)
        {
            try
            {
                conn.Open();
                System.Data.DataTable tbl = new System.Data.DataTable();
                SqlDataAdapter adp = new SqlDataAdapter("Delete from TblCalisan  where Id='" + TxtPrsId.Text + "'", conn);
                adp.Fill(tbl);
                conn.Close();
                GrdPersonel.DataSource = tbl;
                PersonelGridYenile();
                PersonelEkranTemizle();
                PrsKaydet.Enabled = false;
            }
            catch (Exception)
            {

                throw;
            }

        }

        private void GrdPersonel_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            PrsKaydet.Enabled = false;
            int rowprs = GrdPersonel.CurrentRow.Index;
            if (rowprs >= 0)
            {
                TxtPrsId.Text = GrdPersonel[0, rowprs].Value.ToString();
                TxtPrsAdSoyadi.Text = GrdPersonel[1, rowprs].Value.ToString();
                cmbPrsUnvan.Text = GrdPersonel[2, rowprs].Value.ToString();
                TxtPrsGorevYeri.Text = GrdPersonel[3, rowprs].Value.ToString();
            }
        }

        private void GrdPersonel_DataBindingComplete(object sender, DataGridViewBindingCompleteEventArgs e)
        {
            DataGridView gridView = sender as DataGridView;
            if (null != gridView)
            {

                foreach (DataGridViewRow r in gridView.Rows)
                {
                    gridView.Rows[r.Index].HeaderCell.Value = (r.Index + 1).ToString();
                    gridView.Rows[r.Index].ReadOnly = true;
                }
            }
        }

        private void BtnTurnikeGirisCikis_Click(object sender, EventArgs e)
        {
            OpenFileDialog TurnikeDialog = new OpenFileDialog();
            TurnikeDialog.Title = "Dosya Seçiniz !";
            TurnikeDialog.Filter = "Excel Dosyası |*.xlsx| Excel Dosyası|*.xls";
            TurnikeDialog.InitialDirectory = "C:\\Users\\arzu.kurban\\Desktop\\TURNIKE PERSONEL GIRIS-CIKIS";
            if (TurnikeDialog.ShowDialog() == DialogResult.OK)
            {
                label11.Text = TurnikeDialog.FileName;
                excelturnike(TurnikeDialog.FileName);
                TurnikeGridYenile();
                MessageBox.Show("Aktarma işlemi tamamlanmıştır.", "Bilgi", MessageBoxButtons.OK);
            }
            else
                label11.Text = "";
        }

        private static System.Data.DataTable excelturnike(string path)
        {
            try
            {
                SqlConnection conn = new SqlConnection("Data Source=MUTRF3240VYQ; Initial Catalog=VDKOtomasyon; User Id=.; password=Marmara23");
                System.Data.DataTable dtexcel = new System.Data.DataTable();
                bool hasHeaders = false;
                string HDR = hasHeaders ? "Yes" : "No";
                string strConn;
                if (path.Substring(path.LastIndexOf('.')).ToLower() == ".xlsx")
                    strConn = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + path + ";Extended Properties=\"Excel 12.0;HDR=" + HDR + ";IMEX=0\"";
                else
                    strConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + path + ";Extended Properties=\"Excel 8.0;HDR=" + HDR + ";IMEX=0\"";
                OleDbConnection connect = new OleDbConnection(strConn);
                connect.Open();
                System.Data.DataTable schemaTable = connect.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, new object[] { null, null, null, "TABLE" });
                DataRow schemaRow = schemaTable.Rows[0];
                string sheet = schemaRow["TABLE_NAME"].ToString();
                if (!sheet.EndsWith("_"))
                {
                    string query = "SELECT  * FROM [data$]";
                    OleDbDataAdapter daexcel = new OleDbDataAdapter(query, connect);
                    dtexcel.Locale = CultureInfo.CurrentCulture;
                    daexcel.Fill(dtexcel);
                    conn.Open();
                    for (int i = 0; i < dtexcel.Rows.Count; i++)
                    {
                        String TarihSaat = dtexcel.Rows[i]["Tarih ve Saat"].ToString();
                        String PersonelNo = dtexcel.Rows[i]["Personel Numarasi"].ToString();
                        String Isim = dtexcel.Rows[i]["Isim"].ToString();
                        String Soyadi = dtexcel.Rows[i]["Soyadi"].ToString();
                        String KartNo = dtexcel.Rows[i]["Kart Numarasi"].ToString();
                        String Cihaz = dtexcel.Rows[i]["Cihaz Adi"].ToString();
                        String Olay = dtexcel.Rows[i]["Olay Noktasi"].ToString();
                        String GirisCikisDurumu = dtexcel.Rows[i]["Giris/Cikis Durumu"].ToString();
                        String OlayTanim = dtexcel.Rows[i]["Olay Tanimi"].ToString();

                        SqlCommand kmtselect = new SqlCommand("Select COUNT(*) FROM  TblGirisCikis where  KartNo= '" + KartNo + "' and TarihSaat = Convert(datetime,'" + TarihSaat + "',103) AND OlayTanim='" + OlayTanim + "'", conn);
                        //if records exists
                        string rowsAffected = kmtselect.ExecuteScalar().ToString();

                        if (Int32.Parse(rowsAffected) >= 1)
                        {
                            MessageBox.Show(Isim + " " + Soyadi + " kullanıcısı " + TarihSaat + " tarihinde zaten turnikeden geçiş yapmıştır", "Bilgi", MessageBoxButtons.OK);
                        }
                        else
                        {
                            SqlCommand kmtekle = new SqlCommand("Insert into  TblGirisCikis(TarihSaat,PersonelNo,Isim,Soyadi,KartNo,Cihaz,Olay,GirisCikisDurumu,OlayTanim) values (Convert(datetime,'" + TarihSaat + "',103),'" + PersonelNo + "','" + Isim + "','" + Soyadi + "','" + KartNo + "','" + Cihaz + "','" + Olay + "','" + GirisCikisDurumu + "','" + OlayTanim + "' )", conn);
                            kmtekle.ExecuteNonQuery();
                        }

                    }
                    conn.Close();
                }
                connect.Close();
                return dtexcel;
            }
            catch (Exception)
            {
                throw;
            }
        }

        private void TurnikeGridYenile()
        {
            try
            {
                conn.Open();
                System.Data.DataTable tbl = new System.Data.DataTable();
                SqlDataAdapter adptr = new SqlDataAdapter("select * From V_TurnikeGirisCikis ", conn);
                adptr.Fill(tbl);
                conn.Close();
                GrdTurnike.DataSource = tbl;
            }
            catch (Exception)
            {
                throw;
            }
        }

        private void btnFiltrele_Click(object sender, EventArgs e)
        {
            conn.Open();
            SqlCommand cmd = new SqlCommand("TurnikeGCKisitli", conn);
            cmd.CommandType = CommandType.Text;
            cmd.CommandType = CommandType.StoredProcedure;
            DateTime date = Convert.ToDateTime(TxtDatepck.Text);
            String tarih = String.Format("{0:[yyyy-MM-dd]}", date);
            String durum = "";
            if (rbGelen.Checked)
                durum = "1";
            else if (rbGelmeyen.Checked)
                durum = "0";
            else if (rbTumu.Checked)
                durum = " ";
            String GorevYeri = TxtGorevYeriFilt.Text;
            cmd.Parameters.AddWithValue("@deg", tarih);
            cmd.Parameters.AddWithValue("@durum", durum);
            cmd.Parameters.AddWithValue("@GorevYeri", GorevYeri);
            System.Data.DataTable dt = new System.Data.DataTable();
            SqlDataAdapter adap = new SqlDataAdapter(cmd);
            try
            {

                adap.Fill(dt);
                conn.Close();
                GrdFiltreli.DataSource = dt;
                GrdFiltreliStyleOlustur();
            }
            catch (Exception)
            {
                conn.Close();
                MessageBox.Show("Bu tarihte turnike giriş çıkış bulunamamıştır.", "Uyarı", MessageBoxButtons.OK);

            }


        }

        private void GrdFiltreliStyleOlustur()
        {

            GrdFiltreli.Columns[0].Width = 350;
            GrdFiltreli.Columns[1].Width = 250;
            GrdFiltreli.Columns[2].Width = 100;
            GrdFiltreli.Columns[3].Width = 380;
        }

        private void label20_Click(object sender, EventArgs e)
        {

        }

        private void GrdFiltreli_DataBindingComplete(object sender, DataGridViewBindingCompleteEventArgs e)
        {
            if (null != GrdFiltreli)
            {

                foreach (DataGridViewRow r in GrdFiltreli.Rows)
                {
                    GrdFiltreli.Rows[r.Index].HeaderCell.Value = (r.Index + 1).ToString();
                    GrdFiltreli.Rows[r.Index].ReadOnly = true;
                }
            }
        }

        private void btnTüm_Click(object sender, EventArgs e)
        {
            conn.Open();
            SqlCommand cmd = new SqlCommand("TurnikeGirisCikis", conn);
            cmd.CommandType = CommandType.Text;
            cmd.CommandType = CommandType.StoredProcedure;
            System.Data.DataTable dt = new System.Data.DataTable();
            SqlDataAdapter adap = new SqlDataAdapter(cmd);
            adap.Fill(dt);
            conn.Close();
            GrdFiltreli.DataSource = dt;
        }

        private void BtnExcel_Click(object sender, EventArgs e)
        {
            ExceleAktar(GrdFiltreli);
        }

        //private void ExceleAktar(DataGridView grd)
        //{
        //    Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();

        //    excel.Visible = true; //Daha fazla bilgi için : www.gorselprogramlama.com

        //    Microsoft.Office.Interop.Excel.Workbook workbook = excel.Workbooks.Add(System.Reflection.Missing.Value);


        //    Microsoft.Office.Interop.Excel.Worksheet sheet1 = (Microsoft.Office.Interop.Excel.Worksheet)workbook.Sheets[1];


        //    int StartCol = 1;

        //    int StartRow = 1; //Daha fazla bilgi için : www.gorselprogramlama.com

        //    for (int j = 0; j < grd.Columns.Count; j++)
        //    {

        //        Microsoft.Office.Interop.Excel.Range myRange = (Microsoft.Office.Interop.Excel.Range)sheet1.Cells[StartRow, StartCol + j];

        //        myRange.Value2 = grd.Columns[j].HeaderText;

        //    }

        //    StartRow++;

        //    for (int i = 0; i < grd.Rows.Count; i++)
        //    {

        //        for (int j = 0; j < grd.Columns.Count; j++)
        //        { //Daha fazla bilgi için : www.gorselprogramlama.com

        //            try
        //            {
        //                Microsoft.Office.Interop.Excel.Range myRange = (Microsoft.Office.Interop.Excel.Range)sheet1.Cells[StartRow + i, StartCol + j];
        //                myRange.Value2 = grd[j, i].Value == null ? "" : grd[j, i].Value;

        //            }
        //            catch
        //            {
        //                ;
        //            }

        //        } //Daha fazla bilgi için : www.gorselprogramlama.com

        //    }
        //    string name = "C:\\Users\\arzu.kurban\\Documents\\MyExcelTestTest.xls";
        //    workbook.SaveAs(name, Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookNormal, System.Reflection.Missing.Value, System.Reflection.Missing.Value, false, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlShared, false, false, System.Reflection.Missing.Value, System.Reflection.Missing.Value, System.Reflection.Missing.Value);

        //}

        private void ExceleAktar(DataGridView grd)
        {
            grd.AllowUserToAddRows = false;
            System.Globalization.CultureInfo dil = System.Threading.Thread.CurrentThread.CurrentCulture;
            System.Threading.Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("en-us");
            Microsoft.Office.Interop.Excel.Application Tablo = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook kitap = Tablo.Workbooks.Add(true);
            Microsoft.Office.Interop.Excel.Worksheet sayfa = (Microsoft.Office.Interop.Excel.Worksheet)Tablo.ActiveSheet;
            System.Threading.Thread.CurrentThread.CurrentCulture = dil;
            Tablo.Visible = true;
            sayfa = (Worksheet)kitap.ActiveSheet;
            System.Threading.Thread.Sleep(10000);


            try
            {
                for (int i = 0; i < grd.Rows.Count; i++)
                {
                    for (int j = 0; j < grd.ColumnCount; j++)
                    {
                        if (i == 0)
                        {
                            Tablo.Cells[1, j + 1] = grd.Columns[j].HeaderText;
                        }
                        Tablo.Cells[i + 2, j + 1] = grd.Rows[i].Cells[j].Value.ToString();
                    }
                }
                Tablo.Visible = true;
                Tablo.UserControl = true;
            }
            catch (Exception)
            {
                ;
            }



        }

        private void GrdFiltreli_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Control && e.KeyCode == Keys.E)
                ExceleAktar(GrdFiltreli);
        }

        private void BtnTurnikeExcel_Click(object sender, EventArgs e)
        {
            ExceleAktar(GrdTurnike);
        }

        private void GrdTurnike_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Control && e.KeyCode == Keys.E)
                ExceleAktar(GrdTurnike);
        }

        private void GrdTurnike_DataBindingComplete(object sender, DataGridViewBindingCompleteEventArgs e)
        {
            if (null != GrdTurnike)
            {

                foreach (DataGridViewRow r in GrdTurnike.Rows)
                {
                    GrdTurnike.Rows[r.Index].HeaderCell.Value = (r.Index + 1).ToString();
                    GrdTurnike.Rows[r.Index].ReadOnly = true;
                }
            }
        }

        private void btnPersonelGirisCikis_Click(object sender, EventArgs e)
        {
            TurnikeGridYenile();
        }

        private void btnTurnikeFlt_Click(object sender, EventArgs e)
        {
            conn.Open();
            System.Data.DataTable tbl = new System.Data.DataTable();
            String kisit = "";

            if (TxtTarihFltPck.Text != " ")
            {
                DateTime date = Convert.ToDateTime(TxtTarihFltPck.Text);
                String tarihflt = String.Format("{0:yyyy-MM-dd}", date);
                kisit += " Tarih>='" + tarihflt + "'";
            }
            if (TxtAdSoyadFlt.Text != "")
            {
                kisit += " and AdSoyad like '%" + TxtAdSoyadFlt.Text + "%'";
            }
            if (TxtUnvanFlt.Text != "")
            {
                kisit += " and Unvani like '%" + TxtUnvanFlt.Text + "%'";
            }
            if (txtCihazFlt.Text != "")
            {
                kisit += " and Cihaz like '%" + txtCihazFlt.Text + "%'";
            }

            if (TxtGirisCikisFlt.Text != "")
            {
                kisit += "and GirisCikisDurumu like '%" + TxtGirisCikisFlt.Text + "%'";
            }
            if (txtSaatFlt.Text != "")
            {
                kisit += "and Saat like '%" + txtSaatFlt.Text + "%'";
            }
            if (TxtOlayFlt.Text != "")
            {
                kisit += "and Olay like '%" + TxtOlayFlt.Text + "%'";
            }
            else
            {
            }
            SqlDataAdapter adptr = new SqlDataAdapter("select * From V_TurnikeGirisCikis where " + kisit + "", conn);
            adptr.Fill(tbl);
            conn.Close();
            GrdTurnike.DataSource = tbl;
        }


    }
}
