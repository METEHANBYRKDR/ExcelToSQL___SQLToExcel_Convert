using System.Collections;
using System.Data.SqlClient;
using System.Linq.Expressions;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelVTEntegrasyonProjesi
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        SqlConnection sqlConnection = new SqlConnection(@"Data Source=DESKTOP-ITDJ42N\SQLEXPRESS;Initial Catalog=ProjelerVT;Integrated Security=True");

        private void btn_VTdenOku_Click(object sender, EventArgs e)
        {
            Excel.Application excelUygulama = new Excel.Application();
            excelUygulama.Visible = true;
            Excel.Workbook workBook = excelUygulama.Workbooks.Add(System.Reflection.Missing.Value);
            Excel.Worksheet sayfa1 = workBook.Sheets[1];

            string[] basliklar = { "Personel no", "Ad", "Soyad", "Semt", "Sehir" };
            Excel.Range range;
            for (int i = 0; i < basliklar.Length; i++)
            {
                range = sayfa1.Cells[1, (i + 1)];
                range.Value2 = basliklar[i];

            }







            try
            {
                sqlConnection.Open();
                string sqlCumlesi = "SELECT * FROM PERSONEL";
                SqlCommand cmd = new SqlCommand(sqlCumlesi, sqlConnection);
                SqlDataReader reader = cmd.ExecuteReader();

                int satýr = 2;
                while (reader.Read())
                {
                    string Personel_No = reader[0].ToString();
                    string P_Ad = reader[1].ToString();
                    string P_Soyad = reader[2].ToString();
                    string P_Semt = reader[3].ToString();
                    string P_Sehir = reader[4].ToString();
                    richTextBox1.Text = richTextBox1.Text + Personel_No + " " + P_Ad + " " + P_Soyad + " " + P_Semt + " " + P_Sehir + "\n ";
                    range = sayfa1.Cells[satýr, 1];
                    range.Value2 = Personel_No;
                    range = sayfa1.Cells[satýr, 2];
                    range.Value2 = P_Ad;
                    range = sayfa1.Cells[satýr, 3];
                    range.Value2 = P_Soyad;
                    range = sayfa1.Cells[satýr, 4];
                    range.Value2 = P_Semt;
                    range = sayfa1.Cells[satýr, 5];
                    range.Value2 = P_Sehir;

                    satýr++;

                }


            }
            catch (Exception ex)
            {
                MessageBox.Show("Sql baðlantý hatasý yaþandý Hata Kodu =SQLREAD1 \n " + ex.ToString());
            }

            finally
            {
                if (sqlConnection != null)
                    sqlConnection.Close();
            }


        }

        private void btn_ExceldenOku_Click(object sender, EventArgs e)
        {
            Excel.Application exlApp;
            Excel.Workbook exlWorkBook;
            Excel.Worksheet exlWorkSheet;
            Excel.Range range;
            int rCnt = 0;
            int cCnt = 0;
            exlApp = new Excel.Application();
            exlWorkBook = exlApp.Workbooks.Open("C:\\test\\test.xlsx");
            exlWorkSheet = exlWorkBook.Worksheets.get_Item(1);
            range = exlWorkSheet.UsedRange;

            // rich textbox içini temizle
            richTextBox2.Clear();
            // ilk satýr baþlýk olduðu için rCnt = 2 den baþlamasý lazým


            for (rCnt = 2; rCnt <= range.Rows.Count; rCnt++)
            {
                ArrayList list = new ArrayList();

                for (cCnt = 1; cCnt <= range.Columns.Count; cCnt++)
                {
                    string okunanHucre = Convert.ToString((range.Cells[rCnt, cCnt] as Excel.Range).Value2);
                    richTextBox2.Text = richTextBox2.Text + okunanHucre + " ";
                    list.Add(okunanHucre);

                }
                richTextBox2.Text = richTextBox2.Text + "\n";

                try
                {
                    sqlConnection.Open();
                    SqlCommand sqlCommand = new SqlCommand("INSERT INTO Personel (PersonelNo ,Ad ,Soyad ,Semt ,Sehir) " + "VALUES (@P1 ,@P2 ,@P3 ,@P4 ,@ P5)", sqlConnection);
                    sqlCommand.Parameters.AddWithValue("P1", list[0]);
                    sqlCommand.Parameters.AddWithValue("P2", list[1]);
                    sqlCommand.Parameters.AddWithValue("P3", list[2]);
                    sqlCommand.Parameters.AddWithValue("P4", list[3]);
                    sqlCommand.Parameters.AddWithValue("P5", list[4]);
                    sqlCommand.ExecuteNonQuery();


                }
                catch(Exception ex)

                {
                    MessageBox.Show("Baðlantýyý veri tabanýna yazarken bir hata oluþtu! HATA KODU :SQLWRITE01\n" + ex.ToString());

                }
                finally
                {
                    if (sqlConnection != null)
                    {
                        sqlConnection.Close();

                    }

                }

            }
            exlApp.Quit();
            ReleaseObject(exlWorkSheet);
            ReleaseObject(exlWorkBook);
            ReleaseObject(exlApp);


        }

        private void ReleaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;

            }
            finally { GC.Collect(); }
        }


    }
}