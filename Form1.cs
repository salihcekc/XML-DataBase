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
using System.Data.OleDb;

namespace XML_7
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        OleDbCommand cmd;
        OleDbConnection con;
        string baglanti = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=SınıfListesi.accdb";
        string sorgu;

        string dosyaYolu1 = Application.StartupPath + "\\SınıfListesi.xml";

        private void button1_Click(object sender, EventArgs e)
        {
            DataSet ds = new DataSet();
            DataTable dt = new DataTable();
            ds.ReadXml(dosyaYolu1);
            dataGridView1.DataSource = ds.Tables[0];

            dt = ds.Tables[0];

            for (int i = 0; i < dt.Rows.Count; i++)
            {
                int final = Convert.ToInt32(dt.Rows[i]["final"]);
                int vize = Convert.ToInt32(dt.Rows[i]["vize"]);
                double ortalama = vize * 0.4 + final * 0.6;

                dt.Rows[i]["ortalama"] = ortalama;
                if (final < 45)
                {
                    dt.Rows[i]["durum"] = "KALDI";
                }
                else if (final >= 45)
                {
                    if (ortalama >= 60)
                    {
                        dt.Rows[i]["durum"] = "GEÇTİ";
                    }
                    else if (ortalama < 60)
                    {
                        dt.Rows[i]["durum"] = "KALDI";
                    }
                }
            }

            for (int i = 0; i < dt.Rows.Count; i++)
            {
                string val1 = dataGridView1.Rows[i].Cells[0].Value.ToString();
                int val2 = Int32.Parse(dataGridView1.Rows[i].Cells[1].Value.ToString());
                int val3 = Int32.Parse(dataGridView1.Rows[i].Cells[2].Value.ToString());
                double val4 = Double.Parse(dataGridView1.Rows[i].Cells[3].Value.ToString());
                string val5 = dataGridView1.Rows[i].Cells[4].Value.ToString();
                sorgu = "insert into Tablo1 (İsim,Vize,Final,Ortalama,Durum) values ('" +
                    val1 + "','" +
                    val2 + "','" +
                    val3 + "','" +
                    val4 + "','" +
                    val5 + "')";
                 
                con = new OleDbConnection(baglanti);
                cmd = new OleDbCommand(sorgu, con);
                cmd.Connection.Open();
                cmd.ExecuteNonQuery();
                cmd.Connection.Close();
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {

        }
    }
}
