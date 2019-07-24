using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml.Serialization;

namespace PgmDataCollector
{
    public partial class Form1 : Form
    {

		static public SYSSET SETDATA = new SYSSET();
		ComboBox[] sleeveCombo;
		NumericUpDown[] sleeveNumeric;

        public Form1()
        {
            InitializeComponent();

            ReadDataFromXml();

//フォームが最大化されないようにする
			this.MaximizeBox = false;

#if false //書き込み用 start
			WriteDataToXml();
#endif //書き込み用 end

			//号機番号の設定
			SETDATA.goukiCount = 0;
			for(int i = 0; i < SETDATA.goukiName.Length; i++)
			{
				if(SETDATA.goukiName[i] == "")
				{
					break;
				}
				comboBox1.Items.Add(SETDATA.goukiName[i]);
				SETDATA.goukiCount++;
			}

			//製品形名の設定
			SETDATA.seihinNameCount = 0;
			for(int i = 0; i < SETDATA.seihinName.Length; i++)
			{
				if(SETDATA.seihinName[i] == "")
				{
					break;
				}
				comboBox2.Items.Add(SETDATA.seihinName[i]);
				SETDATA.seihinNameCount++;
			}

			//型の取り外し理由の設定
			ComboBox[] sleeveTable;
			sleeveTable = new ComboBox[] {this.comboBox3, this.comboBox4, this.comboBox5, this.comboBox6, this.comboBox7, this.comboBox8, this.comboBox9, this.comboBox10};
			for(int j = 0; j < sleeveTable.Length; j++)
			{
                SETDATA.changeCauseCount = 0;
                for (int i = 0; i < SETDATA.changeCause.Length; i++)
				{
					if(SETDATA.changeCause[i] == "")
					{
						break;
					}

					sleeveTable[j].Items.Add(SETDATA.changeCause[i]);
					SETDATA.changeCauseCount++;
				}
			}

			//使用型数
            label12.Text = GetEnableSleeveCount().ToString();

			//sleeve Noの設定
			sleeveCombo = new ComboBox[] {this.comboBox11, this.comboBox12, this.comboBox13, this.comboBox14, this.comboBox15, this.comboBox16, this.comboBox17, this.comboBox18};
			for(int i = 0; i < sleeveCombo.Length; i++)
			{
				sleeveCombo[i].Items.Add(" ");
				sleeveCombo[i].Items.Add("A");
				sleeveCombo[i].Items.Add("B");
				sleeveCombo[i].Items.Add("C");
				sleeveCombo[i].Items.Add("D");
				sleeveCombo[i].Items.Add("E");
				sleeveCombo[i].Items.Add("F");
				sleeveCombo[i].Items.Add("G");
			}

			sleeveNumeric = new NumericUpDown[] {this.numericUpDown1, this.numericUpDown2, this.numericUpDown3, this.numericUpDown4, this.numericUpDown5, this.numericUpDown6, this.numericUpDown7, this.numericUpDown8};
			for(int i = 0; i < sleeveNumeric.Length; i++)
			{
				sleeveNumeric[i].Value = i + 1;
			}

		}

        static private void ReadDataFromXml()
		{
            SETDATA.load(ref Form1.SETDATA);
		}

		static public void WriteDataToXml()
		{
            SETDATA.save(Form1.SETDATA);
		}

		public void WriteDataToCsv()
		{
            label12.Text = GetEnableSleeveCount().ToString();

			string machineName = "";
			GetMachineString(ref machineName);
	
			string selectedInfo = "";
			GetLogString(ref selectedInfo);

            try
            {
		        // appendをtrueにすると，既存のファイルに追記
		        //         falseにすると，ファイルを新規作成する
		        var append = false;
		        // 出力用のファイルを開く
                string stCurrentDir = System.IO.Directory.GetCurrentDirectory();
                string path = stCurrentDir + "\\PgmOutData_" + machineName + ".csv";

                string buf = "";
                if(System.IO.File.Exists(path))//既にファイルが存在する
				{
					append = true;
				}
				
		        using(var sw = new System.IO.StreamWriter(path, append, System.Text.Encoding.Default))
		        {
					if(!append)
					{
						buf = string.Format("日付");
						buf += string.Format(",時刻");
						buf += string.Format(",成型機");
						buf += string.Format(",型名");
						buf += string.Format(",使用型数");

						for(int i = 0; i < sleeveNumeric.Length; i++)
						{
							buf += "," + sleeveCombo[i].Text + sleeveNumeric[i].Value.ToString();
						}

	                    sw.WriteLine(buf);

						buf = "";
		                DateTime dt = DateTime.Now;
						buf = string.Format("{0}", dt.ToString("yyyy/MM/dd"));//日付
						buf += string.Format(",{0}", dt.ToString("HH:mm:ss"));//時間
						buf += selectedInfo;
	                    sw.WriteLine(buf);
					}
					else
					{
		                DateTime dt = DateTime.Now;
						buf = string.Format("{0}", dt.ToString("yyyy/MM/dd"));//日付
						buf += string.Format(",{0}", dt.ToString("HH:mm:ss"));//時間
						buf += selectedInfo;
	                    sw.WriteLine(buf);
					}
		        }
		    }
		    catch (System.Exception e)
		    {
		        // ファイルを開くのに失敗したときエラーメッセージを表示
		        System.Console.WriteLine(e.Message);
		    }

		}

		public class SYSSET:System.ICloneable
		{
			public string[] goukiName =		{"", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "",
											 "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "",
											 "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""};

			public string[] seihinName =	{"", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "",
											 "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "",
											 "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "",
											 "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "",
											 "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "",
											 "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "",
											 "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "",
											 "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""};

			public string[] changeCause =	{"", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""};

			public int userMode;
			public int intervalMin;
			public int intervalMax;

			public int goukiCount;
			public int seihinNameCount;
			public int changeCauseCount;
			public int sleeveNoCount = 8;//最大8

            public bool load(ref SYSSET ss)
			{
                string stCurrentDir = System.IO.Directory.GetCurrentDirectory();
                string path = stCurrentDir + "\\SettingData.xml";
				bool ret = false;
				try {
					XmlSerializer sz = new XmlSerializer(typeof(SYSSET));
					System.IO.StreamReader fs = new System.IO.StreamReader(path, System.Text.Encoding.Default);
					SYSSET obj;
					obj = (SYSSET)sz.Deserialize(fs);
					fs.Close();
					obj = (SYSSET)obj.Clone();
					ss = obj;
					ret = true;
				}
				catch (Exception /*ex*/) {
				}
				return(ret);
			}

			public Object Clone()
			{
				SYSSET cln = (SYSSET)this.MemberwiseClone();
				return (cln);
			}

			public bool save(SYSSET ss)
			{
                string stCurrentDir = System.IO.Directory.GetCurrentDirectory();
                string path = stCurrentDir + "\\SettingData.xml";
				bool ret = false;
				try {
					XmlSerializer sz = new XmlSerializer(typeof(SYSSET));
					System.IO.StreamWriter fs = new System.IO.StreamWriter(path, false, System.Text.Encoding.Default);
					sz.Serialize(fs, ss);
					fs.Close();
					ret = true;
				}
				catch (Exception /*ex*/) {
				}
				return (ret);
			}
		}

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
			if(SETDATA.userMode == 1)
			{
				timer1.Enabled = true;
			}
        }

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
			if(SETDATA.userMode == 1)
			{
				timer1.Enabled = false;
			}
        }

        private void comboBox3_SelectedIndexChanged(object sender, EventArgs e)
        {
			int index = comboBox3.SelectedIndex;
			if(index == 0)
			{
				comboBox11.BackColor = Color.Green;
				numericUpDown1.BackColor = Color.Green;
				textBox1.BackColor = Color.Green;
			}
			else
			{
				comboBox11.BackColor = Color.Red;
				numericUpDown1.BackColor = Color.Red;
				textBox1.BackColor = Color.Red;
			}

            WriteDataToCsv();
        }

        private void comboBox4_SelectedIndexChanged(object sender, EventArgs e)
        {
			int index = comboBox4.SelectedIndex;
			if(index == 0)
			{
				comboBox12.BackColor = Color.Green;
				numericUpDown2.BackColor = Color.Green;
				textBox2.BackColor = Color.Green;
			}
			else
			{
				comboBox12.BackColor = Color.Red;
				numericUpDown2.BackColor = Color.Red;
				textBox2.BackColor = Color.Red;
			}

            WriteDataToCsv();
        }

        private void comboBox5_SelectedIndexChanged(object sender, EventArgs e)
        {
			int index = comboBox5.SelectedIndex;
			if(index == 0)
			{
				comboBox13.BackColor = Color.Green;
				numericUpDown3.BackColor = Color.Green;
				textBox3.BackColor = Color.Green;
			}
			else
			{
				comboBox13.BackColor = Color.Red;
				numericUpDown3.BackColor = Color.Red;
				textBox3.BackColor = Color.Red;
			}

            WriteDataToCsv();
        }

        private void comboBox6_SelectedIndexChanged(object sender, EventArgs e)
        {
			int index = comboBox6.SelectedIndex;
			if(index == 0)
			{
				comboBox14.BackColor = Color.Green;
				numericUpDown4.BackColor = Color.Green;
				textBox4.BackColor = Color.Green;
			}
			else
			{
				comboBox14.BackColor = Color.Red;
				numericUpDown4.BackColor = Color.Red;
				textBox4.BackColor = Color.Red;
			}

            WriteDataToCsv();
        }

        private void comboBox7_SelectedIndexChanged(object sender, EventArgs e)
        {
			int index = comboBox7.SelectedIndex;
			if(index == 0)
			{
				comboBox15.BackColor = Color.Green;
				numericUpDown5.BackColor = Color.Green;
				textBox5.BackColor = Color.Green;
			}
			else
			{
				comboBox15.BackColor = Color.Red;
				numericUpDown5.BackColor = Color.Red;
				textBox5.BackColor = Color.Red;
			}

            WriteDataToCsv();
        }

        private void comboBox8_SelectedIndexChanged(object sender, EventArgs e)
        {
			int index = comboBox8.SelectedIndex;
			if(index == 0)
			{
				comboBox16.BackColor = Color.Green;
				numericUpDown6.BackColor = Color.Green;
				textBox6.BackColor = Color.Green;
			}
			else
			{
				comboBox16.BackColor = Color.Red;
				numericUpDown6.BackColor = Color.Red;
				textBox6.BackColor = Color.Red;
			}

            WriteDataToCsv();
        }

        private void comboBox9_SelectedIndexChanged(object sender, EventArgs e)
        {
			int index = comboBox9.SelectedIndex;
			if(index == 0)
			{
				comboBox17.BackColor = Color.Green;
				numericUpDown7.BackColor = Color.Green;
				textBox7.BackColor = Color.Green;
			}
			else
			{
				comboBox17.BackColor = Color.Red;
				numericUpDown7.BackColor = Color.Red;
				textBox7.BackColor = Color.Red;
			}

            WriteDataToCsv();
        }

        private void comboBox10_SelectedIndexChanged(object sender, EventArgs e)
        {
			int index = comboBox10.SelectedIndex;
			if(index == 0)
			{
				comboBox18.BackColor = Color.Green;
				numericUpDown8.BackColor = Color.Green;
				textBox8.BackColor = Color.Green;
			}
			else
			{
				comboBox18.BackColor = Color.Red;
				numericUpDown8.BackColor = Color.Red;
				textBox8.BackColor = Color.Red;
			}

            WriteDataToCsv();
        }
        
        private int GetEnableSleeveCount()
        {
			int okCount = 0;
			ComboBox[] sleeveTable;
			sleeveTable = new ComboBox[] {this.comboBox3, this.comboBox4, this.comboBox5, this.comboBox6, this.comboBox7, this.comboBox8, this.comboBox9, this.comboBox10};
			for(int i = 0; i < sleeveTable.Length; i++)
			{
				if(sleeveTable[i].SelectedIndex == 0)
				{
					okCount++;
				}
			}
			return okCount;
		}


		public void GetMachineString(ref string buf)
		{
            if (comboBox1.SelectedItem == null)
            {
                buf = "";
            }
            else
            {
                buf = comboBox1.SelectedItem.ToString();
            }
		}

		public void GetLogString(ref string buf)
		{
			string setbuf = "";
			if(comboBox1.SelectedItem != null)//成型機
			{
				setbuf = comboBox1.SelectedItem.ToString();
			}
			buf += string.Format(",{0}", setbuf);

			setbuf = "";
			if(comboBox2.SelectedItem != null)//型名
			{
				setbuf = comboBox2.SelectedItem.ToString();
			}
			buf += string.Format(",{0}", setbuf);

			buf += string.Format(",{0}", label12.Text);//型数

			setbuf = "";
			if(comboBox3.SelectedItem != null)//SL-1
			{
				setbuf = comboBox3.SelectedItem.ToString();
			}
			buf += string.Format(",{0}", setbuf);

			setbuf = "";
			if(comboBox4.SelectedItem != null)//SL-2
			{
				setbuf = comboBox4.SelectedItem.ToString();
			}
			buf += string.Format(",{0}", setbuf);

			setbuf = "";
			if(comboBox5.SelectedItem != null)//SL-3
			{
				setbuf = comboBox5.SelectedItem.ToString();
			}
			buf += string.Format(",{0}", setbuf);

			setbuf = "";
			if(comboBox6.SelectedItem != null)//SL-4
			{
				setbuf = comboBox6.SelectedItem.ToString();
			}
			buf += string.Format(",{0}", setbuf);

			setbuf = "";
			if(comboBox7.SelectedItem != null)//SL-5
			{
				setbuf = comboBox7.SelectedItem.ToString();
			}
			buf += string.Format(",{0}", setbuf);

			setbuf = "";
			if(comboBox8.SelectedItem != null)//SL-6
			{
				setbuf = comboBox8.SelectedItem.ToString();
			}
			buf += string.Format(",{0}", setbuf);

			setbuf = "";
			if(comboBox9.SelectedItem != null)//SL-7
			{
				setbuf = comboBox9.SelectedItem.ToString();
			}
			buf += string.Format(",{0}", setbuf);

			setbuf = "";
			if(comboBox10.SelectedItem != null)//SL-8
			{
				setbuf = comboBox10.SelectedItem.ToString();
			}
			buf += string.Format(",{0}", setbuf);
		}

        private void timer1_Tick(object sender, EventArgs e)
        {
            Random r = new System.Random();

			int baseNo = 3;
            int sleeveNo = r.Next(baseNo, baseNo + SETDATA.sleeveNoCount);
			int statusNo = r.Next(0, SETDATA.changeCauseCount);

			int timerInterval = r.Next(SETDATA.intervalMin, SETDATA.intervalMax);
			timer1.Interval = timerInterval;

			switch(sleeveNo)
			{
				case 3:
		            comboBox3.SelectedIndex = statusNo;
		            comboBox3_SelectedIndexChanged(sender, e);
					break;
				case 4:
		            comboBox4.SelectedIndex = statusNo;
		            comboBox4_SelectedIndexChanged(sender, e);
					break;
				case 5:
		            comboBox5.SelectedIndex = statusNo;
		            comboBox5_SelectedIndexChanged(sender, e);
					break;
				case 6:
		            comboBox6.SelectedIndex = statusNo;
		            comboBox6_SelectedIndexChanged(sender, e);
					break;
				case 7:
		            comboBox7.SelectedIndex = statusNo;
		            comboBox7_SelectedIndexChanged(sender, e);
					break;
				case 8:
		            comboBox8.SelectedIndex = statusNo;
		            comboBox8_SelectedIndexChanged(sender, e);
					break;
				case 9:
		            comboBox9.SelectedIndex = statusNo;
		            comboBox9_SelectedIndexChanged(sender, e);
					break;
				case 10:
		            comboBox10.SelectedIndex = statusNo;
		            comboBox10_SelectedIndexChanged(sender, e);
					break;
				default:
					break;
			}

        }

    }
}
