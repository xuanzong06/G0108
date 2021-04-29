敘述夢境，

此事難解，與祖先有關，祖先在搶人(林哲男)

因為祖先都姓「林」，有些事情離不清(牽未離)

此刀剖下去了意思是，要剖清楚

看是這個林家要，還是那個林家要

因為林家這房男生只剩哲男哥，沒其他人了，所以要分清楚(看是要歸到哪方祖先下)

沒分清楚，所有子孫都會出事情，因為兩邊都在搶人，搶人的過程會有些意外發生

發生的事情，不是神明可以插手來管的，因為這算是家務事，家務事神明不能插手

同姓結婚比較少，但已經結了，就不用去探討為什麼結婚，事情已發生，就來處理

向祖先稟報：
1)向父方或母方的祖先稟報，看是要分給哪一方祖先(因為都姓林，往任何一邊都行，)

2)姨丈的意思是，要分清楚、乾淨。

3)看決定好要分給哪一方祖先，那另外一方祖先就要好好安撫，說子孫都是無辜的，祖先在搶人，子孫就不順，就會發生事情，那這都不是大家願意樂見的。(好好安撫林家)

兩方祖先都要好好稟報清楚

4)家中有安神明，但此神明官階不高，祂無法插手處理，但還是可以向祂稟報，了解此事

5)若要處理此事，就要看關聖帝君有沒有要受理此案件，因為不好處理啦。
若帝君有要處理此事，則請帝君處理。
若帝君沒有受理此事，則找台南的玉皇上帝做主

直接找玉皇上帝做主是最快的，因為任何事情都要經過祂
不管今天委託媽祖、王爺、觀音這些神明幫忙，都還是要經過玉皇上帝

住台南的話，離玉皇上帝廟很近，就去稟報請示作主。

6)補庫的原因
6.1：因為祖先搶人，運就不好，會出事，所以補庫來補運的，使他平安
6.2：看是否能幫助哥哥的財運

帝君交代的事情，都是要注意的(與祖先有關)，以防萬一注意。

此次補庫是補庫的大概，但距離圓滿有差一點
以後如果要補庫，補圓滿一點，要另外算，若不是找帝君做主，也可以去找其他宮廟，只要神明願意做主，都可以唷!!

解法：
順序：
稟報玉皇上帝 > (詳細說明 往生者名、往生者農曆生日、何時往生、來託夢給誰、被托夢者名字、托夢內容)向玉皇上帝稟報 > 稟報完畢後 > 擲筊請示玉皇上帝是否能做主 > 因為你們有來問行渡宮的神明，得知此事有關祖先的問題 > 祖先問題是否可以請玉皇上帝來做主 > 若有應筊(聖筊)就可以繼續問 > 問林哲男是要往父方那邊，還是母方那邊 > 往男方、母方的這個議題，也可以家庭討論，可問林哲男本人意願 > 

若家庭討論後決定要往哪邊，再向玉皇上帝稟報(結果)>然後請玉皇上帝來做主、調停 > 莫讓祖先為此事吵鬧，使父親不安寧 > 若玉皇上帝應筊願意做主，就可以放心了 > 因為玉皇上帝已受理。

若不放心，可以向祖先稟報、擲筊詢問是否圓滿，若祖先OK，請祖先庇佑平安快樂賺大錢

生者吵架、亡者吵架有差別的XD



我開了一個Console專案，來模擬兩個cs檔案之間傳值，希望能幫上忙。

//MainCS
namespace testIfElse//請忽略我的專案名稱XD
{
    class Program
    {
        public static bool Monitor = false;//宣告public的全域變數，讓其他CS檔案能夠改變這個值
        static Timer t1; //因為我要模擬一個監控，所以我用timer每兩秒做一次偵測
        static int ResultParam;//該專案模擬兩個接值的方法，這是另外其中一種方法需要的物件。

        static void Main(string[] args)
        {
            ProccesingOliver(1,2);//程式啟動時，執行這個方法

            t1 = new Timer();
            t1.Interval = 2000; //2s
            t1.Elapsed += CheckStatus;//每兩秒呼叫一次，我要模擬你的TCP監控是否有人連線
            t1.AutoReset = true;//每兩秒重做一次
            t1.Enabled = true;//啟動計時
            Console.WriteLine("若要離開，請按任何按鍵");
            Console.ReadLine();//只是為了讓畫面不要消失，你就可以看見每兩秒的監控
        }

        public static int FromOliver(int param)//這個方法是第二個CS檔案呼叫的，然後將傳來的值，給主檔CS上面宣告的整數
        {
            ResultParam = param;
            return ResultParam;
        }

        private static void ProccesingOliver(int param1, int param2)
        {
            Console.WriteLine("請輸入代號");
            int startparam = Int32.Parse(Console.ReadLine());//輸入一個值
            Oliver.getStart(startparam, param1, param2);//這時候會將 輸入的值，跟我給的預設兩個參數另外一個CS檔處理
        }        

        public static void getResult(int result)//結果
        {
            Console.WriteLine("The Result From another Class is : {0}" ,result); //這個result的值是從主檔CS主動去第二個CS檔案拿的
            Console.WriteLine("The Result From FromOliver Function is : {0}", ResultParam); //這個result的值，是第二個CS檔案傳過來給主檔的值。 透過FromOliver(int param)這個方法得到。
            Console.ReadLine();
        }

        private static void CheckStatus(Object source, System.Timers.ElapsedEventArgs e)//每兩秒要做的事情
        {            
            if (Monitor == true)//當布林植改變時。 改變的時機是在另外一個CS檔，改變的
            {
                getResult(Oliver.r1);//條件成立，執行getResult
            }
        }
    }
}

//第二個CS檔
namespace testIfElse
{
    class Oliver
    {
        public static int r1;//宣告public物件，讓主程式可以主動抓

        public static void getStart(int sp, int p1, int p2){  //這裡接從主檔傳來的值
            if (sp == 100) // 如果輸入的代號 = 100，才會執行下面的算術
            {
                doMath(p1, p2); //把我預設的參數丟進去
            }
        }

        private static void doMath(int p1, int p2)
        {
            Console.WriteLine("A+B = {0}" , p1 + p2);//顯示算數結果
            r1 = p1 + p2;
            Program.FromOliver(p1+p2); //我主動從第二個CS檔案傳值回去給主檔
            Program.Monitor = true;//我改變主檔的布林植(為了監控)
        }
    }
}


================================

using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using System.Data.OleDb;
using System.IO;
using SharpZip = ICSharpCode.SharpZipLib.Zip; // this one is important

namespace ReadExcelFileApp
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        public DataTable ReadExcel(string FileName, string FileExt)
        {
            string Source = string.Empty;
            DataTable dtexcewl = new DataTable();
            if(FileExt.CompareTo(".xls") == 0){
                Source = @"provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + FileName + ";Extended Properties='Excel 8.0;HRD=Yes;IMEX=1';"; //for below excel 2007 
            }
            else {
                Source = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + FileName + ";Extended Properties='Excel 12.0;HDR=Yes';"; //for above excel 2007  
            }

            using(OleDbConnection conn = new OleDbConnection(Source)){
                try
                {
                    OleDbDataAdapter adapter = new OleDbDataAdapter("select * from [10908$]", conn);
                    adapter.Fill(dtexcewl);
                }
                catch (Exception ex)
                {
                    string ex_txt = ex.Message.ToString();
                }
            }
            return dtexcewl;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string filePath = string.Empty;
            string fileExt = string.Empty;
            OpenFileDialog file = new OpenFileDialog();
            if(file.ShowDialog() == System.Windows.Forms.DialogResult.OK){
                filePath = file.FileName;
                fileExt = Path.GetExtension(filePath);
                if (fileExt.CompareTo(".xls") == 0 || fileExt.CompareTo(".xlsx") == 0 || fileExt.CompareTo(".ods") == 0)
                {
                    try
                    {
                        DataTable dtExcel = new DataTable();
                        dtExcel = ReadExcel(filePath, fileExt);
                        dataGridView1.DataSource = dtExcel;
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message.ToString());
                    }
                }
                else
                {
                    MessageBox.Show("Please choose .xls or .xlsx file only.", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Error); //custom messageBox to show error  
                }
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                if (e.ColumnIndex != 0)
                {
                    MessageBox.Show("請點擊【第一個】欄位");
                }
                else
                {
                    label1.Text = dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex].Value.ToString();
                    label2.Text = dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex+1].Value.ToString();
                    label3.Text = dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex+2].Value.ToString();
                    label4.Text = dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex+3].Value.ToString();
                    label5.Text = dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex+4].Value.ToString();

                    label6.Text = dataGridView1.Rows[e.RowIndex+1].Cells[e.ColumnIndex].Value.ToString();
                    label7.Text = dataGridView1.Rows[e.RowIndex+1].Cells[e.ColumnIndex + 1].Value.ToString();
                    label8.Text = dataGridView1.Rows[e.RowIndex+1].Cells[e.ColumnIndex + 2].Value.ToString();
                    label9.Text = dataGridView1.Rows[e.RowIndex+1].Cells[e.ColumnIndex + 3].Value.ToString();
                    label10.Text = dataGridView1.Rows[e.RowIndex+1].Cells[e.ColumnIndex + 4].Value.ToString();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("錯誤訊息：" +ex.Message.ToString());
            }
        }
    }
}



=================================================================================================================================================
/*Create_Parents_datas*/
using System;
/**/
using System.Data.OleDb;
using System.Windows.Forms;

namespace GPS
{
    public partial class Create_Parents_datas : Form
    {
        public Create_Parents_datas()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            #region 性別確認
            try
            {
                if (!radioButton1.Checked && !radioButton2.Checked)
                {
                    MessageBox.Show("請點選性別");
                }

                if (radioButton1.Checked)
                {
                    int pgender = 1;
                    create_parentsData(pgender);
                }

                if (radioButton2.Checked)
                {
                    int pgender = 0;
                    create_parentsData(pgender);
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show("建立資料時，出現資料庫連線錯誤，請通知IT \r\n" + ex.Message, ToString());
            }
            #endregion            
        }

        #region 將資料加到資料庫
        private void create_parentsData(int pgender)
        {
            try
            {
                DateTime date1 = DateTime.Now; // yyyy/MM/DD HH:MM:SS
                using (OleDbConnection conn = new OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=GFS.accdb;Persist Security Info=False;"))
                {
                    string sqlcmd = "insert into Parent(parent_name,parent_gender,parent_telephone,parent_cellphone,parent_id,parent_address,parent_ec,parent_ectelephone,parent_eccellphone,parent_createdate) values (@parent_name,@parent_gender,@parent_telephone,@parent_cellphone,@parent_id,@parent_address,@parent_ec,@parent_ectelephone,@parent_eccellphone,@parent_createdate)";
                    conn.Open();
                    OleDbCommand cmd = new OleDbCommand(sqlcmd, conn);

                    #region SQL injection prevent
                    cmd.Parameters.Add("@parent_name", OleDbType.VarChar);
                    cmd.Parameters.Add("@parent_gender", OleDbType.VarChar);
                    cmd.Parameters.Add("@parent_telephone", OleDbType.VarChar);
                    cmd.Parameters.Add("@parent_cellphone", OleDbType.VarChar);
                    cmd.Parameters.Add("@parent_id", OleDbType.VarChar);
                    cmd.Parameters.Add("@parent_address", OleDbType.VarChar);
                    cmd.Parameters.Add("@parent_ec", OleDbType.VarChar);
                    cmd.Parameters.Add("@parent_ectelephone", OleDbType.VarChar);
                    cmd.Parameters.Add("@parent_eccellphone", OleDbType.VarChar);
                    cmd.Parameters.Add("@parent_createdate", OleDbType.VarChar);

                    cmd.Parameters["@parent_name"].Value = textBox1.Text.Trim();
                    cmd.Parameters["@parent_gender"].Value = pgender;
                    cmd.Parameters["@parent_telephone"].Value = textBox2.Text.Trim();
                    cmd.Parameters["@parent_cellphone"].Value = textBox3.Text.Trim();
                    cmd.Parameters["@parent_id"].Value = textBox4.Text.Trim();
                    cmd.Parameters["@parent_address"].Value = textBox7.Text.Trim();
                    cmd.Parameters["@parent_ec"].Value = textBox5.Text.Trim();
                    cmd.Parameters["@parent_ectelephone"].Value = textBox6.Text.Trim();
                    cmd.Parameters["@parent_eccellphone"].Value = textBox8.Text.Trim();
                    cmd.Parameters["@parent_createdate"].Value = date1;
                    #endregion

                    cmd.ExecuteNonQuery();

                    if (MessageBox.Show("家長資料建立成功！ \r\n繼續建立寵物基本資料請按「是」，要離開請按「否」", "訊息提示", MessageBoxButtons.YesNo, MessageBoxIcon.Information) == DialogResult.Yes)
                    {
                        int parent_index = Get_Parent_Index();
                        Create_Pets_datas createdatas2 = new Create_Pets_datas(parent_index);
                        createdatas2.Show();
                        this.Close();
                    }
                    else
                    {
                        this.Close();
                    }
                    this.Close();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString());
            }
        }
        #endregion        

        #region 獲得家長編號
        private int Get_Parent_Index()
        {
            int parent_index = 0;
            using (OleDbConnection conn = new OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=GFS.accdb;Persist Security Info=False;"))
            {
                string sqlcmd = "select parent_index from Parent where parent_name = '" + textBox1.Text + "' and parent_cellphone = '" + textBox3.Text + "'";
                conn.Open();
                OleDbCommand cmd = new OleDbCommand(sqlcmd, conn);
                OleDbDataReader dr = cmd.ExecuteReader();
                while(dr.Read()){
                    parent_index = dr.GetInt32(0);
                }
                label9.Text = "編號(成功建檔後產生)：" + parent_index.ToString(); //顯示在畫面上，可能要先取消 this.close()用來看有沒有正常顯示
            }
            return parent_index;
        }
        #endregion
    }
}


/*Create_Pets_datas*/
using System;
using System.Windows.Forms;
/**/
using System.Data.OleDb;

namespace GPS
{
    public partial class Create_Pets_datas : Form
    {
        #region 必要物件
        //儲存寵物編號
        int pets_index = 0;
        //Get家長編號
        int parent_index = 0;
        #endregion

        public Create_Pets_datas(int parents_index)
        {
            InitializeComponent();
            //接家長的編號
            parent_index = parents_index;
        }

        private void button1_Click(object sender, EventArgs e)
        {            

        }

        private void create_pets_data(int pets_sex, string pets_sstatus, string precaution, string allergy_txt, string notouch_txt, string exp_txt, int medical_treatment)
        {
            #region 建立寵物基本資料
            try
            {
                // insert Pets
                using (OleDbConnection conn = new OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=GFS.accdb;Persist Security Info=False;"))
                {
                    conn.Open();
                    string sqlcmd = "insert into Pets (pets_names, pets_breed, pets_sex, pets_birthDay, pets_weight, pets_ligation, pets_sstatus, pets_precautions, pets_allergy, pets_notouch, pets_experiences, pets_parent_index, pets_createdate ) values ( @pets_names, @pets_breed, @pets_sex, @pets_birthDay, @pets_weight, @pets_ligation, @pets_sstatus, @pets_precautions, @pets_allergy, @pets_notouch, @pets_experiences, @pets_parent_index, @pets_createdate )";
                    OleDbCommand cmd = new OleDbCommand(sqlcmd, conn);
                    #region SQL injection prevent
                    cmd.Parameters.Add("@pets_names", OleDbType.VarChar);
                    cmd.Parameters.Add("@pets_breed", OleDbType.VarChar);
                    cmd.Parameters.Add("@pets_sex", OleDbType.VarChar);
                    cmd.Parameters.Add("@pets_birthDay", OleDbType.VarChar);
                    cmd.Parameters.Add("@pets_weight", OleDbType.VarChar);
                    cmd.Parameters.Add("@pets_ligation", OleDbType.VarChar);
                    cmd.Parameters.Add("@pets_sstatus", OleDbType.VarChar);
                    cmd.Parameters.Add("@pets_precautions", OleDbType.VarChar);
                    cmd.Parameters.Add("@pets_allergy", OleDbType.VarChar);
                    cmd.Parameters.Add("@pets_notouch", OleDbType.VarChar);
                    cmd.Parameters.Add("@pets_experiences", OleDbType.VarChar);
                    cmd.Parameters.Add("@pets_parent_index", OleDbType.VarChar);
                    cmd.Parameters.Add("@pets_createdate", OleDbType.Date);

                    DateTime date1 = DateTime.Now; // yyyy/MM/DD HH:MM:SS

                    cmd.Parameters["@pets_names"].Value = textBox8.Text.Trim();
                    cmd.Parameters["@pets_breed"].Value = comboBox1.SelectedItem.ToString();
                    cmd.Parameters["@pets_sex"].Value = pets_sex;
                    cmd.Parameters["@pets_birthDay"].Value = textBox1.Text.Trim();
                    cmd.Parameters["@pets_weight"].Value = textBox2.Text.Trim();
                    cmd.Parameters["@pets_ligation"].Value = pets_sex;
                    cmd.Parameters["@pets_sstatus"].Value = pets_sstatus;
                    cmd.Parameters["@pets_precautions"].Value = precaution;
                    cmd.Parameters["@pets_allergy"].Value = allergy_txt;
                    cmd.Parameters["@pets_notouch"].Value = notouch_txt;
                    cmd.Parameters["@pets_experiences"].Value = exp_txt;
                    cmd.Parameters["@pets_parent_index"].Value = parent_index;
                    cmd.Parameters["@pets_createdate"].Value = date1;
                    #endregion
                    cmd.ExecuteNonQuery();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("新增資料到寵物基本資料時發生錯誤：" + ex.Message.ToString());
            }
            #endregion

            #region 獲得寵物的編號
            try
            {
                int p_i = Get_pet_index();
                pets_index = Get_pet_index();
            }
            catch (Exception ex)
            {
                MessageBox.Show("獲得寵物編號時發生錯誤：" + ex.Message.ToString());
            }
            #endregion

            #region 建立寵物健康狀況
            try
            {
                //insert Pets_info
                using (OleDbConnection conn = new OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=GFS.accdb;Persist Security Info=False;"))
                {
                    string sqlcmd = " insert into Pets_info (petsinfo_hospital, petsinfo_doctors, petsinfo_telephone, petsinfo_address, petsinfo_medicalHistory, petsinfo_medicalTreatement, petsinfo_vacciendate, petsinfo_rabiesvdate, petsinfo_pets_index, petsinfo_createdate) values ('" + textBox10.Text + "','" + textBox11.Text + "','" + textBox12.Text + "','" + textBox13.Text + "','" + textBox14.Text + "',true or false'" + textBox15.Text + "','" + textBox16.Text + "')";
                    OleDbCommand cmd = new OleDbCommand(sqlcmd, conn);

                    #region SQL injection prevent
                    cmd.Parameters.Add("@petsinfo_hospital", OleDbType.VarChar);
                    cmd.Parameters.Add("@petsinfo_doctors", OleDbType.VarChar);
                    cmd.Parameters.Add("@petsinfo_telephone", OleDbType.VarChar);
                    cmd.Parameters.Add("@petsinfo_address", OleDbType.VarChar);
                    cmd.Parameters.Add("@petsinfo_medicalHistory", OleDbType.VarChar);
                    cmd.Parameters.Add("@petsinfo_medicalTreatement", OleDbType.VarChar);
                    cmd.Parameters.Add("@petsinfo_vacciendate", OleDbType.VarChar);
                    cmd.Parameters.Add("@petsinfo_rabiesvdate", OleDbType.VarChar);
                    cmd.Parameters.Add("@petsinfo_pets_index", OleDbType.VarChar);
                    cmd.Parameters.Add("@petsinfo_createdate", OleDbType.VarChar);

                    DateTime date1 = DateTime.Now; // yyyy/MM/DD HH:MM:SS

                    cmd.Parameters["@petsinfo_hospital"].Value = textBox10.Text.Trim();
                    cmd.Parameters["@petsinfo_doctors"].Value = textBox11.Text.Trim();
                    cmd.Parameters["@petsinfo_telephone"].Value = textBox12.Text.Trim();
                    cmd.Parameters["@petsinfo_address"].Value = textBox13.Text.Trim();
                    cmd.Parameters["@petsinfo_medicalHistory"].Value = textBox14.Text.Trim();
                    cmd.Parameters["@petsinfo_medicalTreatement"].Value = medical_treatment;
                    cmd.Parameters["@petsinfo_vacciendate"].Value = textBox15.Text.Trim();
                    cmd.Parameters["@petsinfo_rabiesvdate"].Value = textBox16.Text.Trim();
                    cmd.Parameters["@petsinfo_pets_index"].Value = pets_index;
                    cmd.Parameters["@petsinfo_createdate"].Value = date1;
                    #endregion

                    cmd.ExecuteNonQuery();

                    if (MessageBox.Show("寵物基本資料建立成功！ \r\n繼續建立寵物美容資料請按「是」，要離開請按「否」", "訊息提示", MessageBoxButtons.YesNo, MessageBoxIcon.Information) == DialogResult.Yes)
                    {
                        Create_PetsSalon_datas createdatas3 = new Create_PetsSalon_datas(pets_index, textBox8.Text.Trim());
                        createdatas3.Show();
                        this.Close();
                    }
                    else
                    {
                        this.Close();
                    }
                    this.Close();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("新增寵物健康狀況時發生錯誤：" + ex.Message.ToString());
            }
            #endregion            
        }

        #region 檢核欄位並產生值
        private void status_check()
        {
            int pets_sex = 2;
            string pets_sstatus = null;
            string precaution = null;
            string allergy_txt = null;
            string notouch_txt = null;
            string exp_txt = null;
            int medical_treatment = 2;

            try
            {
                if (!radioButton3.Checked && !radioButton4.Checked)
                {
                    MessageBox.Show("請點選性別");
                }

                if (radioButton4.Checked) //radioButton4 = 公
                {
                    pets_sex = 1;
                }

                if (radioButton3.Checked)//radioButton3 = 母
                {
                    pets_sex = 0;
                }

                if (checkBox2.Checked)
                {
                    pets_sstatus = "良好";
                }
                else
                {
                    pets_sstatus = textBox3.Text;
                }

                if (checkBox3.Checked)
                {
                    precaution = "無";
                }
                else
                {
                    precaution = textBox4.Text;
                }

                if (checkBox4.Checked)
                {
                    allergy_txt = "無";
                }
                else if (checkBox5.Checked)
                {
                    allergy_txt += "電剪;";
                }
                else if (checkBox6.Checked)
                {
                    allergy_txt += "洗毛精;";
                }
                else if (checkBox7.Checked)
                {
                    allergy_txt += "其他:";
                    allergy_txt += textBox5.Text;
                }

                if (checkBox11.Checked)
                {
                    notouch_txt = "無";
                }
                else if (checkBox10.Checked)
                {
                    notouch_txt += "腳;";
                }
                else if (checkBox10.Checked)
                {
                    notouch_txt += "耳朵;";
                }
                else if (checkBox10.Checked)
                {
                    notouch_txt += "嘴;";
                }
                else if (checkBox10.Checked)
                {
                    notouch_txt += "其他:";
                    notouch_txt += textBox6.Text;
                }

                if (checkBox17.Checked)
                {
                    exp_txt = "無";
                }
                else
                {
                    exp_txt = textBox7.Text;
                }

                if (checkBox13.Checked)
                {
                    medical_treatment = 1;
                }
                else
                {
                    medical_treatment = 0;
                }

                create_pets_data(pets_sex, pets_sstatus, precaution, allergy_txt, notouch_txt, exp_txt, medical_treatment);
            }
            catch (Exception ex)
            {
                MessageBox.Show("建立資料時，出現錯誤，請通知IT \r\n" + ex.Message, ToString());
            }
        }
        #endregion        

        #region 針對寵物生日欄位格式化
        private void textBox1_Leave(object sender, EventArgs e)
        {
            /*使用者輸入 20201010，離開欄位後自動轉換 2020/10/10 */
            DateTime dt = DateTime.ParseExact(textBox1.Text, "yyyyMMdd", System.Globalization.CultureInfo.InvariantCulture);
            var result = dt.ToString("yyyy/MM/dd");
            textBox1.Text = result;
        }
        #endregion

        #region 獲得寵物編號
        private int Get_pet_index()
        {
            int _pets_index = 0;
            try
            {
                using (OleDbConnection conn = new OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=GFS.accdb;Persist Security Info=False;"))
                {
                    conn.Open();
                    string sqlcmd = "select pets_index from Pets where pets_names ='" + textBox8 + "' and pets_parent_index = '" + parent_index + "'";
                    OleDbCommand cmd = new OleDbCommand(sqlcmd, conn);
                    OleDbDataReader dr = cmd.ExecuteReader();
                    while (dr.Read())
                    {
                        _pets_index = dr.GetInt32(0);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("獲得寵物編號時發生錯誤：" + ex.Message.ToString());
            }

            return _pets_index;
        }
        #endregion
        
    }
}



/*Create_PetsSalon_datas*/
using System;
using System.Windows.Forms;
/**/
using System.Data.OleDb;

namespace GPS
{
    public partial class Create_Pets_datas : Form
    {
        #region 必要物件
        //儲存寵物編號
        int pets_index = 0;
        //Get家長編號
        int parent_index = 0;
        #endregion

        public Create_Pets_datas(int parents_index)
        {
            InitializeComponent();
            //接家長的編號
            parent_index = parents_index;
        }

        private void button1_Click(object sender, EventArgs e)
        {            

        }

        private void create_pets_data(int pets_sex, string pets_sstatus, string precaution, string allergy_txt, string notouch_txt, string exp_txt, int medical_treatment)
        {
            #region 建立寵物基本資料
            try
            {
                // insert Pets
                using (OleDbConnection conn = new OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=GFS.accdb;Persist Security Info=False;"))
                {
                    conn.Open();
                    string sqlcmd = "insert into Pets (pets_names, pets_breed, pets_sex, pets_birthDay, pets_weight, pets_ligation, pets_sstatus, pets_precautions, pets_allergy, pets_notouch, pets_experiences, pets_parent_index, pets_createdate ) values ( @pets_names, @pets_breed, @pets_sex, @pets_birthDay, @pets_weight, @pets_ligation, @pets_sstatus, @pets_precautions, @pets_allergy, @pets_notouch, @pets_experiences, @pets_parent_index, @pets_createdate )";
                    OleDbCommand cmd = new OleDbCommand(sqlcmd, conn);
                    #region SQL injection prevent
                    cmd.Parameters.Add("@pets_names", OleDbType.VarChar);
                    cmd.Parameters.Add("@pets_breed", OleDbType.VarChar);
                    cmd.Parameters.Add("@pets_sex", OleDbType.VarChar);
                    cmd.Parameters.Add("@pets_birthDay", OleDbType.VarChar);
                    cmd.Parameters.Add("@pets_weight", OleDbType.VarChar);
                    cmd.Parameters.Add("@pets_ligation", OleDbType.VarChar);
                    cmd.Parameters.Add("@pets_sstatus", OleDbType.VarChar);
                    cmd.Parameters.Add("@pets_precautions", OleDbType.VarChar);
                    cmd.Parameters.Add("@pets_allergy", OleDbType.VarChar);
                    cmd.Parameters.Add("@pets_notouch", OleDbType.VarChar);
                    cmd.Parameters.Add("@pets_experiences", OleDbType.VarChar);
                    cmd.Parameters.Add("@pets_parent_index", OleDbType.VarChar);
                    cmd.Parameters.Add("@pets_createdate", OleDbType.Date);

                    DateTime date1 = DateTime.Now; // yyyy/MM/DD HH:MM:SS

                    cmd.Parameters["@pets_names"].Value = textBox8.Text.Trim();
                    cmd.Parameters["@pets_breed"].Value = comboBox1.SelectedItem.ToString();
                    cmd.Parameters["@pets_sex"].Value = pets_sex;
                    cmd.Parameters["@pets_birthDay"].Value = textBox1.Text.Trim();
                    cmd.Parameters["@pets_weight"].Value = textBox2.Text.Trim();
                    cmd.Parameters["@pets_ligation"].Value = pets_sex;
                    cmd.Parameters["@pets_sstatus"].Value = pets_sstatus;
                    cmd.Parameters["@pets_precautions"].Value = precaution;
                    cmd.Parameters["@pets_allergy"].Value = allergy_txt;
                    cmd.Parameters["@pets_notouch"].Value = notouch_txt;
                    cmd.Parameters["@pets_experiences"].Value = exp_txt;
                    cmd.Parameters["@pets_parent_index"].Value = parent_index;
                    cmd.Parameters["@pets_createdate"].Value = date1;
                    #endregion
                    cmd.ExecuteNonQuery();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("新增資料到寵物基本資料時發生錯誤：" + ex.Message.ToString());
            }
            #endregion

            #region 獲得寵物的編號
            try
            {
                int p_i = Get_pet_index();
                pets_index = Get_pet_index();
            }
            catch (Exception ex)
            {
                MessageBox.Show("獲得寵物編號時發生錯誤：" + ex.Message.ToString());
            }
            #endregion

            #region 建立寵物健康狀況
            try
            {
                //insert Pets_info
                using (OleDbConnection conn = new OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=GFS.accdb;Persist Security Info=False;"))
                {
                    string sqlcmd = " insert into Pets_info (petsinfo_hospital, petsinfo_doctors, petsinfo_telephone, petsinfo_address, petsinfo_medicalHistory, petsinfo_medicalTreatement, petsinfo_vacciendate, petsinfo_rabiesvdate, petsinfo_pets_index, petsinfo_createdate) values ('" + textBox10.Text + "','" + textBox11.Text + "','" + textBox12.Text + "','" + textBox13.Text + "','" + textBox14.Text + "',true or false'" + textBox15.Text + "','" + textBox16.Text + "')";
                    OleDbCommand cmd = new OleDbCommand(sqlcmd, conn);

                    #region SQL injection prevent
                    cmd.Parameters.Add("@petsinfo_hospital", OleDbType.VarChar);
                    cmd.Parameters.Add("@petsinfo_doctors", OleDbType.VarChar);
                    cmd.Parameters.Add("@petsinfo_telephone", OleDbType.VarChar);
                    cmd.Parameters.Add("@petsinfo_address", OleDbType.VarChar);
                    cmd.Parameters.Add("@petsinfo_medicalHistory", OleDbType.VarChar);
                    cmd.Parameters.Add("@petsinfo_medicalTreatement", OleDbType.VarChar);
                    cmd.Parameters.Add("@petsinfo_vacciendate", OleDbType.VarChar);
                    cmd.Parameters.Add("@petsinfo_rabiesvdate", OleDbType.VarChar);
                    cmd.Parameters.Add("@petsinfo_pets_index", OleDbType.VarChar);
                    cmd.Parameters.Add("@petsinfo_createdate", OleDbType.VarChar);

                    DateTime date1 = DateTime.Now; // yyyy/MM/DD HH:MM:SS

                    cmd.Parameters["@petsinfo_hospital"].Value = textBox10.Text.Trim();
                    cmd.Parameters["@petsinfo_doctors"].Value = textBox11.Text.Trim();
                    cmd.Parameters["@petsinfo_telephone"].Value = textBox12.Text.Trim();
                    cmd.Parameters["@petsinfo_address"].Value = textBox13.Text.Trim();
                    cmd.Parameters["@petsinfo_medicalHistory"].Value = textBox14.Text.Trim();
                    cmd.Parameters["@petsinfo_medicalTreatement"].Value = medical_treatment;
                    cmd.Parameters["@petsinfo_vacciendate"].Value = textBox15.Text.Trim();
                    cmd.Parameters["@petsinfo_rabiesvdate"].Value = textBox16.Text.Trim();
                    cmd.Parameters["@petsinfo_pets_index"].Value = pets_index;
                    cmd.Parameters["@petsinfo_createdate"].Value = date1;
                    #endregion

                    cmd.ExecuteNonQuery();

                    if (MessageBox.Show("寵物基本資料建立成功！ \r\n繼續建立寵物美容資料請按「是」，要離開請按「否」", "訊息提示", MessageBoxButtons.YesNo, MessageBoxIcon.Information) == DialogResult.Yes)
                    {
                        Create_PetsSalon_datas createdatas3 = new Create_PetsSalon_datas(pets_index, textBox8.Text.Trim());
                        createdatas3.Show();
                        this.Close();
                    }
                    else
                    {
                        this.Close();
                    }
                    this.Close();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("新增寵物健康狀況時發生錯誤：" + ex.Message.ToString());
            }
            #endregion            
        }

        #region 檢核欄位並產生值
        private void status_check()
        {
            int pets_sex = 2;
            string pets_sstatus = null;
            string precaution = null;
            string allergy_txt = null;
            string notouch_txt = null;
            string exp_txt = null;
            int medical_treatment = 2;

            try
            {
                if (!radioButton3.Checked && !radioButton4.Checked)
                {
                    MessageBox.Show("請點選性別");
                }

                if (radioButton4.Checked) //radioButton4 = 公
                {
                    pets_sex = 1;
                }

                if (radioButton3.Checked)//radioButton3 = 母
                {
                    pets_sex = 0;
                }

                if (checkBox2.Checked)
                {
                    pets_sstatus = "良好";
                }
                else
                {
                    pets_sstatus = textBox3.Text;
                }

                if (checkBox3.Checked)
                {
                    precaution = "無";
                }
                else
                {
                    precaution = textBox4.Text;
                }

                if (checkBox4.Checked)
                {
                    allergy_txt = "無";
                }
                else if (checkBox5.Checked)
                {
                    allergy_txt += "電剪;";
                }
                else if (checkBox6.Checked)
                {
                    allergy_txt += "洗毛精;";
                }
                else if (checkBox7.Checked)
                {
                    allergy_txt += "其他:";
                    allergy_txt += textBox5.Text;
                }

                if (checkBox11.Checked)
                {
                    notouch_txt = "無";
                }
                else if (checkBox10.Checked)
                {
                    notouch_txt += "腳;";
                }
                else if (checkBox10.Checked)
                {
                    notouch_txt += "耳朵;";
                }
                else if (checkBox10.Checked)
                {
                    notouch_txt += "嘴;";
                }
                else if (checkBox10.Checked)
                {
                    notouch_txt += "其他:";
                    notouch_txt += textBox6.Text;
                }

                if (checkBox17.Checked)
                {
                    exp_txt = "無";
                }
                else
                {
                    exp_txt = textBox7.Text;
                }

                if (checkBox13.Checked)
                {
                    medical_treatment = 1;
                }
                else
                {
                    medical_treatment = 0;
                }

                create_pets_data(pets_sex, pets_sstatus, precaution, allergy_txt, notouch_txt, exp_txt, medical_treatment);
            }
            catch (Exception ex)
            {
                MessageBox.Show("建立資料時，出現錯誤，請通知IT \r\n" + ex.Message, ToString());
            }
        }
        #endregion        

        #region 針對寵物生日欄位格式化
        private void textBox1_Leave(object sender, EventArgs e)
        {
            /*使用者輸入 20201010，離開欄位後自動轉換 2020/10/10 */
            DateTime dt = DateTime.ParseExact(textBox1.Text, "yyyyMMdd", System.Globalization.CultureInfo.InvariantCulture);
            var result = dt.ToString("yyyy/MM/dd");
            textBox1.Text = result;
        }
        #endregion

        #region 獲得寵物編號
        private int Get_pet_index()
        {
            int _pets_index = 0;
            try
            {
                using (OleDbConnection conn = new OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=GFS.accdb;Persist Security Info=False;"))
                {
                    conn.Open();
                    string sqlcmd = "select pets_index from Pets where pets_names ='" + textBox8 + "' and pets_parent_index = '" + parent_index + "'";
                    OleDbCommand cmd = new OleDbCommand(sqlcmd, conn);
                    OleDbDataReader dr = cmd.ExecuteReader();
                    while (dr.Read())
                    {
                        _pets_index = dr.GetInt32(0);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("獲得寵物編號時發生錯誤：" + ex.Message.ToString());
            }

            return _pets_index;
        }
        #endregion
        
    }
}



/*Create_PetsSalon_datas*/
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.OleDb;

namespace GPS
{
    public partial class Create_PetsSalon_datas : Form
    {
        int pets_index = 0;
        string pets_name = null;

        public Create_PetsSalon_datas(int Pets_index, string Pets_name)
        {
            InitializeComponent();
            pets_index = Pets_index;
            pets_name = Pets_name;
            textBox1.Text = pets_name;//寵物的名子
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Save_Pets_Feature();
            Save_Cost_Items();
        }

        #region 寵物的特徵資料
        private void Save_Pets_Feature()
        {
            try
            {
                using (OleDbConnection conn = new OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=GFS.accdb;Persist Security Info=False;"))
                {
                    conn.Open();
                    string sqlcmd = "insert into Pets (pets_colors, pets_lotion, pets_skinConditions, pets_style, pets_comments, pets_pets_updatedate) values (@pets_colors, @pets_lotion, @pets_skinConditions, @pets_style, @pets_comments, @pets_pets_updatedate)";
                    OleDbCommand cmd = new OleDbCommand(sqlcmd, conn);

                    #region SQL Injection prevent
                    cmd.Parameters.Add("@pets_colors", OleDbType.VarChar);
                    cmd.Parameters.Add("@pets_lotion", OleDbType.VarChar);
                    cmd.Parameters.Add("@pets_skinConditions", OleDbType.VarChar);
                    cmd.Parameters.Add("@pets_style", OleDbType.VarChar);
                    cmd.Parameters.Add("@pets_comments", OleDbType.VarChar);
                    cmd.Parameters.Add("@pets_pets_updatedate", OleDbType.VarChar);

                    DateTime date1 = DateTime.Now; // yyyy/MM/DD HH:MM:SS

                    cmd.Parameters["@pets_colors"].Value = textBox2.Text.Trim();
                    cmd.Parameters["@pets_lotion"].Value = textBox3.Text.Trim();
                    cmd.Parameters["@pets_skinConditions"].Value = textBox4.Text.Trim();
                    cmd.Parameters["@pets_style"].Value = textBox6.Text.Trim();
                    cmd.Parameters["@pets_comments"].Value = textBox5.Text.Trim();
                    cmd.Parameters["@pets_pets_updatedate"].Value = date1;
                    #endregion
                    
                    cmd.ExecuteNonQuery();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("建立特徵資料時出現錯誤");
            }
        }
        #endregion

        #region 費用紀錄
        private void Save_Cost_Items()
        {
            try
            {
                using (OleDbConnection conn = new OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=GFS.accdb;Persist Security Info=False;"))
                {
                    conn.Open();
                    string sqlcmd = "insert into Costs (costs_washOnce, costs_salonOnce, costs_monthlySubScription, costs_bigSalon, costs_noSalon, costs_spatickets, costs_waterHealtickets, costs_pets_index, costs_createdate) values (@costs_washOnce, @costs_salonOnce, @costs_monthlySubScription, @costs_bigSalon, @costs_noSalon, @costs_spatickets, @costs_waterHealtickets, @costs_pets_index, @costs_createdate)";
                    OleDbCommand cmd = new OleDbCommand(sqlcmd, conn);
                    #region SQL injection prevent
                    cmd.Parameters.Add("@costs_washOnce", OleDbType.VarChar);
                    cmd.Parameters.Add("@costs_salonOnce", OleDbType.VarChar);
                    cmd.Parameters.Add("@costs_monthlySubScription", OleDbType.VarChar);
                    cmd.Parameters.Add("@costs_bigSalon", OleDbType.VarChar);
                    cmd.Parameters.Add("@costs_noSalon", OleDbType.VarChar);
                    cmd.Parameters.Add("@costs_spatickets", OleDbType.VarChar);
                    cmd.Parameters.Add("@costs_waterHealtickets", OleDbType.VarChar);
                    cmd.Parameters.Add("@costs_pets_index", OleDbType.VarChar);
                    cmd.Parameters.Add("@costs_createdate", OleDbType.VarChar);

                    DateTime date1 = DateTime.Now; // yyyy/MM/DD HH:MM:SS

                    cmd.Parameters["@costs_washOnce"].Value = textBox7.Text.Trim();
                    cmd.Parameters["@costs_salonOnce"].Value = textBox8.Text.Trim();
                    cmd.Parameters["@costs_monthlySubScription"].Value = textBox9.Text.Trim();
                    cmd.Parameters["@costs_bigSalon"].Value = textBox10.Text.Trim();
                    cmd.Parameters["@costs_noSalon"].Value = textBox11.Text.Trim();
                    cmd.Parameters["@costs_spatickets"].Value = textBox12.Text.Trim();
                    cmd.Parameters["@costs_waterHealtickets"].Value = textBox13.Text.Trim();
                    cmd.Parameters["@costs_pets_index"].Value = pets_index;
                    cmd.Parameters["@costs_createdate"].Value = date1;
                    #endregion
                    cmd.ExecuteNonQuery();
                }
            }
            catch (Exception ex)
            {

            }
        }
        #endregion 
    }
}


/*Login*/

using System;
using System.Data.OleDb;
using System.Windows.Forms;

namespace GPS
{
    public partial class Login : Form
    {
        MainPanel mainpanel1 = null;
        public static string username = null; 
        public Login(MainPanel fr1)
        {
            InitializeComponent();
            mainpanel1 = fr1;
        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            Register frm1 = new Register();
            frm1.Show();
        }

        private void ConnectAccess()
        {
            //OleDbConnection cn = new OleDbConnection(@"Provider=Microsoft.Ace.OLEDB.12.0;Data Source=C:\\Users\\YGE\\Desktop\\MIS\\MIS\\bin\\Debug\\GFS.accdb");
            var DBPath = "Provider=Microsoft.Ace.OLEDB.12.0;Data Source="+ Application.StartupPath + "\\GFS.accdb";
            using (OleDbConnection conn = new OleDbConnection(DBPath))
            {
                try
                {
                    conn.Open();
                    string sqlcmd = "select * from Employee where employee_account = '" + textBox1.Text + "'";
                    //string sqlcmd = "select * from Employee where employee_account = @account"

                    OleDbCommand cmd = new OleDbCommand(sqlcmd, conn);
                    cmd.Parameters.Add("@account", OleDbType.VarChar);
                    cmd.Parameters["@account"].Value = textBox1.Text;


                    OleDbDataReader dr = cmd.ExecuteReader();
                    if (dr.Read())
                    {
                        string confirm = dr.GetString(3);
                        username = dr.GetString(1);
                        if (confirm == textBox2.Text.Trim())
                        {
                            MessageBox.Show("Hi !  " + username);
                            mainpanel1.Controls["label2"].Text = username;
                            mainpanel1.Show();
                            this.Hide();
                        }
                        else
                        {
                            MessageBox.Show("密碼輸入錯誤");
                        }
                    }
                    else
                    {
                        MessageBox.Show("帳號密碼不存在，請重新輸入");
                        textBox1.Text = "";
                        textBox2.Text = "";
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("資料庫連線錯誤 \r\n" + ex.Message.ToString());
                }
                finally
                {

                }
            }
        }
        private void button1_Click(object sender, EventArgs e)
        {
            ConnectAccess();
        }
        private void Login_FormClosing(object sender, FormClosingEventArgs e)
        {
            mainpanel1.Close();
        }
    }
}
 
 
 /*MainPanel*/
 using System;
using System.Windows.Forms;

namespace GPS
{
    public partial class MainPanel : Form
    {
        public string username = null;
        
        public MainPanel()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void Form1_Shown(object sender, EventArgs e)
        {
            Login logfrm = new Login(this);
            logfrm.Show();
            this.Hide();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Create_Parents_datas createdatasfrm = new Create_Parents_datas();
            createdatasfrm.Show();
            //this.Hide();
        }
    }
}


/*Register*/

using System;
using System.Windows.Forms;
/**/
using System.Data.OleDb;

namespace GPS
{
    public partial class Register : Form
    {
        public Register()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string auth_level = null;
            switch (textBox4.Text)
            {
                case "GFS1":
                    //DO something
                    auth_level = "L1";
                    connectDB(auth_level);
                    break;
                case "GFS2":
                    //DO something
                    auth_level = "L2";
                    connectDB(auth_level);
                    break;
                case "GFS3":
                    //DO something
                    auth_level = "L3";
                    connectDB(auth_level);
                    break;
                case "":
                    MessageBox.Show("權限密碼欄位，不可以為空 \r\n 相關資訊請通知俊億");
                    break;
            }
        }

        public void connectDB(string auth_level)
        {
            try
            {
                using (OleDbConnection conn = new OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=GFS.accdb;Persist Security Info=False;"))
                {
                    string sqlcmd = "insert into Employee (employee_name, employee_account, employee_password, employee_authlevel) values('" + textBox3.Text.Trim() + "','" + textBox1.Text.Trim() + "','" + textBox2.Text.Trim() + "','" + auth_level + "')";
                    //textBox5.Text = sqlcmd;
                    conn.Open();
                    OleDbCommand cmd = new OleDbCommand(sqlcmd, conn);
                    cmd.ExecuteNonQuery();
                    MessageBox.Show("資料建立成功！ \r\n 請關閉註冊視窗，嘗試登入");
                    this.Close();
                }
            }
            catch(Exception ex)
            {
                MessageBox.Show("建立資料時，出現資料庫連線錯誤，請通知俊億 \r\n" +ex.Message,ToString());
            }
        }


    }
}
