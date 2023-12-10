using System;
using System.Data;
using System.Data.SqlClient;
using System.Globalization;
using System.Reflection;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
namespace BaseData1
{
    public partial class Form1 : Form
    {
        
        
        public Form1()
        {
            InitializeComponent();//Инициализация копонентов
        }


        

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }
        private void Form1_Load(object sender, EventArgs e)
        {
            

        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {

        }

        private void Budget_click_Click(object sender, EventArgs e)
        {
            string que = "select *,'DELETE' as Command from albums";
            Budget budget = new Budget(que, "albums");
            budget.Show();
            
            
        }

        private void Client_Click_Click(object sender, EventArgs e)
        { 
            string que = "select *,'DELETE' as Command from albums_songs";
            Budget budget = new Budget(que, "albums_songs");
            budget.Show();
        }

        private void Sotrudnik_Click_Click(object sender, EventArgs e)
        {
            string que = "select *,'DELETE' as Command from artist_in_group";
            Budget budget = new Budget(que, "artist_in_group");
            budget.Show();
        }

        private void Fillial_Click_Click(object sender, EventArgs e)
        {
            string que = "select *,'DELETE' as Command from genres";
            Budget budget = new Budget(que, "genres");
            budget.Show();
        }

        private void Costing_Click_Click(object sender, EventArgs e)
        {
            string que = "select *,'DELETE' as Command from Isp_Obor";
            Budget budget = new Budget(que, "Isp_Obor");
            budget.Show();
        }

        private void Sdelka_Click_Click(object sender, EventArgs e)
        { 
            string que = "select *,'DELETE' as Command from groups";
            Budget budget = new Budget(que, "groups");
            budget.Show();
        }

        private void Produkt_Click_Click(object sender, EventArgs e)
        {
            string que = "select *,'DELETE' as Command from Oborudovanie";
            Budget budget = new Budget(que, "Oborudovanie");
            budget.Show();
        }

        private void Pred_Prev_Budget_Click(object sender, EventArgs e)
        {
            string que = "select songs_title,album_title,genre_name from songs s join albums_songs sa on sa.songs_id=s.songs_id join albums al on al.album_id=sa.album_id join artists a on a.artist_id=al.artist_id join genres g on g.genre_id=a.genre_id";
            Budget budget = new Budget(que,"query");
            budget.Show();
        }

        private void Get_Kons_Sdelk_Click(object sender, EventArgs e)
        {
            string que = "select group_name,Oborudovanie,salary from groups g join Isp_Obor i on i.group_id=g.group_id join Oborudovanie ob on ob.Obor_id=i.Obor_id join salary s on s.group_id=g.group_id";
            Budget budget = new Budget(que,"query");
            budget.Show();
        }

        private void Rash_Pred_Click_Click(object sender, EventArgs e)
        {
            string que = "select group_name,last_name,first_name,birthday,pol from groups g join artist_in_group ag on ag.group_id=g.group_id join artists a on a.artist_id=ag.artist_id join persons p on p.artist_id=a.artist_id where LOWER(g.group_name)like 'сектор газа'";
            Budget budget = new Budget(que, "query");
            budget.Show();
        }

        private void Rash_Pred_and_Budget_Click(object sender, EventArgs e)
        {
            string que = "select group_name,album_title,[album\r\nalbum_tracks] from groups g join artist_in_group ag on ag.group_id=g.group_id join artists a on a.artist_id=ag.artist_id join albums al on al.artist_id=a.artist_id where Lower(group_name) like 'Ramschtein'";
            Budget budget = new Budget(que,"query");
            budget.Show();
        }

        private void button1_Click_1(object sender, EventArgs e)
        {

        }

        private void button1_Click_2(object sender, EventArgs e)
        {
            string que = "select country_name,genre_name,last_name,first_name,birthday,pol from artists a join countries c on c.country_id=a.country_id join genres g on g.genre_id=a.genre_id join persons p on p.artist_id=a.artist_id";
            Budget budget = new Budget(que,"query");
            budget.Show();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            string que = "select *,'DELETE' as Command from countries";
            Budget budget = new Budget(que, "countries");
            budget.Show();
        }

        private void button6_Click(object sender, EventArgs e)
        {
            string que = "select *,'DELETE' as Command from persons";
            Budget budget = new Budget(que, "persons");
            budget.Show();
        }

        private void button5_Click(object sender, EventArgs e)
        {
            string que = "select *,'DELETE' as Command from rider";
            Budget budget = new Budget(que, "rider");
            budget.Show();
        }

        private void button7_Click(object sender, EventArgs e)
        {
            string que = "select *,'DELETE' as Command from salary";
            Budget budget = new Budget(que, "salary");
            budget.Show();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            string que = "select *,'DELETE' as Command from songs";
            Budget budget = new Budget(que, "songs");
            budget.Show();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            string que = "select *,'DELETE' as Command from artists";
            Budget budget = new Budget(que, "artists");
            budget.Show();
        }

        private void button8_Click(object sender, EventArgs e)
        {
            string que = "select group_name, last_name, first_name, rider from groups g join artist_in_group ag on g.group_id = ag.group_id join artists a on a.artist_id = ag.artist_id join rider r on r.artist_id = a.artist_id join persons p on p.artist_id = a.artist_id";
            Budget budget = new Budget(que, "query");
            budget.Show();
            
        }

        private void button9_Click(object sender, EventArgs e)
        {
            string que = "select group_name,country_name from groups g join artist_in_group ag on ag.group_id=g.group_id join artists a on a.artist_id=ag.artist_id join countries c on c.country_id=a.country_id";
            Budget budget = new Budget(que, "query");
            budget.Show();
        }
    }
}