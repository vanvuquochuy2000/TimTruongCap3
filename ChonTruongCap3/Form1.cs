using ExcelDataReader;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ChonTruongCap3
{
    public partial class Form1 : Form
    {
        string s_file_ds_tab1 = "";
        string s_file_ds_tab2 = "";
        static int id_truong = 1;

        static int id_truong_tab2 = 1;

        List<tt_hoc_sinh> lst_tt_hoc_sinh_tab2 = new List<tt_hoc_sinh>();
        List<tt_truong> lst_tt_truong = new List<tt_truong>();
        List<tt_truong> lst_tt_truong_tab2 = new List<tt_truong>();
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (String.IsNullOrEmpty(s_file_ds_tab1))
            {
                MessageBox.Show("Vui lòng chọn danh sách học sinh");
                return;
            }
            progressBar1.Visible = true;
            listView1.Enabled = false;

            //clear item 
            listView1.Items.Clear();

            nhap_du_lieu_tu_excel(s_file_ds_tab1, "tab1");

            progressBar1.Visible = false;
            listView1.Enabled = true;

            groupBox2.Text = "Danh sách học sinh: " + listView1.Items.Count;
        }

        private void tb_ten_file_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Title = "Chọn danh sách học sinh";
            //openFileDialog.InitialDirectory = "c:\\";
            openFileDialog.Filter = "excel (*.xlsx,*.xls)|*.xlsx;*.xls";
            openFileDialog.FilterIndex = 2;
            openFileDialog.RestoreDirectory = true;

            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                //Get the path of specified file
                //filePathImportData = openFileDialog.FileName;
                tb_ten_file.Text = openFileDialog.FileName;
                s_file_ds_tab1 = openFileDialog.FileName;
            }
        }


        private void nhap_du_lieu_tu_excel(string s_file, string tab)
        {
            try
            {
                using (var stream = File.Open(s_file, FileMode.Open, FileAccess.Read))
                {

                    // Auto-detect format, supports:
                    //  - Binary Excel files (2.0-2003 format; *.xls)
                    //  - OpenXml Excel files (2007 format; *.xlsx)
                    using (var reader = ExcelReaderFactory.CreateReader(stream))
                    {

                        //config bỏ 9 sheet đầu
                        var conf = new ExcelDataSetConfiguration
                        {
                            ConfigureDataTable = _ => new ExcelDataTableConfiguration
                            {
                                //UseHeaderRow = true,
                                ReadHeaderRow = rowReader =>
                                {
                                    //next 9 rows
                                    rowReader.Read();
                                    rowReader.Read();
                                    rowReader.Read();
                                    rowReader.Read();
                                    rowReader.Read();
                                    rowReader.Read();
                                    rowReader.Read();
                                    rowReader.Read();
                                    rowReader.Read();
                                }
                            }
                        };

                        //lấy số lượng sheet
                        int so_luong_sheet = reader.AsDataSet().Tables.Count;

                        //đọc dữ liệu theo từng sheet để bỏ qua 9 rows đầu
                        for (int i = 0; i < so_luong_sheet; i++)
                        {
                            var dataTable = reader.AsDataSet(conf).Tables[i];

                            foreach (var row in dataTable.Rows)
                            {
                                string gioi_tinh = ((System.Data.DataRow)row)[3].ToString();
                                if (gioi_tinh.ToUpper() == "NAM" || gioi_tinh.ToUpper() == "NỮ")
                                {
                                    tt_hoc_sinh tt = new tt_hoc_sinh();
                                    tt.ho_ten = ((System.Data.DataRow)row)[2].ToString();
                                    tt.gioi_tinh = gioi_tinh;
                                    tt.toan = ((System.Data.DataRow)row)[5].ToString();
                                    tt.li = ((System.Data.DataRow)row)[6].ToString();
                                    tt.hoa = ((System.Data.DataRow)row)[7].ToString();
                                    tt.sinh = ((System.Data.DataRow)row)[8].ToString();
                                    tt.van = ((System.Data.DataRow)row)[9].ToString();
                                    tt.lich_su = ((System.Data.DataRow)row)[10].ToString();
                                    tt.dia_li = ((System.Data.DataRow)row)[11].ToString();
                                    tt.anh_van = ((System.Data.DataRow)row)[12].ToString();
                                    tt.gdcd = ((System.Data.DataRow)row)[13].ToString();
                                    tt.cong_nghe = ((System.Data.DataRow)row)[14].ToString();
                                    tt.the_duc = ((System.Data.DataRow)row)[15].ToString();
                                    tt.my_thuat = ((System.Data.DataRow)row)[16].ToString();
                                    tt.tu_chon = ((System.Data.DataRow)row)[17].ToString();
                                    tt.tbcm = ((System.Data.DataRow)row)[18].ToString();
                                    tt.xl_hl = ((System.Data.DataRow)row)[19].ToString();

                                    //vào listview tab1
                                    if (tab == "tab1")
                                    {
                                        listView1.Items.Add(new ListViewItem(new string[]
                                    {
                                            tt.ho_ten, tt.gioi_tinh, tt.toan, tt.li, tt.hoa, tt.sinh, tt.van, tt.lich_su, tt.dia_li, tt.anh_van, tt.gdcd,
                                            tt.cong_nghe, tt.the_duc, tt.my_thuat, tt.tu_chon, tt.tbcm, tt.xl_hl
                                    }));
                                    }

                                    //vào listview tab2
                                    if (tab == "tab2")
                                    {
                                        listView1_tab2.Items.Add(new ListViewItem(new string[]
                                        {
                                            tt.ho_ten, tt.gioi_tinh, tt.toan, tt.li, tt.hoa, tt.sinh, tt.van, tt.lich_su, tt.dia_li, tt.anh_van, tt.gdcd,
                                            tt.cong_nghe, tt.the_duc, tt.my_thuat, tt.tu_chon, tt.tbcm, tt.xl_hl
                                        }));

                                        //add vào list hs
                                        lst_tt_hoc_sinh_tab2.Add(tt);
                                    }

                                }
                            }
                        }

                    }

                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString(), "Lỗi import dữ liệu");
            }

        }

        private void listView1_SelectedIndexChanged(object sender, EventArgs e)
        {
            string ho_ten = null;
            double toan = 0.0;
            double van = 0.0;
            double tong_diem = 0.0;
            if (listView1.FocusedItem != null)
            {
                if (listView1.SelectedItems.Count == 0)
                {
                    return;
                }

                if (lst_tt_truong.Count == 0)
                {
                    MessageBox.Show("Vui lòng nhập dữ liệu các Trường");
                    return;
                }
                ho_ten = listView1.SelectedItems[0].SubItems[0].Text;
                toan = Double.Parse(listView1.SelectedItems[0].SubItems[2].Text);
                van = Double.Parse(listView1.SelectedItems[0].SubItems[6].Text);
                tong_diem = (toan * 2) + (van * 2);


                tb_ket_qua.Text = "";

                //lấy kết quả theo định nghĩa của thuật toán Kmeans.
                var tt_ket_qua = lst_tt_truong.Where(n => n.diem_chuan <= tong_diem).OrderByDescending(n => n.diem_chuan).FirstOrDefault();

                if (tt_ket_qua != null)
                {
                    tb_ket_qua.ForeColor = Color.Green;
                    tb_ket_qua.Text = ho_ten + " - " + tong_diem + "đ" + Environment.NewLine + tt_ket_qua.ten_truong + " - " + tt_ket_qua.diem_chuan + "đ";
                }
                else
                {
                    tb_ket_qua.ForeColor = Color.Red;
                    tb_ket_qua.Text = ho_ten + ":" + tong_diem + " - Không có trường phù hợp";
                }

            }

        }

        private void btn_them_truong_Click(object sender, EventArgs e)
        {
            if (String.IsNullOrEmpty(tb_ten_truong.Text))
            {
                MessageBox.Show("Vui lòng nhập tên trường");
                return;
            }
            if (String.IsNullOrEmpty(tb_diem_chuan.Text))
            {
                MessageBox.Show("Vui lòng nhập điểm chuẩn");
                return;
            }
            string ten_truong = tb_ten_truong.Text;
            string diem_chuan = tb_diem_chuan.Text;

            ////thêm vào listview
            //listView2.Items.Add(new ListViewItem(new string[] { id_truong.ToString(), ten_truong, diem_chuan }));            

            //tạo list thông tin trường 
            tt_truong tt = new tt_truong();
            tt.id = id_truong;
            tt.ten_truong = ten_truong;
            tt.diem_chuan = Convert.ToDouble(diem_chuan);

            //thêm vào list
            lst_tt_truong.Add(tt);

            //order by diem asc
            lst_tt_truong = lst_tt_truong.OrderBy(n => n.diem_chuan).ToList();

            //load listview2
            load_listview2("tab1");

            id_truong++;

            groupBox3.Text = "Danh sách trường: " + lst_tt_truong.Count;

        }

        private void load_listview2(string tab)
        {
            if (tab == "tab1")
            {
                if (lst_tt_truong.Count != 0)
                {
                    listView2.Items.Clear();
                    foreach (var item in lst_tt_truong)
                    {
                        listView2.Items.Add(new ListViewItem(new string[] { item.id.ToString(), item.ten_truong, item.diem_chuan.ToString() }));
                    }
                }
                else
                {
                    listView2.Items.Clear();
                }
            }

            if (tab == "tab2")
            {
                if (lst_tt_truong_tab2.Count != 0)
                {
                    listView2_tab2.Items.Clear();
                    foreach (var item in lst_tt_truong_tab2)
                    {
                        listView2_tab2.Items.Add(new ListViewItem(new string[] { item.id.ToString(), item.ten_truong, item.diem_chuan.ToString() }));
                    }
                }
                else
                {
                    listView2_tab2.Items.Clear();
                }
            }

        }

        private void btn_xoa_truong_Click(object sender, EventArgs e)
        {
            if (listView2.SelectedItems.Count != 0)
            {
                int id = Int32.Parse(listView2.SelectedItems[0].SubItems[0].Text);

                ////xóa trên listview
                //listView2.SelectedItems[0].Remove();

                //xóa trường trong list
                var tt_truong = lst_tt_truong.Where(n => n.id == id).FirstOrDefault();
                if (tt_truong != null)
                {
                    lst_tt_truong.Remove(tt_truong);

                    load_listview2("tab1");

                    groupBox3.Text = "Danh sách trường: " + lst_tt_truong.Count;
                }
            }
        }

        private void tb_ten_file_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                e.Effect = DragDropEffects.Copy;
            }
        }

        private void tb_ten_file_DragDrop(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                string[] fileNames = (string[])e.Data.GetData(DataFormats.FileDrop);
                if (fileNames[0].Split('.').Last().ToString().ToUpper() == "XLS" || fileNames[0].Split('.').Last().ToString().ToUpper() == "XLSX")
                {
                    tb_ten_file.Lines = fileNames;
                    s_file_ds_tab1 = fileNames[0];
                }
                else
                {
                    MessageBox.Show("Vui lòng chọn file Excel");
                }

            }
        }

        private void tb_diem_chuan_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) && (e.KeyChar != '.'))
            {
                e.Handled = true;
            }
        }

        private void tb_ten_file_tab2_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Title = "Chọn danh sách học sinh";
            //openFileDialog.InitialDirectory = "c:\\";
            openFileDialog.Filter = "excel (*.xlsx,*.xls)|*.xlsx;*.xls";
            openFileDialog.FilterIndex = 2;
            openFileDialog.RestoreDirectory = true;

            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                //Get the path of specified file
                //filePathImportData = openFileDialog.FileName;
                tb_ten_file_tab2.Text = openFileDialog.FileName;
                s_file_ds_tab2 = openFileDialog.FileName;
            }
        }

        private void button1_tab2_Click(object sender, EventArgs e)
        {
            if (String.IsNullOrEmpty(s_file_ds_tab2))
            {
                MessageBox.Show("Vui lòng chọn danh sách học sinh");
                return;
            }

            //clear
            listView1_tab2.Items.Clear();
            lst_tt_hoc_sinh_tab2.Clear();

            progressBar1_tab2.Visible = true;
            listView1_tab2.Enabled = false;
            nhap_du_lieu_tu_excel(s_file_ds_tab2, "tab2");

            progressBar1_tab2.Visible = false;
            listView1_tab2.Enabled = true;

            groupBox4.Text = "Danh sách học sinh: " + listView1_tab2.Items.Count;
        }

        private void btn_them_truong_tab2_Click(object sender, EventArgs e)
        {
            if (String.IsNullOrEmpty(cbb_truong_tab2.Text))
            {
                MessageBox.Show("Vui lòng chọn trường");
                return;
            }
            if (String.IsNullOrEmpty(tb_diem_chuan_tab2.Text))
            {
                MessageBox.Show("Vui lòng nhập điểm chuẩn");
                return;
            }
            string ten_truong = cbb_truong_tab2.Text;
            string diem_chuan = tb_diem_chuan_tab2.Text;

            //tạo list thông tin trường 
            tt_truong tt = new tt_truong();
            tt.id = id_truong_tab2;
            tt.ten_truong = ten_truong;
            tt.diem_chuan = Convert.ToDouble(diem_chuan);

            //thêm vào list
            lst_tt_truong_tab2.Add(tt);

            //order by diem asc
            lst_tt_truong_tab2 = lst_tt_truong_tab2.OrderBy(n => n.diem_chuan).ToList();

            //load listview2
            load_listview2("tab2");

            id_truong_tab2++;

            groupBox5.Text = "Danh sách trường: " + lst_tt_truong_tab2.Count;
        }

        private void btn_xoa_truong_tab2_Click(object sender, EventArgs e)
        {
            if (listView2_tab2.SelectedItems.Count != 0)
            {
                int id = Int32.Parse(listView2_tab2.SelectedItems[0].SubItems[0].Text);

                ////xóa trên listview
                //listView2.SelectedItems[0].Remove();

                //xóa trường trong list
                var tt_truong = lst_tt_truong_tab2.Where(n => n.id == id).FirstOrDefault();
                if (tt_truong != null)
                {
                    lst_tt_truong_tab2.Remove(tt_truong);

                    load_listview2("tab2");

                    groupBox5.Text = "Danh sách trường: " + lst_tt_truong_tab2.Count;
                }
            }
        }

        private void tb_ten_file_tab2_DragDrop(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                string[] fileNames = (string[])e.Data.GetData(DataFormats.FileDrop);
                if (fileNames[0].Split('.').Last().ToString().ToUpper() == "XLS" || fileNames[0].Split('.').Last().ToString().ToUpper() == "XLSX")
                {
                    tb_ten_file_tab2.Lines = fileNames;
                    s_file_ds_tab2 = fileNames[0];
                }
                else
                {
                    MessageBox.Show("Vui lòng chọn file Excel");
                }

            }
        }

        private void tb_ten_file_tab2_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                e.Effect = DragDropEffects.Copy;
            }

        }

        private void btn_thuc_hien_tab2_Click(object sender, EventArgs e)
        {
            if (lst_tt_hoc_sinh_tab2.Count == 0)
            {
                MessageBox.Show("Vui lòng nhập dữ liệu học sinh");
                return;
            }

            if (lst_tt_truong_tab2.Count == 0)
            {
                MessageBox.Show("Vui lòng nhập dữ liệu trường");
                return;
            }

            tb_tan_binh_tab2.Text = "";
            tb_tan_phu_tab2.Text = "";
            tb_tay_thanh_tab2.Text = "";

            //lọc kết quả theo định nghĩa của thuật toán Kmeans.
            foreach (var item in lst_tt_hoc_sinh_tab2)
            {
                double tong_diem = (Convert.ToDouble(item.toan) * 2) + (Convert.ToDouble(item.van) * 2);
                var tt = lst_tt_truong_tab2.Where(n => n.diem_chuan <= tong_diem).OrderByDescending(n => n.diem_chuan).FirstOrDefault();
                if (tt != null)
                {
                    //có kết quả trường

                    switch (tt.ten_truong)
                    {
                        case "Trường THPT Tân Bình":
                            tb_tan_binh_tab2.Text += item.ho_ten + " - " + tong_diem + "đ" + Environment.NewLine;
                            break;
                        case "Trường THPT Tân Phú":
                            tb_tan_phu_tab2.Text += item.ho_ten + " - " + tong_diem + "đ" + Environment.NewLine;
                            break;
                        case "Trường THPT Tây Thạnh":
                            tb_tay_thanh_tab2.Text += item.ho_ten + " - " + tong_diem + "đ" + Environment.NewLine;
                            break;
                    }
                }
                else
                {
                    //không có trường
                    tb_khac_tab2.Text += item.ho_ten + " - " + tong_diem + "đ" + Environment.NewLine;
                }
            }

            lbl_sl_tan_binh.Text = tb_tan_binh_tab2.Lines.Count() != 0 ? (tb_tan_binh_tab2.Lines.Count() - 1).ToString() : "0";
            lbl_sl_tan_phu.Text = tb_tan_phu_tab2.Lines.Count() != 0 ? (tb_tan_phu_tab2.Lines.Count() - 1).ToString() : "0";
            lbl_sl_tay_thanh.Text = tb_tay_thanh_tab2.Lines.Count() != 0 ? (tb_tay_thanh_tab2.Lines.Count() - 1).ToString() : "0";
            lbl_khac_tab2.Text = tb_khac_tab2.Lines.Count() != 0 ? (tb_khac_tab2.Lines.Count() - 1).ToString() : "0";

        }
    }
}
