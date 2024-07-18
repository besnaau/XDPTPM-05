using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.Runtime.InteropServices;


namespace QuanLyCay
{
    public partial class frmDMCay : Form
    {
        DataTable tblTree;
        public frmDMCay()
        {
            InitializeComponent();
        }

        private void frmDMCay_Load(object sender, EventArgs e)
        {
            txtMaCay.Enabled = false;
            LoadDataGridView();
            txtMaCay.Enabled = false;
            btnLuu.Enabled = false;
            btnBoQua.Enabled = false;
            ResetValues();
        }
        private void LoadDataGridView()
        {
            string sql;
            sql = "SELECT * FROM tblTree";
            tblTree = Functions.GetDataToTable(sql);
            dgvCay.DataSource = tblTree;
            dgvCay.Columns[0].HeaderText = "Mã Cây";
            dgvCay.Columns[1].HeaderText = "Tên Cây";
            dgvCay.Columns[2].HeaderText = "Số Lượng";
            dgvCay.Columns[3].HeaderText = "Đơn Giá Nhập";
            dgvCay.Columns[4].HeaderText = "Đơn Giá Bán";
            dgvCay.Columns[5].HeaderText = "Ảnh";
            dgvCay.Columns[6].HeaderText = "Ghi Chú";
            dgvCay.Columns[0].Width = 100;
            dgvCay.Columns[1].Width = 150;
            dgvCay.Columns[2].Width = 100;
            dgvCay.Columns[3].Width = 100;
            dgvCay.Columns[4].Width = 100;
            dgvCay.Columns[5].Width = 200;
            dgvCay.Columns[6].Width = 200;
            dgvCay.AllowUserToAddRows = false;
            dgvCay.EditMode = DataGridViewEditMode.EditProgrammatically;

        }

        private void ResetValues()
        {
            txtMaCay.Text = "";
            txtTenCay.Text = "";
            txtSLuong.Text = "0";
            txtDGNhap.Text = "0";
            txtDGBan.Text = "0";
            txtSLuong.Enabled = true;
            txtDGNhap.Enabled = false;
            txtDGBan.Enabled = false;
            txtAnh.Text = "";
            picAnh.Image = null;
            txtGhiChu.Text = "";
        }

        private void dgvCay_Click(object sender, EventArgs e)
        {
            string sql;
            if (btnThem.Enabled == false)
            {
                MessageBox.Show("Đang ở chế độ thêm mới!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                txtMaCay.Focus();
                return;
            }
            if (tblTree.Rows.Count == 0)
            {
                MessageBox.Show("Không có dữ liệu!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            txtMaCay.Text = dgvCay.CurrentRow.Cells["MaCay"].Value.ToString();
            txtTenCay.Text = dgvCay.CurrentRow.Cells["TenCay"].Value.ToString();
            txtSLuong.Text = dgvCay.CurrentRow.Cells["SoLuong"].Value.ToString();
            txtDGNhap.Text = dgvCay.CurrentRow.Cells["DonGiaNhap"].Value.ToString();
            txtDGBan.Text = dgvCay.CurrentRow.Cells["DonGiaBan"].Value.ToString();
            sql = "SELECT Anh FROM tblTree WHERE MaCay=N'" + txtMaCay.Text + "'";
            txtAnh.Text = Functions.GetFieldValues(sql);
            picAnh.Image = Image.FromFile(txtAnh.Text);
            sql = "SELECT GhiChu FROM tblTree WHERE MaCay = N'" + txtMaCay.Text + "'";
            txtGhiChu.Text = Functions.GetFieldValues(sql);
            btnSua.Enabled = true;
            btnXoa.Enabled = true;
            btnBoQua.Enabled = true;
        }

        private void btnThem_Click(object sender, EventArgs e)
        {
            btnSua.Enabled = false;
            btnXoa.Enabled = false;
            btnBoQua.Enabled = true;
            btnLuu.Enabled = true;
            btnThem.Enabled = false;
            ResetValues();
            txtMaCay.Enabled = true;
            txtMaCay.Focus();
            txtSLuong.Enabled = true;
            txtDGNhap.Enabled = true;
            txtDGBan.Enabled = true;
        }

        private void btnLuu_Click(object sender, EventArgs e)
        {
            string sql;
            if (txtMaCay.Text.Trim().Length == 0)
            {
                MessageBox.Show("Bạn phải nhập mã cây", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                txtMaCay.Focus();
                return;
            }
            if (txtTenCay.Text.Trim().Length == 0)
            {
                MessageBox.Show("Bạn phải nhập tên cây", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                txtTenCay.Focus();
                return;
            }
            
            sql = "SELECT MaCay FROM tblTree WHERE MaCay=N'" + txtMaCay.Text.Trim() + "'";
            if (Functions.CheckKey(sql))
            {
                MessageBox.Show("Mã cây này đã tồn tại, bạn phải chọn mã cây khác", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                txtMaCay.Focus();
                return;
            }
            sql = "INSERT INTO tblTree(MaCay,TenCay,SoLuong,DonGiaNhap, DonGiaBan,Anh,Ghichu) VALUES(N'"
                + txtMaCay.Text.Trim() + "',N'" + txtTenCay.Text.Trim() +
                "'," + txtSLuong.Text.Trim() + "," + txtDGNhap.Text +
                "," + txtDGBan.Text + ",'" + txtAnh.Text + "',N'" + txtGhiChu.Text.Trim() + "')";

            Functions.RunSQL(sql);
            LoadDataGridView();
            //ResetValues();
            btnXoa.Enabled = true;
            btnThem.Enabled = true;
            btnSua.Enabled = true;
            btnBoQua.Enabled = false;
            btnLuu.Enabled = false;
            txtMaCay.Enabled = false;
        }

        private void btnSua_Click(object sender, EventArgs e)
        {
            string sql;
            if (tblTree.Rows.Count == 0)
            {
                MessageBox.Show("Không còn dữ liệu", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            if (txtMaCay.Text == "")
            {
                MessageBox.Show("Bạn chưa chọn bản ghi nào", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                txtMaCay.Focus();
                return;
            }
            if (txtTenCay.Text.Trim().Length == 0)
            {
                MessageBox.Show("Bạn phải nhập tên cây", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                txtTenCay.Focus();
                return;
            }
            if (txtAnh.Text.Trim().Length == 0)
            {
                MessageBox.Show("Bạn phải chọn ảnh minh hoạ cho cây", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                txtAnh.Focus();
                return;
            }
            sql = "UPDATE tblTree SET TenCay=N'" + txtTenCay.Text.Trim().ToString() +
                "',SoLuong=" + txtSLuong.Text +
                ",Anh='" + txtAnh.Text + "',GhiChu=N'" + txtGhiChu.Text + "' WHERE MaCay=N'" + txtMaCay.Text + "'";
            Functions.RunSQL(sql);
            LoadDataGridView();
            ResetValues();
            btnBoQua.Enabled = false;
        }

        private void btnXoa_Click(object sender, EventArgs e)
        {
            string sql;
            if (tblTree.Rows.Count == 0)
            {
                MessageBox.Show("Không còn dữ liệu!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            if (txtMaCay.Text == "")
            {
                MessageBox.Show("Bạn chưa chọn bản ghi nào", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            if (MessageBox.Show("Bạn có muốn xoá bản ghi này không?", "Thông báo", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                sql = "DELETE tblTree WHERE MaCay=N'" + txtMaCay.Text + "'";
                Functions.RunSqlDel(sql);
                LoadDataGridView();
                ResetValues();
            }
        }

        private void btnBoQua_Click(object sender, EventArgs e)
        {
            ResetValues();
            btnXoa.Enabled = true;
            btnSua.Enabled = true;
            btnThem.Enabled = true;
            btnBoQua.Enabled = false;
            btnLuu.Enabled = false;
            txtMaCay.Enabled = false;
        }

        private void btnMo_Click(object sender, EventArgs e)
        {
            OpenFileDialog dlgOpen = new OpenFileDialog();
            dlgOpen.Filter = "Bitmap(*.bmp)|*.bmp|JPEG(*.jpg)|*.jpg|GIF(*.gif)|*.gif|All files(*.*)|*.*";
            dlgOpen.FilterIndex = 2;
            dlgOpen.Title = "Chọn ảnh minh hoạ cho cây";
            if (dlgOpen.ShowDialog() == DialogResult.OK)
            {
                picAnh.Image = Image.FromFile(dlgOpen.FileName);
                txtAnh.Text = dlgOpen.FileName;
            }
        }

        private void btnDong_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
