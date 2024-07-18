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
using COMExcel = Microsoft.Office.Interop.Excel;

namespace QuanLyCay
{

    public partial class frmHoaDonBan : Form
    {
        DataTable tblChiTietHDBan;
        public frmHoaDonBan()
        {
            InitializeComponent();
        }

        private void label7_Click(object sender, EventArgs e)
        {

        }

        private void label6_Click(object sender, EventArgs e)
        {

        }

        private void splitContainer1_Panel1_Paint(object sender, PaintEventArgs e)
        {

        }

        private void frmHoaDonBan_Load(object sender, EventArgs e)
        {
            btnThemHoaDon.Enabled = true;
            btnLuu.Enabled = false;
            btnHuyHoaDon.Enabled = false;
            btnInHoaDon.Enabled = false;
            txtMaHDBan.ReadOnly = true;
            txtTenNV.ReadOnly = true;
            txtTenKH.ReadOnly = true;
            txtDiaChi.ReadOnly = true;
            mtbDienThoai.ReadOnly = true;
            txtTenCay.ReadOnly = true;
            txtDonGia.ReadOnly = true;
            txtThanhTien.ReadOnly = true;
            txtTongTien.ReadOnly = true;
            txtTongTien.Text = "0";

            Functions.FillCombo("SELECT MaKhach, TenKhach FROM tblKhach", cbxMaKH, "MaKhach", "MaKhach");
            cbxMaKH.SelectedIndex = -1;
            Functions.FillCombo("SELECT MaNV, TenNV FROM tblNhanVien", cbxMaNV, "MaNV", "TenKhach");
            cbxMaNV.SelectedIndex = -1;
            Functions.FillCombo("SELECT MaCay, TenCay FROM tblTree", cbxMaCay, "MaCay", "MaCay");
            cbxMaCay.SelectedIndex = -1;
            //Hiển thị thông tin của một hóa đơn được gọi từ form tìm kiếm
            if (txtMaHDBan.Text != "")
            {
                LoadInforHoaDon();
                btnHuyHoaDon.Enabled = true;
                btnInHoaDon.Enabled = true;
            }
            LoadDataGridView();
        }

        private void ResetValues()
        {
            txtMaHDBan.Text = "";
            dtpNgayBan.Value = DateTime.Now;
            cbxMaNV.Text = "";
            cbxMaKH.Text = "";
            txtTongTien.Text = "0";
            cbxMaCay.Text = "";
            txtSLuong.Text = "";
            txtThanhTien.Text = "0";

        }

        private void LoadDataGridView()
        {
            string sql;
            sql = "SELECT a.MaCay, b.TenCay, a.SoLuong, b.DonGiaBan ,a.ThanhTien FROM tblChiTietHDBan AS a, tblTree AS b WHERE a.MaHDBan = N'" + txtMaHDBan.Text + "' AND a.MaCay=b.MaCay";
            tblChiTietHDBan = Functions.GetDataToTable(sql);
            dgvHDBanHang.DataSource = tblChiTietHDBan;
            dgvHDBanHang.Columns[0].HeaderText = "Mã Cây";
            dgvHDBanHang.Columns[1].HeaderText = "Tên Cây";
            dgvHDBanHang.Columns[2].HeaderText = "Số lượng";
            dgvHDBanHang.Columns[3].HeaderText = "Đơn giá";
            dgvHDBanHang.Columns[4].HeaderText = "Thành tiền";
            dgvHDBanHang.Columns[0].Width = 100;
            dgvHDBanHang.Columns[1].Width = 130;
            dgvHDBanHang.Columns[2].Width = 100;
            dgvHDBanHang.Columns[3].Width = 100;
            dgvHDBanHang.Columns[4].Width = 100;
            dgvHDBanHang.AllowUserToAddRows = false;
            dgvHDBanHang.EditMode = DataGridViewEditMode.EditProgrammatically;
        }
       

        private void btnInHoaDon_Click(object sender, EventArgs e)
        {
            // Khởi động chương trình Excel
            COMExcel.Application exApp = new COMExcel.Application();
            COMExcel.Workbook exBook; //Trong 1 chương trình Excel có nhiều Workbook
            COMExcel.Worksheet exSheet; //Trong 1 Workbook có nhiều Worksheet
            COMExcel.Range exRange;
            string sql;
            int hang = 0, cot = 0;
            DataTable tblThongtinHD, tblThongtinHang;
            exBook = exApp.Workbooks.Add(COMExcel.XlWBATemplate.xlWBATWorksheet);
            exSheet = exBook.Worksheets[1];
            // Định dạng chung
            exRange = exSheet.Cells[1, 1];
            exRange.Range["A1:Z300"].Font.Name = "Times new roman"; //Font chữ
            exRange.Range["A1:B3"].Font.Size = 10;
            exRange.Range["A1:B3"].Font.Bold = true;
            exRange.Range["A1:B3"].Font.ColorIndex = 5; //Màu xanh da trời
            exRange.Range["A1:A1"].ColumnWidth = 7;
            exRange.Range["B1:B1"].ColumnWidth = 15;
            exRange.Range["A1:B1"].MergeCells = true;
            exRange.Range["A1:B1"].HorizontalAlignment = COMExcel.XlHAlign.xlHAlignCenter;
            exRange.Range["A1:B1"].Value = "Nhà Vườn ABC";
            exRange.Range["A2:B2"].MergeCells = true;
            exRange.Range["A2:B2"].HorizontalAlignment = COMExcel.XlHAlign.xlHAlignCenter;
            exRange.Range["A2:B2"].Value = "Gia Lâm - Hà Nội";
            exRange.Range["A3:B3"].MergeCells = true;
            exRange.Range["A3:B3"].HorizontalAlignment = COMExcel.XlHAlign.xlHAlignCenter;
            exRange.Range["A3:B3"].Value = "Điện thoại: (09)38526419";
            exRange.Range["C2:E2"].Font.Size = 16;
            exRange.Range["C2:E2"].Font.Bold = true;
            exRange.Range["C2:E2"].Font.ColorIndex = 3; //Màu đỏ
            exRange.Range["C2:E2"].MergeCells = true;
            exRange.Range["C2:E2"].HorizontalAlignment = COMExcel.XlHAlign.xlHAlignCenter;
            exRange.Range["C2:E2"].Value = "HÓA ĐƠN BÁN";
            // Biểu diễn thông tin chung của hóa đơn bán
            sql = "SELECT a.MaHDBan, a.NgayBan, a.TongTien, b.TenKhach, b.DiaChi, b.DienThoai, c.TenNV FROM tblHDBanHang AS a, tblKhach AS b, tblNhanVien AS c WHERE a.MaHDBan = N'" + txtMaHDBan.Text + "' AND a.MaKhach = b.MaKhach AND a.MaNV = c.MaNV";
            tblThongtinHD = Functions.GetDataToTable(sql);
            exRange.Range["B6:C9"].Font.Size = 12;
            exRange.Range["B6:B6"].Value = "Mã hóa đơn:";
            exRange.Range["C6:E6"].MergeCells = true;
            exRange.Range["C6:E6"].Value = tblThongtinHD.Rows[0][0].ToString();
            exRange.Range["B7:B7"].Value = "Khách hàng:";
            exRange.Range["C7:E7"].MergeCells = true;
            exRange.Range["C7:E7"].Value = tblThongtinHD.Rows[0][3].ToString();
            exRange.Range["B8:B8"].Value = "Địa chỉ:";
            exRange.Range["C8:E8"].MergeCells = true;
            exRange.Range["C8:E8"].Value = tblThongtinHD.Rows[0][4].ToString();
            exRange.Range["B9:B9"].Value = "Điện thoại:";
            exRange.Range["C9:E9"].MergeCells = true;
            exRange.Range["C9:E9"].Value = tblThongtinHD.Rows[0][5].ToString();
            //Lấy thông tin các mặt hàng
            sql = "SELECT b.TenCay, a.SoLuong, b.DonGiaBan, a.ThanhTien " +
                  "FROM tblChiTietHDBan AS a , tblTree AS b WHERE a.MaHDBan = N'" +
                  txtMaHDBan.Text + "' AND a.MaCay = b.MaCay";
            tblThongtinHang = Functions.GetDataToTable(sql);
            //Tạo dòng tiêu đề bảng
            exRange.Range["A11:F11"].Font.Bold = true;
            exRange.Range["A11:F11"].HorizontalAlignment = COMExcel.XlHAlign.xlHAlignCenter;
            exRange.Range["C11:F11"].ColumnWidth = 12;
            exRange.Range["A11:A11"].Value = "STT";
            exRange.Range["B11:B11"].Value = "Tên cây";
            exRange.Range["C11:C11"].Value = "Số lượng";
            exRange.Range["D11:D11"].Value = "Đơn giá";
            exRange.Range["E11:E11"].Value = "Thành tiền";
            for (hang = 0; hang < tblThongtinHang.Rows.Count; hang++)
            {
                //Điền số thứ tự vào cột 1 từ dòng 12
                exSheet.Cells[1][hang + 12] = hang + 1;
                for (cot = 0; cot < tblThongtinHang.Columns.Count; cot++)
                //Điền thông tin hàng từ cột thứ 2, dòng 12
                {
                    exSheet.Cells[cot + 2][hang + 12] = tblThongtinHang.Rows[hang][cot].ToString();
                    if (cot == 3) exSheet.Cells[cot + 2][hang + 12] = tblThongtinHang.Rows[hang][cot].ToString() ;
                }
            }
            exRange = exSheet.Cells[cot][hang + 14];
            exRange.Font.Bold = true;
            exRange.Value2 = "Tổng tiền:";
            exRange = exSheet.Cells[cot + 1][hang + 14];
            exRange.Font.Bold = true;
            exRange.Value2 = tblThongtinHD.Rows[0][2].ToString();
            exRange = exSheet.Cells[1][hang + 15]; //Ô A1 
            exRange.Range["A1:F1"].MergeCells = true;
            exRange.Range["A1:F1"].Font.Bold = true;
            exRange.Range["A1:F1"].Font.Italic = true;
            exRange.Range["A1:F1"].HorizontalAlignment = COMExcel.XlHAlign.xlHAlignRight;
            exRange = exSheet.Cells[4][hang + 17]; //Ô A1 
            exRange.Range["A1:C1"].MergeCells = true;
            exRange.Range["A1:C1"].Font.Italic = true;
            exRange.Range["A1:C1"].HorizontalAlignment = COMExcel.XlHAlign.xlHAlignCenter;
            DateTime d = Convert.ToDateTime(tblThongtinHD.Rows[0][1]);
            exRange.Range["A1:C1"].Value = "Hà Nội, ngày " + d.Day + " tháng " + d.Month + " năm " + d.Year;
            exRange.Range["A2:C2"].MergeCells = true;
            exRange.Range["A2:C2"].Font.Italic = true;
            exRange.Range["A2:C2"].HorizontalAlignment = COMExcel.XlHAlign.xlHAlignCenter;
            exRange.Range["A2:C2"].Value = "Nhân viên bán hàng";
            exRange.Range["A6:C6"].MergeCells = true;
            exRange.Range["A6:C6"].Font.Italic = true;
            exRange.Range["A6:C6"].HorizontalAlignment = COMExcel.XlHAlign.xlHAlignCenter;
            exRange.Range["A6:C6"].Value = tblThongtinHD.Rows[0][6];
            exSheet.Name = "Hóa đơn nhập";
            exApp.Visible = true;
        }

        private void btnThemHoaDon_Click(object sender, EventArgs e)
        {
            btnHuyHoaDon.Enabled = false;
            btnLuu.Enabled = true;
            btnInHoaDon.Enabled = false;
            btnThemHoaDon.Enabled = false;
            ResetValues();
            txtMaHDBan.Text = Functions.CreateKey("HDB");
            LoadDataGridView();
        }

        private void btnLuu_Click(object sender, EventArgs e)
        {
            string sql;
            double sl, SLcon, tong, Tongmoi;
            sql = "SELECT MaHDBan FROM tblHDBanHang WHERE MaHDBan=N'" + txtMaHDBan.Text + "'";
            if (!Functions.CheckKey(sql))
            {
                // Mã hóa đơn chưa có, tiến hành lưu các thông tin chung
                // Mã HDBan được sinh tự động do đó không có trường hợp trùng khóa

                if (cbxMaNV.Text.Length == 0)
                {
                    MessageBox.Show("Bạn phải nhập nhân viên", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    cbxMaNV.Focus();
                    return;
                }
                if (cbxMaKH.Text.Length == 0)
                {
                    MessageBox.Show("Bạn phải nhập khách hàng", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    cbxMaKH.Focus();
                    return;
                }
                sql = "INSERT INTO tblHDBanHang(MaHDBan, NgayBan, MaNV, MaKhach, TongTien) VALUES (N'" + txtMaHDBan.Text.Trim() + "','" +
                        dtpNgayBan.Value + "',N'" + cbxMaNV.SelectedValue + "',N'" +
                        cbxMaKH.SelectedValue + "'," + txtTongTien.Text + ")";
                Functions.RunSQL(sql);
            }
            // Lưu thông tin của các mặt hàng
            if (cbxMaCay.Text.Trim().Length == 0)
            {
                MessageBox.Show("Bạn phải nhập mã cây", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                cbxMaCay.Focus();
                return;
            }
            if ((txtSLuong.Text.Trim().Length == 0) || (txtSLuong.Text == "0"))
            {
                MessageBox.Show("Bạn phải nhập số lượng", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                txtSLuong.Text = "";
                txtSLuong.Focus();
                return;
            }
            sql = "SELECT MaCay FROM tblChiTietHDBan WHERE MaCay=N'" + cbxMaCay.SelectedValue + "' AND MaHDBan = N'" + txtMaHDBan.Text.Trim() + "'";
            if (Functions.CheckKey(sql))
            {
                MessageBox.Show("Mã cây này đã có, bạn phải nhập mã khác", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                ResetValuesCay();
                cbxMaCay.Focus();
                return;
            }
            // Kiểm tra xem số lượng cây trong kho còn đủ để cung cấp không?
            sl = Convert.ToDouble(Functions.GetFieldValues("SELECT SoLuong FROM tblTree WHERE MaCay = N'" + cbxMaCay.SelectedValue + "'"));
            if (Convert.ToDouble(txtSLuong.Text) > sl)
            {
                MessageBox.Show("Số lượng cây này chỉ còn " + sl, "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                txtSLuong.Text = "";
                txtSLuong.Focus();
                return;
            }
            sql = "INSERT INTO tblChiTietHDBan(MaHDBan,MaCay,SoLuong,DonGia,ThanhTien) VALUES(N'" + txtMaHDBan.Text.Trim() + "',N'" + cbxMaCay.SelectedValue + "'," + txtSLuong.Text + "," + txtDonGia.Text + "," + txtThanhTien.Text + ")";
            Functions.RunSQL(sql);
            LoadDataGridView();
            // Cập nhật lại số lượng của cây vào bảng tblTree
            SLcon = sl - Convert.ToDouble(txtSLuong.Text);
            sql = "UPDATE tblTree SET SoLuong =" + SLcon + " WHERE MaCay= N'" + cbxMaCay.SelectedValue + "'";
            Functions.RunSQL(sql);
            // Cập nhật lại tổng tiền cho hóa đơn bán
            tong = Convert.ToDouble(Functions.GetFieldValues("SELECT TongTien FROM tblHDBanHang WHERE MaHDBan = N'" + txtMaHDBan.Text + "'"));
            Tongmoi = tong + Convert.ToDouble(txtThanhTien.Text);
            sql = "UPDATE tblHDBanHang SET TongTien =" + Tongmoi + " WHERE MaHDBan = N'" + txtMaHDBan.Text + "'";
            Functions.RunSQL(sql);
            txtTongTien.Text = Tongmoi.ToString();
            ResetValuesCay();
            btnHuyHoaDon.Enabled = true;
            btnThemHoaDon.Enabled = true;
            btnInHoaDon.Enabled = true;
            LoadDataGridView();
        }

        private void ResetValuesCay()
        {
            cbxMaCay.Text = "";
            txtSLuong.Text = "";
            txtThanhTien.Text = "0";
        }

        private void cbxMaCay_SelectedIndexChanged(object sender, EventArgs e)
        {
            string str;
            if (cbxMaCay.Text == "")
            {
                txtTenCay.Text = "";
                txtDonGia.Text = "";
            }
            // Khi chọn mã hàng thì các thông tin về hàng hiện ra
            str = "SELECT TenCay FROM tblTree WHERE MaCay =N'" + cbxMaCay.SelectedValue + "'";
            txtTenCay.Text = Functions.GetFieldValues(str);
            str = "SELECT DonGiaBan FROM tblTree WHERE MaCay =N'" + cbxMaCay.SelectedValue + "'";
            txtDonGia.Text = Functions.GetFieldValues(str);
        }

        private void cbxMaKH_SelectedIndexChanged(object sender, EventArgs e)
        {
            string str;
            if (cbxMaKH.Text == "")
            {
                txtTenKH.Text = "";
                txtDiaChi.Text = "";
                mtbDienThoai.Text = "";
            }
            //Khi chọn Mã khách hàng thì các thông tin của khách hàng sẽ hiện ra
            str = "Select TenKhach from tblKhach where MaKhach = N'" + cbxMaKH.SelectedValue + "'";
            txtTenKH.Text = Functions.GetFieldValues(str);
            str = "Select DiaChi from tblKhach where MaKhach = N'" + cbxMaKH.SelectedValue + "'";
            txtDiaChi.Text = Functions.GetFieldValues(str);
            str = "Select DienThoai from tblKhach where MaKhach= N'" + cbxMaKH.SelectedValue + "'";
            mtbDienThoai.Text = Functions.GetFieldValues(str);
        }

        private void txtSLuong_TextChanged(object sender, EventArgs e)
        {
            //Khi thay đổi số lượng thì thực hiện tính lại thành tiền
            double tt, sl, dg;
            if (txtSLuong.Text == "")
                sl = 0;
            else
                sl = Convert.ToDouble(txtSLuong.Text);
            if (txtDonGia.Text == "")
                dg = 0;
            else
                dg = Convert.ToDouble(txtDonGia.Text);
            tt = sl * dg ;
            txtThanhTien.Text = tt.ToString();
        }

        private void btnDong_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void btnSearch_Click(object sender, EventArgs e)
        {
            if (cboMaHDBan.Text == "")
            {
                MessageBox.Show("Bạn phải chọn một mã hóa đơn để tìm", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                cboMaHDBan.Focus();
                return;
            }
            txtMaHDBan.Text = cboMaHDBan.Text;
            LoadInforHoaDon();
            LoadDataGridView();
            btnHuyHoaDon.Enabled = true;
            btnLuu.Enabled = true;
            btnInHoaDon.Enabled = true;
            cboMaHDBan.SelectedIndex = -1; 
        }

        private void LoadInforHoaDon()
        {
            string str;
            str = "SELECT NgayBan FROM tblHDBanHang WHERE MaHDBan = N'" + txtMaHDBan.Text + "'";
            dtpNgayBan.Text = Functions.ConvertDateTime(Functions.GetFieldValues(str));
            str = "SELECT MaNV FROM tblHDBanHang WHERE MaHDBan = N'" + txtMaHDBan.Text + "'";
            cbxMaNV.Text = Functions.GetFieldValues(str);
            str = "SELECT MaKhach FROM tblHDBanHang WHERE MaHDBan = N'" + txtMaHDBan.Text + "'";
            cbxMaKH.Text = Functions.GetFieldValues(str);
            str = "SELECT TongTien FROM tblHDBanHang WHERE MaHDBan = N'" + txtMaHDBan.Text + "'";
            txtTongTien.Text = Functions.GetFieldValues(str);
            LoadDataGridView();
        }

        private void btnHuyHoaDon_Click(object sender, EventArgs e)
        {
            ResetValues();
            LoadDataGridView();
        }

        private void cboMaHDBan_DropDown(object sender, EventArgs e)
        {
            Functions.FillCombo("SELECT MaHDBan FROM tblHDBanHang", cboMaHDBan, "MaHDBan", "MaHDBan");
            cboMaHDBan.SelectedIndex = -1;
        }

        private void frmHoaDonBan_FormClosing(object sender, FormClosingEventArgs e)
        {
            //Xóa dữ liệu trong các điều khiển trước khi đóng Form
            ResetValues();
        }

        private void txtSLuong_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (((e.KeyChar >= '0') && (e.KeyChar <= '9')) || (Convert.ToInt32(e.KeyChar) == 8))
                e.Handled = false;
            else e.Handled = true;
        }

        private void cboMaHDBan_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void txtTongTien_TextChanged(object sender, EventArgs e)
        {

        }
    }
}
