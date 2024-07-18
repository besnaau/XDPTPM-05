using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;


namespace QuanLyCay
{
    public partial class frmMain : Form
    {
        public frmMain()
        {
            InitializeComponent();
        }

        private void chấtLiệuToolStripMenuItem_Click(object sender, EventArgs e)
        {
            frmDMNhanVien frm = new frmDMNhanVien();
            frm.MdiParent = this;
            frm.Show();
        }

        private void frmMain_Load(object sender, EventArgs e)
        {
            Functions.Connect();
        }

        private void mnuThoat_Click(object sender, EventArgs e)
        {
            Functions.Disconnect();
            Application.Exit();
        }

        private void mnuCay_Click(object sender, EventArgs e)
        {
            frmDMCay frm = new frmDMCay();
            frm.MdiParent = this;
            frm.Show();
        }

        private void mnuKhachHang_Click(object sender, EventArgs e)
        {
            frmDMKhachHang frm = new frmDMKhachHang(); 
            frm.MdiParent = this;
            frm.Show();
        }

        private void mnuHoaDonBan_Click(object sender, EventArgs e)
        {
        }

        private void hóaĐơnToolStripMenuItem_Click(object sender, EventArgs e)
        {
           
        }

        private void hóaĐơnToolStripMenuItem1_Click(object sender, EventArgs e)
        {

        }

        private void mnuHDBH_Click(object sender, EventArgs e)
        {
            frmHoaDonBan frm =new frmHoaDonBan();
            frm.MdiParent = this;
            frm.Show();
        }

        private void trợGiúpToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        private void mnuTKHoaDon_Click(object sender, EventArgs e)
        {
            frmTimKiemHD frm = new frmTimKiemHD();
            frm.MdiParent = this;   
            frm.Show();
        }

        private void danhMụcToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }
    }

      
}
