namespace BTL_DPT
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void btnLogin_Click(object sender, EventArgs e)
        {
            // Kiểm tra tên người dùng và mật khẩu
            string username = txtUsername.Text;
            string password = txtPassword.Text;

            // Điều kiện kiểm tra: nếu tên = "admin" và mật khẩu = "12345"
            if (username == "admin" && password == "12345")
            {
                // Nếu đúng, chuyển sang form Admin
                admin adminForm = new admin();
                adminForm.Show();  // Mở form Admin
                this.Hide();  // Ẩn form đăng nhập
            }
            else
            {
                // Nếu sai, thông báo lỗi
                MessageBox.Show("Tên người dùng hoặc mật khẩu không đúng!", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnUpload_Click(object sender, EventArgs e)
        {
            
        }
    }
}
