using System;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using MongoDB.Driver;
using MongoDB.Bson;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using PdfiumViewer;
using iText.Kernel.Pdf;
using iText.Kernel.Pdf.Canvas.Parser.Listener;
using iText.Kernel.Pdf.Canvas.Parser;

namespace BTL_DPT
{
    public partial class admin : Form
    {
        private IMongoCollection<BsonDocument> _reportsCollection;

        public admin()
        {
            InitializeComponent();
            ConnectToMongoDB();
            LoadReportsToListView();
        }

        // Kết nối với MongoDB
        private void ConnectToMongoDB()
        {
            var client = new MongoClient("mongodb://localhost:27017"); // Kết nối với MongoDB
            var database = client.GetDatabase("ReportsDB");  // Chọn database
            _reportsCollection = database.GetCollection<BsonDocument>("Reports");  // Chọn collection
        }

        // Quay lại form chính
        private void btnBack_Click(object sender, EventArgs e)
        {
            Form1 Form1 = new Form1();
            Form1.Show();
            this.Hide();
        }

        // Thêm báo cáo vào MongoDB
        private void btnThem_Click(object sender, EventArgs e)
        {
            // Mở file dialog để chọn file
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Word Documents (*.docx)|*.docx"; // Chọn .docx 
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                string filePath = openFileDialog.FileName;
                string fileName = Path.GetFileName(filePath); // Lấy tên file
                byte[] fileData = File.ReadAllBytes(filePath); // Đọc file dưới dạng byte

                // Kiểm tra xem tên file đã tồn tại trong MongoDB chưa
                if (IsFileNameExist(fileName))
                {
                    MessageBox.Show("File với tên này đã tồn tại trong MongoDB!", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                else
                {
                    // Lưu file vào MongoDB
                    SaveFileToMongoDB(fileName, fileData);
                }
            }
        }

        // Kiểm tra tên file đã tồn tại trong MongoDB chưa
        private bool IsFileNameExist(string fileName)
        {
            var filter = Builders<BsonDocument>.Filter.Eq("FileName", fileName);
            var existingFile = _reportsCollection.Find(filter).FirstOrDefault();
            return existingFile != null;
        }

        // Lưu file vào MongoDB
        private void SaveFileToMongoDB(string fileName, byte[] fileData)
        {
            try
            {
                // Tạo document MongoDB
                var reportDocument = new BsonDocument
                {
                    { "FileName", fileName },
                    { "FileData", new BsonBinaryData(fileData) }
                };

                // Chèn document vào MongoDB
                _reportsCollection.InsertOne(reportDocument);

                // Cập nhật ListView sau khi tải lên thành công
                LoadReportsToListView();

                MessageBox.Show("Báo cáo đã được tải lên MongoDB!");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Lỗi khi tải báo cáo: {ex.Message}");
            }
        }

        // Tải danh sách báo cáo từ MongoDB vào ListView
        private void LoadReportsToListView()
        {
            try
            {
                // Xóa tất cả item hiện tại trong ListView
                LoadReports.Items.Clear();

                // Lấy tất cả báo cáo từ MongoDB
                var reports = _reportsCollection.Find(new BsonDocument()).ToList();

                // Cập nhật cấu trúc cột của ListView nếu chưa có
                LoadReports.Columns.Clear();
                LoadReports.Columns.Add("File Name", 200); // Thêm cột tên file (có thể thay đổi kích thước tùy ý)
                LoadReports.Columns.Add("Size (Bytes)", 100);

                foreach (var report in reports)
                {
                    // Lấy tên file từ MongoDB
                    string fileName = report["FileName"].AsString;

                    // Kiểm tra trường "FileData" có kiểu dữ liệu là BsonBinaryData không
                    if (report["FileData"] is BsonBinaryData fileBinaryData)
                    {
                        // Lấy dữ liệu file (dưới dạng byte[])
                        byte[] fileData = fileBinaryData.Bytes;

                        // Lấy kích thước file (số byte)
                        int fileSize = fileData.Length;

                        // Thêm item vào ListView với tên file và kích thước file
                        var listViewItem = new ListViewItem(fileName);  // Tạo item với tên file
                        listViewItem.SubItems.Add(fileSize.ToString());  // Thêm kích thước file vào item
                        LoadReports.Items.Add(listViewItem);  // Thêm item vào ListView
                    }
                    else
                    {
                        MessageBox.Show($"Trường 'FileData' của báo cáo '{fileName}' không phải là kiểu BsonBinaryData.");
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Lỗi khi tải báo cáo: {ex.Message}");
            }
        }

        // Sự kiện khi người dùng nhấn vào nút Xóa
        private void btnXoa_Click(object sender, EventArgs e)
        {
            if (LoadReports.SelectedItems.Count > 0)
            {
                // Lấy tên file từ item được chọn trong ListView
                string selectedFileName = LoadReports.SelectedItems[0].Text;

                // Xác nhận xóa
                var confirmResult = MessageBox.Show($"Bạn có chắc chắn muốn xóa báo cáo '{selectedFileName}'?", "Xác nhận xóa", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
                if (confirmResult == DialogResult.Yes)
                {
                    // Xóa báo cáo trong MongoDB
                    DeleteFileFromMongoDB(selectedFileName);

                    // Cập nhật lại ListView
                    LoadReportsToListView();

                    MessageBox.Show("Báo cáo đã được xóa!");
                }
            }
            else
            {
                MessageBox.Show("Vui lòng chọn báo cáo để xóa.");
            }
        }

        // Xóa file khỏi MongoDB
        private void DeleteFileFromMongoDB(string fileName)
        {
            try
            {
                // Tạo filter để tìm báo cáo có tên trùng với tên file
                var filter = Builders<BsonDocument>.Filter.Eq("FileName", fileName);

                // Xóa báo cáo khỏi MongoDB
                _reportsCollection.DeleteOne(filter);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Lỗi khi xóa báo cáo: {ex.Message}");
            }
        }

        // Xử lý sự kiện khi người dùng chọn báo cáo trong ListView
        private void LoadReports_SelectedIndexChanged(object sender, EventArgs e)
        {
            // Không cần xử lý gì ở đây
        }

        // Đọc file .docx từ byte[]
        private string ReadDocxFile(byte[] fileData)
        {
            using (MemoryStream memoryStream = new MemoryStream(fileData))
            {
                using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(memoryStream, false))
                {
                    StringBuilder text = new StringBuilder();
                    Body body = wordDoc.MainDocumentPart.Document.Body;

                    // Duyệt qua tất cả các phần tử trong body và lấy text
                    foreach (var para in body.Elements<Paragraph>())
                    {
                        foreach (var run in para.Elements<Run>())
                        {
                            foreach (var textElement in run.Elements<Text>())
                            {
                                text.Append(textElement.Text);
                            }
                        }
                        text.AppendLine();  // Thêm xuống dòng sau mỗi đoạn văn
                    }
                    return text.ToString();  // Trả về nội dung văn bản
                }
            }
        }

        // Sự kiện khi nhấn nút "Xem"
        private void btnXem_Click(object sender, EventArgs e)
        {
            if (LoadReports.SelectedItems.Count > 0)
            {
                // Lấy tên file từ item được chọn trong ListView
                string selectedFileName = LoadReports.SelectedItems[0].Text;

                // Tìm báo cáo trong MongoDB theo tên file
                var report = _reportsCollection.Find(new BsonDocument("FileName", selectedFileName)).FirstOrDefault();

                if (report != null)
                {
                    // Kiểm tra xem trường "FileData" có kiểu dữ liệu là BsonBinaryData không
                    if (report["FileData"] is BsonBinaryData fileBinaryData)
                    {
                        // Lấy nội dung file (dưới dạng byte[]) từ MongoDB
                        byte[] fileData = fileBinaryData.Bytes;

                        string fileText = string.Empty;
                        
                        // Nếu là file docx, sử dụng phương thức ReadDocxFile
                        fileText = ReadDocxFile(fileData);
                       

                        // Hiển thị nội dung file trong RichTextBox hoặc TextBox
                        txtReportContent.Text = fileText;
                    }
                    else
                    {
                        MessageBox.Show("Dữ liệu không hợp lệ, không thể đọc file.");
                    }
                }
                else
                {
                    MessageBox.Show("Không tìm thấy báo cáo này trong MongoDB.");
                }
            }
            else
            {
                MessageBox.Show("Vui lòng chọn báo cáo để xem.");
            }
        }
    }
}
