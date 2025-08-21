using System;
using System.Data.SQLite;
using System.Drawing;
using System.IO;
using System.Media;
using System.Windows.Forms;
using AForge.Video;
using AForge.Video.DirectShow;
using ZXing;
using ClosedXML.Excel;

namespace BarcodeReaderWinForms
{
    public class MainForm : Form
    {
        private FilterInfoCollection? videoDevices;
        private VideoCaptureDevice? videoSource;
        private Timer timer;
        private string? lastBarcode;
        private DateTime lastBeepTime = DateTime.MinValue;
        private SQLiteConnection? conn;

        private PictureBox videoBox;
        private TextBox resultText;
        private Button exportButton;
        private DateTimePicker fromDatePicker;
        private DateTimePicker toDatePicker;

        public MainForm()
        {
            InitializeComponent();
            InitDatabase();
            StartCamera();
        }

        private void InitializeComponent()
        {
            Text = "Ứng dụng Đọc Barcode";
            Width = 800;
            Height = 600;

            videoBox = new PictureBox { Dock = DockStyle.Top, Height = 360, SizeMode = PictureBoxSizeMode.Zoom };
            resultText = new TextBox { Dock = DockStyle.Top, ReadOnly = true };

            var panel = new FlowLayoutPanel { Dock = DockStyle.Top, Height = 40 };
            fromDatePicker = new DateTimePicker();
            toDatePicker = new DateTimePicker();
            exportButton = new Button { Text = "Xuất Excel" };
            panel.Controls.Add(fromDatePicker);
            panel.Controls.Add(toDatePicker);
            panel.Controls.Add(exportButton);

            Controls.Add(panel);
            Controls.Add(resultText);
            Controls.Add(videoBox);

            timer = new Timer { Interval = 100 };
            timer.Tick += Timer_Tick;

            exportButton.Click += ExportButton_Click;
            FormClosing += MainForm_FormClosing;
        }
        private void StartCamera()
        {
            videoDevices = new FilterInfoCollection(FilterCategory.VideoInputDevice);
            if (videoDevices.Count == 0)
            {
                MessageBox.Show("Không tìm thấy camera");
                return;
            }
            videoSource = new VideoCaptureDevice(videoDevices[0].MonikerString);
            videoSource.NewFrame += VideoSource_NewFrame;
            videoSource.Start();
            timer.Start();
        }

        private void VideoSource_NewFrame(object sender, NewFrameEventArgs eventArgs)
        {
            Bitmap bitmap = (Bitmap)eventArgs.Frame.Clone();
            videoBox.Image = bitmap;
        }

        private void Timer_Tick(object? sender, EventArgs e)
        {
            if (videoBox.Image == null) return;
            var reader = new BarcodeReader();
            var result = reader.Decode((Bitmap)videoBox.Image);
            if (result != null)
            {
                if (result.Text != lastBarcode || (DateTime.Now - lastBeepTime).TotalSeconds > 1)
                {
                    lastBarcode = result.Text;
                    lastBeepTime = DateTime.Now;
                    resultText.Text = result.Text;
                    try { Clipboard.SetText(result.Text); } catch { }
                    SystemSounds.Beep.Play();
                    SaveScan(result.Text);
                }
            }
        }

        private void InitDatabase()
        {
            string dbPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "barcodes.db");
            bool createTable = !File.Exists(dbPath);
            conn = new SQLiteConnection($"Data Source={dbPath};Version=3;");
            conn.Open();
            if (createTable)
            {
                using var cmd = new SQLiteCommand(@"CREATE TABLE scans (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    content TEXT NOT NULL,
                    scanned_at TEXT NOT NULL,
                    scan_date TEXT,
                    scan_time TEXT
                )", conn);
                cmd.ExecuteNonQuery();
            }
        }

        private void SaveScan(string content)
        {
            if (conn == null) return;
            var now = DateTime.Now;
            using var cmd = new SQLiteCommand("INSERT INTO scans(content, scanned_at, scan_date, scan_time) VALUES(@c, @dt, @d, @t)", conn);
            cmd.Parameters.AddWithValue("@c", content);
            cmd.Parameters.AddWithValue("@dt", now.ToString("yyyy-MM-dd HH:mm:ss"));
            cmd.Parameters.AddWithValue("@d", now.ToString("yyyy-MM-dd"));
            cmd.Parameters.AddWithValue("@t", now.ToString("HH:mm:ss"));
            cmd.ExecuteNonQuery();
        }

        private void ExportButton_Click(object? sender, EventArgs e)
        {
            if (conn == null) return;
            DateTime from = fromDatePicker.Value.Date;
            DateTime to = toDatePicker.Value.Date;
            using var cmd = new SQLiteCommand("SELECT content, scan_date, scan_time FROM scans WHERE scan_date >= @from AND scan_date <= @to ORDER BY scan_date, scan_time", conn);
            cmd.Parameters.AddWithValue("@from", from.ToString("yyyy-MM-dd"));
            cmd.Parameters.AddWithValue("@to", to.ToString("yyyy-MM-dd"));
            using var reader = cmd.ExecuteReader();
            var wb = new XLWorkbook();
            var ws = wb.AddWorksheet("Scans");
            ws.Cell(1,1).Value = "Nội dung";
            ws.Cell(1,2).Value = "Ngày";
            ws.Cell(1,3).Value = "Giờ";
            int row = 2;
            while (reader.Read())
            {
                ws.Cell(row,1).Value = reader.GetString(0);
                ws.Cell(row,2).Value = reader.GetString(1);
                ws.Cell(row,3).Value = reader.GetString(2);
                row++;
            }
            using var sfd = new SaveFileDialog { Filter = "Excel files|*.xlsx", FileName = "barcode_scans.xlsx" };
            if (sfd.ShowDialog() == DialogResult.OK)
            {
                wb.SaveAs(sfd.FileName);
                MessageBox.Show("Đã xuất Excel");
            }
        }

        private void MainForm_FormClosing(object? sender, FormClosingEventArgs e)
        {
            timer.Stop();
            if (videoSource != null && videoSource.IsRunning)
            {
                videoSource.SignalToStop();
                videoSource.NewFrame -= VideoSource_NewFrame;
            }
            conn?.Close();
        }
    }
}
