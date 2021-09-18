
using NLog;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Famis_Tool
{
    public partial class MainForm : Form
    {
        private static Logger logger = LogManager.GetCurrentClassLogger();
        List<data> dsdata;
        public MainForm()
        {
            InitializeComponent();
        }

        private void bt_choice_dir_Click(object sender, EventArgs e)
        {

            int sofile = 1;
            DialogResult result = FBD_loadFile.ShowDialog();

            if (result == DialogResult.OK && !string.IsNullOrWhiteSpace(FBD_loadFile.SelectedPath))
            {
                string[] files = Directory.GetFiles(FBD_loadFile.SelectedPath, "*.txt");
                dsdata = new List<data>();
               
                foreach (var f in files)
                {
                    Encoding vnmese = Encoding.GetEncoding("windows-1258");
                    using (var reader = new StreamReader(File.OpenRead(f), vnmese))
                    {
                        var index = 0;
                        var sodong = 0;
                        while (!reader.EndOfStream)
                        {
                            var line = reader.ReadLine();
                            if (line[0].ToString() == "|")
                            {
                                if (index > 4)
                                {
                                    var values = line.Split('|');
                                    dsdata.Add(new data()
                                    {
                                        thua = values[1].Trim(),
                                        tam_x = Convert.ToDouble(values[2].Trim()),
                                        tam_y = Convert.ToDouble(values[3].Trim()),
                                        Dientich = Convert.ToDouble(values[4].Trim()),
                                        DT_PL = values[5].Trim(),
                                        LoaiDat = values[6].Trim(),
                                        MDSD2003 = values[7].Trim(),
                                        TenChuSuDung = (values[8].Trim()),
                                        Diachi = (values[9].Trim()),
                                        SHTam = values[10].Trim(),
                                        XuDong = values[11].Trim(),
                                        SoTo = f.Replace(".txt","")

                                    });
                                    sodong++;
                                }
                            }
                            index++;
                        }
                        logger.Info(string.Format("file {0} - so dong {1}", f, sodong));
                    }
                    sofile++;
                    
                }
                dataGridView1.DataSource = dsdata;

                //System.Windows.Forms.MessageBox.Show("Files found: " + files.Length.ToString(), "Message");
            }
            MessageBox.Show(string.Format("Đọc được {0} file", sofile));
        }
        public static void OpenExplorer(string dir)
        {
            var result = MessageBox.Show($"Xuất Bc thành công \n File Lưu tại {dir} \n Bạn Có muốn mở file", @"OpenFile", MessageBoxButtons.YesNo,
                                   MessageBoxIcon.Question);

            // If the no button was pressed ...
            if (result == DialogResult.Yes)
            {
                Process.Start(new ProcessStartInfo
                {
                    FileName = dir,
                    UseShellExecute = true,
                    Verb = "open"
                });
            }
        }
        private void bt_export_Click(object sender, EventArgs e)
        {
            if (dsdata.Count == 0)
            {
                MessageBox.Show("Ban chưa load file txt");
            }
            else
            {
                if (saveFileDialog1.ShowDialog() == DialogResult.OK)
                {

                    using (FileStream stream = new FileStream(saveFileDialog1.FileName, FileMode.Create, FileAccess.Write))
                    {
                        IWorkbook wb = new XSSFWorkbook();
                        ISheet sheet = wb.CreateSheet("Sheet1");
                     
                        // Ghi tên cột ở row 0
                        var row1 = sheet.CreateRow(0);
                        row1.CreateCell(0).SetCellValue("Thửa");
                        row1.CreateCell(1).SetCellValue("Tâm x");
                        row1.CreateCell(2).SetCellValue("Tâm Y");
                        row1.CreateCell(3).SetCellValue("Diện tích");
                        row1.CreateCell(4).SetCellValue("D/tích PL");
                        row1.CreateCell(5).SetCellValue("Loại đất");
                        row1.CreateCell(6).SetCellValue("MDSD2003");
                        row1.CreateCell(7).SetCellValue("Tên chủ sử dụng");
                        row1.CreateCell(8).SetCellValue("Địa chỉ");
                        row1.CreateCell(9).SetCellValue("SHTam");
                        row1.CreateCell(10).SetCellValue("Xu dong");
                        row1.CreateCell(11).SetCellValue("Số Tờ");
                        // bắt đầu duyệt mảng và ghi tiếp tục
                        int rowIndex = 1;
                        foreach (var item in dsdata)
                        {
                            // tao row mới
                            var newRow = sheet.CreateRow(rowIndex);
                            // set giá trị
                            newRow.CreateCell(0).SetCellValue(item.thua);
                            newRow.CreateCell(1).SetCellValue(item.tam_x);
                            newRow.CreateCell(2).SetCellValue(item.tam_y);
                            newRow.CreateCell(3).SetCellValue(item.Dientich);
                            newRow.CreateCell(4).SetCellValue(item.DT_PL);
                            newRow.CreateCell(5).SetCellValue(item.LoaiDat);
                            newRow.CreateCell(6).SetCellValue(item.MDSD2003);
                            newRow.CreateCell(7).SetCellValue(item.TenChuSuDung);
                            newRow.CreateCell(8).SetCellValue(item.Diachi);
                            newRow.CreateCell(9).SetCellValue(item.SHTam);
                            newRow.CreateCell(10).SetCellValue(item.XuDong);
                            newRow.CreateCell(11).SetCellValue(item.SoTo);


                            // tăng index
                            rowIndex++;
                        }

                        wb.Write(stream);
                    }
                }
                OpenExplorer(saveFileDialog1.FileName);
            }
        }
    }
}
