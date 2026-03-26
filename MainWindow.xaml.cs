using System;
using Ivi.Visa;
using System.Globalization;
using System.Data;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Threading;
using ClosedXML.Excel; // 引用 ClosedXML

namespace TC_Measuring_and_I2C_Reading
{
    public partial class MainWindow : Window
    {
        // === 1. 外部 DLL 引用 (DllImport) ===
        [DllImport("libMPSSE.dll", CallingConvention = CallingConvention.Cdecl)]
        public static extern int I2C_GetNumChannels(out uint numChannels);              // 取得FT2232 I2C 通道數
                
        [DllImport("libMPSSE.dll", CallingConvention = CallingConvention.Cdecl)]
        public static extern int I2C_OpenChannel(uint index, out IntPtr handle);        // 開啟指定通道
                
        [DllImport("libMPSSE.dll", CallingConvention = CallingConvention.Cdecl)]
        public static extern int I2C_InitChannel(IntPtr handle, ref ChannelConfig config);  // 初始化通道配置

        // I2C 寫入
        [DllImport("libMPSSE.dll", CallingConvention = CallingConvention.Cdecl)]
        public static extern int I2C_DeviceWrite(IntPtr handle, uint deviceAddress, uint sizeToTransfer, byte[] buffer, out uint sizeTransferred, uint options);

        // I2C 讀取
        [DllImport("libMPSSE.dll", CallingConvention = CallingConvention.Cdecl)]
        public static extern int I2C_DeviceRead(IntPtr handle, uint deviceAddress, uint sizeToTransfer, byte[] buffer, out uint sizeTransferred, uint options);
                
        [DllImport("libMPSSE.dll", CallingConvention = CallingConvention.Cdecl)]
        public static extern int I2C_CloseChannel(IntPtr handle);                       // 關閉通道 (釋放資源關鍵)

        // === 2. 結構與型別定義 ===
        [StructLayout(LayoutKind.Sequential)]
        public struct ChannelConfig
        {            
            public uint ClockRate;      // I2C 速率 (例如 100000 = 100kHz)
            public byte LatencyTimer;   // 延遲計時器 (建議設為 2-16)
            public uint Options;        // 選項 (通常設為 0)
        }

        // === 3. 私有變數 (Fields) ===
        // --- Excel 與 Timer 相關變數 ---
        private DispatcherTimer _testTimer;
        private DataTable _testData;
        private int _currentRow = 2;    // 從第 2 列開始寫（第 1 列留給標題）
        // I2C 傳輸選項 (Options bitmask)
        private const uint I2C_TRANSFER_OPTIONS_START_BIT = 0x01;
        private const uint I2C_TRANSFER_OPTIONS_STOP_BIT = 0x02;
        private const uint I2C_TRANSFER_OPTIONS_NACK_LAST_BYTE = 0x08;
        // 設備 USB 識別名稱
        private string _resourceNameDmm2 = "USB0::0x05E6::0x6500::04437055::INSTR"; // DMM6500, Serial-Number:04437055

        // === 4. 建構函式 (Constructor) ===
        public MainWindow()
        {
            InitializeComponent();  // Interaction logic for MainWindow.xaml    
            // 可以在這裡初始化你的變數
            _testTimer = new DispatcherTimer();
            _testData = new DataTable();
        }

        // === 5. 其他方法 (Methods) ===
        private void btnStartTest_Click(object sender, RoutedEventArgs e)
        {
            // --- 初始化記憶體中的資料表 ---
            _testData = new DataTable();
            _testData.Columns.Add("Time", typeof(string));
            _testData.Columns.Add("Reg Value", typeof(string));
            _testData.Columns.Add("TC Value", typeof(string));

            // --- 啟動 ---
            int intveralTime = 3;
            _testTimer = new DispatcherTimer { Interval = TimeSpan.FromSeconds(intveralTime) };
            _testTimer.Tick += OnTimerTick;
            _testTimer.Start();

            btnStartTest.IsEnabled = false;
            btnStopTest.IsEnabled = true;
            txtStatus.Text = $"測試開始：每 {intveralTime} 秒記錄一次...";
        }


        private void btnStopTest_Click(object sender, RoutedEventArgs e)
        {
            _testTimer?.Stop();

            try
            {
                // 1. 處理目錄路徑：專案執行目錄下的 Datalog 子目錄
                DirectoryInfo dir = new DirectoryInfo(AppDomain.CurrentDomain.BaseDirectory);
                string projectRoot = dir.Parent?.Parent?.Parent?.FullName ??        // 回到專案根目錄
                    throw new InvalidOperationException("Unable to determine project root directory.");   


                string folderPath = Path.Combine(projectRoot, "Datalog");
                if (!Directory.Exists(folderPath))
                {
                    Directory.CreateDirectory(folderPath);
                }

                // 2. 建立檔名：Datalog_日期_時間.xlsx
                string fileName = $"Datalog_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx";
                string fullPath = Path.Combine(folderPath, fileName);

                // 3. 使用 ClosedXML 將 DataTable 轉換為 Excel 存檔
                using (var workbook = new XLWorkbook())
                {
                    var worksheet = workbook.Worksheets.Add("TestResults");
                    // 將 DataTable 直接匯入 Worksheet (從第一列第一格開始)
                    worksheet.Cell(1, 1).InsertTable(_testData);
                    worksheet.Columns().AdjustToContents(); // 自動調整欄寬

                    workbook.SaveAs(fullPath);
                }

                MessageBox.Show($"測試停止！檔案已儲存至：\n{fullPath}", "存檔成功");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"存檔過程發生錯誤: {ex.Message}");
            }
            finally
            {
                btnStartTest.IsEnabled = true;
                btnStopTest.IsEnabled = false;
                txtStatus.Text = "準備就緒";
            }
        }


        // 定時觸發的讀取動作
        private async void OnTimerTick(object? sender, EventArgs e)
        {
            //var usbData = await Task.Run(() => GetChromaMeasurement()); // 使用原本寫好的 ReadChromaData
            double temperature = 0;
            try
            {
                temperature = ReadTemperature(_resourceNameDmm2);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[ERROR] {ex.Message}");
            }
            //var usbData = temperature.ToString("F2");
            var i2cData = await Task.Run(() => PerformI2COperation());  // 透過 i2C 讀取 IC 溫度

            // 將資料加入 DataTable
            _testData.Rows.Add(DateTime.Now.ToString("HH:mm:ss"), i2cData, "25"); //, usbData);
            //dmmDataReading.Text = $"Reading = {usbData} C";
            i2cDataReading.Text = $"Reading = {i2cData}";
            txtStatus.Text = $"已記錄 {_testData.Rows.Count} 筆";
        }


        private string PerformI2COperation()
        {
            IntPtr handle = IntPtr.Zero; // 初始化為空
            uint status;

            try
            {
                // 檢查裝置數量
                uint numChannels;
                I2C_GetNumChannels(out numChannels);
                if (numChannels == 0) return "找不到 FT2232H";

                // 開啟通道 (Channel 0 通常是 FT2232H 的 Port A)
                status = (uint)I2C_OpenChannel(0, out handle);
                if (status != 0 || handle == IntPtr.Zero) return $"開啟通道失敗 (Error: {status})";

                // 初始化配置
                ChannelConfig config = new ChannelConfig
                {
                    ClockRate = 100000,     // 100kHz
                    LatencyTimer = 2,
                    Options = 0
                };
                status = (uint)I2C_InitChannel(handle, ref config);
                if (status != 0) return $"初始化失敗 (Error: {status})";

                // 執行讀取 (從 0x61 的 reg 0x8B 讀回 2 Bytes)
                // 注意：這裡我們只送出 0x8B，並且 options 只給 START (0x01)，不給 STOP
                uint bytesWritten;
                byte[] regAddr = new byte[] { 0x8B };
                status = (uint)I2C_DeviceWrite(handle, 0x61,
                         (uint)regAddr.Length, regAddr, out bytesWritten,
                         I2C_TRANSFER_OPTIONS_START_BIT); // 只有 Start
                byte[] readBuf = new byte[2];
                uint bytesRead;
                // Options 設為 START | STOP | NACK_LAST (0x01 | 0x02 | 0x08 = 11)
                // 這樣會發送 Restart -> 讀取 -> 最後一個 Byte 給 NACK -> Stop
                status = (uint)I2C_DeviceRead(handle, 0x61,
                         (uint)readBuf.Length, readBuf, out bytesRead,
                         I2C_TRANSFER_OPTIONS_START_BIT | I2C_TRANSFER_OPTIONS_STOP_BIT | I2C_TRANSFER_OPTIONS_NACK_LAST_BYTE);

                if (status == 0)
                {
                    // 假設 readBuffer[0] 是 Lower Byte, readBuffer[1] 是 Higher Byte
                    byte lowerByte = readBuf[0];
                    byte higherByte = readBuf[1];

                    // 組合回 16-bit 數值 (例如: Higher << 8 | Lower)
                    ushort combinedValue = (ushort)(((higherByte & 0b0000_0111) << 8) | lowerByte);

                    return combinedValue.ToString();
                }
                else
                {
                    return $"讀取失敗 (Error: {status})";
                }
            }
            catch (Exception ex)
            {
                return $"錯誤: {ex.Message}";
            }
            finally
            {
                //  F. 無論成功或失敗，只要 handle 不是空的就關閉它
                if (handle != IntPtr.Zero)
                {
                    I2C_CloseChannel(handle);
                }
            }
        }


        private static double ReadTemperature(string visaAddress)
        {
            try
            {
                using IMessageBasedSession session = GlobalResourceManager.Open(visaAddress) as IMessageBasedSession
                    ?? throw new InvalidOperationException("Unable to open IMessageBasedSession.");

                // --- VISA 通訊設定 ---
                session.TimeoutMilliseconds = 5000;
                session.TerminationCharacterEnabled = true;
                session.TerminationCharacter = (byte)'\n';

                var io = session.FormattedIO;

                // --- SCPI 流程 ---
                //io.WriteLine("*RST");
                io.WriteLine("TEMP:TC:TYPE J"); // 若需指定熱電偶
                io.WriteLine("MEAS:TEMP?");
                string response = io.ReadLine();

                // --- 解析 ---
                if (!double.TryParse(
                        response.Trim(),
                        NumberStyles.Float,
                        CultureInfo.InvariantCulture,
                        out double temperature))
                {
                    throw new FormatException(
                        $"Invalid temperature response: '{response}'");
                }

                return temperature;
            }
            catch (VisaException visaEx)
            {
                // ✅ VISA 專屬錯誤（最重要）
                throw new Exception(
                    $"VISA Error : {visaEx.Message}", visaEx);
            }
            catch (TimeoutException)
            {
                throw new Exception("VISA timeout while communicating with instrument.");
            }
            catch (Exception)
            {
                // 其餘錯誤原樣往外丟
                throw;
            }
        }





        
        

    }
}