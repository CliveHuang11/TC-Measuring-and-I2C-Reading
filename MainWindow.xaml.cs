using System;
//using Ivi.Visa;
using NationalInstruments.Visa;
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
        // I2C register cache / table
        private readonly Dictionary<byte, ushort> _i2cRegisters = new();
        const byte AP72054Q_ADDR = 0x61;

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
            _testData.Columns.Add("Reg Tj", typeof(string));
            _testData.Columns.Add("TC Value", typeof(string));
            _testData.Columns.Add("Reg VIN", typeof(string));
            _testData.Columns.Add("Reg VOUT", typeof(string));
            _testData.Columns.Add("Reg StatusWord", typeof(string));
            _testData.Columns.Add("Reg StatusTemp.", typeof(string));

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
                //Console.WriteLine($"[ERROR] {ex.Message}");
                string usbErr = $"[ERROR] {ex.Message}";
                dmmDataReading.Text = usbErr;
            }
            var usbData = temperature.ToString("F2");
            var i2cData = await Task.Run(() => PerformI2COperation());  // 透過 i2C 讀取 IC 溫度

            // 將資料加入 DataTable
            _i2cRegisters.TryGetValue(0x8D, out var val8D);
            double tJ = val8D * 0.5;
            _i2cRegisters.TryGetValue(0x88, out var val88);
            double vin = val88 / 16.0;
            _i2cRegisters.TryGetValue(0x8B, out var val8B);
            double vout = val8B / 1024.0;
            _i2cRegisters.TryGetValue(0x79, out var statusWord);
            _i2cRegisters.TryGetValue(0x7D, out var statusTemperature);
            _testData.Rows.Add(DateTime.Now.ToString("HH:mm:ss"), tJ, usbData, vin, vout, statusWord, statusTemperature); //, usbData);
            //dmmDataReading.Text = $"Reading = {usbData} C";
            i2cDataReading.Text = $"{i2cData}";
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

                // ===== 讀取 reg 0x88, VIN, 2-bytes =====
                if (!ReadRegister(handle, AP72054Q_ADDR, 0x88, 2, out var buf88, out var err))
                    return err;
                byte low88 = buf88[0];
                byte high88 = buf88[1];
                ushort val88 = (ushort)(((high88 & 0b0000_0111) << 8) | low88);     // 只保留 bits[10:0]
                _i2cRegisters[0x88] = val88;

                // ===== 讀取 reg 0x8B, VOUT, 2-bytes =====
                if (!ReadRegister(handle, AP72054Q_ADDR, 0x8B, 2, out var buf8B, out err))  
                    return err;
                byte low8B = buf8B[0];
                byte high8B = buf8B[1];
                ushort val8B = (ushort)((high8B << 8) | low8B);
                _i2cRegisters[0x8B] = val8B;

                // ===== 讀取 reg 0x8D, TEMPERATURE, 2-bytes =====
                if (!ReadRegister(handle, AP72054Q_ADDR, 0x8D, 2, out var buf8D, out err))
                    return err;
                byte low8D = buf8D[0];
                byte high8D = buf8D[1];
                ushort val8D = (ushort)(((high8D & 0b0000_0111) << 8) | low8D);     // 只保留 bits[10:0]
                _i2cRegisters[0x8D] = val8D;

                // ===== Sends CLEAR_FAULT command before reading STATUS registers
                byte[] cmd = { 0x03 };
                status = (uint)I2C_DeviceWrite(handle, AP72054Q_ADDR, 1, cmd, out var bytesWritten, I2C_TRANSFER_OPTIONS_START_BIT | I2C_TRANSFER_OPTIONS_STOP_BIT);
                if (status != 0 || bytesWritten != 1)
                {
                    return $"CLEAR_FAULT failed (Status:{status}, BW:{bytesWritten})";
                }

                // ===== 讀取 reg 0x79, STATUS_WORD, 2-bytes =====
                if (!ReadRegister(handle, AP72054Q_ADDR, 0x79, 2, out var buf79, out err))
                    return err;
                byte low79 = buf79[0];
                byte high79 = buf79[1];
                ushort val79 = (ushort)((high79 << 8) | low79);
                _i2cRegisters[0x79] = val79;

                // ===== 讀取 reg 0x7D, STATUS_TEMPERATURE, 1-bytes =====
                if (!ReadRegister(handle, AP72054Q_ADDR, 0x7D, 1, out var buf7D, out err))
                    return err;
                _i2cRegisters[0x7D] = buf7D[0];


                // ===== 一起回傳 =====
                return
                    $"Temp. = {(val8D * 0.5):F1} C \n(HiByte = 0x{high8D:X2}, LoByte = 0x{low8D:X2})\n\n" +     // N = -1
                    $"VIN   = {(val88 / 16.0):F3} V \n(HiByte = 0x{high88:X2}, LoByte = 0x{low88:X2})\n\n" +      // N = -4
                    $"VOUT  = {(val8B / 1024.0):F3} V \n(HiByte = 0x{high8B:X2}, LoByte = 0x{low8B:X2})\n\n" +    // N = -10
                    $"STATUS_WORD = 0x{val79:X4}  \n(HiByte = 0x{high79:X2}, LoByte = 0x{low79:X2})\n\n" +
                    $"STATUS_TEMPERATURE = 0x{buf7D[0]:X2}";
            }
            catch (Exception ex)
            {
                return $"錯誤: {ex.Message}";
            }
            finally
            {
                if (handle != IntPtr.Zero)
                    I2C_CloseChannel(handle);
            }
        }

        private bool ReadRegister(IntPtr handle, byte slaveAddr, byte reg, int readLen, out byte[] data, out string error)
        {
            data = new byte[readLen];
            error = string.Empty;

            uint status;
            uint bytesWritten, bytesRead;

            // Write register address
            status = (uint)I2C_DeviceWrite(handle, slaveAddr, 1, new byte[] { reg }, 
                    out bytesWritten, I2C_TRANSFER_OPTIONS_START_BIT);

            if (status != 0 || bytesWritten != 1)
            {
                error = $"Write reg 0x{reg:X2} failed (Status:{status}, BW:{bytesWritten})";
                return false;
            }

            // Read data
            status = (uint)I2C_DeviceRead(handle, slaveAddr, (uint)readLen, data, 
                    out bytesRead, I2C_TRANSFER_OPTIONS_START_BIT | I2C_TRANSFER_OPTIONS_STOP_BIT | I2C_TRANSFER_OPTIONS_NACK_LAST_BYTE);

            if (status != 0 || bytesRead != readLen)
            {
                error = $"Read reg 0x{reg:X2} failed (Status:{status}, BR:{bytesRead})";
                return false;
            }

            return true;
        }

        private void UpdateRegister(byte reg, ushort value)
        {
            _i2cRegisters[reg] = value;
        }

        private bool TryGetRegister(byte reg, out ushort value)
        {
            return _i2cRegisters.TryGetValue(reg, out value);
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