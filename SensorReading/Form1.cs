﻿using System;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO.Ports;
using System.Globalization;
using System.Reflection;
using System.Collections.Generic;
using Excel = Microsoft.Office.Interop.Excel;

namespace SensorReading
{
    public partial class MainForm : Form
    {
        public string connectionPort;
        public string oldConnectionPort;
        public bool statusFlag = true;
        static SerialPort _serialPort;
        bool isReadingData = false;
        int currentMouseOver = 0;
        public Dictionary<int, string> spisokPropuskov = new Dictionary<int, string>();
        List<string> results = new List<string>();
        //Предыдущая таблица
        public string lastName = "Все данные";
        //Словарь для хранения данных по каждому шаблону
        Dictionary<string, List<string[]>> templatesData = new Dictionary<string, List<string[]>>();
        Dictionary<int, List<string>> excelOutput = new Dictionary<int, List<string>>();

        //Словарь для хранения настроек шаблонов
        Dictionary<int, string> cellDescriptions = new Dictionary<int, string>
        {
            { 0 , "№"},
            { 1 , "Время" },
            { 2 , " RM3100 Магнитометр X" },
            { 3 , " RM3100 Магнитометр Y" },
            { 4 , " RM3100 Магнитометр Z" },
            { 5 , " MTI temp" },
            { 6 , " MTI Магнитометр X" },
            { 7 , " MTI Магнитометр Y" },
            { 8 , " MTI Магнитометр Z" },
            { 9 , " MTI Акселерометр X" },
            { 10 , " MTI Акселерометр Y" },
            { 11 , " MTI Акселерометр Z" },
            { 12 , " PNI (Курс)" },
            { 13 , " PNI (Тангаж)" },
            { 14 , " PNI (Крен)" },
            { 15 , " PNI Акселерометр X" },
            { 16 , " PNI Акселерометр Y" },
            { 17 , " PNI Акселерометр Z" },
            { 18 , " PNI Магнитометр X" },
            { 19 , " PNI Магнитометр Y" },
            { 20 , " PNI Магнитометр Z" },
            { 21 , " ADIS Акселерометр X" },
            { 22 , " ADIS Акселерометр Y" },
            { 23 , " ADIS Акселерометр Z" },
            { 24 , " ADIS Магнитометр X" },
            { 25 , " ADIS Магнитометр Y" },
            { 26 , " ADIS Магнитометр Z" },
            { 27 , " ADIS Гироскоп X" },
            { 28 , " ADIS Гироскоп Y" },
            { 29 , " ADIS Гироскоп Z" },
            { 30 , " ADIS temp" },
            { 31 , " RM3100 MAz" },
            { 32 , " MTI MAz" },
            { 33 , " PNI MAz" },
            { 34 , " ADIS MAz" },
            { 35 , " ADIS (Курс)" },
            { 36 , " ADIS (Тангаж)" },
            { 37 , " ADIS (Крен)" },
            { 38 , " PNI Подсчитано (Курс)" },
            { 39 , " PNI Подсчитано (Тангаж)" },
            { 40 , " PNI Подсчитано (Крен)" },
            { 41 , " MTI (Курс)" },
            { 42 , " MTI (Тангаж)" },
            { 43 , " MTI (Крен)" },
            { 44 , " Маджвик (Курс)" },
            { 45 , " Маджвик (Тангаж)" },
            { 46 , " Маджвик (Крен)" }
        };

        //Шаблоны таблиц
        int[] allColumns = new int[] { 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29, 30 };
        Dictionary<string, int[]> tableTemplates = new Dictionary<string, int[]>
        {
            {"Все данные", new int[] { 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29, 30 } }, //Все колонки
            {"Магнитные азимуты", new int[]{ 0, 1, 31, 32, 33, 34 } },  //Магнитные азимуты
            {"Акселерометры", new int[]{ 0, 1, 9, 10, 11, 15, 16, 17, 21, 22, 23} }, //Акселерометры
            {"Магнитометры", new int[] { 0, 1, 2, 3, 4, 6, 7, 8, 18, 19, 20, 24, 25, 26} }, //Магнитометры
            {"Круги Эйлера", new int[] {0, 1, 12, 13, 14, 38, 39, 40, 35, 36, 37, 41, 42, 43} }, //Курс, Крен, Тангаж
            {"Маджвик", new int[] {0, 1, 44, 45, 46} } //Фильтр Маджвика
        };
        public class MadgwickFilter
        {
            // Кватернионы ориентации
            private double q1 = 1f, q2 = 0f, q3 = 0f, q4 = 0f;

            // Коэффициент фильтра
            public double Beta { get; set; } = 0.1f;

            // Обновление фильтра с использованием данных гироскопа и акселерометра
            public (double Roll, double Pitch, double Yaw) Update(double gx, double gy, double gz, double ax, double ay, double az, int dataPeriod)
            {
                double sampleFreq = 1000f / dataPeriod; // Частота в Гц, рассчитанная из периода накопления данных в мс
                double recipNorm;
                double s1, s2, s3, s4;
                double qDot1, qDot2, qDot3, qDot4;

                // Нормализация данных акселерометра
                recipNorm = (double)Math.Sqrt(ax * ax + ay * ay + az * az);
                ax /= recipNorm;
                ay /= recipNorm;
                az /= recipNorm;

                // Промежуточные значения для угловой скорости гироскопа
                qDot1 = 0.5f * (-q2 * gx - q3 * gy - q4 * gz);
                qDot2 = 0.5f * (q1 * gx + q3 * gz - q4 * gy);
                qDot3 = 0.5f * (q1 * gy - q2 * gz + q4 * gx);
                qDot4 = 0.5f * (q1 * gz + q2 * gy - q3 * gx);

                // Градиент для коррекции ошибки
                s1 = 2f * (q1 * q3 - q2 * q4) - ax;
                s2 = 2f * (q1 * q2 + q3 * q4) - ay;
                s3 = 1f - 2f * (q2 * q2 + q3 * q3) - az;
                s4 = q2 * q3 - q1 * q4;

                // Нормализация градиента для предотвращения деления на ноль
                recipNorm = (double)Math.Sqrt(s1 * s1 + s2 * s2 + s3 * s3 + s4 * s4);
                s1 /= recipNorm;
                s2 /= recipNorm;
                s3 /= recipNorm;
                s4 /= recipNorm;

                // Применение градиента к скоростям изменения кватернионов
                qDot1 -= Beta * s1;
                qDot2 -= Beta * s2;
                qDot3 -= Beta * s3;
                qDot4 -= Beta * s4;

                // Интегрирование для получения новых кватернионов
                q1 += qDot1 * (1.0f / sampleFreq);
                q2 += qDot2 * (1.0f / sampleFreq);
                q3 += qDot3 * (1.0f / sampleFreq);
                q4 += qDot4 * (1.0f / sampleFreq);

                // Нормализация кватернионов
                recipNorm = (double)Math.Sqrt(q1 * q1 + q2 * q2 + q3 * q3 + q4 * q4);
                q1 /= recipNorm;
                q2 /= recipNorm;
                q3 /= recipNorm;
                q4 /= recipNorm;

                return ConvertQuaternionToEuler();
            }

            // Аналогично реализуйте другие перегруженные версии метода Update

            private (float Roll, float Pitch, float Yaw) ConvertQuaternionToEuler()
            {
                // Расчет углов Эйлера из кватернионов
                float roll = (float)Math.Atan2(2f * (q1 * q2 + q3 * q4), 1 - 2f * (q2 * q2 + q3 * q3));
                float pitch = (float)Math.Asin(2f * (q1 * q3 - q4 * q2));
                float yaw = (float)Math.Atan2(2f * (q1 * q4 + q2 * q3), 1 - 2f * (q3 * q3 + q4 * q4));

                return (roll * 180.0f / (float)Math.PI, pitch * 180.0f / (float)Math.PI, yaw * 180.0f / (float)Math.PI);
            }
        }

        public MainForm()
        {
            InitializeComponent();
            EnableDoubleBuffering(SensorGridView);
            ConnectGrin.BackColor = Color.DarkGreen; //Lime\DarkGreen
            ConnectGreenText.Font = new Font(ConnectGreenText.Font, FontStyle.Regular);
            ConnectGreenText.ForeColor = Color.Gray;
            ConnectYellow.BackColor = Color.Yellow; //Yellow\Olive
            ConnectYellowText.Font = new Font(ConnectYellowText.Font, FontStyle.Bold);
            ConnectRed.BackColor = Color.Maroon; //Red\Maroon
            ConnectRedText.Font = new Font(ConnectRedText.Font, FontStyle.Regular);
            ConnectRedText.ForeColor = Color.Gray;

            ConnectGrinMini.Hide();
            ConnectGrinMini.BackColor = Color.DarkGreen;
            ConnectYellowMini.Hide();
            ConnectYellowMini.BackColor = Color.Yellow;
            ConnectRedMini.Hide();
            ConnectRedMini.BackColor = Color.Maroon;

            UpdateComPortsList();

            MergeHeaderGrid();

            SetRoundedShape(ConnectGrin, 27);
            SetRoundedShape(ConnectYellow, 27);
            SetRoundedShape(ConnectRed, 27);
            SetRoundedShape(ConnectGrinMini, 22);
            SetRoundedShape(ConnectYellowMini, 22);
            SetRoundedShape(ConnectRedMini, 22);
            SensorGridView.ClearSelection();

            button1.MouseEnter += Button_MouseEnter;
            button1.MouseLeave += Button_MouseLeave;
            button2.MouseEnter += Button_MouseEnter;
            button2.MouseLeave += Button_MouseLeave;
            OpenFormFull.MouseEnter += Button_MouseEnter;
            OpenFormFull.MouseLeave += Button_MouseLeave;
        }

        private void UpdateComPortsList()
        {
            ComConnectorsList.Items.Clear();
            foreach (string s in SerialPort.GetPortNames())
            {
                ComConnectorsList.Items.Add(s);
            }
        }

        private void SaveCurrentGridViewData(string templateName)
        {
            if (!templatesData.ContainsKey(templateName))
            {
                templatesData[templateName] = new List<string[]>();
            }

            templatesData[templateName].Clear();

            //Сохраняем все строки кроме последней
            for (int rowIndex = 0; rowIndex < SensorGridView.Rows.Count - 1; rowIndex++)
            {
                DataGridViewRow row = SensorGridView.Rows[rowIndex];
                string[] rowData = new string[row.Cells.Count];
                for (int i = 0; i < row.Cells.Count; i++)
                {
                    rowData[i] = row.Cells[i].Value?.ToString() ?? "";
                }

                templatesData[templateName].Add(rowData);
            }
        }
        private void LoadGridViewData(string templateName)
        {
            if (!templatesData.ContainsKey(templateName)) return;

            SensorGridView.Rows.Clear();

            foreach (var rowData in templatesData[templateName])
            {
                SensorGridView.Rows.Add(rowData);
            }
        }
        private void LoadDataIntoGridView(int[] template)
        {
            SensorGridView.ColumnCount = template.Length;
            string[,] matrix = new string[3, template.Length];

            for (int i = 0; i < template.Length; i++)
            {
                if (cellDescriptions.TryGetValue(template[i], out string description))
                {
                    string[] parts = description.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
                    FillMatrixColumn(matrix, i, parts);
                }
            }
            LoadMatrixIntoGridView(matrix);
        }
        private void FillMatrixColumn(string[,] matrix, int columnIndex, string[] parts)
        {
            switch (parts.Length)
            {
                case 1:
                    matrix[0, columnIndex] = "";
                    matrix[1, columnIndex] = "";
                    matrix[2, columnIndex] = parts[0];
                    break;
                case 2:
                    matrix[0, columnIndex] = parts[0];
                    matrix[1, columnIndex] = "";
                    matrix[2, columnIndex] = parts[1];
                    break;
                case 3:
                    matrix[0, columnIndex] = parts[0];
                    matrix[1, columnIndex] = parts[1];
                    matrix[2, columnIndex] = parts[2];
                    break;
            }
        }
        private void LoadMatrixIntoGridView(string[,] matrix)
        {
            SensorGridView.Rows.Clear();
            SensorGridView.Rows.Insert(0,3);
            for (int i = 0; i < matrix.GetLength(1); i++)
            {
                SensorGridView[i, 0].Value = matrix[0, i];
                SensorGridView[i, 1].Value = matrix[1, i];
                SensorGridView[i, 2].Value = matrix[2, i];
            }
        }
        private void RepairDataInGridView()
        {
            for (int row = 0; row < 3; row++)
            {
                int startCol = 0;
                while (startCol < SensorGridView.ColumnCount)
                {
                    string currentValue = SensorGridView[startCol, row].Value?.ToString();
                    int count = 1;
                    int col = startCol + 1;

                    //Подсчитываем количество одинаковых ячеек подряд
                    while (col < SensorGridView.ColumnCount && SensorGridView[col, row].Value?.ToString() == currentValue)
                    {
                        count++;
                        col++;
                    }

                    //Если есть повторы, удаляем все, кроме среднего значения
                    if (count > 1)
                    {
                        int middle = startCol + count / 2;
                        for (int i = startCol; i < startCol + count; i++)
                        {
                            if (i != middle)
                            {
                                SensorGridView[i, row].Style.ForeColor = Color.Transparent;
                            }
                        }
                    }
                    startCol += count;
                }
            }
        }
        public void EnableDoubleBuffering(DataGridView dgv)
        {
            Type dgvType = dgv.GetType();
            PropertyInfo pi = dgvType.GetProperty("DoubleBuffered", BindingFlags.Instance | BindingFlags.NonPublic);
            pi.SetValue(dgv, true, null);
        }

        static void SetRoundedShape(Control control, int radius)
        {
            System.Drawing.Drawing2D.GraphicsPath path = new System.Drawing.Drawing2D.GraphicsPath();
            path.AddLine(radius, 0, control.Width - radius, 0);
            path.AddArc(control.Width - radius, 0, radius, radius, 270, 90);
            path.AddLine(control.Width, radius, control.Width, control.Height - radius);
            path.AddArc(control.Width - radius, control.Height - radius, radius, radius, 0, 90);
            path.AddLine(control.Width - radius, control.Height, radius, control.Height);
            path.AddArc(0, control.Height - radius, radius, radius, 90, 90);
            path.AddLine(0, control.Height - radius, 0, radius);
            path.AddArc(0, 0, radius, radius, 180, 90);
            control.Region = new Region(path);
        }

        private void MergeHeaderGrid()
        {
            SensorGridView.ColumnCount = 31;

            // Очищаем строки в DataGridView
            SensorGridView.Rows.Clear();

            // Отключаем строку заголовков
            SensorGridView.ColumnHeadersVisible = false;

            LoadDataIntoGridView(allColumns);
            RepairDataInGridView();

            SensorGridView.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            for (int col = 0; col < SensorGridView.Columns.Count; col++)
            {
                SensorGridView.Columns[col].Width = 115;
            }
            SensorGridView.Columns[0].Width = 50;
            SensorGridView.Columns[1].Width = 80;
        }

        //Общий обработчик MouseEnter для кнопок
        private void Button_MouseEnter(object sender, EventArgs e)
        {
            Button button = sender as Button;
            if (button != null)
            {
                //Установка цвета в зависимости от идентификатора кнопки
                button.BackColor = button.Name == "button1" ? Color.Red :
                                   button.Name == "button2" ? Color.RoyalBlue :
                                   Color.LightBlue;
            }
        }

        //Общий обработчик MouseLeave для кнопок
        private void Button_MouseLeave(object sender, EventArgs e)
        {
            Button button = sender as Button;
            if(button != null)
            {
                button.BackColor = Color.Transparent;
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            WindowState = FormWindowState.Minimized;
        }

        private void panel1_MouseDown(object sender, MouseEventArgs e)
        {
            panel1.Capture = false;
            Message m = Message.Create(Handle, 0xa1, new IntPtr(2), IntPtr.Zero);
            WndProc(ref m);
        }

        private string SendAndReadData(string dataPeriod)
        {
            if (!_serialPort.IsOpen) return null;

            try
            {
                _serialPort.Write(dataPeriod);
                try
                {
                    return _serialPort.ReadLine();
                }
                catch (TimeoutException e)
                {
                    Console.WriteLine($"Данные с датчика не поступают: {e.Message}");
                    return null;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                return null;
            }
        } 

        public void WriteGridView(List<string> results, string selectedTemplate, int rowIndex)
        {
            if (!tableTemplates.TryGetValue(selectedTemplate, out int[] template)) return;

            int maxIndex = Math.Min(template.Length, results.Count);
            for (int i = 2; i < maxIndex; i++)
            {
                if (double.TryParse(results[template[i] - 1], NumberStyles.Any, CultureInfo.InvariantCulture, out double value))
                {
                    SensorGridView.Rows[rowIndex].Cells[i].Value = Math.Round(value, 2);
                }
            }
        }

        private double MagAz(double Pitch, double Roll, double X, double Y, double Z)
        {
            // Расчет магнитного азимута
            double azimuth = Math.Atan2(Z * Math.Sin(Roll) - Y * Math.Cos(Roll), X * Math.Cos(Pitch) + Y * Math.Sin(Pitch) * Math.Sin(Roll) + Z * Math.Sin(Pitch) * Math.Cos(Roll));
            // Преобразование из радиан в градусы
            azimuth *= 180 / Math.PI;
            //Корректировка ортогональных преобразований
            if (azimuth < 0)
            {
                azimuth += 360;
            }
            else if (azimuth >= 360)
            {
                azimuth -= 360;
            }
            return azimuth;
        }
        // Акселерометр - Крен Тангаж
        // Магнитометр - Магнитный азимут
       
        private (double roll, double pitch, double yaw) AccMagCounting(double X, double Y, double Z, double mx, double my)
        {
            double toAngle = 180 / Math.PI;
            double roll = Math.Atan(Y / Z) * toAngle;
            double pitch = Math.Atan(-X / Math.Sqrt(Y * Y + Z * Z)) * toAngle;
            double yaw = 360 - Math.Atan2(my , mx) * toAngle;
            if (yaw < 0)
            {
                yaw += 360;
            }
            else if (yaw >= 360)
            {
                yaw -= 360;
            }

            return (roll, pitch, yaw);
        }

        private double GetData(string data)
        {
            double result = Math.Round(double.Parse(data, CultureInfo.InvariantCulture), 2);
            return result;
        }

        // Функция для преобразования измерений в строку с заменой запятой на точку
        string FormatMeasurement(double measurement)
        {
            return measurement.ToString().Replace(',', '.');
        }

        // Обработка данных сенсоров и добавление результатов в список
        void ProcessSensorData(double acsX, double acsY, double acsZ, double magX, double magY)
        {
            // Расчет углов Эйлера из данных акселерометра и магнитометра
            (double roll, double pitch, double yaw) = AccMagCounting(acsX, acsY, acsZ, magX, magY);

            // Добавление углов Эйлера в результаты с форматированием
            results.Add($"{FormatMeasurement(yaw)}");
            results.Add($"{FormatMeasurement(pitch)}");
            results.Add($"{FormatMeasurement(roll)}");
        }

        private async void StartDataReading()
        {
            while (isReadingData) //метка для остановки чтения
            {
                string dataPeriod = await GetSelectedDataPeriodAsync();

                string response = await Task.Run(() => SendAndReadData(dataPeriod));

                if (!string.IsNullOrEmpty(response))
                {
                    await UpdateUIAsync(response);
                }
            }
        }

        private async Task<string> GetSelectedDataPeriodAsync()
        {
            return await Task.Run(() =>
            {
                string dataPeriod = string.Empty;
                Invoke(new Action(() =>
                {
                    dataPeriod = DataPeriodBox.SelectedItem.ToString();
                }));
                return dataPeriod;
            });
        }

        private async Task UpdateUIAsync(string response)
        {
            //Здесь используется Task.Run для выполнения тяжелых операций в отдельном потоке.
            await Task.Run(() =>
            {
                results.Clear();
                results.AddRange(response.Split(';'));

                //Обновляем SensorGridView в основном потоке
                int rowIndex = 3;
                Invoke((MethodInvoker)(() =>
                {
                    SensorGridView.Rows.Insert(rowIndex, 1);
                }));
                
                SensorGridView.Rows[rowIndex].Cells[0].Value = SensorGridView.Rows.Count - 3;

                int totalTimeInSeconds = int.Parse(results[0], CultureInfo.InvariantCulture) / 1000;
                // Преобразование общего времени в секундах в часы, минуты и секунды
                TimeSpan timeSpan = TimeSpan.FromSeconds(totalTimeInSeconds);
                // Форматирование времени в формате "HH:mm:ss"
                string formattedTime = timeSpan.ToString(@"hh\:mm\:ss");
                SensorGridView.Rows[rowIndex].Cells[1].Value = formattedTime;

                //Общие значения (Не трогать)
                double Pitch = GetData(results[12]) * Math.PI / 180;// значение Тангажа (например, G3);
                double Roll = GetData(results[13]) * Math.PI / 180;// значение Крена (например, H3);
                                                                    //RM3100 2-4 Магнит
                                                                    //MTI 6-8 магнит, 9-12 Акс
                                                                    //PNI 12-14 Эйлер, 15-17 Акс, 18-20 Магнит
                                                                    //ADIS 21-23 Акс, 24-26 Магнит, 27-29 Гиро
                                                                    //Для RM3100
                double rm3100MagX = GetData(results[1]);// значение X (например, B3);
                double rm3100MagY = GetData(results[2]);// значение Y (например, C3);
                double rm3100MagZ = GetData(results[3]);// значение Z (например, D3);

                //Для PNI
                double pniYaw = GetData(results[11]);
                double pniPitch = GetData(results[12]);
                double pniRoll = GetData(results[13]);

                double pniAcsY = GetData(results[14]);
                double pniAcsX = -GetData(results[15]);
                double pniAcsZ = GetData(results[16]);

                double pniMagY = GetData(results[17]);
                double pniMagX = -GetData(results[18]);
                double pniMagZ = GetData(results[19]);

                //Для ADIS
                double adisAcsY = GetData(results[20]);
                double adisAcsX = GetData(results[21]);
                double adisAcsZ = -GetData(results[22]);

                double adisMagY = GetData(results[23]);
                double adisMagX = GetData(results[24]);
                double adisMagZ = -GetData(results[25]);

                double adisGiroY = GetData(results[26]);
                double adisGiroX = GetData(results[27]);
                double adisGiroZ = -GetData(results[28]);

                //Для MTI
                double mtiMagY = GetData(results[5]);
                double mtiMagX = GetData(results[6]);
                double mtiMagZ = -GetData(results[7]);

                double mtiAcsY = GetData(results[8]);
                double mtiAcsX = GetData(results[9]);
                double mtiAcsZ = -GetData(results[10]);

                // Расчет магнитного азимута
                double azimuthRM3100 = MagAz(Pitch, Roll, rm3100MagX, rm3100MagY, rm3100MagZ);
                double azimuthPNI = MagAz(Pitch, Roll, pniMagX, pniMagY, pniMagZ);
                double azimuthMTI = MagAz(Pitch, Roll, mtiMagX, mtiMagY, mtiMagZ);
                double azimuthADIS = MagAz(Pitch, Roll, adisMagX, adisMagY, adisMagZ);

                results.Add(azimuthRM3100.ToString().Replace(',', '.'));
                results.Add(azimuthPNI.ToString().Replace(',', '.'));
                results.Add(azimuthMTI.ToString().Replace(',', '.'));
                results.Add(azimuthADIS.ToString().Replace(',', '.'));

                // Обработка данных для каждого сенсора
                ProcessSensorData(adisAcsX, adisAcsY, adisAcsZ, adisMagX, adisMagY);    //ADIS
                ProcessSensorData(pniAcsX, pniAcsY, pniAcsZ, pniMagX, pniMagY);         //PNI
                ProcessSensorData(mtiAcsX, mtiAcsY, mtiAcsZ, mtiMagX, mtiMagY);         //MTI   

                //Маджвик
                MadgwickFilter mgF = new MadgwickFilter();

                var eulerAngles = mgF.Update(adisGiroX, adisGiroY, adisGiroZ, adisAcsX, adisAcsY, adisAcsZ, int.Parse("3000"));

                (double roll, double pitch, double yaw) = eulerAngles;
                results.Add(roll.ToString().Replace(',', '.'));
                results.Add(pitch.ToString().Replace(',', '.'));
                results.Add(yaw.ToString().Replace(',', '.'));

                //Теперь возвращаемся в поток UI для обновления пользовательского интерфейса
                Invoke((MethodInvoker)(() =>
                {
                    string selectedTemplate = selectTableBox.SelectedItem.ToString();
                    WriteGridView(results, selectedTemplate, rowIndex);
                }));

                excelOutput.Add(excelOutput.Count(), new List<string>(results));
            });
        }

        private bool IsListBoxSelectionValid(ListBox listBox, string errorMessage)
        {
            if (listBox.SelectedIndex == -1)
            {
                MessageBox.Show(errorMessage, "Выбор не сделан", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return false;
            }
            return true;
        }

        private async void ConnectComPort_Click(object sender, EventArgs e)
        {
            if (!IsListBoxSelectionValid(ComConnectorsList, "Укажите COM-порт для чтения.") ||
            !IsSelectionValid(DataPeriodBox, "Укажите период накопления данных."))
            {
                return;
            }

            string selectedPortName = ComConnectorsList.SelectedItem.ToString();
            if (!SerialPort.GetPortNames().Contains(selectedPortName))
            {
                ShowPortError(selectedPortName, true);
                return;
            }

            _serialPort = new SerialPort
            {
                PortName = selectedPortName,
                BaudRate = 115200,
                DataBits = 8,
                StopBits = StopBits.One,
                Parity = Parity.None,
                Handshake = Handshake.None,
                ReadTimeout = 6000,
                WriteTimeout = 6000
            };

            try
            {
                _serialPort.Open();
                UpdateUIBasedOnPortStatus(_serialPort.IsOpen);
                if (_serialPort.IsOpen)
                {
                    isReadingData = true;
                    await Task.Run(() => StartDataReading());
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Произошла ошибка при подключении к COM-порту: {ex.Message}", "Ошибка подключения", MessageBoxButtons.OK, MessageBoxIcon.Error);
                ShowPortError(selectedPortName, false);
            }
        }

        private bool IsSelectionValid(ComboBox comboBox, string errorMessage)
        {
            if (comboBox.SelectedIndex == -1)
            {
                MessageBox.Show(errorMessage, "Выбор не сделан", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return false;
            }
            return true;
        }

        private void ShowPortError(string portName, bool isNotExistError)
        {
            string message = isNotExistError ? $"Выбранный COM-порт '{portName}' не существует. Пожалуйста, выберите корректный порт." :
            "Не удалось открыть COM-порт. Пожалуйста, проверьте подключение.";
            MessageBox.Show(message, "Ошибка подключения", MessageBoxButtons.OK, MessageBoxIcon.Error);
            UpdateUIBasedOnPortStatus(false);
        }

        private void UpdateUIBasedOnPortStatus(bool isOpen)
        {
            // Обновление состояния для зеленого индикатора
            ConnectGrin.BackColor = isOpen ? Color.Lime : Color.DarkGreen;
            ConnectGreenText.Font = new Font(ConnectGreenText.Font, isOpen ? FontStyle.Bold : FontStyle.Regular);
            ConnectGreenText.ForeColor = isOpen ? Color.Black : Color.Gray;

            // Обновление состояния для зеленого индикатора
            ConnectYellow.BackColor = isOpen ? Color.Olive : Color.Yellow;
            ConnectYellowText.Font = new Font(ConnectYellowText.Font, isOpen ? FontStyle.Regular : FontStyle.Bold);
            ConnectYellowText.ForeColor = isOpen ? Color.Black : Color.Gray;

            // Обновление состояния для зеленого индикатора
            ConnectRed.BackColor = isOpen ? Color.Maroon : Color.Red;
            ConnectRedText.Font = new Font(ConnectRedText.Font, isOpen ? FontStyle.Regular : FontStyle.Bold);
            ConnectRedText.ForeColor = isOpen ? Color.Black : Color.Gray;

            // Обновление состояния для зеленого индикатора
            ConnectGrinMini.BackColor = isOpen ? Color.Lime : Color.DarkGreen;
            ConnectYellowMini.BackColor = isOpen ? Color.Olive : Color.Yellow;
            ConnectRedMini.BackColor = isOpen ? Color.Maroon : Color.Red;
        }

        public Image RotateImage(Image someImage, bool flag)
        {
            RotateFlipType rotateType = flag ? RotateFlipType.Rotate270FlipNone : RotateFlipType.Rotate90FlipNone;
            someImage.RotateFlip(rotateType);
            return someImage;
        }

        private void ToggleVisibility(bool status)
        {
            //Управление видимостью компонентов
            ConnectGreenText.Visible = !status;
            ConnectYellowText.Visible = !status;
            ConnectRedText.Visible = !status;
            ConnectGrin.Visible = !status;
            ConnectYellow.Visible = !status;
            ConnectRed.Visible = !status;
            ConnectGrinMini.Visible = status;
            ConnectYellowMini.Visible = status;
            ConnectRedMini.Visible = status;

            //Установка позиций
            label1.Location = status ? new Point(21, 81) : new Point(335, 48);
            label2.Location = status ? new Point(126, 81) : new Point(440, 48);
            ComConnectorsList.Location = status ? new Point(21, 100) : new Point(335, 67);
            DataPeriodBox.Location = status ? new Point(129, 100) : new Point(443, 67);
            ConnectComPort.Location = status ? new Point(335, 48) : new Point(335, 139);
            CloseComPort.Location = status ? new Point(460, 48) : new Point(460, 139);
            SaveDataToFile.Location = status ? new Point(585, 48) : new Point(585, 139);
            StopReading.Location = status ? new Point(335, 83) : new Point(335, 174);
            ContinueReading.Location = status ? new Point(460, 83) : new Point(460, 174);
            SensorGridView.Location = status ? new Point(16, 173) : new Point(16, 223);
            ClearButton.Location = status ? new Point(585, 83) : new Point(585, 174);

            //Изменение размера SensorGridView
            SensorGridView.Height += status ? 50 : -50;
        }
        private void StatusListing_Click(object sender, EventArgs e)
        {
            Image someImage;
            someImage = StatusListing.Image;
            StatusListing.Image = RotateImage(someImage, statusFlag);
            ToggleVisibility(statusFlag);
            statusFlag = statusFlag ? false : true;
        }

        private void CloseComPort_Click(object sender, EventArgs e)
        {
            isReadingData = false;

            if (_serialPort != null && _serialPort.IsOpen)
            {
                _serialPort.Close();
                ConnectGrin.BackColor = Color.DarkGreen; //Lime\DarkGreen
                ConnectGreenText.Font = new Font(ConnectGreenText.Font, FontStyle.Regular);
                ConnectGreenText.ForeColor = Color.Gray;
                ConnectYellow.BackColor = Color.Yellow; //Yellow\Olive
                ConnectYellowText.Font = new Font(ConnectYellowText.Font, FontStyle.Bold);
                ConnectYellowText.ForeColor = Color.Black;
                ConnectRed.BackColor = Color.Maroon; //Red\Maroon
                ConnectRedText.Font = new Font(ConnectRedText.Font, FontStyle.Regular);
                ConnectRedText.ForeColor = Color.Gray;

                ConnectGrinMini.BackColor = Color.DarkGreen;
                ConnectYellowMini.BackColor = Color.Yellow;
                ConnectRedMini.BackColor = Color.Maroon;
            }
        }

        private void StopReading_Click(object sender, EventArgs e)
        {
            isReadingData = false;
        }

        private async void ContinueReading_Click(object sender, EventArgs e)
        {
            if (_serialPort.IsOpen)
            {
                isReadingData = true;
                await Task.Run(() => StartDataReading());
            }
            else
            {
                MessageBox.Show("Не удалось открыть COM-порт. Пожалуйста, проверьте подключение.");
            }
        }

        private void SensorGridView_CellPainting_1(object sender, DataGridViewCellPaintingEventArgs e)
        {
            if (e.RowIndex >= 0 && e.RowIndex <= 2)
            {
                //Проверка на пустую ячейку
                if (string.IsNullOrEmpty(e.Value?.ToString()))
                {
                    e.AdvancedBorderStyle.Bottom = DataGridViewAdvancedCellBorderStyle.None;
                }

                //Управление границами между повторяющимися ячейками
                if (e.ColumnIndex > 0 && e.ColumnIndex < SensorGridView.ColumnCount - 1)
                {
                    var currentValue = e.Value?.ToString();
                    var previousValue = SensorGridView[e.ColumnIndex - 1, e.RowIndex].Value?.ToString();
                    var nextValue = SensorGridView[e.ColumnIndex + 1, e.RowIndex].Value?.ToString();

                    if ((currentValue == previousValue) && !string.IsNullOrEmpty(currentValue))
                    {
                        e.AdvancedBorderStyle.Left = DataGridViewAdvancedCellBorderStyle.None;
                    }

                    if ((currentValue == nextValue) && !string.IsNullOrEmpty(currentValue))
                    {
                        e.AdvancedBorderStyle.Right = DataGridViewAdvancedCellBorderStyle.None;
                    }
                }
            }
        }

        private void SaveDataToFile_Click(object sender, EventArgs e)
        {
            SaveFileDialog saveFileDialog = new SaveFileDialog
            {
                Filter = "Excel Files|*.xlsx|All Files|*.*",
                Title = "Сохранить данные как Excel файл",
                FileName = "test.xlsx"
            };

            if (saveFileDialog.ShowDialog() == DialogResult.OK)
            {
                Excel.Application excel = null;
                Excel._Worksheet worksheet = null;
                try
                {
                    // Создаём экземпляр приложения Excel
                    excel = new Excel.Application();

                    // Добавляем новую книгу
                    excel.Workbooks.Add();

                    // Получаем активный лист
                    worksheet = (Excel._Worksheet)excel.ActiveSheet;
                    // Устанавливаем название листа
                    worksheet.Name = "Данные с платы";

                    // Заполняем заголовки столбцов
                    for (int i = 1; i < cellDescriptions.Count(); i++)
                    {
                        worksheet.Cells[1, i] = cellDescriptions[i];
                    }

                    // Заполняем данные
                    for (int i = 0; i < excelOutput.Count(); i++)
                    {
                        for (int j = 0; j < excelOutput[i].Count(); j++)
                        {
                            worksheet.Cells[i + 2, j + 1] = excelOutput[i][j];
                        }
                    }

                    // Выравниваем данные по центру и автоматически подгоняем ширину столбцов
                    worksheet.UsedRange.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                    worksheet.Columns.AutoFit();

                    // Сохраняем лист
                    worksheet.SaveAs(saveFileDialog.FileName);
                }
                catch (Exception ex)
                {
                    // Обработка исключений
                    MessageBox.Show(ex.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                finally
                {
                    // Очищаем ресурсы
                    if (worksheet != null)
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(worksheet);
                    if (excel != null)
                    {
                        excel.Quit();
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(excel);
                    }

                    MessageBox.Show("Данные успешно экспортированы в Excel.", "Экспорт завершен", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
        }

        private void GoFullscreen(bool fullscreen)
        {
            Screen currentScreen = Screen.FromControl(this);

            int monitorWid = currentScreen.Bounds.Width;
            int monitorHei = currentScreen.Bounds.Height;

            //Установка состояния окна и изменение ширины панели
            this.WindowState = fullscreen ? FormWindowState.Normal : FormWindowState.Maximized;
            panel1.Width = fullscreen ? panel1.Width / 3 : panel1.Width * 3;

            //Устанока позиций кнопок
            button1.Location = fullscreen ? new Point(766) : new Point(monitorWid - 35);
            button2.Location = fullscreen ? new Point(766-70) : new Point(monitorWid - 105);
            OpenFormFull.Location = fullscreen ? new Point(766 - 35) : new Point(monitorWid - 70);

            //Настройка размеров SensorGridView
            SensorGridView.Width = fullscreen ? 772 : monitorWid - 32;
            SensorGridView.Height = fullscreen ? (statusFlag ? 300 : 350) : (statusFlag ? monitorHei - 250 : monitorHei - 200);
        }
        private void OpenFormFull_Click(object sender, EventArgs e)
        {
            GoFullscreen(this.WindowState == FormWindowState.Maximized);
        }

        private void ClearButton_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < SensorGridView.Rows.Count; i++)
            {
                SensorGridView.Rows.RemoveAt(3);
            }
        }

        private Dictionary<(int Row, int Column), object> originalValues = new Dictionary<(int Row, int Column), object>();
        private void CloseColumn(object sender, System.EventArgs e)
        {
            if (SensorGridView[currentMouseOver, 1].Value == null)
            {
                
            }
            else
            {
                spisokPropuskov[currentMouseOver] = "true";
                for (int i = 3; i < SensorGridView.Rows.Count; i++)
                {
                    var cell = SensorGridView.Rows[i].Cells[currentMouseOver];
                    if (cell.Value != null)
                    {
                        originalValues[(i, currentMouseOver)] = cell.Value;
                        cell.Value = "";
                    }
                }
            }
        }

        private void ReturnAllColumns(object sender, System.EventArgs e)
        {
            int row = SensorGridView.Rows.Count - 3;
            foreach (var key in originalValues.Keys)
            {
                SensorGridView.Rows[row - (row-key.Row)].Cells[key.Column].Value = originalValues[key];
            }

            originalValues.Clear();
            spisokPropuskov.Clear();
        }

        private void SensorGridView_MouseClick(object sender, MouseEventArgs e)
        {
            if (isReadingData == false)
            {
                if (e.Button == MouseButtons.Right)
                {
                    var currentMO = SensorGridView.HitTest(e.X, e.Y);
                    currentMouseOver = currentMO.ColumnIndex;
                    if (currentMO.RowIndex >= 3 && currentMO.ColumnIndex >=2 && currentMO.ColumnIndex != 5 && currentMO.ColumnIndex != 13 && currentMO.ColumnIndex != 23 && currentMO.ColumnIndex != 34)
                    {
                        ContextMenu m = new ContextMenu();
                        MenuItem hideMenuItem = new MenuItem();
                        MenuItem returnAllMenuItem = new MenuItem();

                        hideMenuItem.Text = "&Спрятать";
                        returnAllMenuItem.Text = "&Вернуть все столбцы";

                        m.MenuItems.Add(hideMenuItem);
                        m.MenuItems.Add(returnAllMenuItem);

                        hideMenuItem.Click += new System.EventHandler(this.CloseColumn);
                        returnAllMenuItem.Click += new System.EventHandler(this.ReturnAllColumns);

                        m.Show(SensorGridView, new Point(e.X, e.Y));
                    }
                }
            }
            else
            {
                if (e.Button == MouseButtons.Right)
                {
                    ContextMenu m = new ContextMenu();
                    MenuItem hideMenuItem = new MenuItem();

                    hideMenuItem.Text = "&Для работы с таблицей остановите чтение.";

                    m.MenuItems.Add(hideMenuItem);

                    m.Show(SensorGridView, new Point(e.X, e.Y));
                }
            }
        }

        private void selectTableBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            string selectedTemplate = selectTableBox.SelectedItem.ToString();
            if (!tableTemplates.TryGetValue(selectedTemplate, out int[] template))
            {
                return; // Если шаблон не найден, прерываем выполнение
            }
            SaveCurrentGridViewData(lastName);
            LoadDataIntoGridView(tableTemplates[$"{selectedTemplate}"]);

            if (selectedTemplate == "Магнитные азимуты")
            {
                SensorGridView.Rows.RemoveAt(1);
                SensorGridView.Rows.Insert(2, 1);
            }

            LoadGridViewData(selectedTemplate);
            RepairDataInGridView();
            lastName = selectedTemplate;
        }

        private void label1_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            UpdateComPortsList();
        }
    }
}
