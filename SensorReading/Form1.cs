using System;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO.Ports;
using System.Globalization;
using System.Reflection;
using System.Collections.Generic;

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
        //Предыдущая таблица
        public string lastName = "Все данные";
        //Словарь для хранения данных по каждому шаблону
        Dictionary<string, List<string[]>> templatesData = new Dictionary<string, List<string[]>>();

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
            { 37 , " ADIS (Крен)" }
        };

        //Шаблоны таблиц
        int[] allColumns = new int[] { 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29, 30 };
        Dictionary<string, int[]> tableTemplates = new Dictionary<string, int[]>
        {
            {"Все данные", new int[] { 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29, 30 } }, //Все колонки
            {"Магнитные азимуты", new int[]{ 0, 1, 31, 32, 33, 34 } },  //Магнитные азимуты
            {"Акселерометры", new int[]{ 0, 1, 9, 10, 11, 15, 16, 17, 21, 22, 23} }, //Акселерометры
            {"Магнитометры", new int[] { 0, 1, 2, 3, 4, 6, 7, 8, 18, 19, 20, 24, 25, 26} }, //Магнитометры
            {"Круги Эйлера", new int[] {0, 1, 12, 13, 14, 35, 36, 37} } //Курс, Крен, Тангаж
        };
        public class MadgwickFilter
        {
            // Кватернионы ориентации
            private float q1 = 1f, q2 = 0f, q3 = 0f, q4 = 0f;

            // Коэффициент фильтра
            public float Beta { get; set; } = 0.1f;

            // Обновление фильтра с использованием данных гироскопа и акселерометра
            public (float Roll, float Pitch, float Yaw) Update(float gx, float gy, float gz, float ax, float ay, float az, int dataPeriod)
            {
                float sampleFreq = 1000f / dataPeriod; // Частота в Гц, рассчитанная из периода накопления данных в мс
                float recipNorm;
                float s1, s2, s3, s4;
                float qDot1, qDot2, qDot3, qDot4;

                // Нормализация данных акселерометра
                recipNorm = (float)Math.Sqrt(ax * ax + ay * ay + az * az);
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
                recipNorm = (float)Math.Sqrt(s1 * s1 + s2 * s2 + s3 * s3 + s4 * s4);
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
                recipNorm = (float)Math.Sqrt(q1 * q1 + q2 * q2 + q3 * q3 + q4 * q4);
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
            EnableDoubleBuffering(this.SensorGridView);
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
                                //SensorGridView[i, row].Value = "";
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

            int totalwidth = SensorGridView.RowHeadersWidth + 1;

            for (int i = 0; i < SensorGridView.Columns.Count; i++)
            {
                totalwidth += SensorGridView.Columns[i].Width;
            }
        }

        private void MainForm_FormClosed(object sender, FormClosedEventArgs e)
        {
            Application.Exit();
        }
        
        private void button1_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void button1_MouseEnter(object sender, EventArgs e)
        {
            button1.BackColor = Color.Red;
        }

        private void button1_MouseLeave(object sender, EventArgs e)
        {
            button1.BackColor = Color.Transparent;
        }

        private void button2_MouseEnter(object sender, EventArgs e)
        {
            button2.BackColor = Color.RoyalBlue;
        }

        private void button2_MouseLeave(object sender, EventArgs e)
        {
            button2.BackColor = Color.Transparent;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.WindowState = FormWindowState.Minimized;
        }

        private void panel1_MouseDown(object sender, MouseEventArgs e)
        {
            panel1.Capture = false;
            Message m = Message.Create(base.Handle, 0xa1, new IntPtr(2), IntPtr.Zero);
            this.WndProc(ref m);
        }

        private string SendAndReadData(string dataPeriod)
        {
            try
            {
                if (_serialPort.IsOpen)
                {
                    //отправляем данные на плату
                    _serialPort.Write(dataPeriod);
                    //читаем ответ
                    string answer = "";
                    try
                    {
                        answer = _serialPort.ReadLine();
                    }
                    catch (TimeoutException e)
                    {
                        Console.WriteLine($"Данные с датчика не поступают: {e.Message}");
                    }
                    return answer;
                }
                else
                {
                    //COM-порт не открыт
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
            if (!tableTemplates.TryGetValue(selectedTemplate, out int[] template))
            {
                return; // Если шаблон не найден, прерываем выполнение
            }
            for (int i = 2; i < template.Length && i < results.Count; i++)
            {
                int column = i;
                double value;
                if (double.TryParse(results[template[i]-1], NumberStyles.Any, CultureInfo.InvariantCulture, out value))
                {
                    SensorGridView.Rows[rowIndex].Cells[column].Value = Math.Round(value, 2);
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
            double roll = Math.Atan(Y / Z) * 180 / Math.PI -90;
            double pitch = -Math.Atan(-X / Math.Sqrt(Y * Y + Z * Z)) * 180 / Math.PI;
            double yaw = Math.Atan2(my , mx) * 180 / Math.PI+90;
            
            return (roll, pitch, yaw);
        }
        private async void StartDataReading()
        {
            while (isReadingData) //метка для остановки чтения
            {
                string dataPeriod = string.Empty;

                Invoke(new Action(() =>
                {
                    dataPeriod = DataPeriodBox.SelectedItem.ToString();
                }));

                string response = await Task.Run(() => SendAndReadData(dataPeriod));

                if (!string.IsNullOrEmpty(response))
                {//Excel -atan2(x;y) C# -atan2(y;x)
                 //X - X * COS(Pitch) + Y * SIN(Pitch) * SIN(Roll) + Z * SIN(Pitch) * COS(Roll)
                 //Y - Z * SIN(Roll) - Y * COS(Roll)
                    List<string> results = new List<string>(response.Split(';'));
                    int len = SensorGridView.Rows.Count;
                    //Обновляем SensorGridView в основном потоке
                    BeginInvoke(new Action(() =>
                    {
                        int rowIndex = 3;
                        SensorGridView.Rows.Insert(rowIndex,1);
                        SensorGridView.Rows[rowIndex].Cells[0].Value = len-3;
                        int totalTimeInSeconds = int.Parse(results[0], CultureInfo.InvariantCulture) / 1000;
                        // Преобразование общего времени в секундах в часы, минуты и секунды
                        TimeSpan timeSpan = TimeSpan.FromSeconds(totalTimeInSeconds);
                        // Форматирование времени в формате "HH:mm:ss"
                        string formattedTime = timeSpan.ToString(@"hh\:mm\:ss");
                        SensorGridView.Rows[rowIndex].Cells[1].Value = formattedTime;

                        //Общие значения (Не трогать)
                        double Pitch = Math.Round(double.Parse(results[12], CultureInfo.InvariantCulture), 2) * Math.PI / 180;// значение Тангажа (например, G3);
                        double Roll = Math.Round(double.Parse(results[13], CultureInfo.InvariantCulture), 2) * Math.PI / 180;// значение Крена (например, H3);

                        //Для RM3100
                        double X = Math.Round(double.Parse(results[1], CultureInfo.InvariantCulture), 2);// значение X (например, B3);
                        double Y = Math.Round(double.Parse(results[2], CultureInfo.InvariantCulture), 2);// значение Y (например, C3);
                        double Z = Math.Round(double.Parse(results[3], CultureInfo.InvariantCulture), 2);// значение Z (например, D3);
                        //Для PNI
                        double YPNI = Math.Round(double.Parse(results[17], CultureInfo.InvariantCulture), 2);//
                        double XPNI = -Math.Round(double.Parse(results[18], CultureInfo.InvariantCulture), 2);//
                        double ZPNI = Math.Round(double.Parse(results[19], CultureInfo.InvariantCulture), 2);//
                        //Для ADIS
                        double YADIS = Math.Round(double.Parse(results[23], CultureInfo.InvariantCulture), 2);
                        double XADIS = Math.Round(double.Parse(results[24], CultureInfo.InvariantCulture), 2);
                        double ZADIS = -Math.Round(double.Parse(results[25], CultureInfo.InvariantCulture), 2);
                        //Для MTI
                        double YMTI = Math.Round(double.Parse(results[5], CultureInfo.InvariantCulture), 2);
                        double XMTI = Math.Round(double.Parse(results[6], CultureInfo.InvariantCulture), 2);
                        double ZMTI = -Math.Round(double.Parse(results[7], CultureInfo.InvariantCulture), 2);

                        // Расчет магнитного азимута
                        double azimuthRM3100 = MagAz(Pitch, Roll, X, Y, Z);
                        double azimuthPNI = MagAz(Pitch, Roll, XPNI, YPNI, ZPNI);
                        double azimuthMTI = MagAz(Pitch, Roll, XMTI, YMTI, ZMTI);
                        double azimuthADIS = MagAz(Pitch, Roll, XADIS, YADIS, ZADIS);

                        results.Add(azimuthRM3100.ToString().Replace(',', '.'));
                        results.Add(azimuthPNI.ToString().Replace(',', '.'));
                        results.Add(azimuthMTI.ToString().Replace(',', '.'));
                        results.Add(azimuthADIS.ToString().Replace(',', '.'));

                        MadgwickFilter mgF = new MadgwickFilter();

                        float ax = float.Parse(results[21], CultureInfo.InvariantCulture);
                        float ay = float.Parse(results[22], CultureInfo.InvariantCulture);
                        float az = float.Parse(results[23], CultureInfo.InvariantCulture);
                        float gx = float.Parse(results[27], CultureInfo.InvariantCulture);
                        float gy = float.Parse(results[28], CultureInfo.InvariantCulture);
                        float gz = float.Parse(results[29], CultureInfo.InvariantCulture);
                        var eulerAngles = mgF.Update( gx, gy, gz, ax, ay, az, int.Parse(dataPeriod));

                        (float roll, float pitch, float yaw) = eulerAngles;
                        //results.Add(roll.ToString());
                        //results.Add(pitch.ToString());
                        //results.Add(yaw.ToString());

                        (double r, double p, double y) = AccMagCounting(ax, ay, az, double.Parse(results[24], CultureInfo.InvariantCulture), double.Parse(results[25], CultureInfo.InvariantCulture));
                        results.Add(y.ToString().Replace(',','.'));
                        results.Add(p.ToString().Replace(',', '.'));
                        results.Add(r.ToString().Replace(',', '.'));
                        string selectedTemplate = selectTableBox.SelectedItem.ToString();
                        WriteGridView(results, selectedTemplate, rowIndex);
                    }));
                }
            }
        }

        private async void ConnectComPort_Click(object sender, EventArgs e)
        {
            if (ComConnectorsList.SelectedIndex == -1)
            {
                MessageBox.Show("Укажите COM-порт для чтения.","COM-порт не указан", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            if (DataPeriodBox.SelectedIndex == -1)
            {
                MessageBox.Show("Укажите период накопления данных.", "Период накопления данных не указан", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            string selectedPortName = ComConnectorsList.SelectedItem.ToString();
            string[] availablePorts = SerialPort.GetPortNames();
            if (availablePorts.Contains(selectedPortName))
            {
                try
                {
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

                    _serialPort.Open();

                    if (_serialPort.IsOpen)
                    {
                        isReadingData = true;
                        await Task.Run(() => StartDataReading());
                        ConnectGrin.BackColor = Color.Lime; //Lime\DarkGreen
                        ConnectGreenText.Font = new Font(ConnectGreenText.Font, FontStyle.Bold);
                        ConnectGreenText.ForeColor = Color.Black;
                        ConnectYellow.BackColor = Color.Olive; //Yellow\Olive
                        ConnectYellowText.Font = new Font(ConnectYellowText.Font, FontStyle.Regular);
                        ConnectYellowText.ForeColor = Color.Gray;
                        ConnectRed.BackColor = Color.Maroon;
                        ConnectRedText.ForeColor = Color.Gray;
                        ConnectRedText.Font = new Font(ConnectRedText.Font, FontStyle.Regular);

                        ConnectGrinMini.BackColor = Color.Lime;
                        ConnectYellowMini.BackColor = Color.Olive;
                        ConnectRedMini.BackColor = Color.Maroon;
                    }
                    else
                    {
                        MessageBox.Show("Не удалось открыть COM-порт. Пожалуйста, проверьте подключение.", "Ошибка подключения", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        ConnectGrin.BackColor = Color.DarkGreen; //Lime\DarkGreen
                        ConnectGreenText.Font = new Font(ConnectGreenText.Font, FontStyle.Regular);
                        ConnectGreenText.ForeColor = Color.Gray;
                        ConnectYellow.BackColor = Color.Olive; //Yellow\Olive
                        ConnectYellowText.Font = new Font(ConnectYellowText.Font, FontStyle.Regular);
                        ConnectYellowText.ForeColor = Color.Gray;
                        ConnectRed.BackColor = Color.Red; //Red\Maroon
                        ConnectRedText.Font = new Font(ConnectRedText.Font, FontStyle.Bold);
                        ConnectRedText.ForeColor = Color.Black;

                        ConnectGrinMini.BackColor = Color.DarkGreen;
                        ConnectYellowMini.BackColor = Color.Olive;
                        ConnectRedMini.BackColor = Color.Red;
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Произошла ошибка при подключении к COM-порту: {ex.Message}", "Ошибка подключения", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    ConnectGrin.BackColor = Color.DarkGreen; //Lime\DarkGreen
                    ConnectGreenText.Font = new Font(ConnectGreenText.Font, FontStyle.Regular);
                    ConnectGreenText.ForeColor = Color.Gray;
                    ConnectYellow.BackColor = Color.Olive; //Yellow\Olive
                    ConnectYellowText.Font = new Font(ConnectYellowText.Font, FontStyle.Regular);
                    ConnectYellowText.ForeColor = Color.Gray;
                    ConnectRed.BackColor = Color.Red; //Red\Maroon
                    ConnectRedText.Font = new Font(ConnectRedText.Font, FontStyle.Bold);
                    ConnectRedText.ForeColor = Color.Black;
                    ConnectGrinMini.BackColor = Color.DarkGreen;
                    ConnectYellowMini.BackColor = Color.Olive;
                    ConnectRedMini.BackColor = Color.Red;
                }
            }
            else
            {
                MessageBox.Show($"Выбранный COM-порт '{selectedPortName}' не существует. Пожалуйста, выберите корректный порт.");
                ConnectGrin.BackColor = Color.DarkGreen; //Lime\DarkGreen
                ConnectGreenText.Font = new Font(ConnectGreenText.Font, FontStyle.Regular);
                ConnectGreenText.ForeColor = Color.Gray;
                ConnectYellow.BackColor = Color.Olive; //Yellow\Olive
                ConnectYellowText.Font = new Font(ConnectYellowText.Font, FontStyle.Regular);
                ConnectYellowText.ForeColor = Color.Gray;
                ConnectRed.BackColor = Color.Red; //Red\Maroon
                ConnectRedText.Font = new Font(ConnectRedText.Font, FontStyle.Bold);
                ConnectRedText.ForeColor = Color.Black;
                ConnectGrinMini.BackColor = Color.DarkGreen;
                ConnectYellowMini.BackColor = Color.Olive;
                ConnectRedMini.BackColor = Color.Red;
            }
        }
        public Image rotateImage(Image someImage, bool flag)
        {
            if (flag == true)
            {
                someImage.RotateFlip(RotateFlipType.Rotate270FlipNone);
                return someImage;
            }
            else
            {
                someImage.RotateFlip(RotateFlipType.Rotate90FlipNone);
                return someImage;
            }
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
            StatusListing.Image = rotateImage(someImage, statusFlag);
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
            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.Filter = "Excel Files|*.xlsx|All Files|*.*";
            saveFileDialog.Title = "Сохранить данные как Excel файл";
            saveFileDialog.FileName = "test.xlsx";

            if (saveFileDialog.ShowDialog() == DialogResult.OK)
            {
                // Создаём новый экземпляр Excel и книгу
                Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();
                excel.Workbooks.Add();
                Microsoft.Office.Interop.Excel._Worksheet worksheet = (Microsoft.Office.Interop.Excel._Worksheet)excel.ActiveSheet;

                for (int i = 0; i < SensorGridView.Rows.Count - 1; i++)
                {
                    for (int j = 0; j < SensorGridView.Columns.Count; j++)
                    {
                        worksheet.Cells[i + 1, j + 1] = SensorGridView.Rows[i].Cells[j].Value.ToString();
                    }
                }

                worksheet.SaveAs(saveFileDialog.FileName);
                excel.Quit();

                // Очищаем ресурсы
                System.Runtime.InteropServices.Marshal.ReleaseComObject(worksheet);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excel);

                MessageBox.Show("Данные успешно экспортированы в Excel.", "Экспорт завершен", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void ConnectGrin_Paint(object sender, PaintEventArgs e)
        {

        }

        private void MainForm_Load(object sender, EventArgs e)
        {

        }

        private void OpenFormFull_MouseEnter(object sender, EventArgs e)
        {
            OpenFormFull.BackColor = Color.LightBlue;
        }

        private void OpenFormFull_MouseLeave(object sender, EventArgs e)
        {
            OpenFormFull.BackColor = Color.Transparent;
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
                SensorGridView.Rows.RemoveAt(i);
            }
        }

        private void GridViewBar_Scroll(object sender, ScrollEventArgs e)
        {
            SensorGridView.HorizontalScrollingOffset = e.NewValue;
        }

        private void SensorGridView_Scroll(object sender, ScrollEventArgs e)
        {
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

        private void SensorGridView_ColumnRemoved(object sender, DataGridViewColumnEventArgs e)
        {
            
        }

        private void MagAzimuth_CheckedChanged(object sender, EventArgs e)
        {
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
