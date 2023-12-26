using System;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO.Ports;
using System.Globalization;
using ClosedXML.Excel;
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
            { 2 , " RM3100 магнит X" },
            { 3 , " RM3100 магнит Y" },
            { 4 , " RM3100 магнит Z" },
            { 5 , " MTI temp" },
            { 6 , " MTI магнит X" },
            { 7 , " MTI магнит Y" },
            { 8 , " MTI магнит Z" },
            { 9 , " MTI акс X" },
            { 10 , " MTI акс Y" },
            { 11 , " MTI акс Z" },
            { 12 , " PNI (Курс)" },
            { 13 , " PNI (Тангаж)" },
            { 14 , " PNI (Крен)" },
            { 15 , " PNI акс X" },
            { 16 , " PNI акс Y" },
            { 17 , " PNI акс Z" },
            { 18 , " PNI магнит X" },
            { 19 , " PNI магнит Y" },
            { 20 , " PNI магнит Z" },
            { 21 , " ADIS акс X" },
            { 22 , " ADIS акс Y" },
            { 23 , " ADIS акс Z" },
            { 24 , " ADIS магнит X" },
            { 25 , " ADIS магнит Y" },
            { 26 , " ADIS магнит Z" },
            { 27 , " ADIS гироскоп X" },
            { 28 , " ADIS гироскоп Y" },
            { 29 , " ADIS гироскоп Z" },
            { 30 , " ADIS temp" },
            { 31 , " RM3100 MAz" },
            { 32 , " MTI MAz" },
            { 33 , " PNI MAz" },
            { 34 , " ADIS MAz" }
        };

        //Шаблоны таблиц
        int[] allColumns = new int[] { 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29, 30 };
        Dictionary<string, int[]> tableTemplates = new Dictionary<string, int[]>
        {
            {"Все данные", new int[] { 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29, 30 } }, //allColumns
            {"Магнитные азимуты", new int[]{ 0, 1, 31, 32, 33, 34 } }  //Mag Azimuth
        };
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

            int i = 0;
            foreach (string s in SerialPort.GetPortNames())
            {
                ComConnectorsList.Items.Insert(i, s);
                i++;
            }
            MergeHeaderGrid();
            SetRoundedShape(ConnectGrin, 27);
            SetRoundedShape(ConnectYellow, 27);
            SetRoundedShape(ConnectRed, 27);
            SetRoundedShape(ConnectGrinMini, 22);
            SetRoundedShape(ConnectYellowMini, 22);
            SetRoundedShape(ConnectRedMini, 22);
            SensorGridView.ClearSelection();
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
            foreach (DataGridViewRow row in SensorGridView.Rows)
            {

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
                SensorGridView.Columns[col].Width = 70;
                SensorGridView.Columns[col].Width = 70;
            }
            SensorGridView.Columns[0].Width = 80;
            SensorGridView.Columns[0].Width = 80;
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
                {
                    List<string> results = new List<string>(response.Split(';'));
                    int len = SensorGridView.Rows.Count;
                    //Обновляем SensorGridView в основном потоке
                    BeginInvoke(new Action(() =>
                    {
                        int rowIndex = 3;
                        SensorGridView.Rows.Insert(rowIndex,1);
                        SensorGridView.Rows[rowIndex].Cells[0].Value = len-3;
                        string hours = $"{int.Parse(results[0], CultureInfo.InvariantCulture) / 1000 / 60 / 60}";
                        string minutes = $"{int.Parse(results[0], CultureInfo.InvariantCulture) / 1000 / 60 % 60}";
                        string seconds = $"{int.Parse(results[0], CultureInfo.InvariantCulture) / 1000 % 60}";
                        if (seconds.Length < 2)
                        {
                            seconds = "0" + seconds;
                        }
                        if (minutes.Length < 2)
                        {
                            minutes = "0" + minutes;
                        }
                        if (hours.Length < 2)
                        {
                            hours = "0" + hours;
                        }
                        string times = $"{hours}:{minutes}:{seconds}";
                        SensorGridView.Rows[rowIndex].Cells[1].Value = times;
                        //Общие значения (Не трогать)
                        double Pitch = Math.Round(double.Parse(results[12], CultureInfo.InvariantCulture), 2) * Math.PI / 180;// значение Тангажа (например, G3);
                        double Roll = Math.Round(double.Parse(results[13], CultureInfo.InvariantCulture), 2) * Math.PI / 180;// значение Крена (например, H3);

                        //Для RM3100
                        double X = Math.Round(double.Parse(results[1], CultureInfo.InvariantCulture), 2);// значение X (например, B3);
                        double Y = Math.Round(double.Parse(results[2], CultureInfo.InvariantCulture), 2);// значение Y (например, C3);
                        double Z = Math.Round(double.Parse(results[3], CultureInfo.InvariantCulture), 2);// значение Z (например, D3);
                        //Для PNI
                        double XPNITEMP = Math.Round(double.Parse(results[17], CultureInfo.InvariantCulture), 2);//
                        double YPNITEMP = Math.Round(double.Parse(results[18], CultureInfo.InvariantCulture), 2);//
                        double ZPNITEMP = Math.Round(double.Parse(results[19], CultureInfo.InvariantCulture), 2);//
                        //Для ADIS
                        double XADISTEMP = Math.Round(double.Parse(results[23], CultureInfo.InvariantCulture), 2);
                        double YADISTEMP = Math.Round(double.Parse(results[24], CultureInfo.InvariantCulture), 2);
                        double ZADISTEMP = Math.Round(double.Parse(results[25], CultureInfo.InvariantCulture), 2);
                        //Для MTI
                        double XMTITEMP = Math.Round(double.Parse(results[5], CultureInfo.InvariantCulture), 2);
                        double YMTITEMP = Math.Round(double.Parse(results[6], CultureInfo.InvariantCulture), 2);
                        double ZMTITEMP = Math.Round(double.Parse(results[7], CultureInfo.InvariantCulture), 2);

                        //Таблицы переходов
                        //PNI
                        double XPNI = -YPNITEMP;
                        double YPNI = XPNITEMP;
                        double ZPNI = ZPNITEMP;
                        //MTI
                        double XMTI = YMTITEMP;
                        double YMTI = XMTITEMP;
                        double ZMTI = -ZMTITEMP;
                        //ADIS
                        double XADIS = YADISTEMP;
                        double YADIS = XADISTEMP;
                        double ZADIS = -ZADISTEMP;

                        //Excel -atan2(x;y) C# -atan2(y;x)
                        //X - X * COS(Pitch) + Y * SIN(Pitch) * SIN(Roll) + Z * SIN(Pitch) * COS(Roll)
                        //Y - Z * SIN(Roll) - Y * COS(Roll)

                        // Расчет магнитного азимута
                        double azimuthRM3100 = Math.Atan2(Z * Math.Sin(Roll) - Y * Math.Cos(Roll), X * Math.Cos(Pitch) + Y * Math.Sin(Pitch) * Math.Sin(Roll) + Z * Math.Sin(Pitch) * Math.Cos(Roll));
                        double azimuthPNI = Math.Atan2(ZPNI * Math.Sin(Roll) - YPNI * Math.Cos(Roll), XPNI * Math.Cos(Pitch) + YPNI * Math.Sin(Pitch) * Math.Sin(Roll) + ZPNI * Math.Sin(Pitch) * Math.Cos(Roll));
                        double azimuthMTI = Math.Atan2(ZMTI * Math.Sin(Roll) - YMTI * Math.Cos(Roll), XMTI * Math.Cos(Pitch) + YMTI * Math.Sin(Pitch) * Math.Sin(Roll) + ZMTI * Math.Sin(Pitch) * Math.Cos(Roll));
                        double azimuthADIS = Math.Atan2(ZADIS * Math.Sin(Roll) - YADIS * Math.Cos(Roll), XADIS * Math.Cos(Pitch) + YADIS * Math.Sin(Pitch) * Math.Sin(Roll) + ZADIS * Math.Sin(Pitch) * Math.Cos(Roll));

                        // Преобразование из радиан в градусы
                        azimuthRM3100 *= 180 / Math.PI;
                        azimuthPNI *= 180 / Math.PI;
                        azimuthMTI *= 180 / Math.PI;
                        azimuthADIS *= 180 / Math.PI;

                        //Корректировки для ортогональных преобразований
                        if (azimuthRM3100 < 0)
                        {
                            azimuthRM3100 += 360;
                        } else if (azimuthRM3100 >= 360)
                        {
                            azimuthRM3100 -= 360;
                        }

                        if (azimuthPNI < 0)
                        {
                            azimuthPNI += 360;
                        }
                        else if (azimuthPNI >= 360)
                        {
                            azimuthPNI -= 360;
                        }

                        if (azimuthMTI < 0)
                        {
                            azimuthMTI += 360;
                        }
                        else if (azimuthMTI >= 360)
                        {
                            azimuthMTI -= 360;
                        }

                        if (azimuthADIS < 0)
                        {
                            azimuthADIS += 360;
                        }
                        else if (azimuthADIS >= 360)
                        {
                            azimuthADIS -= 360;
                        }
                        string selectedTemplate = selectTableBox.SelectedItem.ToString();
                        results.Add(azimuthRM3100.ToString().Replace(',', '.'));
                        results.Add(azimuthPNI.ToString().Replace(',', '.'));
                        results.Add(azimuthMTI.ToString().Replace(',', '.'));
                        results.Add(azimuthADIS.ToString().Replace(',', '.'));

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
            int monitorWid = Screen.PrimaryScreen.Bounds.Width;
            int monitorHei = Screen.PrimaryScreen.Bounds.Height;

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
    }
}
