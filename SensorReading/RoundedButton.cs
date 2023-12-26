using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace SensorReading
{
    public class RoundedButton : Button
    {
        protected override void OnPaint(PaintEventArgs pevent)
        {
            base.OnPaint(pevent);

            // Рисуем кнопку с скругленными углами
            Graphics g = pevent.Graphics;
            using (var path = new System.Drawing.Drawing2D.GraphicsPath())
            {
                int radius = 10; // Задай радиус скругления
                path.AddArc(0, 0, radius * 2, radius * 2, 180, 90);
                path.AddArc(this.Width - radius * 2, 0, radius * 2, radius * 2, 270, 90);
                path.AddArc(this.Width - radius * 2, this.Height - radius * 2, radius * 2, radius * 2, 0, 90);
                path.AddArc(0, this.Height - radius * 2, radius * 2, radius * 2, 90, 90);
                g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;
                this.Region = new Region(path);
                using (var pen = new Pen(this.FlatAppearance.BorderColor, 1.75f))
                {
                    g.DrawPath(pen, path);
                }
            }
        }
    }
}
