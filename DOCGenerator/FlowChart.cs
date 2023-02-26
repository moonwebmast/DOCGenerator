using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.IO;
using System.Text;

namespace DOCGenerator
{
    public class FlowChart
    {
        /// <summary>
        /// 绘制流程图
        /// </summary>
        /// <param name="items">流程节点</param>
        /// <returns>图片内容</returns>
        public static byte[] Draw(string[] items)
        {


            using (MemoryStream stream = new MemoryStream())
            {
                Bitmap bitMap = new Bitmap(600, 1800);
                var g = Graphics.FromImage(bitMap);

                #region 测试图片高度
                var testPoint = Measure(g, "开始", 300, 50, 100);
                testPoint = Measure(g, testPoint);
                foreach (var item in items)
                {
                    testPoint = Measure(g, item.Trim(), testPoint);
                    testPoint = Measure(g, testPoint);
                }
                testPoint = Measure(g, "结束", testPoint, 100);
                #endregion


                bitMap = new Bitmap(600, (int)testPoint.Y + 50);
                g = Graphics.FromImage(bitMap);
                PointF nextPoint = DrawTag(g, "开始", 300, 50, 100, 100);

                nextPoint = DrawLine(g, nextPoint);

                for (int i = 0; i < items.Length; i++)
                {
                    nextPoint = DrawTag(g, items[i].Trim(), nextPoint);
                    nextPoint = DrawLine(g, nextPoint);
                }

                nextPoint = DrawTag(g, "结束", nextPoint, 100, 100);

                bitMap.Save(stream, ImageFormat.Png);
                return stream.ToArray();
            }
        }


        /// <summary>
        /// 预测流程图节点下标位置
        /// </summary>
        /// <param name="g"></param>
        /// <param name="text"></param>
        /// <param name="point"></param>
        /// <param name="width"></param>
        /// <returns></returns>
        private static PointF Measure(Graphics g, string text, PointF point, int width = 200)
        {
            return Measure(g, text, point.X, point.Y, width);
        }

        /// <summary>
        /// 预测流程图节点下标位置
        /// </summary>
        /// <param name="g"></param>
        /// <param name="text"></param>
        /// <param name="x"></param>
        /// <param name="y"></param>
        /// <param name="width"></param>
        /// <returns></returns>
        private static PointF Measure(Graphics g, string text, float x, float y, int width = 200)
        {
            var size = g.MeasureString(text, SystemFonts.DefaultFont, width);

            float h = size.Height + 30;
            float w = width + 40;

            return new PointF(x, y + h);
        }

        /// <summary>
        /// 绘制流程节点
        /// </summary>
        /// <param name="g"></param>
        /// <param name="text"></param>
        /// <param name="point"></param>
        /// <param name="width"></param>
        /// <param name="cornerRadius"></param>
        /// <returns></returns>
        private static PointF DrawTag(Graphics g, string text, PointF point, int width = 200, int cornerRadius = 0)
        {
            return DrawTag(g, text, point.X, point.Y, width, cornerRadius);
        }

        /// <summary>
        /// 绘制流程节点
        /// </summary>
        /// <param name="g"></param>
        /// <param name="text"></param>
        /// <param name="x"></param>
        /// <param name="y"></param>
        /// <param name="width"></param>
        /// <param name="cornerRadius"></param>
        /// <returns></returns>
        private static PointF DrawTag(Graphics g, string text, float x, float y, int width = 200, int cornerRadius = 0)
        {
            //g.DrawRectangle(Pens.Black, 300, 50, 200, 80);
            var size = g.MeasureString(text, SystemFonts.DefaultFont);

            float h = size.Height + 30;
            float w = Math.Max(width, size.Width) + 40;

            g.DrawString(text, SystemFonts.DefaultFont, Brushes.Black, x - size.Width / 2, y + h / 2 - size.Height / 2);

            if (cornerRadius > 0)
            {
                var path = CreateRoundedRectanglePath(new RectangleF(x - w / 2, y, w, h), (int)h / 2);
                g.DrawPath(Pens.Black, path);
            }
            else
            {
                g.DrawRectangle(Pens.Black, x - w / 2, y, w, h);
            }

            return new PointF(x, y + h);
        }

        /// <summary>
        /// 测试箭头长度
        /// </summary>
        /// <param name="g"></param>
        /// <param name="point"></param>
        /// <param name="length"></param>
        /// <returns></returns>
        private static PointF Measure(Graphics g, PointF point, float length = 30)
        {
            return new PointF(point.X, point.Y + length);
        }

        /// <summary>
        /// 画箭头
        /// </summary>
        /// <param name="g"></param>
        /// <param name="point"></param>
        /// <param name="length"></param>
        /// <returns></returns>
        private static PointF DrawLine(Graphics g, PointF point, float length = 30)
        {
            return DrawLine(g, point.X, point.Y, length);
        }

        

        /// <summary>
        /// 画箭头
        /// </summary>
        /// <param name="g"></param>
        /// <param name="x"></param>
        /// <param name="from"></param>
        /// <param name="length"></param>
        /// <returns></returns>
        private static PointF DrawLine(Graphics g, float x, float from, float length = 30)
        {
            g.DrawLine(Pens.Black, x, from, x, from + length);
            g.DrawLine(Pens.Black, x - 3, from + length - 5, x, from + length);
            g.DrawLine(Pens.Black, x + 3, from + length - 5, x, from + length);
            return new PointF(x, from + length);
        }

        /// <summary>
        /// 生成圆角矩形路径
        /// </summary>
        /// <param name="rect"></param>
        /// <param name="cornerRadius"></param>
        /// <returns></returns>
        private static GraphicsPath CreateRoundedRectanglePath(RectangleF rect, int cornerRadius)
        {
            float w = Math.Min(rect.Width, rect.Height);
            if (cornerRadius > w / 2)
            {
                cornerRadius = (int)w / 2;
            }

            GraphicsPath roundedRect = new GraphicsPath();
            // 左上角
            roundedRect.AddArc(rect.X, rect.Y, cornerRadius * 2, cornerRadius * 2, 180, 90);
            //上边线
            if (cornerRadius < (int)rect.Width / 2)
                roundedRect.AddLine(rect.X + cornerRadius, rect.Y, rect.Right - cornerRadius * 2, rect.Y);
            //右上角
            roundedRect.AddArc(rect.X + rect.Width - cornerRadius * 2, rect.Y, cornerRadius * 2, cornerRadius * 2, 270, 90);
            //右边
            if (cornerRadius < (int)rect.Height / 2)
                roundedRect.AddLine(rect.Right, rect.Y + cornerRadius * 2, rect.Right, rect.Y + rect.Height - cornerRadius * 2);

            roundedRect.AddArc(rect.X + rect.Width - cornerRadius * 2, rect.Y + rect.Height - cornerRadius * 2, cornerRadius * 2, cornerRadius * 2, 0, 90);

            if (cornerRadius < (int)rect.Width / 2)
                roundedRect.AddLine(rect.Right - cornerRadius * 2, rect.Bottom, rect.X + cornerRadius * 2, rect.Bottom);

            roundedRect.AddArc(rect.X, rect.Bottom - cornerRadius * 2, cornerRadius * 2, cornerRadius * 2, 90, 90);

            if (cornerRadius < (int)rect.Height / 2)
                roundedRect.AddLine(rect.X, rect.Bottom - cornerRadius * 2, rect.X, rect.Y + cornerRadius * 2);

            roundedRect.CloseFigure();
            return roundedRect;
        }
    }
}
