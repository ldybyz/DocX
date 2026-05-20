using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http.Headers;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;
using Microsoft.SqlServer.Server;
using SkiaSharp;
using ScottPlot;
using MathNet.Numerics;
using System.Globalization;
using System.Numerics;
using ScottPlot.Plottables;
using ScottPlot.Colormaps;
using System.IO.Pipes;


namespace LimsFileHelper
{
    public class ImageHelper
    {
        public string test()
        {
            return "1";
        }

        public string ConvertHttpPngToJpg(string imageUrl, string pngfilePath, string jpgFilePath)
        {
            string pngFilePath = DownloadImage(imageUrl, pngfilePath);

            // 将PNG转换为白底JPG
            return ConvertPngToJpg(pngFilePath, jpgFilePath);
        }
        /// <summary>
        /// 下载图片
        /// </summary>
        /// <param name="imageUrl">图片URL</param>
        /// <returns>本地图片路径</returns>
        public string DownloadImage(string imageUrl, string filePath)
        {
            using (WebClient client = new WebClient())
            {
                client.DownloadFile(imageUrl, filePath);
            }

            return filePath;
        }


        public string UploadFile(string apiUrl, string filePath, string hearderName, string headerValue)
        {
            if (!File.Exists(filePath))
            {
                throw new FileNotFoundException($"文件未找到: {filePath}");
            }

            try
            {
                using (var content = new MultipartFormDataContent())
                using (var fileStream = File.OpenRead(filePath))
                using (var streamContent = new StreamContent(fileStream))
                {
                    streamContent.Headers.ContentType = new MediaTypeHeaderValue("application/octet-stream");
                    //streamContent.Headers.Add(hearderName, headerValue);
                    content.Add(streamContent, "file", Path.GetFileName(filePath));
                   
                    
                    using (var request = new HttpRequestMessage(HttpMethod.Post, apiUrl)
                    {
                        Content = content
                    })
                    using (var _httpClient = new HttpClient())
                    {
                        request.Headers.Add(hearderName, headerValue);
                        using (var response = _httpClient.SendAsync(request).Result)
                        {
                            response.EnsureSuccessStatusCode();

                            // 6. 读取响应内容
                            string responseBody = response.Content.ReadAsStringAsync().Result;

                            return responseBody;
                        }
                    }
                }
            }
            catch (HttpRequestException e)
            {
                return $"请求发生错误: {e.Message}";

            }
            catch (FileNotFoundException e)
            {
                return $"文件系统错误: {e.Message}";
            }
            catch (Exception e)
            {
                return $"发生未知错误: {e.Message}";
            }
        }

        /// <summary>
        /// 将PNG转换为白底JPG
        /// </summary>
        /// <param name="pngFilePath">PNG图片路径</param>
        /// <returns>JPG图片路径</returns>

        public string ConvertPngToJpg(string pngFilePath, string jpgFilePath)
        {
            using (Bitmap pngImage = new Bitmap(pngFilePath))
            {
                // 关键步骤：检查并根据 EXIF 方向元数据修正图片方向
                // EXIF 方向标签的 ID 是 0x0112
                if (pngImage.PropertyIdList.Contains(0x0112))
                {
                    var orientation = (int)pngImage.GetPropertyItem(0x0112).Value[0];
                    switch (orientation)
                    {
                        case 2:
                            pngImage.RotateFlip(RotateFlipType.RotateNoneFlipX); // 水平翻转
                            break;
                        case 3:
                            pngImage.RotateFlip(RotateFlipType.Rotate180FlipNone); // 旋转180度
                            break;
                        case 4:
                            pngImage.RotateFlip(RotateFlipType.Rotate180FlipX); // 垂直翻转
                            break;
                        case 5:
                            pngImage.RotateFlip(RotateFlipType.Rotate90FlipX); // 顺时针旋转90度后水平翻转
                            break;
                        case 6:
                            pngImage.RotateFlip(RotateFlipType.Rotate90FlipNone); // 顺时针旋转90度
                            break;
                        case 7:
                            pngImage.RotateFlip(RotateFlipType.Rotate270FlipX); // 顺时针旋转270度后水平翻转
                            break;
                        case 8:
                            pngImage.RotateFlip(RotateFlipType.Rotate270FlipNone); // 顺时针旋转270度
                            break;
                            // case 1: 默认，无需操作
                    }
                    // 修正方向后，移除方向标签，防止保存到新图片中造成二次旋转
                    pngImage.RemovePropertyItem(0x0112);
                }

                // 注意：旋转90/270度后，图片的 Width 和 Height 会交换，
                // 所以创建新画布的操作必须在旋转之后。
                using (Bitmap jpgImage = new Bitmap(pngImage.Width, pngImage.Height))
                {
                    // 额外优化：同步DPI，防止缩放问题
                    jpgImage.SetResolution(pngImage.HorizontalResolution, pngImage.VerticalResolution);

                    using (Graphics graphic = Graphics.FromImage(jpgImage))
                    {
                        // 设置白色背景
                        graphic.Clear(System.Drawing.Color.White);

                        // 设置高质量绘图模式
                        graphic.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
                        graphic.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.HighQuality;
                        graphic.PixelOffsetMode = System.Drawing.Drawing2D.PixelOffsetMode.HighQuality;
                        graphic.CompositingQuality = System.Drawing.Drawing2D.CompositingQuality.HighQuality;

                        graphic.DrawImage(pngImage, 0, 0, pngImage.Width, pngImage.Height);
                    }

                    // 为了获得更好的JPG质量，可以设置编码器参数
                    ImageCodecInfo jpgEncoder = ImageCodecInfo.GetImageDecoders().First(c => c.FormatID == System.Drawing.Imaging.ImageFormat.Jpeg.Guid);
                    EncoderParameters encoderParams = new EncoderParameters(1);
                    encoderParams.Param[0] = new EncoderParameter(System.Drawing.Imaging.Encoder.Quality, 95L); // 质量设为95

                    jpgImage.Save(jpgFilePath, jpgEncoder, encoderParams);
                }
            }

            return jpgFilePath;
        }


        /// <summary>
        /// 绘制颗粒级配曲线
        /// </summary>
        /// <param name="dataX"></param>
        /// <param name="dataY"></param>
        /// <param name="lableX"></param>
        /// <param name="lableY"></param>
        /// <param name="witdth"></param>
        /// <param name="height"></param>
        /// <param name="filepath"></param>
        public void DrawParticleTthencheunakCurvePlot(string strdataX, string strdataY, string lableX, string lableY, int witdth, int height, string filepath,int lableXsize, int lableYsize)
        {
            float? f_lableXsize;
            float? f_lableYsize;
            if (lableXsize ==0)
            {
                f_lableXsize = null;
            }
            else
            {
                f_lableXsize = lableXsize;
            }

            if (lableYsize == 0)
            {
                f_lableYsize = null;
            }
            else
            {
                f_lableYsize = lableYsize;
            }


            double[] dataX = strdataX.Split(',').Select(s => double.Parse(s, CultureInfo.InvariantCulture)).ToArray();
            double[] dataY = strdataY.Split(',').Select(s => double.Parse(s, CultureInfo.InvariantCulture)).ToArray();

            double[] logXs = dataX.Select(Math.Log10).ToArray();
            ScottPlot.Plot myPlot = new ScottPlot.Plot();
            myPlot.Add.Palette = new ScottPlot.Palettes.Redness();
            var scatter = myPlot.Add.Scatter(logXs, dataY);

            // 反转
            myPlot.Axes.AutoScaler.InvertedX = true;
            scatter.Smooth = true;


            // create a minor tick generator that places log-distributed minor ticks
            ScottPlot.TickGenerators.LogMinorTickGenerator minorTickGen = new ScottPlot.TickGenerators.LogMinorTickGenerator();

            // create a numeric tick generator that uses our custom minor tick generator
            ScottPlot.TickGenerators.NumericAutomatic tickGen = new ScottPlot.TickGenerators.NumericAutomatic();
            tickGen.MinorTickGenerator = minorTickGen;

            // tell our major tick generator to only show major ticks that are whole integers
            tickGen.IntegerTicksOnly = true;

            // tell our custom tick generator to use our new label formatter
            tickGen.LabelFormatter = (double x) => $"{Math.Pow(10, x):0.#######}";

            // tell the left axis to use our custom tick generator
            myPlot.Axes.Bottom.TickGenerator = tickGen;


            // show grid lines for minor ticks
            myPlot.Grid.MajorLineColor = ScottPlot.Colors.Black.WithOpacity(.15);
            myPlot.Grid.MinorLineColor = ScottPlot.Colors.Black.WithOpacity(.05);
            myPlot.Grid.MinorLineWidth = 1;

            //添加x轴，y轴名称
            myPlot.XLabel(lableX, f_lableXsize);
            myPlot.YLabel(lableY, f_lableYsize);
            myPlot.Font.Automatic();
            myPlot.SavePng(filepath, witdth, height);
        }

        /// <summary>
        /// 绘制单条击实曲线
        /// </summary>
        /// <param name="waterContents"></param>
        /// <param name="dryDensities"></param>
        /// <param name="lableX"></param>
        /// <param name="lableY"></param>
        /// <param name="witdth"></param>
        /// <param name="height"></param>
        /// <param name="filepath"></param>
        public void DrawSingleCompactionCurvePlot(string strwaterContents, string strdryDensities, string lableX, string lableY, int witdth, int height, string filepath, int lableXsize, int lableYsize)
        {
            float? f_lableXsize;
            float? f_lableYsize;
            if (lableXsize == 0)
            {
                f_lableXsize = null;
            }
            else
            {
                f_lableXsize = lableXsize;
            }

            if (lableYsize == 0)
            {
                f_lableYsize = null;
            }
            else
            {
                f_lableYsize = lableYsize;
            }

            double[] waterContents = strwaterContents.Split(',').Select(s => double.Parse(s, CultureInfo.InvariantCulture)).ToArray();
            double[] dryDensities = strdryDensities.Split(',').Select(s => double.Parse(s, CultureInfo.InvariantCulture)).ToArray();
            // 多项式函数计算
            double[] coefficients = MathNet.Numerics.Fit.Polynomial(waterContents, dryDensities, 4);
            double e_ = coefficients[0];
            double d_ = coefficients[1];
            double c_ = coefficients[2];
            double b_ = coefficients[3];
            double a_ = coefficients[4];
            Func<double, double?> fittedCurveFunc = (x) => a_ * Math.Pow(x, 4) + b_ * Math.Pow(x, 3) + c_ * x * x + d_ * x + e_;

            // ScottPlot 绘图
            ScottPlot.Plot myPlotFit = new ScottPlot.Plot();


            // 绘制原始数据点
            var markers = myPlotFit.Add.Scatter(waterContents, dryDensities);
            markers.LineWidth = 0;
            markers.MarkerSize = 5;
            markers.Color = ScottPlot.Colors.Red;

            // 绘制拟合的抛物线
            myPlotFit.Add.Palette = new ScottPlot.Palettes.Redness();
            double parabola(double x) => a_ * Math.Pow(x, 4) + b_ * Math.Pow(x, 3) + c_ * x * x + d_ * x + e_;

            double minX = waterContents.Min();
            double maxX = waterContents.Max();
            var funcPlot = myPlotFit.Add.Function(parabola);
            funcPlot.MinX = minX;
            funcPlot.MaxX = maxX;

            string peak = this.GetCompactionCurvePeakNew(strwaterContents, strdryDensities);

            double peakX = Convert.ToDouble(peak.Split(',')[0]);
            double peakY = Convert.ToDouble(peak.Split(',')[1]);


            // 标记计算出的精确峰值
            //myPlotFit.Add.Marker(
            //    x: optimumWaterContent,
            //    y: maxDryDensity,
            //    size: 15
            //);

            // 设置图表样式... (同方法一)
            //myPlotFit.Title("击实曲线 (二次拟合)");
            myPlotFit.XLabel(lableX, f_lableXsize);
            myPlotFit.YLabel(lableY, f_lableYsize);
            myPlotFit.Legend.IsVisible = true;
            //自适应图表
            AxesLimit axesLimit = this.GetAxesLimits(minX, maxX, dryDensities.Min(), peakY);
            myPlotFit.Axes.SetLimits(axesLimit.MinX, axesLimit.MaxX, axesLimit.MinY, axesLimit.MaxY);

            //自适应字体
            myPlotFit.Font.Automatic();

            // 保存图像
            myPlotFit.SavePng(filepath, witdth, height);
        }


        /// <summary>
        /// 绘制双条击实曲线
        /// </summary>
        /// <param name="waterContents1"></param>
        /// <param name="dryDensities1"></param>
        /// <param name="waterContents2"></param>
        /// <param name="dryDensities2"></param>
        /// <param name="lableX"></param>
        /// <param name="lableY"></param>
        /// <param name="witdth"></param>
        /// <param name="height"></param>
        /// <param name="filepath"></param>
        public void DrawTwoCompactionCurvePlot(string strwaterContents1, string strdryDensities1, string strwaterContents2, string strdryDensities2, string lableX, string lableY, int witdth, int height, string filepath, int lableXsize, int lableYsize)
        {

            float? f_lableXsize;
            float? f_lableYsize;
            if (lableXsize == 0)
            {
                f_lableXsize = null;
            }
            else
            {
                f_lableXsize = lableXsize;
            }

            if (lableYsize == 0)
            {
                f_lableYsize = null;
            }
            else
            {
                f_lableYsize = lableYsize;
            }


            double[] waterContents1 = strwaterContents1.Split(',').Select(s => double.Parse(s, CultureInfo.InvariantCulture)).ToArray();
            double[] dryDensities1 = strdryDensities1.Split(',').Select(s => double.Parse(s, CultureInfo.InvariantCulture)).ToArray();
            double[] waterContents2 = strwaterContents2.Split(',').Select(s => double.Parse(s, CultureInfo.InvariantCulture)).ToArray();
            double[] dryDensities2 = strdryDensities2.Split(',').Select(s => double.Parse(s, CultureInfo.InvariantCulture)).ToArray();
            // ScottPlot 绘图
            ScottPlot.Plot myPlotFit = new ScottPlot.Plot();
            //myPlotFit.Add.Palette = new ScottPlot.Palettes.Redness();


            // 绘制原始数据点
            var markers1 = myPlotFit.Add.Scatter(waterContents1, dryDensities1);
            markers1.LineWidth = 0;
            markers1.MarkerSize = 5;
            markers1.Color = ScottPlot.Colors.Red;

            var markers2 = myPlotFit.Add.Scatter(waterContents2, dryDensities2);
            markers2.LineWidth = 0;
            markers2.MarkerSize = 5;
            markers2.Color = ScottPlot.Colors.Blue;



            // 拟合曲线
            double[] coefficients1 = MathNet.Numerics.Fit.Polynomial(waterContents1, dryDensities1, 4);
            double e1 = coefficients1[0];
            double d1 = coefficients1[1];
            double c1 = coefficients1[2];
            double b1 = coefficients1[3];
            double a1 = coefficients1[4];
            double parabola1(double x) => a1 * Math.Pow(x, 4) + b1 * Math.Pow(x, 3) + c1 * x * x + d1 * x + e1;
            var funcPlot1 = myPlotFit.Add.Function(parabola1);
            funcPlot1.MinX = waterContents1.Min();
            funcPlot1.MaxX = waterContents1.Max();
            funcPlot1.LineColor = ScottPlot.Colors.Red;
            funcPlot1.LegendText = "试验一";

            double[] coefficients2 = MathNet.Numerics.Fit.Polynomial(waterContents2, dryDensities2, 4);
            double e2 = coefficients2[0];
            double d2 = coefficients2[1];
            double c2 = coefficients2[2];
            double b2 = coefficients2[3];
            double a2 = coefficients2[4];
            double parabola2(double x) => a2 * Math.Pow(x, 4) + b2 * Math.Pow(x, 3) + c2 * x * x + d2 * x + e2;
            var funcPlot2 = myPlotFit.Add.Function(parabola2);
            funcPlot2.MinX = waterContents2.Min();
            funcPlot2.MaxX = waterContents2.Max();
            funcPlot2.LineColor = ScottPlot.Colors.Blue;
            funcPlot2.LegendText = "试验二";


            string peak1 = this.GetCompactionCurvePeakNew(strwaterContents1, strdryDensities1);
            double peakX1 = Convert.ToDouble(peak1.Split(',')[0]);
            double peakY1 = Convert.ToDouble(peak1.Split(',')[1]);

            string peak2 = this.GetCompactionCurvePeakNew(strwaterContents2, strdryDensities2);
            double peakX2 = Convert.ToDouble(peak2.Split(',')[0]);
            double peakY2 = Convert.ToDouble(peak2.Split(',')[1]);

            double minX = Math.Min(waterContents1.Min(), waterContents2.Min());
            double maxX = Math.Max(waterContents1.Max(), waterContents2.Max());
            double minY = Math.Min(dryDensities1.Min(), dryDensities2.Min());
            double maxY = Math.Max(peakY1, peakY2);

            myPlotFit.XLabel(lableX, f_lableXsize);
            myPlotFit.YLabel(lableY, f_lableYsize);
            myPlotFit.Legend.IsVisible = true;
            //自适应图表
            AxesLimit axesLimit = this.GetAxesLimits(minX, maxX, minY, maxY);
            myPlotFit.Axes.SetLimits(axesLimit.MinX, axesLimit.MaxX, axesLimit.MinY, axesLimit.MaxY);

            //自适应字体
            myPlotFit.Font.Automatic();

            myPlotFit.ShowLegend(ScottPlot.Alignment.LowerCenter, ScottPlot.Orientation.Horizontal);
            myPlotFit.ShowLegend(ScottPlot.Edge.Bottom);

            // 保存图像
            myPlotFit.SavePng(filepath, witdth, height);
        }

        /// <summary>
        /// 绘制散点图
        /// </summary>
        /// <param name="dataX"></param>
        /// <param name="dataY"></param>
        /// <param name="lableX"></param>
        /// <param name="lableY"></param>
        /// <param name="witdth"></param>
        /// <param name="height"></param>
        /// <param name="filepath"></param>
        /// <param name="smoothTension"></param>
        public void CommonScatterPlot(string strdataX, string strdataY, string lableX, string lableY, int witdth, int height, string filepath, double smoothTension, int lableXsize, int lableYsize)
        {
            float? f_lableXsize;
            float? f_lableYsize;
            if (lableXsize == 0)
            {
                f_lableXsize = null;
            }
            else
            {
                f_lableXsize = lableXsize;
            }

            if (lableYsize == 0)
            {
                f_lableYsize = null;
            }
            else
            {
                f_lableYsize = lableYsize;
            }

            double[] dataX = strdataX.Split(',').Select(s => double.Parse(s, CultureInfo.InvariantCulture)).ToArray();
            double[] dataY = strdataY.Split(',').Select(s => double.Parse(s, CultureInfo.InvariantCulture)).ToArray();

            ScottPlot.Plot myPlot = new ScottPlot.Plot();
            myPlot.Add.Palette = new ScottPlot.Palettes.Redness();
            var scatter = myPlot.Add.Scatter(dataX, dataY);
            //修改曲线曲率
            scatter.SmoothTension = smoothTension;

            myPlot.XLabel(lableX, f_lableXsize);
            myPlot.YLabel(lableY, f_lableYsize);

            //自适应字体
            myPlot.Font.Automatic();
            myPlot.SavePng(filepath, witdth, height);
        }

        /// <summary>
        /// 搜索法
        /// </summary>
        /// <param name="waterContents"></param>
        /// <param name="dryDensities"></param>
        /// <returns></returns>
        public string GetCompactionCurvePeak(string strwaterContents, string strdryDensities)
        {

            double[] waterContents = strwaterContents.Split(',').Select(s => double.Parse(s, CultureInfo.InvariantCulture)).ToArray();
            double[] dryDensities = strdryDensities.Split(',').Select(s => double.Parse(s, CultureInfo.InvariantCulture)).ToArray();
            // 多项式函数计算
            double[] coefficients = MathNet.Numerics.Fit.Polynomial(waterContents, dryDensities, 4);
            double e_ = coefficients[0];
            double d_ = coefficients[1];
            double c_ = coefficients[2];
            double b_ = coefficients[3];
            double a_ = coefficients[4];
            Func<double, double?> fittedCurveFunc = (x) => a_ * Math.Pow(x, 4) + b_ * Math.Pow(x, 3) + c_ * x * x + d_ * x + e_;

            double minX = waterContents.Min();
            double maxX = waterContents.Max();
            double peakX = minX;
            double peakY = fittedCurveFunc(minX) ?? double.MinValue;
            int steps = 1000; // 搜索精度，可以增加
            double stepSize = (maxX - minX) / steps;

            stepSize = (double)GetLastDigitUnitFromString(minX);
            steps = (int)((maxX - minX) / stepSize);

            for (int i = 1; i <= steps; i++)
            {
                double currentX = minX + i * stepSize;
                double currentY = fittedCurveFunc(currentX) ?? double.MinValue;
                if (currentY > peakY)
                {
                    peakY = currentY;
                    peakX = currentX;
                }
            }

            Console.WriteLine("通过四次多项式拟合 + 数值搜索找到的峰值:");
            Console.WriteLine($"最佳含水率 (OWC): {peakX:F5} %");
            Console.WriteLine($"最大干密度 (MDD): {peakY:F5} g/cm³");

            return $"{peakX:F5},{peakY:F5}";
        }


        /// <summary>
        /// 求导法
        /// </summary>
        /// <param name="waterContents"></param>
        /// <param name="dryDensities"></param>
        /// <returns></returns>
        public string GetCompactionCurvePeakNew(string strwaterContents, string strdryDensities)
        {
            double[] waterContents = strwaterContents.Split(',').Select(s => double.Parse(s, CultureInfo.InvariantCulture)).ToArray();
            double[] dryDensities = strdryDensities.Split(',').Select(s => double.Parse(s, CultureInfo.InvariantCulture)).ToArray();
            // 多项式函数计算
            double[] coefficients = MathNet.Numerics.Fit.Polynomial(waterContents, dryDensities, 4);
            double e_ = coefficients[0];
            double d_ = coefficients[1];
            double c_ = coefficients[2];
            double b_ = coefficients[3];
            double a_ = coefficients[4];

            var p = new Polynomial(e_, d_, c_, b_, a_);

            var p_prime = p.Differentiate();

            Complex[] roots = p_prime.Roots();


            double peakX = 0;
            double peakY = 0;
            double minX = waterContents.Min();
            double maxX = waterContents.Max();

            foreach (var root in roots)
            {
                // 我们通常只关心实数解，所以检查虚部是否接近于0
                // IsAlmostZero() 是一个扩展方法，可以处理浮点数的精度问题
                //if (MathNet.Numerics.Precision.AlmostEqual(root.Real, 0.0))
                //{
                double x = root.Real;
                if (x >= minX && x <= maxX)
                {
                    double y = p.Evaluate(x);

                    if (y > peakY)
                    {
                        peakY = y;
                        peakX = x;
                    }

                    //Console.WriteLine($"找到一个驻点: (x, y) = ({x:F4}, {y:F4})");
                }
                //}
            }

            Console.WriteLine("通过四次多项式拟合 + 数值搜索找到的峰值:");
            Console.WriteLine($"最佳含水率 (OWC): {peakX:F5} %");
            Console.WriteLine($"最大干密度 (MDD): {peakY:F5} g/cm³");

            return $"{peakX:F5},{peakY:F5}";
        }

        public void TestScottPlot()
        {
            double[] dataX = { 1, 2, 3, 4, 5 };
            double[] dataY = { 1, 4, 9, 16, 25 };

            ScottPlot.Plot myPlot = new ScottPlot.Plot();
            myPlot.Add.Scatter(dataX, dataY);

            myPlot.SavePng("quickstart.png", 400, 300);
        }
        /// <summary>
        /// 颗粒级配曲线
        /// </summary>
        public void TestScottPlot_42()
        {
            double[] dataX = { 60, 40, 20, 10, 5, 2, 1, 0.5, 0.25, 0.075 };
            double[] dataY = { 100.0, 100.0, 94.1, 91.3, 88.8, 82.1, 73.2, 55.3, 31.4, 9.4 };

            double[] logXs = dataX.Select(Math.Log10).ToArray();


            ScottPlot.Plot myPlot = new ScottPlot.Plot();
            myPlot.Add.Palette = new ScottPlot.Palettes.Redness();
            var scatter = myPlot.Add.Scatter(logXs, dataY);

            // 反转
            myPlot.Axes.AutoScaler.InvertedX = true;
            scatter.Smooth = true;

            //myPlot.Axes.Margins(bottom: 0);


            // create a minor tick generator that places log-distributed minor ticks
            ScottPlot.TickGenerators.LogMinorTickGenerator minorTickGen = new ScottPlot.TickGenerators.LogMinorTickGenerator();

            // create a numeric tick generator that uses our custom minor tick generator
            ScottPlot.TickGenerators.NumericAutomatic tickGen = new ScottPlot.TickGenerators.NumericAutomatic();
            tickGen.MinorTickGenerator = minorTickGen;

            // tell our major tick generator to only show major ticks that are whole integers
            tickGen.IntegerTicksOnly = true;

            // tell our custom tick generator to use our new label formatter
            tickGen.LabelFormatter = (double x) => $"{Math.Pow(10, x):0.#######}";

            // tell the left axis to use our custom tick generator
            myPlot.Axes.Bottom.TickGenerator = tickGen;


            // show grid lines for minor ticks
            myPlot.Grid.MajorLineColor = ScottPlot.Colors.Black.WithOpacity(.15);
            myPlot.Grid.MinorLineColor = ScottPlot.Colors.Black.WithOpacity(.05);
            myPlot.Grid.MinorLineWidth = 1;


            //ScottPlot.TickGenerators.NumericManual yticks = new ScottPlot.TickGenerators.NumericManual();
            //yticks.AddMajor(0, "0.0");
            //yticks.AddMajor(10, "10.0");
            //yticks.AddMajor(20, "20.0");
            //yticks.AddMajor(30, "30.0");
            //yticks.AddMajor(40, "40.0");
            //yticks.AddMajor(50, "50.0");
            //yticks.AddMajor(60, "60.0");
            //yticks.AddMajor(70, "70.0");
            //yticks.AddMajor(80, "80.0");
            //yticks.AddMajor(90, "90.0");
            //yticks.AddMajor(100, "100.0");
            //myPlot.Axes.Left.TickGenerator = yticks;


            //添加x轴，y轴名称
            myPlot.XLabel("孔径(mm)");
            myPlot.YLabel("小于该孔径土占总土质量百分比(%)");
            myPlot.Font.Automatic();

            myPlot.SavePng("quickstart.png", 400, 300);
        }


        /// <summary>
        /// 干密度-含水率曲线
        /// </summary>
        public void TestScottPlot_43()
        {
            double[] dataX = { 9.2, 11.4, 13.2, 15.3, 17.2 };
            double[] dataY = { 1.72, 1.84, 1.89, 1.85, 1.76 };

            ScottPlot.Plot myPlot = new ScottPlot.Plot();
            myPlot.Add.Palette = new ScottPlot.Palettes.Redness();
            var scatter = myPlot.Add.Scatter(dataX, dataY);
            //修改曲线曲率
            scatter.SmoothTension = 1;

            //添加x轴，y轴名称
            myPlot.XLabel("含水率(%)");
            myPlot.Axes.Bottom.Label.FontName = SKFontManager.Default.MatchCharacter('汉').FamilyName;
            //myPlot.Axes.Bottom.Label.FontSize = 24;
            //myPlot.Axes.Bottom.Label.Bold = false;

            myPlot.YLabel("干密度(g/cm3)");
            myPlot.Axes.Left.Label.FontName = SKFontManager.Default.MatchCharacter('汉').FamilyName;
            myPlot.Axes.Left.Label.Alignment = ScottPlot.Alignment.LowerLeft;
            //myPlot.Axes.Left.Label.FontSize = 24;
            //myPlot.Axes.Left.Label.Bold = false;

            myPlot.SavePng("TestScottPlot_43.png", 400, 300);
        }


        public void TestScottPlot_43_math()
        {
            double[] waterContents = { 9.2, 11.4, 13.2, 15.3, 17.2 };
            double[] dryDensities = { 1.72, 1.84, 1.89, 1.85, 1.76 };

            int maxIndex = Array.IndexOf(dryDensities, dryDensities.Max());

            if (maxIndex > 0 && maxIndex < waterContents.Length - 1)
            {
                // 选取最高点及其前后各一个点
                double[] fitX = new[] { waterContents[maxIndex - 1], waterContents[maxIndex], waterContents[maxIndex + 1] };
                double[] fitY = new[] { dryDensities[maxIndex - 1], dryDensities[maxIndex], dryDensities[maxIndex + 1] };

                // 3. 使用 MathNet.Numerics 进行二次多项式拟合
                // 返回的数组是 [c, b, a] 对应 y = c + bx + ax^2
                double[] coefficients = MathNet.Numerics.Fit.Polynomial(fitX, fitY, 2);
                double c = coefficients[0];
                double b = coefficients[1];
                double a = coefficients[2];

                // 4. 计算抛物线顶点
                double optimumWaterContent = -b / (2 * a);
                double maxDryDensity = (a * optimumWaterContent * optimumWaterContent) + (b * optimumWaterContent) + c;

                Console.WriteLine("\n通过拟合计算的峰值点:");
                Console.WriteLine($"最佳含水率 (OWC): {optimumWaterContent:F2} %");
                Console.WriteLine($"最大干密度 (MDD): {maxDryDensity:F2} g/cm³");

                // --- 接下来是 ScottPlot 绘图 ---
                ScottPlot.Plot myPlotFit = new ScottPlot.Plot();

                // 绘制原始数据点
                //myPlotFit.Add.Scatter(waterContents, dryDensities);

                // 绘制拟合的抛物线，让可视化效果更好
                //Func<double, double?> parabola = (x) => a * x * x + b * x + c;

                double parabola(double x) => a * x * x + b * x + c;

                double minX = waterContents.Min();
                double maxX = waterContents.Max();

                var funcPlot = myPlotFit.Add.Function(parabola);
                funcPlot.MinX = minX;
                funcPlot.MaxX = maxX;

                myPlotFit.Axes.SetLimits(9, 17, 1.55, 2.06);
                var markers = myPlotFit.Add.Scatter(waterContents, dryDensities);
                markers.LineWidth = 0; // 不画连线，只画点
                markers.MarkerSize = 2;
                //markers.MarkerShape = ScottPlot.MarkerShape.OpenCircle;
                markers.Color = ScottPlot.Colors.Red;

                // 标记计算出的精确峰值
                //myPlotFit.Add.Marker(
                //    x: optimumWaterContent,
                //    y: maxDryDensity,
                //    size: 15
                //);

                // 设置图表样式... (同方法一)
                myPlotFit.Title("击实曲线 (二次拟合)");
                myPlotFit.XLabel("含水率 (%)");
                myPlotFit.YLabel("干密度 (g/cm³)");
                myPlotFit.Legend.IsVisible = true;


                myPlotFit.Font.Automatic();

                // 保存图像
                myPlotFit.SavePng("TestScottPlot_43_math.png", 400, 300);
            }


            //ScottPlot.Plot myPlot = new ScottPlot.Plot();
            //myPlot.Add.Palette = new ScottPlot.Palettes.Redness();
            //var scatter = myPlot.Add.Scatter(dataX, dataY);
            ////修改曲线曲率
            //scatter.SmoothTension = 1;

            ////添加x轴，y轴名称
            //myPlot.XLabel("含水率(%)");
            //myPlot.Axes.Bottom.Label.FontName = SKFontManager.Default.MatchCharacter('汉').FamilyName;
            ////myPlot.Axes.Bottom.Label.FontSize = 24;
            ////myPlot.Axes.Bottom.Label.Bold = false;

            //myPlot.YLabel("干密度(g/cm3)");
            //myPlot.Axes.Left.Label.FontName = SKFontManager.Default.MatchCharacter('汉').FamilyName;
            //myPlot.Axes.Left.Label.Alignment = ScottPlot.Alignment.LowerLeft;
            ////myPlot.Axes.Left.Label.FontSize = 24;
            ////myPlot.Axes.Left.Label.Bold = false;

            //myPlot.SavePng("TestScottPlot_43.png", 400, 300);
        }


        public void TestScottPlot_43_math2(double[] waterContents, double[] dryDensities)
        {
            //double[] waterContents = { 9.2, 11.4, 13.2, 15.3, 17.2 };
            //double[] dryDensities = { 1.72, 1.84, 1.89, 1.85, 1.76 };




            double[] coefficients = MathNet.Numerics.Fit.Polynomial(waterContents, dryDensities, 4);
            double e_ = coefficients[0];
            double d_ = coefficients[1];
            double c_ = coefficients[2];
            double b_ = coefficients[3];
            double a_ = coefficients[4];

            // 创建拟合函数
            Func<double, double?> fittedCurveFunc = (x) => a_ * Math.Pow(x, 4) + b_ * Math.Pow(x, 3) + c_ * x * x + d_ * x + e_;

            // --- 接下来是 ScottPlot 绘图 ---
            ScottPlot.Plot myPlotFit = new ScottPlot.Plot();


            // 绘制原始数据点
            var markers = myPlotFit.Add.Scatter(waterContents, dryDensities);
            markers.LineWidth = 0;
            markers.MarkerSize = 5;
            markers.Color = ScottPlot.Colors.Red;

            // 绘制拟合的抛物线
            myPlotFit.Add.Palette = new ScottPlot.Palettes.Redness();
            double parabola(double x) => a_ * Math.Pow(x, 4) + b_ * Math.Pow(x, 3) + c_ * x * x + d_ * x + e_;

            double minX = waterContents.Min();
            double maxX = waterContents.Max();
            var funcPlot = myPlotFit.Add.Function(parabola);
            funcPlot.MinX = minX;
            funcPlot.MaxX = maxX;


            //myPlotFit.Axes.AutoScale();


            double peakX = minX;
            double peakY = fittedCurveFunc(minX) ?? double.MinValue;
            int steps = 1000; // 搜索精度，可以增加
            double stepSize = (maxX - minX) / steps;

            for (int i = 1; i <= steps; i++)
            {
                double currentX = minX + i * stepSize;
                double currentY = fittedCurveFunc(currentX) ?? double.MinValue;
                if (currentY > peakY)
                {
                    peakY = currentY;
                    peakX = currentX;
                }

            }

            Console.WriteLine("通过四次多项式拟合 + 数值搜索找到的峰值:");
            Console.WriteLine($"最佳含水率 (OWC): {peakX:F5} %");
            Console.WriteLine($"最大干密度 (MDD): {peakY:F5} g/cm³");

            // 标记计算出的精确峰值
            //myPlotFit.Add.Marker(
            //    x: optimumWaterContent,
            //    y: maxDryDensity,
            //    size: 15
            //);

            // 设置图表样式... (同方法一)
            //myPlotFit.Title("击实曲线 (二次拟合)");
            myPlotFit.XLabel("含水率 (%)");
            myPlotFit.YLabel("干密度 (g/cm³)");
            myPlotFit.Legend.IsVisible = true;

            AxesLimit axesLimit = this.GetAxesLimits(minX, maxX, dryDensities.Min(), peakY);
            myPlotFit.Axes.SetLimits(axesLimit.MinX, axesLimit.MaxX, axesLimit.MinY, axesLimit.MaxY);


            myPlotFit.Font.Automatic();



            // 保存图像
            myPlotFit.SavePng("TestScottPlot_43_math2.png", 800, 600);

        }
        /// <summary>
        /// 干密度-含水量(p d-w)关系曲线
        /// </summary>
        public void TestScottPlot_44()
        {
            //试验一
            double[] dataX1 = { 5.04, 5.74, 6.65, 7.44, 8.38 };
            double[] dataY1 = { 2.1376, 2.2035, 2.2564, 2.2355, 2.1621 };

            //试验二
            double[] dataX2 = { 5.14, 5.76, 6.67, 7.41, 8.44 };
            double[] dataY2 = { 2.1141, 2.1803, 2.2534, 2.2082, 2.1387 };

            //string fontName = SKFontManager.Default.MatchCharacter('汉').FamilyName;


            ScottPlot.Plot myPlot = new ScottPlot.Plot();
            myPlot.Add.Palette = new ScottPlot.Palettes.Redness();
            var scatter1 = myPlot.Add.Scatter(dataX1, dataY1);
            scatter1.LegendText = "试验一";
            scatter1.Color = ScottPlot.Colors.Red;
            var scatter2 = myPlot.Add.Scatter(dataX2, dataY2);
            scatter2.LegendText = "试验二";
            scatter2.Color = ScottPlot.Colors.Blue;

            myPlot.ShowLegend(ScottPlot.Alignment.LowerCenter, ScottPlot.Orientation.Horizontal);
            myPlot.ShowLegend(ScottPlot.Edge.Bottom);


            //添加x轴，y轴名称
            myPlot.XLabel("含水率(%)");
            // myPlot.Axes.Bottom.Label.FontName = fontName;
            //myPlot.Axes.Bottom.Label.FontSize = 24;
            //myPlot.Axes.Bottom.Label.Bold = false;

            myPlot.YLabel("干密度(g/cm3)");
            //myPlot.Axes.Left.Label.FontName = fontName;
            //myPlot.Axes.Left.Label.Alignment = ScottPlot.Alignment.LowerLeft;
            //myPlot.Axes.Left.Label.FontSize = 24;
            //myPlot.Axes.Left.Label.Bold = false;

            myPlot.Font.Automatic();

            myPlot.SavePng("TestScottPlot_44.png", 400, 300);
        }

        public void TestScottPlot_45()
        {
            double[] dataX = { 30.1, 27.6, 20.9 };
            double[] dataY = { 7.4, 4.6, 1.0 };

            ScottPlot.Plot myPlot = new ScottPlot.Plot();
            myPlot.Add.Palette = new ScottPlot.Palettes.Redness();
            var scatter = myPlot.Add.Scatter(dataX, dataY);
            //修改曲线曲率
            scatter.SmoothTension = 1;

            //添加x轴，y轴名称
            myPlot.XLabel("含水率(%)");
            myPlot.Axes.Bottom.Label.FontName = SKFontManager.Default.MatchCharacter('汉').FamilyName;
            //myPlot.Axes.Bottom.Label.FontSize = 24;
            //myPlot.Axes.Bottom.Label.Bold = false;

            myPlot.YLabel("锥入深度mm");
            myPlot.Axes.Left.Label.FontName = SKFontManager.Default.MatchCharacter('汉').FamilyName;
            myPlot.Axes.Left.Label.Alignment = ScottPlot.Alignment.UpperLeft;
            //myPlot.Axes.Left.Label.FontSize = 24;
            //myPlot.Axes.Left.Label.Bold = false;

            myPlot.SavePng("TestScottPlot_45.png", 400, 300);
        }


        public void TestScottPlot_46_1()
        {
            //试验一
            double[] dataX = { 50.6, 42.8, 31.2 };
            double[] dataY = { 19.95, 12.25, 4.70 };

            ScottPlot.Plot myPlot = new ScottPlot.Plot();
            myPlot.Add.Palette = new ScottPlot.Palettes.Redness();
            var scatter = myPlot.Add.Scatter(dataX, dataY);
            //修改曲线曲率
            scatter.SmoothTension = 1;

            //添加x轴，y轴名称
            myPlot.XLabel("含水率(%)");
            myPlot.Axes.Bottom.Label.FontName = SKFontManager.Default.MatchCharacter('汉').FamilyName;
            //myPlot.Axes.Bottom.Label.FontSize = 24;
            //myPlot.Axes.Bottom.Label.Bold = false;

            myPlot.YLabel("锥入深度mm");
            myPlot.Axes.Left.Label.FontName = SKFontManager.Default.MatchCharacter('汉').FamilyName;
            myPlot.Axes.Left.Label.Alignment = ScottPlot.Alignment.UpperLeft;
            //myPlot.Axes.Left.Label.FontSize = 24;
            //myPlot.Axes.Left.Label.Bold = false;

            myPlot.SavePng("TestScottPlot_46_1.png", 400, 300);
        }


        public void TestScottPlot_46_2()
        {

            double[] dataX = { 50.3, 42.4, 31.1 };
            double[] dataY = { 20.00, 12.50, 4.65 };

            ScottPlot.Plot myPlot = new ScottPlot.Plot();
            myPlot.Add.Palette = new ScottPlot.Palettes.Redness();
            var scatter = myPlot.Add.Scatter(dataX, dataY);
            //修改曲线曲率
            scatter.SmoothTension = 1;

            //添加x轴，y轴名称
            myPlot.XLabel("含水率(%)");
            myPlot.Axes.Bottom.Label.FontName = SKFontManager.Default.MatchCharacter('汉').FamilyName;
            //myPlot.Axes.Bottom.Label.FontSize = 24;
            //myPlot.Axes.Bottom.Label.Bold = false;



            string verticalYLabel = string.Join("\n", "锥入深度mm".ToCharArray());
            //myPlot.Axes.Left.Label.LineSpacing = 0.5f;
            myPlot.YLabel(verticalYLabel);

            myPlot.Axes.Left.Label.FontName = SKFontManager.Default.MatchCharacter('汉').FamilyName;
            myPlot.Axes.Left.Label.Rotation = 0;
            myPlot.Axes.Left.MinimumSize = 60;
            myPlot.Axes.Left.MaximumSize = 60;
            myPlot.Axes.Left.Label.Alignment = ScottPlot.Alignment.MiddleRight;
            myPlot.Axes.Left.Label.OffsetX = 20;
            myPlot.Axes.Left.Label.OffsetY = -40;

            //myPlot.Axes.Left.Label.FontSize = 24;
            //myPlot.Axes.Left.Label.Bold = false;

            myPlot.SavePng("TestScottPlot_46_2.png", 400, 300);
        }



        public AxesLimit GetAxesLimits(double minX, double maxX, double minY, double maxY)
        {
            double yRange = maxY - minY;
            double yMinWithPadding = minY - yRange * 0.1;
            double yMaxWithPadding = maxY + yRange * 0.1;

            return new AxesLimit(minX, maxX, yMinWithPadding, yMaxWithPadding);

        }


        public static decimal GetLastDigitUnitFromString(double number)
        {
            // 使用 InvariantCulture 确保小数点是 '.'
            string s = number.ToString(CultureInfo.InvariantCulture);

            // 查找小数点的位置
            int decimalPointIndex = s.IndexOf('.');

            // 如果没有小数点，说明是整数，单位是 1
            if (decimalPointIndex == -1)
            {
                return 1m;
            }

            // 小数部分的长度就是小数位数
            int decimalPlaces = s.Length - decimalPointIndex - 1;

            // 计算 1 / (10 的 decimalPlaces 次方)
            // 使用循环可以避免 double 转换带来的精度问题
            decimal result = 1m;
            for (int i = 0; i < decimalPlaces; i++)
            {
                result /= 10m;
            }
            return result;
        }
    }


    public class AxesLimit
    {
        public AxesLimit(double minX, double maxX, double minY, double maxY)
        {
            MinX = minX;
            MaxX = maxX;
            MinY = minY;
            MaxY = maxY;
        }
        public double MinX { get; set; }
        public double MaxX { get; set; }
        public double MinY { get; set; }
        public double MaxY { get; set; }
    }
}
