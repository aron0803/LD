using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Runtime.InteropServices;
using System.Diagnostics;
using System.Threading;
using System.Text;

public class DebugMatch
{
    [DllImport("kernel32.dll")]
    private static extern int GetPrivateProfileString(string section, string key, string def, StringBuilder retVal, int size, string filePath);

    public static void Main(string[] args)
    {
        Console.WriteLine("=== DebugMatch Tool ===");
        
        string[] testImages = { "diary.png", "ball.png" };
        
        foreach (string imgPath in testImages)
        {
            if (!File.Exists(imgPath))
            {
                Console.WriteLine(string.Format("Warning: {0} not found, skipping...", imgPath));
                continue;
            }

            Console.WriteLine(string.Format("\n========== Testing {0} ==========", imgPath));
            
            // 1. Synthetic Test
            Console.WriteLine("\n--- Test 1: Synthetic Image ---");
            RunSyntheticTest(imgPath);

            // 2. Real Screenshot Test
            Console.WriteLine("\n--- Test 2: Real Screenshot ---");
            RunRealTest(imgPath);
        }
    }

    static void RunSyntheticTest(string templatePath)
    {
        try
        {
            using (Bitmap template = new Bitmap(templatePath))
            {
                int w = 800;
                int h = 600;
                using (Bitmap canvas = new Bitmap(w, h, PixelFormat.Format24bppRgb))
                using (Graphics g = Graphics.FromImage(canvas))
                {
                    g.Clear(Color.Black);
                    // Draw template at 100, 100
                    g.DrawImage(template, 100, 100, template.Width, template.Height);
                    
                    Console.WriteLine("Created synthetic image 800x600 with template at 100,100");
                    
                    MatchResult result = FindImage(canvas, template, 0.90);
                    Console.WriteLine(string.Format("Result: Success={0}, Score={1:F4}, Loc={2}", result.Success, result.Score, result.Location));
                    
                    if (result.Success && result.Location.X == 100 && result.Location.Y == 100)
                        Console.WriteLine("PASS: Synthetic match accurate.");
                    else
                        Console.WriteLine("FAIL: Synthetic match failed or inaccurate.");
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine("Synthetic Test Error: " + ex.Message);
        }
    }

    static void RunRealTest(string templatePath)
    {
        string configPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "config.ini");
        string ldPath = ReadIni(configPath, "Settings", "ld_path", @"D:\LDPlayer\LDPlayer9\ld.exe");
        string screenshotDir = ReadIni(configPath, "Settings", "screenshot_dir", @"D:\Screenshots");
        int index = int.Parse(ReadIni(configPath, "Settings", "ld_index", "0"));

        Console.WriteLine(string.Format("Config: LD={0}, Dir={1}, Index={2}", ldPath, screenshotDir, index));

        if (!File.Exists(ldPath))
        {
            Console.WriteLine("LDPlayer executable not found.");
            return;
        }

        Bitmap cap = Screencap(ldPath, screenshotDir, index);
        if (cap == null)
        {
            Console.WriteLine("Failed to capture screenshot.");
            return;
        }

        try
        {
            cap.Save("debug_real_cap.png", ImageFormat.Png);
            Console.WriteLine("Saved debug_real_cap.png");

            using (Bitmap template = new Bitmap(templatePath))
            {
                MatchResult result = FindImage(cap, template, 0.90);
                Console.WriteLine(string.Format("Result: Success={0}, Score={1:F4}, Loc={2}", result.Success, result.Score, result.Location));
            }
        }
        finally
        {
            cap.Dispose();
        }
    }

    static string ReadIni(string path, string section, string key, string def)
    {
        StringBuilder sb = new StringBuilder(255);
        GetPrivateProfileString(section, key, def, sb, 255, path);
        return sb.ToString();
    }

    static Bitmap Screencap(string ldPath, string screenshotDir, int index)
    {
        string filename = string.Format("cap_{0}.png", index);
        string localPath = Path.Combine(screenshotDir, filename);
        string remotePath = "/sdcard/Pictures/" + filename;

        ProcessStartInfo psi = new ProcessStartInfo();
        psi.FileName = ldPath;
        psi.Arguments = string.Format("-s {0} screencap {1}", index, remotePath);
        psi.UseShellExecute = false;
        psi.CreateNoWindow = true;
        using (Process p = Process.Start(psi))
        {
            p.WaitForExit();
        }

        // Wait for file
        for (int i = 0; i < 20; i++)
        {
            if (File.Exists(localPath))
            {
                try
                {
                    using (var fs = new FileStream(localPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                    {
                        return new Bitmap(fs);
                    }
                }
                catch { }
            }
            Thread.Sleep(100);
        }
        return null;
    }

    public struct MatchResult
    {
        public bool Success;
        public Point Location;
        public double Score;
    }

    public static MatchResult FindImage(Bitmap source, Bitmap template, double threshold)
    {
        using (Bitmap source24 = ConvertTo24bpp(source))
        using (Bitmap template24 = ConvertTo24bpp(template))
        {
            return TemplateMatch(source24, template24, threshold);
        }
    }

    private static Bitmap ConvertTo24bpp(Bitmap img)
    {
        Bitmap newImg = new Bitmap(img.Width, img.Height, PixelFormat.Format24bppRgb);
        using (Graphics g = Graphics.FromImage(newImg))
        {
            g.DrawImage(img, new Rectangle(0, 0, img.Width, img.Height));
        }
        return newImg;
    }

    private static MatchResult TemplateMatch(Bitmap source, Bitmap template, double threshold)
    {
        int w = source.Width;
        int h = source.Height;
        int tw = template.Width;
        int th = template.Height;

        if (w < tw || h < th) return new MatchResult { Success = false };

        BitmapData srcData = source.LockBits(new Rectangle(0, 0, w, h), ImageLockMode.ReadOnly, PixelFormat.Format24bppRgb);
        BitmapData tmplData = template.LockBits(new Rectangle(0, 0, tw, th), ImageLockMode.ReadOnly, PixelFormat.Format24bppRgb);

        try
        {
            int srcStride = srcData.Stride;
            int tmplStride = tmplData.Stride;
            
            unsafe
            {
                byte* srcPtr = (byte*)srcData.Scan0;
                byte* tmplPtr = (byte*)tmplData.Scan0;

                double maxVal = -1;
                Point maxLoc = Point.Empty;

                // Debug: Print center pixel of template
                byte* tCenter = tmplPtr + (th/2)*tmplStride + (tw/2)*3;
                Console.WriteLine(string.Format("Template Center Pixel: B={0} G={1} R={2}", tCenter[0], tCenter[1], tCenter[2]));

                for (int y = 0; y <= h - th; y++)
                {
                    for (int x = 0; x <= w - tw; x++)
                    {
                        double matchVal = CompareRegion(srcPtr, srcStride, tmplPtr, tmplStride, tw, th, x, y);
                        if (matchVal > maxVal)
                        {
                            maxVal = matchVal;
                            maxLoc = new Point(x, y);
                        }
                    }
                }

                return new MatchResult 
                { 
                    Success = maxVal >= threshold,
                    Location = maxLoc,
                    Score = maxVal
                };
            }
        }
        finally
        {
            source.UnlockBits(srcData);
            template.UnlockBits(tmplData);
        }
    }

    private static unsafe double CompareRegion(byte* srcBase, int srcStride, byte* tmplBase, int tmplStride, int w, int h, int startX, int startY)
    {
        // Normalized Cross-Correlation (NCC) - similar to OpenCV TM_CCOEFF_NORMED
        long totalPixels = w * h;
        
        // Calculate means
        long sumT_B = 0, sumT_G = 0, sumT_R = 0;
        long sumS_B = 0, sumS_G = 0, sumS_R = 0;
        
        for (int y = 0; y < h; y++)
        {
            byte* tRow = tmplBase + y * tmplStride;
            byte* sRow = srcBase + (startY + y) * srcStride + startX * 3;
            for (int x = 0; x < w; x++)
            {
                sumT_B += tRow[0]; sumT_G += tRow[1]; sumT_R += tRow[2];
                sumS_B += sRow[0]; sumS_G += sRow[1]; sumS_R += sRow[2];
                tRow += 3;
                sRow += 3;
            }
        }
        
        double meanT_B = (double)sumT_B / totalPixels;
        double meanT_G = (double)sumT_G / totalPixels;
        double meanT_R = (double)sumT_R / totalPixels;
        double meanS_B = (double)sumS_B / totalPixels;
        double meanS_G = (double)sumS_G / totalPixels;
        double meanS_R = (double)sumS_R / totalPixels;
        
        // Calculate correlation
        double numerator_B = 0, numerator_G = 0, numerator_R = 0;
        double denomS_B = 0, denomS_G = 0, denomS_R = 0;
        double denomT_B = 0, denomT_G = 0, denomT_R = 0;
        
        for (int y = 0; y < h; y++)
        {
            byte* sRow = srcBase + (startY + y) * srcStride + startX * 3;
            byte* tRow = tmplBase + y * tmplStride;
            
            for (int x = 0; x < w; x++)
            {
                double diffS_B = sRow[0] - meanS_B;
                double diffS_G = sRow[1] - meanS_G;
                double diffS_R = sRow[2] - meanS_R;
                double diffT_B = tRow[0] - meanT_B;
                double diffT_G = tRow[1] - meanT_G;
                double diffT_R = tRow[2] - meanT_R;
                
                numerator_B += diffS_B * diffT_B;
                numerator_G += diffS_G * diffT_G;
                numerator_R += diffS_R * diffT_R;
                denomS_B += diffS_B * diffS_B;
                denomS_G += diffS_G * diffS_G;
                denomS_R += diffS_R * diffS_R;
                denomT_B += diffT_B * diffT_B;
                denomT_G += diffT_G * diffT_G;
                denomT_R += diffT_R * diffT_R;
                
                sRow += 3;
                tRow += 3;
            }
        }
        
        // Compute normalized correlation for each channel
        double corr_B = 0, corr_G = 0, corr_R = 0;
        double denom_B = Math.Sqrt(denomS_B * denomT_B);
        if (denom_B > 1e-5) corr_B = numerator_B / denom_B;
        double denom_G = Math.Sqrt(denomS_G * denomT_G);
        if (denom_G > 1e-5) corr_G = numerator_G / denom_G;
        double denom_R = Math.Sqrt(denomS_R * denomT_R);
        if (denom_R > 1e-5) corr_R = numerator_R / denom_R;
        
        double avgCorr = (corr_B + corr_G + corr_R) / 3.0;
        return Math.Max(0, Math.Min(1, avgCorr));
    }
}
