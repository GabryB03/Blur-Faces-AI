using MetroSuite;
using System.Windows.Forms;
using OpenCvSharp;
using OpenCvSharp.Dnn;
using System.Diagnostics;
using Microsoft.WindowsAPICodePack.Shell;
using System.Threading;
using System.Runtime;
using System;
using System.Runtime.InteropServices;

public partial class MainForm : MetroForm
{
    private Net faceNet = CvDnn.ReadNetFromCaffe("models\\deploy.prototxt",
        "models\\res10_300x300_ssd_iter_140000_fp16.caffemodel");

    [DllImport("psapi.dll")]
    static extern int EmptyWorkingSet(IntPtr hwProc);

    [DllImport("kernel32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool SetProcessWorkingSetSize(IntPtr process, UIntPtr minimumWorkingSetSize, UIntPtr maximumWorkingSetSize);

    public MainForm()
    {
        InitializeComponent();
        ProcessFramesFolder();
        Process.GetCurrentProcess().PriorityClass = ProcessPriorityClass.RealTime;

        Thread clearRamThread = new Thread(ClearRAM);
        clearRamThread.Priority = ThreadPriority.Highest;
        clearRamThread.Start();
    }

    private void guna2Button2_Click(object sender, System.EventArgs e)
    {
        if (openFileDialog1.ShowDialog().Equals(DialogResult.OK))
        {
            guna2TextBox1.Text = openFileDialog1.FileName;
        }
    }

    private void guna2Button1_Click(object sender, System.EventArgs e)
    {
        if (!System.IO.File.Exists(guna2TextBox1.Text))
        {
            MessageBox.Show("The specified file does not exist.", Text, MessageBoxButtons.OK, MessageBoxIcon.Error);
            return;
        }

        if (!System.Web.MimeMapping.GetMimeMapping(guna2TextBox1.Text).StartsWith("image/") &&
            !System.Web.MimeMapping.GetMimeMapping(guna2TextBox1.Text).StartsWith("video/") &&
            !System.Web.MimeMapping.GetMimeMapping(guna2TextBox1.Text).StartsWith("application/octet-stream"))
        {
            MessageBox.Show(System.Web.MimeMapping.GetMimeMapping(guna2TextBox1.Text));
            MessageBox.Show("The specified file is not a image or a video.", Text, MessageBoxButtons.OK, MessageBoxIcon.Error);
            return;
        }

        string extension = System.IO.Path.GetExtension(guna2TextBox1.Text).ToLower().Substring(1);
        saveFileDialog1.Filter = $"{extension.ToUpper()} file (*.{extension})|*.{extension}";

        if (!saveFileDialog1.ShowDialog().Equals(DialogResult.OK))
        {
            return;
        }

        if (System.IO.File.Exists(saveFileDialog1.FileName))
        {
            System.IO.File.Delete(saveFileDialog1.FileName);
        }

        guna2TextBox1.Enabled = false;
        guna2Button2.Enabled = false;

        if (System.Web.MimeMapping.GetMimeMapping(guna2TextBox1.Text).StartsWith("image/"))
        {
            BlurFaceInImage(guna2TextBox1.Text, saveFileDialog1.FileName);
        }
        else
        {
            BlurFaceInVideo(guna2TextBox1.Text, saveFileDialog1.FileName);
        }

        guna2TextBox1.Enabled = true;
        guna2Button2.Enabled = true;

        MessageBox.Show("Processing finished succesfully. Enjoy!", Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
    }

    public void BlurFaceInImage(string inputImage, string outputImage)
    {
        try
        {
            Mat image = Cv2.ImRead(inputImage);

            int frameHeight = image.Rows;
            int frameWidth = image.Cols;

            Mat blob = CvDnn.BlobFromImage(image, 1.0, new Size(300, 300), new Scalar(104, 117, 123), false, false);
            faceNet.SetInput(blob, "data");

            Mat detection = faceNet.Forward("detection_out");
            Mat detectionMat = new Mat(detection.Size(2), detection.Size(3), MatType.CV_32F, detection.Ptr(0));

            for (int i = 0; i < detectionMat.Rows; i++)
            {
                float confidence = detectionMat.At<float>(i, 2);

                if (confidence > 0.5)
                {
                    int x1 = (int)(detectionMat.At<float>(i, 3) * frameWidth);
                    int y1 = (int)(detectionMat.At<float>(i, 4) * frameHeight);
                    int x2 = (int)(detectionMat.At<float>(i, 5) * frameWidth);
                    int y2 = (int)(detectionMat.At<float>(i, 6) * frameHeight);

                    Mat faceImg = new Mat(image, new OpenCvSharp.Range(y1, y2), new OpenCvSharp.Range(x1, x2));

                    Mat faceBlur = new Mat();
                    Cv2.GaussianBlur(faceImg, faceBlur, new Size(67, 67), 90, 60);
                    image[new OpenCvSharp.Range(y1, y2), new OpenCvSharp.Range(x1, x2)] = faceBlur;
                    image.SaveImage(outputImage);
                    faceBlur.Dispose();
                    faceImg.Dispose();
                }
            }

            detection.Dispose();
            detectionMat.Dispose();
            blob.Dispose();
            image.Dispose();
        }
        catch
        {

        }
    }

    public void BlurFaceInVideo(string inputVideo, string outputVideo)
    {
        string extension = System.IO.Path.GetExtension(outputVideo).ToLower().Substring(1);

        ProcessFramesFolder();
        ShellFile shellFile = ShellFile.FromFilePath(inputVideo);
        string FPS = (shellFile.Properties.System.Video.FrameRate.Value / 1000).ToString();
        RunFFMpeg($"-threads {Environment.ProcessorCount} -i \"{inputVideo}\" -vn -acodec copy \"{System.IO.Path.GetFullPath("frames")}\\audio.aac\"");
        RunFFMpeg($"-threads {Environment.ProcessorCount} -i \"{inputVideo}\" -vf fps={FPS} \"{System.IO.Path.GetFullPath("frames")}\\%16d.png\"");

        foreach (string file in System.IO.Directory.GetFiles("frames"))
        {
            BlurFaceInImage(file, file);
        }

        RunFFMpeg($"-threads {Environment.ProcessorCount} -r {FPS} -i \"{System.IO.Path.GetFullPath("frames")}\\%16d.png\" -b:v {FPS}M \"{outputVideo}\"");
        RunFFMpeg($"-threads {Environment.ProcessorCount} -i \"{outputVideo}\" -i \"{System.IO.Path.GetFullPath("frames")}\\audio.aac\" -c copy -map 0:v:0 -map 1:a:0 \"{System.IO.Path.GetFullPath("models")}\\newOutput.{extension}\"");

        while (!System.IO.File.Exists(outputVideo) || !System.IO.File.Exists($"models\\newOutput.{extension}"))
        {
            Thread.Sleep(1);
        }

        System.IO.File.Delete(outputVideo);
        System.IO.File.Move($"models\\newOutput.{extension}", outputVideo);

        ProcessFramesFolder();
    }

    public void ProcessFramesFolder()
    {
        if (!System.IO.Directory.Exists("frames"))
        {
            System.IO.Directory.CreateDirectory("frames");
        }

        foreach (string file in System.IO.Directory.GetFiles("frames"))
        {
            System.IO.File.Delete(file);
        }
    }

    public void RunFFMpeg(string arguments)
    {
        Process.Start(new ProcessStartInfo
        {
            FileName = "ffmpeg.exe",
            Arguments = arguments,
            UseShellExecute = false,
            RedirectStandardOutput = true,
            CreateNoWindow = true,
            WindowStyle = ProcessWindowStyle.Hidden
        }).WaitForExit();
    }

    public void ClearRAM()
    {
        while (true)
        {
            Thread.Sleep(3000);
            EmptyWorkingSet(Process.GetCurrentProcess().Handle);
            GCSettings.LargeObjectHeapCompactionMode = GCLargeObjectHeapCompactionMode.CompactOnce;
            GC.Collect(GC.MaxGeneration);
            GC.WaitForPendingFinalizers();
            SetProcessWorkingSetSize(Process.GetCurrentProcess().Handle, (UIntPtr)0xFFFFFFFF, (UIntPtr)0xFFFFFFFF);
        }
    }

    private void MainForm_FormClosing(object sender, FormClosingEventArgs e)
    {
        Process.GetCurrentProcess().Kill();
    }

    private void guna2TextBox1_DragEnter(object sender, DragEventArgs e)
    {
        if (e.Data.GetDataPresent(DataFormats.FileDrop))
        {
            e.Effect = DragDropEffects.Copy;
        }
    }

    private void guna2TextBox1_DragDrop(object sender, DragEventArgs e)
    {
        if (e.Data.GetDataPresent(DataFormats.FileDrop))
        {
            guna2TextBox1.Text = ((string[])e.Data.GetData(DataFormats.FileDrop))[0];
        }
    }
}