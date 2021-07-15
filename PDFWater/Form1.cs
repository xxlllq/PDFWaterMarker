using Spire.Pdf;
using Spire.Pdf.Graphics;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows.Forms;

namespace PDFWaterMarker
{
    public partial class Form1 : Form
    {

        public Form1()
        {
            InitializeComponent();
        }

        /// <summary>
        /// 添加水印
        /// </summary>
        /// <param name="path">文件夹路径</param>
        /// <param name="wmt">水印文本</param>
        public void ApplyWatermark(string path,string wmt, ref string resultStr)
        {
            var list = new List<FileMoidel>();
            var res = resultStr;

            Director(path, list);
            if (list != null && list.Any())
            {
                list.ForEach(m =>
                {
                    try
                    {
                        //创建PdfDocument对象            
                        PdfDocument pdf = new PdfDocument();
                        //加载现有PDF文档
                        pdf.LoadFromFile(m.FullPath);
                        PdfPageBase pb = pdf.Pages.Add(); //新增一页
                        pdf.Pages.Remove(pb); //去除第一页水印
                        //创建True Type字体            
                        PdfTrueTypeFont font = new PdfTrueTypeFont(new Font("宋体", 22f), true);
                        //水印文字            
                        string text = !string.IsNullOrEmpty(wmt) ? wmt : "仅供个人学习\n请勿用于任何商业用途";
                        //测量文字所占的位置大小，即高宽            
                        SizeF fontSize = font.MeasureString(text);
                        //计算两个偏移量            
                        float offset1 = (float)(fontSize.Width * Math.Sqrt(2) / 4);
                        float offset2 = (float)(fontSize.Height * Math.Sqrt(2) / 4);
                        //遍历文档每一页            
                        foreach (PdfPageBase page in pdf.Pages)
                        {
                            if (page != null && page?.Canvas != null)
                            {
                                //创建PdfTilingBrush对象                
                                PdfTilingBrush brush = new PdfTilingBrush(new SizeF(page.Canvas.Size.Width / 2, page.Canvas.Size.Height / 2));
                                //设置画刷透明度                
                                brush.Graphics.SetTransparency(0.8f);
                                //将画刷中坐标系向右下平移                
                                brush.Graphics.TranslateTransform(brush.Size.Width / 2 - offset1 - offset2, brush.Size.Height / 2 + offset1 - offset2);
                                //将坐标系逆时针旋转45度                
                                brush.Graphics.RotateTransform(-45);
                                //在画刷上绘制文本                
                                brush.Graphics.DrawString(text, font, PdfBrushes.LightGray, 0, 0);
                                //在PDF页面绘制跟页面一样大小的矩形，并使用定义的画刷填充                
                                page.Canvas.DrawRectangle(brush, new RectangleF(new PointF(0, 0), page.Canvas.Size));
                            }
                        }

                        //保存文档   
                        var filePathName = m.FullPath.Replace(path, path + "-副本");
                        pdf.SaveToFile(filePathName);
                        res = "转换成功！";
                    }
                    catch (Exception ex)
                    {
                        res = ex.Message;
                    }
                });
            }
            resultStr = res;
        }

        /// <summary>
        /// 转PDF
        /// </summary>
        /// <param name="path">文件夹路径</param>
        /// <param name="wmt">水印文本</param>
        public void ToPDF(string path,  ref string resultStr)
        {
            var list = new List<FileMoidel>();
            var res = resultStr;

            Director(path, list);
            if (list != null && list.Any())
            {
                list.ForEach(m =>
                {
                    try
                    {
                        var filePathName = m.FullPath.Replace(path, path + "-副本");
                        Aspose.Words.Document doc = new Aspose.Words.Document(m.FullPath);
                        if (!string.IsNullOrEmpty(filePathName)) {
                            filePathName = filePathName.Replace(".docx", ".pdf");
                            filePathName = filePathName.Replace(".doc", ".pdf");
                        }
                        doc.Save(filePathName);
                        res = "转换成功！";
                    }
                    catch (Exception ex)
                    {
                        res = ex.Message;
                    }
                });
            }
            resultStr = res;
        }

        private void toImage_Click(object sender, EventArgs e)
        {
//            var client = new RestClient("https://api.aliyundrive.com/adrive/v2/share_link/get_share_by_anonymous");
//            client.Timeout = -1;
//            var request = new RestRequest(Method.POST);
//            request.AddHeader("Content-Type", "application/json");
//            request.AddHeader("Cookie", "acw_tc=2f6a1f9c16263269778963337e47548f9a053d4d6e3c4fb6b155654f17df7c");
//            var body = @"{
//" + "\n" +
//            @"    ""share_id"":""moB5Z35cViv""
//" + "\n" +
//            @"}";
//            request.AddParameter("application/json", body, ParameterType.RequestBody);
//            IRestResponse response = client.Execute(request);
//            Console.WriteLine(response.Content);
            //Apply();
        }
        public void Apply() {
            var inputFile = @"C:\Users\xxx\source\repos\xxlllq\system_architect\真题(截至2020年)【推荐】\历年真题及解析合并版本（截至2020年）\2009年\2009年系统架构师考试科目三：论文真题.pdf";
            //创建PdfDocument对象            
            PdfDocument doc = new PdfDocument();
            //加载现有PDF文档
            doc.LoadFromFile(inputFile);
            //PdfPageBase pb = doc.Pages.Add(); //新增一页
            //doc.Pages.Remove(pb); //去除第一页水印
            for (int i = 0; i < doc.Pages.Count; i++)
            {
                Image img = doc.SaveAsImage(i);
                string fn1 = String.Format("{0}.png", i + 1);
                var outputFile = inputFile.Substring(0, inputFile.LastIndexOf(".pdf"));
                if (Directory.Exists(outputFile) == false)
                {
                    Directory.CreateDirectory(outputFile);
                }
                string of = Path.Combine(outputFile+"\\", fn1);
                img.Save(of, System.Drawing.Imaging.ImageFormat.Png);
            }
            var sd = 1;
        }

        /// <summary>
        /// 递归获取文件
        /// </summary>
        /// <param name="dir"></param>
        /// <param name="list"></param>
        public void Director(string dir, List<FileMoidel> list,bool chooseDoc = false)
        {
            DirectoryInfo d = new DirectoryInfo(dir);
            FileInfo[] files = d.GetFiles();//文件
            DirectoryInfo[] directs = d.GetDirectories();//文件夹
            foreach (FileInfo f in files)
            {
                if (f?.Extension == ".pdf" || (chooseDoc = true && (f?.Extension == ".docx" || f?.Extension == ".doc")))
                {
                    list.Add(new FileMoidel { FullPath = f.FullName, FolderName = f.Directory.Name, FileName = f.Name });//添加文件名到列表中  
                }
            }
            //获取子文件夹内的文件列表，递归遍历  
            foreach (DirectoryInfo dd in directs)
            {
                Director(dd.FullName, list);
            }
        }

        /// <summary>
        /// 选择文件
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void openFile_Click(object sender, System.EventArgs e)
        {
            FolderBrowserDialog dialog = new FolderBrowserDialog();
            dialog.Description = "请选择文件路径";
            if (dialog.ShowDialog() == DialogResult.OK)
            {
                var path = dialog.SelectedPath;
                var wmt = watermarkerText.Text;
                filePath.Text = path;
                result.Text = "开始转换…";

                Func<string> wait = () =>
                {
                    var resultStr = string.Empty;
                    if (!string.IsNullOrEmpty(path))
                    {
                        ApplyWatermark(path, wmt,ref resultStr);
                    }
                    return resultStr;
                };
                wait.BeginInvoke(new AsyncCallback(res =>
                {
                    string rst = wait.EndInvoke(res);
                    this.Invoke((Action)(() => result.Text = rst));
                }), null);

            }
        }

        /// <summary>
        /// 转PDF
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void toPDF_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog dialog = new FolderBrowserDialog();
            dialog.Description = "请选择文件路径";
            if (dialog.ShowDialog() == DialogResult.OK)
            {
                var path = dialog.SelectedPath;
                textBox1.Text = path;
                result.Text = "开始转换…";

                Func<string> wait = () =>
                {
                    var resultStr = string.Empty;
                    if (!string.IsNullOrEmpty(path))
                    {
                        ToPDF(path, ref resultStr);
                    }
                    return resultStr;
                };
                wait.BeginInvoke(new AsyncCallback(res =>
                {
                    string rst = wait.EndInvoke(res);
                    this.Invoke((Action)(() => result.Text = rst));
                }), null);

            }
        }
    }


    /// <summary>
    /// 文件Model类
    /// </summary>
    public class FileMoidel
    {
        /// <summary>
        /// 全称（全路径）
        /// </summary>
        public string FullPath { set; get; }

        /// <summary>
        /// 父级文件夹
        /// </summary>
        public string FolderName { set; get; }

        /// <summary>
        /// 文件名
        /// </summary>
        public string FileName { set; get; }
    }
}