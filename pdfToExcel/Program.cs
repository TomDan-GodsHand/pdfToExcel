using System.Drawing;
using System.Drawing.Imaging;
using System.Text.RegularExpressions;
using Spire.Pdf;
using Spire.Pdf.Utilities;
using Spire.Xls;

List<string> filePathList = new List<string>();
if (args.Length > 0)
{
    Console.WriteLine("拖拽的文件路径:");
    foreach (string file in args)
    {
        Console.WriteLine(file);
        filePathList.Add(file);
    }
}
else
{
    filePathList.Add("C:\\Users\\TomDan\\Desktop\\2025011104马锦媛提交的云南西南林业大学后勤服务有限公司物业维修单202501111302000178572.pdf");
}
System.Console.WriteLine("按回车继续");
Console.ReadLine();
foreach (string filePath in filePathList)
{
    try
    {

        string dictionPath = Path.GetDirectoryName(filePath);
        string fileName = Path.GetFileNameWithoutExtension(filePath);

        PdfDocument doc = new PdfDocument();
        doc.LoadFromFile(filePath);
        doc.ConvertOptions.SetPdfToXpsOptions(true, true, true);

        Dictionary<string, string> keyValues = new Dictionary<string, string>();
        Stream QRCodeImage = null;
        Stream SignImage = null;
        List<Stream> images = new List<Stream>();
        PdfImageHelper pdfImageHelper = new PdfImageHelper();
        //循环遍历文档中的所有页面
        foreach (PdfPageBase page in doc.Pages)
        {
            PdfImageInfo[] imagesInfo = pdfImageHelper.GetImagesInfo(page);
            //从每个页面提取图像并将其保存到指定的文件路径
            foreach (var imageInfo in imagesInfo)
            {

                using (Image image = Image.FromStream(imageInfo.Image))
                {
                    if (image.Width == image.Height)
                    {
                        QRCodeImage = imageInfo.Image;
                    }
                    else if (image.Width == 100 && image.Height == 99)
                    {
                        continue;
                    }
                    else if (HasTransparency(image))
                    {
                        SignImage = imageInfo.Image;
                    }
                    else
                    {
                        images.Add(imageInfo.Image);
                    }
                }
            }
        }
        // 初始化 PdfTableExtractor 类的实例
        PdfTableExtractor extractor = new PdfTableExtractor(doc);

        // 声明 PdfTable 数组
        PdfTable[]? tableList = null;
        // 循环遍历页面
        for (int pageIndex = 0; pageIndex < doc.Pages.Count; pageIndex++)
        {
            // 从特定页面提取表格
            tableList = extractor.ExtractTable(pageIndex);

            // 判断表格列表是否为空
            if (tableList != null && tableList.Length > 0)
            {

                // 遍历列表中的表格
                foreach (PdfTable table in tableList)
                {
                    // 获取特定表格的行数和列数
                    int row = table.GetRowCount();
                    int column = table.GetColumnCount();
                    // 添加工作表
                    // 遍历行和列
                    for (int i = 0; i < row; i++)
                    {
                        string key = table.GetText(i, 0).Replace("\n", "").Replace("黄照灵", "");
                        key = key.Trim();
                        string Value = table.GetText(i, 1).Replace("\n", "").Replace("黄照灵", "");
                        Value = Value.Trim();
                        if (key.Length > 0)
                            keyValues.Add(key, Value);
                    }
                }
            }
        }

        #region excel操作

        Workbook workbook = new Workbook();
        workbook.Worksheets.Clear();
        Worksheet sheet = workbook.Worksheets.Add(fileName);

        //设置列宽
        sheet.SetColumnWidth(1, 12.38);
        sheet.SetColumnWidth(2, 13.13);
        sheet.SetColumnWidth(3, 11.25);
        sheet.SetColumnWidth(4, 12.63);
        sheet.SetColumnWidth(5, 13.75);
        sheet.SetColumnWidth(6, 20.63);

        //设置页边距
        sheet.PageSetup.TopMargin = 0.1;
        sheet.PageSetup.LeftHeader = "";
        //标题
        sheet.Range["A1:F1"].Merge();
        sheet.SetRowHeight(1, 49);
        sheet.Range["A1"].Text = "云南西南林业大学后勤服务有限公司";
        sheet.Range["A1"].Style.Font.IsBold = true; // 字体加粗
        sheet.Range["A1"].Style.Font.Size = 18; // 字体大小
        sheet.Range["A1"].Style.Font.FontName = "SimSun";
        sheet.Range["A1"].Style.HorizontalAlignment = HorizontalAlignType.Center; // 水平居中
        sheet.Range["A1"].Style.VerticalAlignment = VerticalAlignType.Center; // 垂直居中

        //副标题
        sheet.Range["A2:F2"].Merge();
        sheet.SetRowHeight(2, 49);
        sheet.Range["A2"].Text = "物业部维修单";
        sheet.Range["A2"].Style.Font.IsBold = true; // 字体加粗
        sheet.Range["A2"].Style.Font.Size = 18; // 字体大小
        sheet.Range["A2"].Style.Font.FontName = "SimSun";
        sheet.Range["A2"].Style.HorizontalAlignment = HorizontalAlignType.Center; // 水平居中
        sheet.Range["A2"].Style.VerticalAlignment = VerticalAlignType.Center; // 垂直居中
        sheet.Pictures.Add(2, 6, QRCodeImage, 170, 170);

        //编号行,第三行
        sheet.Range["A3:F3"].Merge();
        sheet.SetRowHeight(3, 49);

        string serialNumber = keyValues["维修单编号"]; // 编号
        var time = keyValues["维修开始时间"];
        // 使用固定格式解析
        DateTime dateTime = DateTime.ParseExact(time, "yyyy-MM-dd HH:mm", null);

        string year = dateTime.Year.ToString();          // 年
        string month = dateTime.Month.ToString();           // 月
        string day = dateTime.Day.ToString();             // 日
        int padding = 45 - serialNumber.Length; // 动态计算填充空格
        string content = $"编 号 ：{serialNumber.PadRight(padding)}{year} 年 {month} 月 {day} 日";
        sheet.Range["A3"].Text = content;
        sheet.Range["A3"].Style.Font.Size = 14; // 字体大小
        sheet.Range["A3"].Style.Font.FontName = "FangSong";
        sheet.Range["A3"].Style.HorizontalAlignment = HorizontalAlignType.Left; // 水平居中
        sheet.Range["A3"].Style.VerticalAlignment = VerticalAlignType.Center; // 垂直居中

        // 第四行
        sheet.SetRowHeight(4, 49);
        sheet.Range["A4:F4"].BorderAround(LineStyleType.Thin, Color.Black);
        sheet.Range["A4:F4"].BorderInside(LineStyleType.Thin, Color.Black);
        sheet.Range["A4:F4"].BorderInside(LineStyleType.Thin, Color.Black);
        sheet.Range["A4:F4"].Style.Font.Size = 12; // 字体大小
        sheet.Range["A4:F4"].Style.Font.FontName = "SimSun";
        sheet.Range["A4"].Text = "报修人";
        sheet.Range["A4:F4"].Style.HorizontalAlignment = HorizontalAlignType.Center;
        sheet.Range["A4:F4"].Style.VerticalAlignment = VerticalAlignType.Center; // 垂直居中
        sheet.Range["B4"].Text = keyValues["报修人"];
        sheet.Range["C4"].Text = "工号/学号";
        sheet.Range["D4"].Text = keyValues["学号/工号"];
        sheet.Range["E4"].Text = "电话";
        sheet.Range["F4"].Text = keyValues["报修人电话"];

        // 第5行
        sheet.SetRowHeight(5, 49);
        sheet.Range["A5:F5"].BorderAround(LineStyleType.Thin, Color.Black); sheet.Range["A5:F5"].BorderInside(LineStyleType.Thin, Color.Black);
        sheet.Range["A5:F5"].Style.Font.Size = 12; // 字体大小
        sheet.Range["A5:F5"].Style.Font.FontName = "SimSun";
        sheet.Range["A5"].Text = "维修时间";
        DateTime startTime = DateTime.Parse(keyValues["维修开始时间"]);
        DateTime endTime = DateTime.Parse(keyValues["维修结束时间"]);
        string timeContent = $"{startTime.ToString("HH:mm")} - {endTime.ToString("HH:mm")}";
        sheet.Range["B5:D5"].Merge();
        sheet.Range["B5"].Text = timeContent;
        sheet.Range["A4:F5"].Style.HorizontalAlignment = HorizontalAlignType.Center;
        sheet.Range["A5:F5"].Style.VerticalAlignment = VerticalAlignType.Center; // 垂直居中
        sheet.Range["E5"].Text = "维修人员";
        sheet.Range["F5"].Text = keyValues["维修工人"];

        // 第6行
        sheet.SetRowHeight(6, 49);
        sheet.Range["A6:F6"].BorderAround(LineStyleType.Thin, Color.Black); sheet.Range["A6:F6"].BorderInside(LineStyleType.Thin, Color.Black);
        sheet.Range["A6:F6"].Style.Font.Size = 12; // 字体大小
        sheet.Range["A6:F6"].Style.Font.FontName = "SimSun";
        sheet.Range["A6:F6"].Style.WrapText = true;
        sheet.Range["C6:F6"].Merge();
        sheet.Range["A6:B6"].Merge();
        sheet.Range["A6"].Text = "报修地址";
        sheet.Range["A6:F6"].Style.HorizontalAlignment = HorizontalAlignType.Center;
        sheet.Range["A6:F6"].Style.VerticalAlignment = VerticalAlignType.Center; // 垂直居中
        sheet.Range["C6"].Text = keyValues["详细地址"];

        // 第7行
        sheet.SetRowHeight(7, 49);
        sheet.Range["A7:F7"].BorderAround(LineStyleType.Thin, Color.Black); sheet.Range["A7:F7"].BorderInside(LineStyleType.Thin, Color.Black);
        sheet.Range["A7:F7"].Style.Font.Size = 12; // 字体大小
        sheet.Range["A7:F7"].Style.Font.FontName = "SimSun";
        sheet.Range["C7:F7"].Merge();
        sheet.Range["A7:B7"].Merge();
        sheet.Range["A7"].Text = "报修内容";
        sheet.Range["A7:F7"].Style.HorizontalAlignment = HorizontalAlignType.Center;
        sheet.Range["A7:F7"].Style.VerticalAlignment = VerticalAlignType.Center; // 垂直居中
        sheet.Range["C7"].Text = keyValues["报修内容"];
        // 第8行
        sheet.SetRowHeight(8, 49);
        sheet.Range["A8:F8"].BorderAround(LineStyleType.Thin, Color.Black); sheet.Range["A8:F8"].BorderInside(LineStyleType.Thin, Color.Black);
        sheet.Range["A8:F8"].Style.Font.Size = 12; // 字体大小
        sheet.Range["A8:F8"].Style.Font.FontName = "SimSun";
        sheet.Range["C8:F8"].Merge();
        sheet.Range["A8:B8"].Merge();
        sheet.Range["A8"].Text = "现场情况";
        sheet.Range["A8:F8"].Style.HorizontalAlignment = HorizontalAlignType.Center;
        sheet.Range["A8:F8"].Style.VerticalAlignment = VerticalAlignType.Center; // 垂直居中
        sheet.Range["C8"].Text = keyValues["现场情况"];
        // 第9行
        sheet.SetRowHeight(9, 129);
        sheet.Range["A9:F9"].BorderAround(LineStyleType.Thin, Color.Black); sheet.Range["A9:F9"].BorderInside(LineStyleType.Thin, Color.Black);
        sheet.Range["A9:F9"].Style.Font.Size = 12; // 字体大小
        sheet.Range["A9:F9"].Style.Font.FontName = "SimSun";
        sheet.Range["A9:F9"].Style.WrapText = true;
        sheet.Range["C9:F9"].Merge();
        sheet.Range["A9:B9"].Merge();
        sheet.Range["A9"].Text = "维修内容";
        sheet.Range["A9"].Style.HorizontalAlignment = HorizontalAlignType.Center;
        sheet.Range["B9:F9"].Style.HorizontalAlignment = HorizontalAlignType.Left;
        sheet.Range["A9"].Style.VerticalAlignment = VerticalAlignType.Center; // 垂直居中
        sheet.Range["B9:F9"].Style.VerticalAlignment = VerticalAlignType.Top; // 垂直居中
        sheet.Range["C9"].Text = keyValues["维修内容"];
        // 第10行
        sheet.SetRowHeight(10, 45);
        sheet.Range["A10:F10"].BorderAround(LineStyleType.Thin, Color.Black);
        sheet.Range["A10:F10"].BorderInside(LineStyleType.Thin, Color.Black);
        sheet.Range["A10:F10"].Style.Font.Size = 12; // 字体大小
        sheet.Range["A10:F10"].Style.Font.FontName = "SimSun";
        sheet.Range["C10:F10"].Merge();
        sheet.Range["A10:B10"].Merge();
        sheet.Range["A10"].Text = "是否产生材料";
        sheet.Range["A10:F10"].Style.HorizontalAlignment = HorizontalAlignType.Center;
        sheet.Range["A10:F10"].Style.VerticalAlignment = VerticalAlignType.Center; // 垂直居中
        var text = keyValues["维修材料"];
        string output = string.Empty;
        if (ContainsArabicDigits(text))
        {
            output = "☑是        □否";
        }
        else
            output = "□是        ☑否";
        sheet.Range["C10"].Text = output;
        // 第11行
        sheet.SetRowHeight(11, 45);
        sheet.Range["A11:F11"].BorderAround(LineStyleType.Thin, Color.Black);
        sheet.Range["A11:F11"].BorderInside(LineStyleType.Thin, Color.Black);
        sheet.Range["A11:F11"].Style.Font.Size = 12; // 字体大小
        sheet.Range["A11:F11"].Style.Font.FontName = "SimSun";
        sheet.Range["C11:F11"].Merge();
        sheet.Range["A11:B11"].Merge();
        sheet.Range["A11"].Text = "维修结果";
        sheet.Range["A11:F11"].Style.HorizontalAlignment = HorizontalAlignType.Center;
        sheet.Range["A11:F11"].Style.VerticalAlignment = VerticalAlignType.Center; // 垂直居中
        sheet.Range["C11"].Text = " ☑ 同 意 验 收           □ 不 同 意 验 收";
        // 第12行
        sheet.SetRowHeight(12, 66);
        sheet.Range["A12:F12"].BorderAround(LineStyleType.Thin, Color.Black); sheet.Range["A12:F12"].BorderInside(LineStyleType.Thin, Color.Black);
        sheet.Range["A12:F12"].Style.Font.Size = 12; // 字体大小
        sheet.Range["A12:F12"].Style.Font.FontName = "SimSun";
        sheet.Range["C12:D12"].Merge();
        var p = sheet.Pictures.Add(12, 3, SignImage, 10, 10);
        p.Left = 250;
        sheet.Range["A12:B12"].Merge();
        sheet.Range["A12"].Text = "验收人签字";
        sheet.Range["E12"].Text = "工号/学号";
        sheet.Range["A12:F12"].Style.HorizontalAlignment = HorizontalAlignType.Center;
        sheet.Range["A12:F12"].Style.VerticalAlignment = VerticalAlignType.Center; // 垂直居中
        sheet.Range["F12"].Text = keyValues["验收人工号/学号"];
        // 第13行
        sheet.SetRowHeight(13, 86.25);
        sheet.Range["A13:F13"].BorderAround(LineStyleType.Thin, Color.Black); sheet.Range["A13:F13"].BorderInside(LineStyleType.Thin, Color.Black);
        sheet.Range["A13:F13"].Style.Font.Size = 12; // 字体大小
        sheet.Range["A13:F13"].Style.Font.FontName = "SimSun";
        sheet.Range["C13:D13"].Merge();
        sheet.Range["A13:B13"].Merge();
        sheet.Range["A13"].Text = "验收人部门/学院";
        sheet.Range["E13"].Text = "验收人电话";
        sheet.Range["A13:F13"].Style.HorizontalAlignment = HorizontalAlignType.Center;
        sheet.Range["A13:F13"].Style.VerticalAlignment = VerticalAlignType.Center; // 垂直居中
        sheet.Range["C13"].Text = keyValues["验收人部门/学院"];
        sheet.Range["F13"].Text = keyValues["验收人电话"];

        int index = 16;
        foreach (var image in images)
        {
            var pic = sheet.Pictures.Add(index, 1, image);
            pic.Width = pic.Width / 5;
            pic.Height = pic.Height / 5;
            var rowCount = pic.Height / 12.75;
            int count = (int)rowCount;
            index += count;
        }

        #endregion
        workbook.SaveToFile(dictionPath + "\\" + fileName + ".xlsx", ExcelVersion.Version2013);
        doc.Close();
        workbook.Dispose();
    }
    catch (Exception ex)
    {
        System.Console.WriteLine(ex.Message);
        Console.ReadLine();
    }
}
Console.WriteLine("完成，按回车退出程序！");
Console.ReadLine();




static bool ContainsArabicDigits(string input)
{
    if (string.IsNullOrEmpty(input))
        return false;

    return Regex.IsMatch(input, @"\d", RegexOptions.None);
}

static bool HasTransparency(Image image)
{
    Bitmap bitmap = new Bitmap(image);

    // 检查图片格式，只有支持透明度的格式才需要检测
    if (bitmap.PixelFormat != PixelFormat.Format32bppArgb)
    {
        return false; // 没有 Alpha 通道，不透明
    }

    // 遍历图片的每个像素，检测透明度
    for (int y = 0; y < bitmap.Height; y++)
    {
        for (int x = 0; x < bitmap.Width; x++)
        {
            Color pixel = bitmap.GetPixel(x, y);
            if (pixel.A < 255) // 透明度小于 255，说明有透明像素
            {
                return true;
            }
        }
    }
    return false;
}

