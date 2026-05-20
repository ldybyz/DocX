using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using LimsDocX;
using System.Drawing;
using System.IO;
using Xceed.Words.NET;
using Xceed.Document.NET;
using LimsFileHelper;

namespace ConsoleApp1
{
	class Program
	{

		static void Main(string[] args)
		{
			//string imageUrl = @"https://demo.cti-soft.net.cn/Resources/UploadFiles/Singnature/202405/e26333dc77a445e985a2c391b1b8b0c6.png";


			ImageHelper imageHelper = new ImageHelper();
            //string pdfFile = @"C:\Users\lindy\Desktop\广州监管平台\lims开发报告记录模版\lims开发报告记录模版\广州新LIMS开发检测报告（1-47）\1.混凝土抗压强度检测（GBT 50081-2019）.pdf";
            //string result = imageHelper.UploadFile("https://ctimalltest.ctimall.com/api/cai/commonFileUpload", pdfFile, "authToken", "1CD917A11739DF6E2D2AC3F2F4D37AF4");
            //Console.WriteLine("上传图片");


            //string jpgFilePath = @"C:\Users\lindy\Desktop\20098唐倩\20098唐倩\20098唐倩_中文.jpg";
            //string pngFilePath = @"C:\Users\lindy\Desktop\20098唐倩\20098唐倩\20098唐倩_中文.png";
            //imageHelper.ConvertPngToJpg(pngFilePath,jpgFilePath);
            //Console.WriteLine("png转jpg");


            //string filedir = @"C:\Users\lindy\Documents\limsdocx\交通基建.docx";

			//var result = imageHelper.UploadFile("https://ctimalltest.ctimall.com/api/cai/commonUpload", filedir, "authToken", "1CD917A11739DF6E2D2AC3F2F4D37AF4");
            //Console.WriteLine(result);
            //generateDoc(filedir);

            //replaceSign();
            //testReplacePicture();
            //replaceFlag();

            //replaceBookmark();


            // insertRow();

            //combineDoc();


            //var doc = DocX.Load(@"C:\Users\lindy\Documents\limsdocx\doc1.docx");

            //doc.Bookmarks["dateCreated"].SetText(DateTime.UtcNow.ToString());
            //doc.Bookmarks["dateModified"].SetText(DateTime.UtcNow.ToString());
            //         doc.Bookmarks["currentDate"].SetText(DateTime.UtcNow.ToString());
            //         doc.SaveAs(@"C:\Users\lindy\Documents\limsdocx\out.docx");
            //imageHelper.TestScottPlot_42();

            //imageHelper.TestScottPlot_43();

            //double[] waterContents = { 9.2, 11.4, 13.2, 15.3, 17.2 };
            //double[] dryDensities = { 1.72, 1.84, 1.89, 1.85, 1.76 };
            //13.4,1.89093  13.44213,1.89075
            //double[] waterContents = { 6.8,7.6,8.6,9.7,10.4};
            //double[] dryDensities = {1.55,1.70,1.79,1.75,1.65 };
            //8.9,1.79374   8.85181,1.79388

            //double[] waterContents = { 5.04, 5.74,6.65,7.74,8.38};
            //double[] dryDensities = { 2.1376, 2.2035, 2.2564, 2.2355, 2.1621 };
            //7.01,2.26172   7.00631,2.26172

            //double[] waterContents = { 5.14, 5.76, 6.67, 7.41, 8.44 };
            //double[] dryDensities = { 2.1141, 2.1803, 2.2534, 2.2082, 2.1387 };
            //6.69,2.25343   6.68861,2.25343

            //double[] waterContents = { 12.3, 14.0, 16.2, 18.2, 20.3 };
            //double[] dryDensities = { 1.67, 1.73, 1.75, 1.70,1.63 };
            //6.69,2.25    15.62353,1.75299

            //double[] waterContents = { 7.4, 9.4, 11.3, 13.4, 15.4 };
            //double[] dryDensities = { 1.67, 1.79, 1.86, 1.81,1.70 };

            string waterContents = "0.0714,0.0714,0.0514,0.0368,0.0238,0.0140,0.0101,0.0072,0.0051,0.0042,0.0015";
            string dryDensities = "56.9,56.9,51.8,47.8,41.2,33.2,27.9,24.0,21.3,18.7,9.3";
			//var result =imageHelper.GetCompactionCurvePeakNew(waterContents, dryDensities);



			imageHelper.DrawParticleTthencheunakCurvePlot(waterContents, dryDensities, "粒径(mm)",
				"小于某径土占总土质量百分比(%)", 500, 180, "DrawParticleTthencheunakCurvePlot.png",10,10);
            //Console.WriteLine(result);
            //imageHelper.TestScottPlot_45();
            //imageHelper.TestScottPlot_46_1();
            //imageHelper.TestScottPlot_46_2();

            //imageHelper.TestScottPlot_42();
            Console.WriteLine("执行成功");
            Console.ReadLine();

        }

		private static void insertRow()
		{
            string filedir = @"C:\Users\lindy\Documents\limsdocx\验收类通用报告模版-测试.docx";
            string filedir2 = @"C:\Users\lindy\Documents\limsdocx\验收类通用报告模版-测试-new.docx";

            npLimsDocX.classLimsDocX LIMSDocX = new npLimsDocX.classLimsDocX();

            var tablexml = @"<?xml version='1.0' ?>
<complexType length='3'>
	<complexType length='9'>
		<string>序号</string>
		<string>检测项目</string>
		<string>le</string>
		<string>le</string>
		<string>单位</string>
		<string>判定依据</string>
		<string>技术要求</string>
		<string>检测结果</string>
		<string>单项评定</string>
	</complexType>
	<complexType length='9'>
		<string>1</string>
		<string>可溶物含量</string>
		<string>le</string>
		<string>le</string>
		<string>g/m²</string>
		<string>GB18242</string>
		<string>≥2100</string>
		<string>2554</string>
		<string>合格</string>
	</complexType>
	<complexType length='9'>
		<string>2</string>
		<string>耐热性</string>
		<string>试验现象</string>
		<string>le</string>
		<string>—</string>
		<string>GB18242</string>
		<string>无流淌、滴落</string>
		<string>无流淌、滴落</string>
		<string>合格</string>
	</complexType>
	<complexType length='9'>
		<string>3</string>
		<string>低温柔性</string>
		<string>le</string>
		<string>le</string>
		<string>—</string>
		<string>GB18242</string>
		<string>-25℃，无裂缝</string>
		<string>无裂缝</string>
		<string>合格</string>
	</complexType>
</complexType>";
			tablexml = getTableXml2();
            using (DocX document = LIMSDocX.Load(filedir))
            {
				//var p = LIMSDocX.GetParagraphByReplaceFlag(document, "[#table]", "left");

                //var table = document.Tables[1];
				var index = 7;
				var tableIndex = 1;

                var table = LIMSDocX.InsertTableRows(document, tablexml, tableIndex, index);

				//            Row row = table.Rows[index-1];
				//            for (int i = 0; i < 5; i++)
				//{

				//                var newRow = table.InsertRow(row,index,true);
				//                newRow.Cells[0].Paragraphs[0].Append("1");
				//                newRow.Cells[1].Paragraphs[0].Append("测试"+(i+1));
				//                newRow.Cells[2].Paragraphs[0].Append("测试" + (i + 1));
				//                newRow.Cells[3].Paragraphs[0].Append("测试" + (i + 1));
				//            }
				//            row.Cells[0].Paragraphs[0].Append("1");
				//            row.Cells[1].Paragraphs[0].Append("2");
				//            row.Cells[2].Paragraphs[0].Append("3");
				//            row.Cells[3].Paragraphs[0].Append("4");
				//table.AutoFit = AutoFit.Contents;
                document.SaveAs(filedir2);
            }
            

        }

		/// <summary>
		/// 创建一个docx文档
		/// </summary>
		private static void generateDoc(string filedir)
		{

			//if (!File.Exists(filedir))
			//{
			//    using (FileStream fs = File.Create(filedir))
			//    {

			//    }
			//}
			npLimsDocX.classLimsDocX c = new npLimsDocX.classLimsDocX();

			//DocX doc = DocX.Load(filedir); //当文档不含有任何字符的时候，该方法报错
			using (DocX doc = DocX.Load(filedir))
			{

				var tablexml = getTableXml();
				tablexml = @"<?xml version='1.0' ?>
<complexType length='3'>
	<complexType length='8'>
		<string>序号</string>
		<string>检测项目</string>
		<string>le</string>
		<string>le</string>
		<string>单位</string>
		<string>技术要求</string>
		<string>检测结果</string>
		<string>单项评定</string>
	</complexType>
	<complexType length='8'>
		<string>1</string>
		<string>外观</string>
		<string>le</string>
		<string>le</string>
		<string>—</string>
		<string>/</string>
		<string>正常</string>
		<string>合格</string>
	</complexType>
	<complexType length='8'>
		<string>2</string>
		<string>面积</string>
		<string>le</string>
		<string>偏差</string>
		<string>m²/卷</string>
		<string>±0.1</string>
		<string>0.09</string>
		<string>合格</string>
	</complexType>
</complexType>";
                //var table = c.GenerateTable(doc, tablexml, "[#TestTable]", true);

                var table = c.GenerateTableWithWidth(doc, tablexml, "[#TestTable]", true, "7,10,10,10,10,24,18,11", "150");


				//c.SetTableColFixedWidth(doc, table, "8,11,11,11,9,20,20,10", "100");
				//c.SetTableCellFont(table, 1, 3, "Times New Roman", 0, false, true);
				//c.SetTableCellFont(table, 1, 3, "Times New Roman", 0, false, true);

				int rowcount = table.RowCount;
				int X = 1;
				int Y = 1;

				while (X++ < rowcount)
				{
					Y = 1;
					while (Y++ < table.ColumnCount)
					{
						c.SetTableCellFont(table, X - 1, Y - 1, "Times New Roman", 0, false, false);
						c.SetTableCellFont(table, X - 1, Y - 2, "Times New Roman", 0, false, false);
						c.SetTableCellStyle(table, X - 1, Y - 1, "BORDER:TOP:BORDERSIZE_FOUR;BORDER:LEFT;BORDER:RIGHT;PARAGRAPHALIGN:CENTER;PARAGRAPHALIGN:VCENTER");
					}
				}

				c.SetTableRowHeight(doc, table, -1, 0.5, false);
				c.SetTableHeader(table, 0);
				//c.SetTableHeader(table, 1);
				//c.SetTableBorderLine(table, "INSIDEV,INSIDEH,LEFT,RIGHT,TOP,BOTTOM");

				table.AutoFit = AutoFit.Fixed;
                //doc.Save();
                doc.SaveAs(@"C:\Users\lindy\Documents\limsdocx\out.docx");

            }
		}

		private static void editDoc(string filedir)
		{
			if (!File.Exists(filedir))
			{
				using (FileStream fs = File.Create(filedir))
				{

				}
			}
			npLimsDocX.classLimsDocX c = new npLimsDocX.classLimsDocX();

			//DocX doc = DocX.Load(filedir); //当文档不含有任何字符的时候，该方法报错
			using (DocX doc = DocX.Create(filedir))
			{

				doc.Save();

			}


			Console.WriteLine("finish!");

		}

		private static void test1()
		{


			string filedir = @"C:\Users\lindy\Documents\limsdocx\111.docx";

			npLimsDocX.classLimsDocX LIMSDocX = new npLimsDocX.classLimsDocX();
			using (DocX document = LIMSDocX.Load(filedir))
			{
				LIMSDocX.InsertParagraphByReplaceFlag(document, "[#TestTable]", "[#TestTable1]\n\n[#TestTable2]\n\n[#TestTable3]", "");
				LIMSDocX.Save(document);
			}

		}

		private static void replaceFlag()
		{


			string filedir = @"C:\Users\lindy\Documents\limsdocx\苏州建材.docx";
            string filedir2 = @"C:\Users\lindy\Documents\limsdocx\苏州建材-out.docx";

            npLimsDocX.classLimsDocX LIMSDocX = new npLimsDocX.classLimsDocX();
			using (DocX document = LIMSDocX.Load(filedir))
			{
                //LIMSDocX.ReplaceFlagByReg(document, @"\[#.*?_1\]", "R001", "left",true);
                LIMSDocX.ReplaceFlagFromDoc(document, @"\[#.*?_1\]", "", "left", true);
                LIMSDocX.ReplaceFlag(document, @"[#FolderInfoTable]", "CO{U|2}", "left");
                LIMSDocX.ReplaceFlag(document, @"[#CONCLUSION]", "456", "left");
                //LIMSDocX.ReplaceFlag(document, "[#REPORTNO]", "R001", "left");
                //            LIMSDocX.ReplaceBookmark(document, "REPORTNO", "R001", "");
                //            LIMSDocX.ReplaceBookmark(document, "REPORTNO1", "R002", "");
                //            LIMSDocX.ReplaceBookmark(document, "REPORTNO2", "R003", "");
                //LIMSDocX.ReplaceBookmark(document, "REPORTNO2", "", "");
                //LIMSDocX.ReplaceBookmark(document, "RM1", "", "");
                //LIMSDocX.ReplaceBookmark(document, "RM2", "", "");
                document.SaveAs(filedir2);
            }

		}

		private static void replaceBookmark()
		{


            //string filedir = @"C:\Users\lindy\Documents\limsdocx\带封面模板测试.docx";
            //string filedir2 = @"C:\Users\lindy\Documents\limsdocx\带封面模板测试-new.docx";

            string filedir = @"C:\Users\lindy\Documents\limsdocx\常规最新（测试）.docx";
            string filedir2 = @"C:\Users\lindy\Documents\limsdocx\常规最新（测试）-new.docx";

            npLimsDocX.classLimsDocX LIMSDocX = new npLimsDocX.classLimsDocX();
			using (DocX document = LIMSDocX.Load(filedir))
			{
				LIMSDocX.ReplaceBookmark(document, "REPORTNO", "num1","");

                LIMSDocX.ReplaceBookmark(document, "REPORTNO2", "num2", "");

                LIMSDocX.ReplaceBookmark(document, "REPORTNO3", "num3", "");


                //LIMSDocX.ReplaceBookmark(document, "DATE", "2024.07.13", "");

                //LIMSDocX.Save(document);
                document.SaveAs(filedir2);

            }

		}

		private static void generateTestMethod()
		{
			//string filedir = @"C:\Users\lindy\Documents\word\1.docx";
			//generateDoc(filedir);

			//Console.WriteLine("finish!");
			//Console.ReadLine();

			string filedir = @"C:\Users\lindy\Documents\limsdocx\method.docx";

			npLimsDocX.classLimsDocX LIMSDocX = new npLimsDocX.classLimsDocX();
			using (DocX document = LIMSDocX.Load(filedir))
			{
				string tablexml = @"<?xml version='1.0' ?>
                <complexType length='2'>
	                <complexType length='2'>
		                <string>1.</string>
                        <string>静液压强度ddd</string>
	                </complexType>
	                <complexType length='2'>
                        <string>2.</string>
		                <string>熔体质量流动速率</string>
	                </complexType>
                </complexType>";

				//合并单元格等功能
				var table = LIMSDocX.GenerateTable(document, tablexml, "[#Methods]", true);

				//LIMSDocX.SetTableColWidth(document, table, "1", "100");
				//LIMSDocX.SetTableCellFont(table, 0, 0, "宋体", 22, true, false);
				//LIMSDocX.SetTableCellFont(table, 0, 2, "Times New Roman", 10.5, false, false);
				// LIMSDocX.SetTableCellStyle(table, 0, 0, "PARAGRAPHALIGN:LEFT;");
				//LIMSDocX.SetTableCellStyle(table, 2, 0, "PARAGRAPHALIGN:CENTER;");
				//LIMSDocX.SetTableCellStyle(table, 2, 0, "PARAGRAPHALIGN:CENTER;");
				LIMSDocX.SetTableColWidth(document, table, "5,95", "100");

				//表格靠左
				table.Alignment = Alignment.left;
				LIMSDocX.Save(document);
			}
		}

		private static void replacePicture()
		{
			//string filedir = @"C:\Users\lindy\Documents\word\1.docx";
			//generateDoc(filedir);

			//Console.WriteLine("finish!");
			//Console.ReadLine();

			string filedir = @"C:\Users\lindy\Documents\limsdocx\111.docx";
			string pic1 = @"C:\Users\lindy\Documents\limsdocx\1.JPG";
			string pic2 = @"C:\Users\lindy\Documents\limsdocx\1.JPG";

			npLimsDocX.classLimsDocX LIMSDocX = new npLimsDocX.classLimsDocX();
			using (DocX document = LIMSDocX.Load(filedir))
			{
				string tablexml = @"<?xml version='1.0' ?>
                <complexType length='5'>
	                <complexType length='1'>
		                <string>样品图片</string>
	                </complexType>
	                <complexType length='1'>
		                <string>[#picture1]</string>
	                </complexType>
                    <complexType length='1'>
		                <string>bed1</string>
	                </complexType>
                    <complexType length='1'>
		                <string>[#picture2]</string>
	                </complexType>
                       <complexType length='1'>
		                <string>bed2</string>
	                </complexType>
                </complexType>";

				//合并单元格等功能
				var table = LIMSDocX.GenerateTable(document, tablexml, "[#Photo of the sample#]", true);

				LIMSDocX.SetTableColWidth(document, table, "1");
				LIMSDocX.SetTableCellFont(table, 0, 0, "宋体", 22, true, false);
				LIMSDocX.SetTableCellFont(table, 0, 2, "Times New Roman", 10.5, false, false);

				LIMSDocX.SetTableCellStyle(table, 0, 0, "PARAGRAPHALIGN:CENTER;");
				LIMSDocX.SetTableCellStyle(table, 2, 0, "PARAGRAPHALIGN:CENTER;");



				LIMSDocX.InsertPicture(document, "[#picture1]", pic1, "center", 80, 106.7);
				LIMSDocX.InsertPicture(document, "[#picture2]", pic2, "center", 80, 106.7);

				LIMSDocX.SetTableCellStyle(table, 2, 0, "PARAGRAPHALIGN:CENTER;");
				//LIMSDocX.ReplaceFlag(document, "[#Photo of the sample#]", "[#picture1]", "left");
				//LIMSDocX.InsertPicture(document, "[#picture1]", pic, "left", 271, 203);
				LIMSDocX.Save(document);
			}
		}

		private static void replaceIdentity()
		{
			//string filedir = @"C:\Users\lindy\Documents\word\1.docx";
			//generateDoc(filedir);

			//Console.WriteLine("finish!");
			//Console.ReadLine();

			string filedir = @"C:\Users\lindy\Documents\limsdocx\222.docx";
			string pic1 = @"C:\Users\lindy\Documents\limsdocx\2.png";
			npLimsDocX.classLimsDocX LIMSDocX = new npLimsDocX.classLimsDocX();
			using (DocX document = LIMSDocX.Load(filedir))
			{
				var Width1 = 90;
				var Height1 = 62;

				var Width2 = 201;
				var Height2 = 57;

				LIMSDocX.InsertPicture(document, "[#CMA]", pic1, "left", Height1, Width1);



				LIMSDocX.Save(document);
			}


		}

		private static void replaceSign()
		{


			string filedir = @"C:\Users\lindy\Documents\limsdocx\1025.docx";
			string pic1 = @"C:\Users\lindy\Documents\limsdocx\s1.png";
			string pic2 = @"C:\Users\lindy\Documents\limsdocx\s2.png";
			string pic3 = @"C:\Users\lindy\Documents\limsdocx\s7.png";

            npLimsDocX.classLimsDocX LIMSDocX = new npLimsDocX.classLimsDocX();
			using (DocX document = LIMSDocX.Load(filedir))
			{
				var Width1 = 100;
				var Height1 = 30;
				LIMSDocX.InsertPicture(document, "[#SIGNEE2]", pic1, "CENTER", Height1, Width1);
				//LIMSDocX.InsertPicture(document, "[#SIGNEE3]", pic2, "CENTER", Height1, Width1);
				//LIMSDocX.InsertPicture(document, "[#SIGNEE5]", pic3, "CENTER", Height1, Width1);

				LIMSDocX.Save(document);
			}


		}


		private static void testReplacePicture()
		{
            npLimsDocX.classLimsDocX c = new npLimsDocX.classLimsDocX();

            string filedir = @"C:\Users\lindy\Documents\limsdocx\测试签名.docx";

            filedir = @"C:\Users\lindy\Documents\limsdocx\doc1.docx";

            filedir = @"C:\Users\lindy\Documents\limsdocx\A04HCC2Q00028C.docx";

            string pic3 = @"C:\Users\lindy\Documents\limsdocx\s1.png";
            npLimsDocX.classLimsDocX LIMSDocX = new npLimsDocX.classLimsDocX();
            using (DocX document = LIMSDocX.Load(filedir))
            {
                var Width1 = 100;
                var Height1 = 30;
                LIMSDocX.InsertPictureByBookMark(document, "SIGNEE2", "", "CENTER", Height1, Width1);
                LIMSDocX.InsertPictureByBookMark(document, "SIGNEE3", "", "CENTER", Height1, Width1);
                LIMSDocX.InsertPictureByBookMark(document, "SIGNEE5", "", "CENTER", Height1, Width1);
                LIMSDocX.InsertPictureByBookMark(document, "cnaspic", pic3, "CENTER", Height1, Width1);
				LIMSDocX.InsertPictureByBookMark(document, "currentDate", pic3, "CENTER", Height1, Width1);
                


                LIMSDocX.Save(document);
            }
        }


		//private static void testPictureArrow()
		//{
		//    //string filedir = @"C:\Users\lindy\Documents\word\1.docx";
		//    //generateDoc(filedir);

		//    //Console.WriteLine("finish!");
		//    //Console.ReadLine();

		//    string filedir = @"C:\Users\lindy\Documents\limsdocx\111.docx";
		//    string pic1 = @"C:\Users\lindy\Documents\limsdocx\CNAS章.png";

		//    npLimsDocX.classLimsDocX LIMSDocX = new npLimsDocX.classLimsDocX();
		//    using (DocX document = LIMSDocX.Load(filedir))
		//    {
		//        LIMSDocX.InsertPicture(document, "[#CNAS]", pic1, "center", 80, 106.7, BlockArrowShapes.upArrow);

		//        LIMSDocX.Save(document);

		//    }
		//}

		private static string getTableXml()
		{
			string tablexml = @"<?xml version='1.0' ?>
<complexType length='32'>
	<complexType length='9'>
		<string>产品名称
Sample</string>
		<string>le</string>
		<string>[#SAMPLENAME]</string>
		<string>le</string>
		<string>le</string>
		<string>le</string>
		<string>规格型号
Model</string>
		<string>[#SPECIFICATION]</string>
		<string>le</string>
	</complexType>
	<complexType length='9'>
		<string>序号
№</string>
		<string>检验项目
Items</string>
		<string>le</string>
		<string>le</string>
		<string>单位
Unit</string>
		<string>检验方法依据
Standards</string>
		<string>标准要求
Specification</string>
		<string>检验结果
Inspection Data</string>
		<string>单项结论
Conclusion</string>
	</complexType>
	<complexType length='9'>
		<string>1</string>
		<string>面积</string>
		<string>le</string>
		<string>le</string>
		<string>—</string>
		<string>GB 23441-2009
 中5.2</string>
		<string>不小于产品面积标记值的99%</string>
		<string>101,101,101,101,101</string>
		<string>合格</string>
	</complexType>
	<complexType length='9'>
		<string>2</string>
		<string>单位面积质量</string>
		<string>le</string>
		<string>le</string>
		<string>kg/m²</string>
		<string>GB 23441-2009
 中5.3</string>
		<string>≥1.5</string>
		<string>2.0,2.0,2.0,2.0,2.0</string>
		<string>合格</string>
	</complexType>
	<complexType length='9'>
		<string>3</string>
		<string>厚度</string>
		<string>平均值</string>
		<string>le</string>
		<string>mm</string>
		<string>GB 23441-2009
 中5.4</string>
		<string>≥1.5</string>
		<string>1.8,1.8,1.7,1.8,1.8</string>
		<string>合格</string>
	</complexType>
	<complexType length='9'>
		<string>up</string>
		<string>up</string>
		<string>最小单值</string>
		<string>le</string>
		<string>mm</string>
		<string>up</string>
		<string>≥1.3</string>
		<string>1.7,1.7,1.7,1.7,1.8</string>
		<string>up</string>
	</complexType>
	<complexType length='9'>
		<string>4</string>
		<string>外观</string>
		<string>le</string>
		<string>le</string>
		<string>—</string>
		<string>GB/T 328.2-2007</string>
		<string>1. 成卷卷材应卷紧卷齐，端面里进外出不得超过20mm。 2. 成卷卷材在4℃~45 ℃任一产品温度下展开，在距卷芯1000mm长度外不应有裂纹或长度10mm以上的粘结。 3. PY类产品,其胎基应浸透，不应有未被漫溃的浅色条纹。 4. 卷材表面应平整，不允许有孔洞、结块、气泡、缺边和裂口，上表面为细砂的，细砂应均匀一致并紧密地粘附于卷材表面。 5. 每卷材接头不应超过一个，较短的一段长度不应少于1000mm，接头应剪切整齐，并加长150 mm。</string>
		<string>1</string>
		<string>合格</string>
	</complexType>
	<complexType length='9'>
		<string>5</string>
		<string>拉伸性能</string>
		<string>拉力</string>
		<string>纵向</string>
		<string>N/mm</string>
		<string>GB 23441-2009
 中5.7</string>
		<string>≥150</string>
		<string>640</string>
		<string>不合格</string>
	</complexType>
	<complexType length='9'>
		<string>up</string>
		<string>up</string>
		<string>up</string>
		<string>横向</string>
		<string>N/mm</string>
		<string>up</string>
		<string>≥150</string>
		<string>811</string>
		<string>up</string>
	</complexType>
	<complexType length='9'>
		<string>up</string>
		<string>up</string>
		<string>最大拉力时延伸率</string>
		<string>纵向</string>
		<string>%</string>
		<string>up</string>
		<string>≥30</string>
		<string>79</string>
		<string>up</string>
	</complexType>
	<complexType length='9'>
		<string>up</string>
		<string>up</string>
		<string>up</string>
		<string>横向</string>
		<string>%</string>
		<string>up</string>
		<string>≥30</string>
		<string>88</string>
		<string>up</string>
	</complexType>
	<complexType length='9'>
		<string>up</string>
		<string>up</string>
		<string>沥青断裂延伸率</string>
		<string>纵向</string>
		<string>%</string>
		<string>up</string>
		<string>≥150</string>
		<string>82</string>
		<string>up</string>
	</complexType>
	<complexType length='9'>
		<string>up</string>
		<string>up</string>
		<string>up</string>
		<string>横向</string>
		<string>%</string>
		<string>up</string>
		<string>≥150</string>
		<string>93</string>
		<string>up</string>
	</complexType>
	<complexType length='9'>
		<string>up</string>
		<string>up</string>
		<string>拉伸时现象</string>
		<string>le</string>
		<string>—</string>
		<string>up</string>
		<string>拉伸过程中，在膜断裂前无沥青涂盖层与膜分离现象</string>
		<string>拉伸过程中，在膜断裂前无沥青涂盖层与膜分离现象</string>
		<string>up</string>
	</complexType>
	<complexType length='9'>
		<string>6</string>
		<string>钉杆撕裂强度</string>
		<string>le</string>
		<string>le</string>
		<string>N</string>
		<string>GB/T 328.18-2007</string>
		<string>≥30</string>
		<string>48</string>
		<string>合格</string>
	</complexType>
	<complexType length='9'>
		<string>7</string>
		<string>耐热性</string>
		<string>le</string>
		<string>le</string>
		<string>—</string>
		<string>GB 23441-2009
 中5.9</string>
		<string>70℃滑移不超过2mm</string>
		<string>滑动：0</string>
		<string>合格</string>
	</complexType>
	<complexType length='9'>
		<string>8</string>
		<string>低温柔性</string>
		<string>le</string>
		<string>le</string>
		<string>—</string>
		<string>GB 23441-2009
 中5.10</string>
		<string>-20℃，无裂纹</string>
		<string>无裂纹</string>
		<string>合格</string>
	</complexType>
	<complexType length='9'>
		<string>9</string>
		<string>不透水性</string>
		<string>le</string>
		<string>le</string>
		<string>—</string>
		<string>GB 23441-2009
 中5.11</string>
		<string>0.2MPa，120min不透水</string>
		<string>不透水</string>
		<string>合格</string>
	</complexType>
	<complexType length='9'>
		<string>10</string>
		<string>剥离强度</string>
		<string>卷材与卷材</string>
		<string>le</string>
		<string>N/mm</string>
		<string>GB 23441-2009
 中5.12</string>
		<string>≥1.0</string>
		<string>0.2</string>
		<string>不合格</string>
	</complexType>
	<complexType length='9'>
		<string>up</string>
		<string>up</string>
		<string>卷材与铝板</string>
		<string>le</string>
		<string>N/mm</string>
		<string>up</string>
		<string>≥1.5</string>
		<string>0.4</string>
		<string>up</string>
	</complexType>
	<complexType length='9'>
		<string>11</string>
		<string>钉杆水密性</string>
		<string>le</string>
		<string>le</string>
		<string>—</string>
		<string>GB 23441-2009
 中5.13</string>
		<string>通过</string>
		<string>不通过</string>
		<string>不合格</string>
	</complexType>
	<complexType length='9'>
		<string>12</string>
		<string>渗油性</string>
		<string>le</string>
		<string>le</string>
		<string>张</string>
		<string>GB 23441-2009
 中5.14</string>
		<string>≤2</string>
		<string>1</string>
		<string>合格</string>
	</complexType>
	<complexType length='9'>
		<string>13</string>
		<string>持粘性</string>
		<string>le</string>
		<string>le</string>
		<string>min</string>
		<string>GB 23441-2009
 中5.15</string>
		<string>≥20</string>
		<string>1</string>
		<string>不合格</string>
	</complexType>
	<complexType length='9'>
		<string>14</string>
		<string>热老化</string>
		<string>拉力保持率</string>
		<string>纵向</string>
		<string>%</string>
		<string>GB 23441-2009
 中5.16</string>
		<string>≥80</string>
		<string>107</string>
		<string>不合格</string>
	</complexType>
	<complexType length='9'>
		<string>up</string>
		<string>up</string>
		<string>up</string>
		<string>横向</string>
		<string>%</string>
		<string>up</string>
		<string>≥80</string>
		<string>102</string>
		<string>up</string>
	</complexType>
	<complexType length='9'>
		<string>up</string>
		<string>up</string>
		<string>最大拉力时延伸率</string>
		<string>纵向</string>
		<string>%</string>
		<string>up</string>
		<string>≥30</string>
		<string>99</string>
		<string>up</string>
	</complexType>
	<complexType length='9'>
		<string>up</string>
		<string>up</string>
		<string>up</string>
		<string>横向</string>
		<string>%</string>
		<string>up</string>
		<string>≥30</string>
		<string>88</string>
		<string>up</string>
	</complexType>
	<complexType length='9'>
		<string>up</string>
		<string>up</string>
		<string>低温柔性</string>
		<string>le</string>
		<string>—</string>
		<string>up</string>
		<string>-18℃，无裂纹</string>
		<string>无裂纹</string>
		<string>up</string>
	</complexType>
	<complexType length='9'>
		<string>up</string>
		<string>up</string>
		<string>剥离强度卷材与铝板</string>
		<string>le</string>
		<string>N/mm</string>
		<string>up</string>
		<string>≥1.5</string>
		<string>0.2</string>
		<string>up</string>
	</complexType>
	<complexType length='9'>
		<string>15</string>
		<string>热稳定性</string>
		<string>尺寸变化</string>
		<string>纵向</string>
		<string>%</string>
		<string>GB 23441-2009
 中5.17</string>
		<string>≤2</string>
		<string>0.6</string>
		<string>合格</string>
	</complexType>
	<complexType length='9'>
		<string>up</string>
		<string>up</string>
		<string>up</string>
		<string>横向</string>
		<string>%</string>
		<string>up</string>
		<string>≤2</string>
		<string>0.2</string>
		<string>up</string>
	</complexType>
	<complexType length='9'>
		<string>up</string>
		<string>up</string>
		<string>外观</string>
		<string>le</string>
		<string>—</string>
		<string>up</string>
		<string>无起鼓、皱褶、滑动、流淌</string>
		<string>无起鼓、褶皱、滑动、流淌</string>
		<string>up</string>
	</complexType>
</complexType>";

			return tablexml;

		}


		private static string getTableXml2()
		{
			string tablexml = @"<?xml version='1.0' ?>
<complexType length='10'>
	<complexType length='9'>
		<string>序号</string>
		<string>检测项目</string>
		<string>le</string>
		<string>le</string>
		<string>单位</string>
		<string>检测依据</string>
		<string>技术要求</string>
		<string>检测结果</string>
		<string>单项评定</string>
	</complexType>
	<complexType length='9'>
		<string>1</string>
		<string>可溶物含量</string>
		<string>le</string>
		<string>le</string>
		<string>g/m²</string>
		<string>GB18242</string>
		<string>≥2100</string>
		<string>2554</string>
		<string>合格</string>
	</complexType>
	<complexType length='9'>
		<string>2</string>
		<string>耐热性</string>
		<string>试验现象</string>
		<string>le</string>
		<string>—</string>
		<string>GB18242</string>
		<string>无流淌、滴落</string>
		<string>无流淌、滴落</string>
		<string>合格</string>
	</complexType>
	<complexType length='9'>
		<string>3</string>
		<string>低温柔性</string>
		<string>le</string>
		<string>le</string>
		<string>—</string>
		<string>GB18242</string>
		<string>-25℃，无裂缝</string>
		<string>无裂缝</string>
		<string>合格</string>
	</complexType>
	<complexType length='9'>
		<string>4</string>
		<string>不透水性</string>
		<string>le</string>
		<string>le</string>
		<string>—</string>
		<string>GB18242</string>
		<string>0.3MPa，30min不透水</string>
		<string>不透水</string>
		<string>合格</string>
	</complexType>
	<complexType length='9'>
		<string>5</string>
		<string>拉力</string>
		<string>最大峰拉力</string>
		<string>纵向</string>
		<string>N/50mm</string>
		<string>GB18242</string>
		<string>≥800</string>
		<string>1035</string>
		<string>合格</string>
	</complexType>
	<complexType length='9'>
		<string>up</string>
		<string>up</string>
		<string>up</string>
		<string>横向</string>
		<string>N/50mm</string>
		<string>GB18242</string>
		<string>≥800</string>
		<string>965</string>
		<string>up</string>
	</complexType>
	<complexType length='9'>
		<string>up</string>
		<string>up</string>
		<string>试验现象</string>
		<string>le</string>
		<string>—</string>
		<string>GB18242</string>
		<string>拉伸过程中，试件中部无沥青涂盖层开裂或与胎基分离现象</string>
		<string>拉伸过程中，试件中部无沥青涂盖层开裂和与胎基分离现象</string>
		<string>up</string>
	</complexType>
	<complexType length='9'>
		<string>6</string>
		<string>延伸率</string>
		<string>最大峰时延伸率</string>
		<string>纵向</string>
		<string>%</string>
		<string>GB18242</string>
		<string>≥40</string>
		<string>45</string>
		<string>合格</string>
	</complexType>
	<complexType length='9'>
		<string>up</string>
		<string>up</string>
		<string>up</string>
		<string>横向</string>
		<string>%</string>
		<string>GB18242</string>
		<string>≥40</string>
		<string>55</string>
		<string>up</string>
	</complexType>
</complexType>";
			return tablexml;

        }

		private static void combineDoc()
		{
            string filedir = @"C:\Users\lindy\Desktop\交通基建集料报告模板\交通基建模板测试.docx";
            string newfiledir = @"C:\Users\lindy\Desktop\交通基建集料报告模板\交通基建模板测试-out.docx";

            string filedir2 = @"C:\Users\lindy\Desktop\交通基建集料报告模板\颗粒级配.docx";
            npLimsDocX.classLimsDocX LIMSDocX = new npLimsDocX.classLimsDocX();
            using (DocX document = LIMSDocX.Load(filedir))
            {
				DocX document2 = LIMSDocX.Load(filedir2);

				LIMSDocX.DocUnitAsOne(document, document2);
                //LIMSDocX.Save(document);
                //document.SaveAs(newfiledir);
            }
        }
	}
}

