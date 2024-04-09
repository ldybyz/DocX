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

namespace ConsoleApp1
{
	class Program
	{

		static void Main(string[] args)
		{
			string filedir = @"C:\Users\lindy\Documents\limsdocx\111.docx";
			//replaceSign();
			generateDoc(filedir);

			//generateTestMethod();

			//test1();


			Console.WriteLine("finish!");
			Console.ReadLine();

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

				string tablexml = @"<?xml version='1.0' ?>
<complexType length='15'>
	<complexType length='8'>
		<string>序号</string>
		<string>检验项目</string>
		<string>le</string>
		<string>le</string>
		<string>单位</string>
		<string>标准要求</string>
		<string>检测结果</string>
		<string>单项评定</string>
	</complexType>
	<complexType length='8'>
		<string>1</string>
		<string>断后伸长率A₅₀ₘₘ</string>
		<string>抗拉强度Rₘ</string>
		<string>规定非比例延伸强度Rₚ₀.₂</string>
		<string>规定非比例延伸强度Rₚ₀.₂</string>
		<string>/</string>
		<string>0.008</string>
		<string>合格</string>
	</complexType>
	<complexType length='8'>
		<string>up</string>
		<string>up</string>
		<string>锰</string>
		<string>le</string>
		<string>μg/L</string>
		<string>≤30.0</string>
		<string>0.0617</string>
		<string>up</string>
	</complexType>
	<complexType length='8'>
		<string>up</string>
		<string>up</string>
		<string>镍</string>
		<string>le</string>
		<string>μg/L</string>
		<string>≤20.0</string>
		<string>1.07</string>
		<string>up</string>
	</complexType>
	<complexType length='8'>
		<string>up</string>
		<string>up</string>
		<string>铜</string>
		<string>le</string>
		<string>μg/L</string>
		<string>≤130.0</string>
		<string>24.85</string>
		<string>up</string>
	</complexType>
	<complexType length='8'>
		<string>2</string>
		<string>抗水压机械性能</string>
		<string>阀芯下游</string>
		<string>le</string>
		<string>/</string>
		<string>阀芯下游的任何零部件无永久性变形</string>
		<string>阀芯上游的任何零部件无永久性变形</string>
		<string>合格</string>
	</complexType>
	<complexType length='8'>
		<string>up</string>
		<string>up</string>
		<string>阀芯上游</string>
		<string>le</string>
		<string>/</string>
		<string>阀芯上游的任何零部件无永久性变形</string>
		<string>阀芯下游的任何零部件无永久性变形</string>
		<string>up</string>
	</complexType>
	<complexType length='8'>
		<string>3</string>
		<string>密封性能</string>
		<string>冷热水隔墙</string>
		<string>le</string>
		<string>/</string>
		<string>出水口及未连接的进水口无渗漏</string>
		<string>出水口及未连接的进水口无渗漏</string>
		<string>合格</string>
	</complexType>
	<complexType length='8'>
		<string>up</string>
		<string>up</string>
		<string>阀芯下游</string>
		<string>le</string>
		<string>/</string>
		<string>阀芯下游任何密封部位无渗漏</string>
		<string>阀芯下游任何密封部位无渗漏</string>
		<string>up</string>
	</complexType>
	<complexType length='8'>
		<string>up</string>
		<string>up</string>
		<string>阀芯上游</string>
		<string>le</string>
		<string>/</string>
		<string>阀芯及上游过水通道无渗漏</string>
		<string>阀芯及上游过水通道无渗漏</string>
		<string>up</string>
	</complexType>
	<complexType length='8'>
		<string>4</string>
		<string>流量</string>
		<string>普通型</string>
		<string>le</string>
		<string>L/min</string>
		<string>3.0~9.0</string>
		<string>3.2</string>
		<string>合格</string>
	</complexType>
	<complexType length='8'>
		<string>5</string>
		<string>抗安装负载</string>
		<string>le</string>
		<string>le</string>
		<string>/</string>
		<string>试验后螺纹应无裂纹、无损坏</string>
		<string>经试验后螺纹无裂纹、无损坏</string>
		<string>合格</string>
	</complexType>
	<complexType length='8'>
		<string>6</string>
		<string>抗使用负载</string>
		<string>le</string>
		<string>le</string>
		<string>/</string>
		<string>水嘴手柄或手轮在开启和关闭方向上施加（6±0.2）N.m后，应无变形或损坏等削弱水嘴功能的情况出现，水嘴阀芯上游密封性能应符合7.6.2的要求;其他水嘴手柄或手轮承受45N的轴向拉力应无松动现象</string>
		<string>无变形和损坏等削弱水嘴功能的情况出现，且阀芯及上游过水通道无渗漏；水嘴手柄承受45N的轴向拉力无松动现象</string>
		<string>合格</string>
	</complexType>
	<complexType length='8'>
		<string>7</string>
		<string>表面耐腐蚀性能</string>
		<string>le</string>
		<string>le</string>
		<string>/</string>
		<string>水嘴按8.6.7进行酸性盐雾试验后，应不低于GB/T 6461-2002标准的表1中外观评级（R{d|A}）9级的要求。</string>
		<string>10级</string>
		<string>合格</string>
	</complexType>
	<complexType length='8'>
		<string>8</string>
		<string>防回流性能</string>
		<string>le</string>
		<string>le</string>
		<string>/</string>
		<string>/</string>
		<string>无虹吸现象产生</string>
		<string>合格</string>
	</complexType>
</complexType>";
				//var table = c.GenerateTable(doc, tablexml, "[#TestTable]", true);

				var table = c.GenerateTableWithWidth(doc, tablexml, "[#TestTable]", true, "8,11,11,11,8,19,20,12", "108");

				//c.SetTableColFixedWidth(doc, table, "8,11,11,11,9,20,20,10", "100");
				//c.SetTableCellFont(table, 1, 3, "Times New Roman", 0, false, true);
				c.SetTableCellFont(table, 1, 3, "Times New Roman", 0, false, true);
				c.SetTableRowHeight(doc, table, -1, 0.5, false);
				c.SetTableHeader(table, 0);
				c.SetTableBorderLine(table, "INSIDEV,INSIDEH,LEFT,RIGHT,TOP,BOTTOM");


				doc.Save();

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
				LIMSDocX.ReplaceFlag(document, "[#TestTable]", "[#TestTable1]\n[#TestTable2]\n[#TestTable3]", "");
				LIMSDocX.Save(document);
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


			string filedir = @"C:\Users\lindy\Documents\limsdocx\测试签名.docx";
			string pic1 = @"C:\Users\lindy\Documents\limsdocx\s1.png";
			string pic2 = @"C:\Users\lindy\Documents\limsdocx\s2.png";
			string pic3 = @"C:\Users\lindy\Documents\limsdocx\s7.png";
			npLimsDocX.classLimsDocX LIMSDocX = new npLimsDocX.classLimsDocX();
			using (DocX document = LIMSDocX.Load(filedir))
			{
				var Width1 = 100;
				var Height1 = 30;
				LIMSDocX.InsertPicture(document, "[#SIGNEE2]", pic1, "CENTER", Height1, Width1);
				LIMSDocX.InsertPicture(document, "[#SIGNEE3]", pic2, "CENTER", Height1, Width1);
				LIMSDocX.InsertPicture(document, "[#SIGNEE5]", pic3, "CENTER", Height1, Width1);
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
<complexType length='50'>
	<complexType length='8'>
		<string>序号</string>
		<string>检验项目</string>
		<string>le</string>
		<string>le</string>
		<string>单位</string>
		<string>标准要求</string>
		<string>检测结果</string>
		<string>单项评定</string>
	</complexType>
<complexType length='8'>
		<string>序号1</string>
		<string>检验项目1</string>
		<string>le</string>
		<string>le</string>
		<string>单位1</string>
		<string>标准要求1</string>
		<string>检测结果1</string>
		<string>单项评定1</string>
	</complexType>
	<complexType length='8'>
		<string>1</string>
		<string>不透水性</string>
		<string>le</string>
		<string>le</string>
		<string>/</string>
		<string>0.3MPa，30min不透水</string>
		<string>透水</string>
		<string>不合格</string>
	</complexType>
	<complexType length='8'>
		<string>2</string>
		<string>低温柔性</string>
		<string>le</string>
		<string>上表面</string>
		<string>/</string>
		<string>-25℃，无裂缝</string>
		<string>通过</string>
		<string>合格</string>
	</complexType>
	<complexType length='8'>
		<string>up</string>
		<string>up</string>
		<string>up</string>
		<string>下表面</string>
		<string>/</string>
		<string>-25℃，无裂缝</string>
		<string>通过</string>
		<string>up</string>
	</complexType>
	<complexType length='8'>
		<string>3</string>
		<string>卷材下表面沥青涂盖层厚度</string>
		<string>le</string>
		<string>le</string>
		<string>mm</string>
		<string>≥1.0</string>
		<string>1.1</string>
		<string>合格</string>
	</complexType>
	<complexType length='8'>
		<string>4</string>
		<string>可溶物含量 </string>
		<string>可溶物含量</string>
		<string>le</string>
		<string>g/㎡</string>
		<string>≥2900</string>
		<string>2900</string>
		<string>合格</string>
	</complexType>
	<complexType length='8'>
		<string>5</string>
		<string>延伸率</string>
		<string>第二峰时延伸率</string>
		<string>横向</string>
		<string>%</string>
		<string>/</string>
		<string>45</string>
		<string>合格</string>
	</complexType>
	<complexType length='8'>
		<string>up</string>
		<string>up</string>
		<string>up</string>
		<string>纵向</string>
		<string>%</string>
		<string>/</string>
		<string>43</string>
		<string>up</string>
	</complexType>
	<complexType length='8'>
		<string>up</string>
		<string>up</string>
		<string>最大峰时延伸率</string>
		<string>横向</string>
		<string>%</string>
		<string>≥40</string>
		<string>42</string>
		<string>up</string>
	</complexType>
	<complexType length='8'>
		<string>up</string>
		<string>up</string>
		<string>up</string>
		<string>纵向</string>
		<string>%</string>
		<string>≥40</string>
		<string>41</string>
		<string>up</string>
	</complexType>
	<complexType length='8'>
		<string>6</string>
		<string>拉力</string>
		<string>次高峰拉力</string>
		<string>横向</string>
		<string>N/50mm</string>
		<string>/</string>
		<string>100</string>
		<string>合格</string>
	</complexType>
	<complexType length='8'>
		<string>up</string>
		<string>up</string>
		<string>up</string>
		<string>纵向</string>
		<string>N/50mm</string>
		<string>/</string>
		<string>800</string>
		<string>up</string>
	</complexType>
	<complexType length='8'>
		<string>up</string>
		<string>up</string>
		<string>最大峰拉力</string>
		<string>横向</string>
		<string>N/50mm</string>
		<string>≥800</string>
		<string>800</string>
		<string>up</string>
	</complexType>
	<complexType length='8'>
		<string>up</string>
		<string>up</string>
		<string>up</string>
		<string>纵向</string>
		<string>N/50mm</string>
		<string>≥800</string>
		<string>800</string>
		<string>up</string>
	</complexType>
	<complexType length='8'>
		<string>up</string>
		<string>up</string>
		<string>试验现象</string>
		<string>le</string>
		<string>/</string>
		<string>拉伸过程中，试件中部无沥青涂盖层厚度开裂或与胎基分离现象</string>
		<string>拉伸过程中，试件中部无沥青涂盖层厚度开裂或与胎基分离现象</string>
		<string>up</string>
	</complexType>
	<complexType length='8'>
		<string>7</string>
		<string>接缝剥离强度</string>
		<string>le</string>
		<string>le</string>
		<string>N/mm</string>
		<string>≥1.5</string>
		<string>1.6</string>
		<string>合格</string>
	</complexType>
	<complexType length='8'>
		<string>8</string>
		<string>热老化</string>
		<string>延伸率保持率</string>
		<string>横向</string>
		<string>%</string>
		<string>≥90</string>
		<string>90</string>
		<string>合格</string>
	</complexType>
	<complexType length='8'>
		<string>up</string>
		<string>up</string>
		<string>up</string>
		<string>纵向</string>
		<string>%</string>
		<string>≥90</string>
		<string>90</string>
		<string>up</string>
	</complexType>
	<complexType length='8'>
		<string>up</string>
		<string>up</string>
		<string>拉力保持率</string>
		<string>横向</string>
		<string>%</string>
		<string>≥80</string>
		<string>80.0</string>
		<string>up</string>
	</complexType>
	<complexType length='8'>
		<string>up</string>
		<string>up</string>
		<string>up</string>
		<string>纵向</string>
		<string>%</string>
		<string>≥80</string>
		<string>80.0</string>
		<string>up</string>
	</complexType>
	<complexType length='8'>
		<string>up</string>
		<string>up</string>
		<string>尺寸变化率</string>
		<string>le</string>
		<string>%</string>
		<string>≤0.7</string>
		<string>0.7</string>
		<string>up</string>
	</complexType>
	<complexType length='8'>
		<string>up</string>
		<string>up</string>
		<string>低温柔性</string>
		<string>上表面</string>
		<string>/</string>
		<string>/</string>
		<string>通过</string>
		<string>up</string>
	</complexType>
	<complexType length='8'>
		<string>up</string>
		<string>up</string>
		<string>up</string>
		<string>下表面</string>
		<string>/</string>
		<string>/</string>
		<string>通过</string>
		<string>up</string>
	</complexType>
	<complexType length='8'>
		<string>up</string>
		<string>up</string>
		<string>质量损失</string>
		<string>le</string>
		<string>%</string>
		<string>≤1.0</string>
		<string>0.9</string>
		<string>up</string>
	</complexType>
	<complexType length='8'>
		<string>9</string>
		<string>耐热性</string>
		<string>le</string>
		<string>le</string>
		<string>/</string>
		<string>≤2mm，无流淌、滴落</string>
		<string>≤2mm，无流淌、滴落</string>
		<string>合格</string>
	</complexType>


<complexType length='8'>
		<string>1</string>
		<string>不透水性</string>
		<string>le</string>
		<string>le</string>
		<string>/</string>
		<string>0.3MPa，30min不透水</string>
		<string>透水</string>
		<string>不合格</string>
	</complexType>
	<complexType length='8'>
		<string>2</string>
		<string>低温柔性</string>
		<string>le</string>
		<string>上表面</string>
		<string>/</string>
		<string>-25℃，无裂缝</string>
		<string>通过</string>
		<string>合格</string>
	</complexType>
	<complexType length='8'>
		<string>up</string>
		<string>up</string>
		<string>up</string>
		<string>下表面</string>
		<string>/</string>
		<string>-25℃，无裂缝</string>
		<string>通过</string>
		<string>up</string>
	</complexType>
	<complexType length='8'>
		<string>3</string>
		<string>卷材下表面沥青涂盖层厚度</string>
		<string>le</string>
		<string>le</string>
		<string>mm</string>
		<string>≥1.0</string>
		<string>1.1</string>
		<string>合格</string>
	</complexType>
	<complexType length='8'>
		<string>4</string>
		<string>可溶物含量 </string>
		<string>可溶物含量</string>
		<string>le</string>
		<string>g/㎡</string>
		<string>≥2900</string>
		<string>2900</string>
		<string>合格</string>
	</complexType>
	<complexType length='8'>
		<string>5</string>
		<string>延伸率</string>
		<string>第二峰时延伸率</string>
		<string>横向</string>
		<string>%</string>
		<string>/</string>
		<string>45</string>
		<string>合格</string>
	</complexType>
	<complexType length='8'>
		<string>up</string>
		<string>up</string>
		<string>up</string>
		<string>纵向</string>
		<string>%</string>
		<string>/</string>
		<string>43</string>
		<string>up</string>
	</complexType>
	<complexType length='8'>
		<string>up</string>
		<string>up</string>
		<string>最大峰时延伸率</string>
		<string>横向</string>
		<string>%</string>
		<string>≥40</string>
		<string>42</string>
		<string>up</string>
	</complexType>
	<complexType length='8'>
		<string>up</string>
		<string>up</string>
		<string>up</string>
		<string>纵向</string>
		<string>%</string>
		<string>≥40</string>
		<string>41</string>
		<string>up</string>
	</complexType>
	<complexType length='8'>
		<string>6</string>
		<string>拉力</string>
		<string>次高峰拉力</string>
		<string>横向</string>
		<string>N/50mm</string>
		<string>/</string>
		<string>100</string>
		<string>合格</string>
	</complexType>
	<complexType length='8'>
		<string>up</string>
		<string>up</string>
		<string>up</string>
		<string>纵向</string>
		<string>N/50mm</string>
		<string>/</string>
		<string>800</string>
		<string>up</string>
	</complexType>
	<complexType length='8'>
		<string>up</string>
		<string>up</string>
		<string>最大峰拉力</string>
		<string>横向</string>
		<string>N/50mm</string>
		<string>≥800</string>
		<string>800</string>
		<string>up</string>
	</complexType>
	<complexType length='8'>
		<string>up</string>
		<string>up</string>
		<string>up</string>
		<string>纵向</string>
		<string>N/50mm</string>
		<string>≥800</string>
		<string>800</string>
		<string>up</string>
	</complexType>
	<complexType length='8'>
		<string>up</string>
		<string>up</string>
		<string>试验现象</string>
		<string>le</string>
		<string>/</string>
		<string>拉伸过程中，试件中部无沥青涂盖层厚度开裂或与胎基分离现象</string>
		<string>拉伸过程中，试件中部无沥青涂盖层厚度开裂或与胎基分离现象</string>
		<string>up</string>
	</complexType>
	<complexType length='8'>
		<string>7</string>
		<string>接缝剥离强度</string>
		<string>le</string>
		<string>le</string>
		<string>N/mm</string>
		<string>≥1.5</string>
		<string>1.6</string>
		<string>合格</string>
	</complexType>
	<complexType length='8'>
		<string>8</string>
		<string>热老化</string>
		<string>延伸率保持率</string>
		<string>横向</string>
		<string>%</string>
		<string>≥90</string>
		<string>90</string>
		<string>合格</string>
	</complexType>
	<complexType length='8'>
		<string>up</string>
		<string>up</string>
		<string>up</string>
		<string>纵向</string>
		<string>%</string>
		<string>≥90</string>
		<string>90</string>
		<string>up</string>
	</complexType>
	<complexType length='8'>
		<string>up</string>
		<string>up</string>
		<string>拉力保持率</string>
		<string>横向</string>
		<string>%</string>
		<string>≥80</string>
		<string>80.0</string>
		<string>up</string>
	</complexType>
	<complexType length='8'>
		<string>up</string>
		<string>up</string>
		<string>up</string>
		<string>纵向</string>
		<string>%</string>
		<string>≥80</string>
		<string>80.0</string>
		<string>up</string>
	</complexType>
	<complexType length='8'>
		<string>up</string>
		<string>up</string>
		<string>尺寸变化率</string>
		<string>le</string>
		<string>%</string>
		<string>≤0.7</string>
		<string>0.7</string>
		<string>up</string>
	</complexType>
	<complexType length='8'>
		<string>up</string>
		<string>up</string>
		<string>低温柔性</string>
		<string>上表面</string>
		<string>/</string>
		<string>/</string>
		<string>通过</string>
		<string>up</string>
	</complexType>
	<complexType length='8'>
		<string>up</string>
		<string>up</string>
		<string>up</string>
		<string>下表面</string>
		<string>/</string>
		<string>/</string>
		<string>通过</string>
		<string>up</string>
	</complexType>
	<complexType length='8'>
		<string>up</string>
		<string>up</string>
		<string>质量损失</string>
		<string>le</string>
		<string>%</string>
		<string>≤1.0</string>
		<string>0.9</string>
		<string>up</string>
	</complexType>
	<complexType length='8'>
		<string>9</string>
		<string>耐热性</string>
		<string>le</string>
		<string>le</string>
		<string>/</string>
		<string>≤2mm，无流淌、滴落</string>
		<string>≤2mm，无流淌、滴落</string>
		<string>合格</string>
	</complexType>
</complexType>";

			return tablexml;

		}
	}
}

