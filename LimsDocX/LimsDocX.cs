

using System.Data;
using System.Xml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.IO;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO.Packaging;
using System.Xml.Linq;
using System.Text.RegularExpressions;
using Xceed.Words.NET;
using Xceed.Document.NET;
using Formatting = Xceed.Document.NET.Formatting;

namespace npLimsDocX
{
    public class classLimsDocX
    {

        private const double OldVersionFactor = 1.333333333333333;
        /// <summary>
        /// 加载DocX
        /// </summary>
        /// <returns></returns>
        public DocX Load(string fileName)
        {
            return DocX.Load(fileName);
        }

        /// <summary>
        /// 保存DocX
        /// </summary>
        /// <param name="doc"></param>
        public void Save(DocX doc)
        {
            doc.Save();
        }

        /// <summary>
        /// 生成一个表格
        /// </summary>
        /// <param name="doc"></param>
        /// <param name="dataTable"></param>
        public Table GenerateTable(DocX doc, string strXml, string replaceFlag, bool bEmptyParagraph)
        {
            DataTable dt = XMLDeserialize(strXml);
            int nRow = dt.Rows.Count;
            int nCol = dt.Columns.Count;

            Paragraph p = GetParagraphByReplaceFlag(doc, replaceFlag, null);

            if (p == null)
            {
                return null;
            }

            Table tbl = p.InsertTableBeforeSelf(nRow, nCol);

            int[] aRowNCol = new int[nRow];//每一行有多少列

            #region 数组：每行有多少列
            //目的是由于合并情况的出现会导致每行列数减少，但是合并单元格以后行不会减少
            for (int i = 0; i < aRowNCol.Length; i++)
            {
                aRowNCol[i] = nCol;
            }
            #endregion 数组：每行有多少列

            #region 合并单元格,从右下角往左上角扫
            int nLe;
            int nUp;
            string sCellText;

            for (int j = nCol - 1; j >= 0; j--)
            {
                for (int i = nRow - 1; i >= 0; i--)
                {
                    nUp = 0;
                    nLe = 0;

                    if (dt.Rows[i][j].ToString().ToUpper().Trim() == "UP" || dt.Rows[i][j].ToString().ToUpper().Trim() == "LE")
                    {
                        continue;
                    }
                    else
                    {
                        sCellText = dt.Rows[i][j].ToString().Replace("*10e", "×10{U|").Replace("*10E", "×10{U|");
                        if (sCellText.IndexOf("×10{U|") > 0 && sCellText.IndexOf("}") < 0)
                        {
                            sCellText += "}";
                        }
                        tbl.Rows[i].Cells[j].Paragraphs[0].Append(sCellText);
                    }

                    for (int n = i + 1; n < nRow; n++)
                    {
                        if (dt.Rows[n][j].ToString().ToUpper().Trim() == "UP")
                        {
                            aRowNCol[n] -= 1;
                            nUp++;
                        }
                        else
                        {
                            break;
                        }
                    }

                    for (int m = j + 1; m < aRowNCol[i]; m++)
                    {
                        if (dt.Rows[i][m].ToString().ToUpper().Trim() == "LE")
                        {
                            nLe++;
                        }
                        else
                        {
                            break;
                        }
                    }

                    #region 合并行
                    if (nUp > 0)
                    {
                        try
                        {
                            tbl.MergeCellsInColumn(j, i, i + nUp);
                        }
                        catch (System.ArgumentOutOfRangeException e)
                        {
                            continue;
                        }
                    }
                    #endregion 合并行

                    if (nLe > 0)
                    {
                        aRowNCol[i] -= nLe;
                        try
                        {
                            tbl.Rows[i].MergeCells(j, j + nLe);
                        }
                        catch (System.ArgumentOutOfRangeException e)
                        {
                            continue;
                        }

                        #region 删除合并单元格中的多余回车
                        //for (int z = 0; z < nLe; z++)
                        //{
                        //    tbl.Rows[i].Cells[j].Paragraphs[tbl.Rows[i].Cells[j].Paragraphs.Count - 1].Remove(false);
                        //}
                        #endregion 删除合并单元格中的多余回车

                        if (nUp > 0)
                        {
                            for (int l = i + 1; l <= i + nUp; l++)
                            {
                                try
                                {
                                    tbl.Rows[l].MergeCells(j, j + nLe);
                                }
                                catch (System.ArgumentOutOfRangeException e)
                                {
                                    continue;
                                }
                            }
                        }
                    }
                    #region 科学计数法 & 上下标
                    int nParagraphs = tbl.Rows[i].Cells[j].Paragraphs.Count;//看这个单元格有多少paragraphs
                    for (int iParagraphs = 0; iParagraphs < nParagraphs; iParagraphs++)
                    {
                        String sComment = tbl.Rows[i].Cells[j].Paragraphs[iParagraphs].Text.ToUpper();

                        if (sComment.IndexOf("E") > 0)
                        {
                            string sPre = sComment.Substring(0, sComment.IndexOf("E"));
                            string sZhishu = sComment.Substring(sComment.IndexOf("E") + 1, sComment.Length - sComment.IndexOf("E") - 1);
                            if ((Regex.IsMatch(sPre, @"^\d+\.\d+$") || Regex.IsMatch(sPre, @"^\d+$") || Regex.IsMatch(sPre, @"^[-]+\d+$") || Regex.IsMatch(sPre, @"^[-]+\d+\.\d+$"))
                                    && (Regex.IsMatch(sZhishu, @"^\d+$") || Regex.IsMatch(sZhishu, @"^[-]+\d+$")))
                            {
                                tbl.Rows[i].Cells[j].Paragraphs[iParagraphs].RemoveText(0, sComment.Length);
                                tbl.Rows[i].Cells[j].Paragraphs[iParagraphs].InsertText(Convert.ToDecimal(sPre).ToString() + "×10{U|" + Int32.Parse(sZhishu).ToString() + "}");
                            }
                        }
                    }
                    for (int iParagraphs = 0; iParagraphs < nParagraphs; iParagraphs++)
                    {
                        #region 上标
                        int iBeginUp = -1;
                        Xceed.Document.NET.Formatting formattingUp = new Xceed.Document.NET.Formatting();
                        formattingUp.Script = Script.superscript;

                        while (tbl.Rows[i].Cells[j].Paragraphs[iParagraphs].Text.ToUpper().IndexOf("{U|") > 0)
                        {
                            iBeginUp = tbl.Rows[i].Cells[j].Paragraphs[iParagraphs].Text.ToUpper().IndexOf("{U|");
                            int iEndUp = -1;
                            for (iEndUp = iBeginUp + 3; iEndUp < tbl.Rows[i].Cells[j].Paragraphs[iParagraphs].Text.Length; iEndUp++)
                            {
                                if (tbl.Rows[i].Cells[j].Paragraphs[iParagraphs].Text.Substring(iEndUp, 1) == "}")
                                {
                                    break;
                                }
                            }
                            if ((iBeginUp + 3) != iEndUp)
                            {
                                string strSub = tbl.Rows[i].Cells[j].Paragraphs[iParagraphs].Text.Substring(iBeginUp + 3, iEndUp - iBeginUp - 3);
                                tbl.Rows[i].Cells[j].Paragraphs[iParagraphs].RemoveText(iBeginUp, iEndUp - iBeginUp + 1);
                                tbl.Rows[i].Cells[j].Paragraphs[iParagraphs].InsertText(iBeginUp, strSub, false, formattingUp);
                            }

                        }
                        #endregion 上标

                        #region 下标
                        int iBeginDown = -1;
                        Xceed.Document.NET.Formatting formattingDown = new Xceed.Document.NET.Formatting();
                        formattingDown.Script = Script.subscript;

                        while (tbl.Rows[i].Cells[j].Paragraphs[iParagraphs].Text.ToUpper().IndexOf("{D|") > 0)
                        {
                            iBeginDown = tbl.Rows[i].Cells[j].Paragraphs[iParagraphs].Text.ToUpper().IndexOf("{D|");
                            int iEndDown = -1;
                            for (iEndDown = iBeginDown + 3; iEndDown < tbl.Rows[i].Cells[j].Paragraphs[iParagraphs].Text.Length; iEndDown++)
                            {
                                if (tbl.Rows[i].Cells[j].Paragraphs[iParagraphs].Text.Substring(iEndDown, 1) == "}")
                                {
                                    break;
                                }
                            }
                            if ((iBeginDown + 3) != iEndDown)
                            {
                                string strSub = tbl.Rows[i].Cells[j].Paragraphs[iParagraphs].Text.Substring(iBeginDown + 3, iEndDown - iBeginDown - 3);
                                tbl.Rows[i].Cells[j].Paragraphs[iParagraphs].RemoveText(iBeginDown, iEndDown - iBeginDown + 1);
                                tbl.Rows[i].Cells[j].Paragraphs[iParagraphs].InsertText(iBeginDown, strSub, false, formattingDown);
                            }
                        }
                        #endregion 下标

                        #region μ转换成Times New Roman
                        int iBegin_u = -1;
                        Xceed.Document.NET.Formatting formatting_u = new Xceed.Document.NET.Formatting();
                        formatting_u.FontFamily = new Xceed.Document.NET.Font("Times New Roman");
                        iBegin_u = tbl.Rows[i].Cells[j].Paragraphs[iParagraphs].Text.ToString().IndexOf("µ");
                        while (iBegin_u >= 0)
                        {
                            tbl.Rows[i].Cells[j].Paragraphs[iParagraphs].RemoveText(iBegin_u, 1);
                            tbl.Rows[i].Cells[j].Paragraphs[iParagraphs].InsertText(iBegin_u, "µ", false, formatting_u);
                            iBegin_u = tbl.Rows[i].Cells[j].Paragraphs[iParagraphs].Text.ToString().IndexOf("µ", iBegin_u + 1);
                        }

                        iBegin_u = tbl.Rows[i].Cells[j].Paragraphs[iParagraphs].Text.ToString().IndexOf("μ");
                        while (iBegin_u >= 0)
                        {
                            tbl.Rows[i].Cells[j].Paragraphs[iParagraphs].RemoveText(iBegin_u, 1);
                            tbl.Rows[i].Cells[j].Paragraphs[iParagraphs].InsertText(iBegin_u, "μ", false, formatting_u);
                            iBegin_u = tbl.Rows[i].Cells[j].Paragraphs[iParagraphs].Text.ToString().IndexOf("μ", iBegin_u + 1);
                        }
                        #endregion μ转换成Times New Roman
                    }

                    #endregion 上下标


                }
            }


            #endregion 合并单元格,从右下角往左上角扫

            for (int i = 0; i < tbl.Rows.Count; i++)
            {
                tbl.Rows[i].BreakAcrossPages = false;
            }

            if (!bEmptyParagraph)
            {
                Paragraph p1 = tbl.InsertParagraphAfterSelf("");

            }

            RemoveParagraphByReplaceFlag(doc, replaceFlag);
            //这里会多一个空白行
            //rp.ReplaceText(new StringReplaceTextOptions { SearchValue= replaceFlag, NewValue=String.Empty});

            //这里不能保存，保存要到最后一个步骤操作
            //Save(doc);

            return tbl;
        }


        public Table GenerateTableWithWidth(DocX doc, string strXml, string replaceFlag, bool bEmptyParagraph, string sPercent, string sPagePercent)
        {
            DataTable dt = XMLDeserialize(strXml);
            int nRow = dt.Rows.Count;
            int nCol = dt.Columns.Count;

            Paragraph p = GetParagraphByReplaceFlag(doc, replaceFlag, null);

            if (p == null)
            {
                return null;
            }

            Table tbl = p.InsertTableBeforeSelf(nRow, nCol);

            //生成表格后就设置表格宽度，合并单元格后再设置无法控制表格宽度
            this.SetTableColWidth(doc, tbl, sPercent, sPagePercent);
            tbl.AutoFit = AutoFit.Fixed;

            int[] aRowNCol = new int[nRow];//每一行有多少列

            #region 数组：每行有多少列
            //目的是由于合并情况的出现会导致每行列数减少，但是合并单元格以后行不会减少
            for (int i = 0; i < aRowNCol.Length; i++)
            {
                aRowNCol[i] = nCol;
            }
            #endregion 数组：每行有多少列

            #region 合并单元格,从右下角往左上角扫
            int nLe;
            int nUp;
            string sCellText;

            for (int j = nCol - 1; j >= 0; j--)
            {
                for (int i = nRow - 1; i >= 0; i--)
                {
                    nUp = 0;
                    nLe = 0;

                    if (dt.Rows[i][j].ToString().ToUpper().Trim() == "UP" || dt.Rows[i][j].ToString().ToUpper().Trim() == "LE")
                    {
                        continue;
                    }
                    else
                    {
                        sCellText = dt.Rows[i][j].ToString().Replace("*10e", "×10{U|").Replace("*10E", "×10{U|");
                        if (sCellText.IndexOf("×10{U|") > 0 && sCellText.IndexOf("}") < 0)
                        {
                            sCellText += "}";
                        }
                        tbl.Rows[i].Cells[j].Paragraphs[0].Append(sCellText);
                    }

                    for (int n = i + 1; n < nRow; n++)
                    {
                        if (dt.Rows[n][j].ToString().ToUpper().Trim() == "UP")
                        {
                            aRowNCol[n] -= 1;
                            nUp++;
                        }
                        else
                        {
                            break;
                        }
                    }

                    for (int m = j + 1; m < aRowNCol[i]; m++)
                    {
                        if (dt.Rows[i][m].ToString().ToUpper().Trim() == "LE")
                        {
                            nLe++;
                        }
                        else
                        {
                            break;
                        }
                    }

                    #region 合并行
                    if (nUp > 0)
                    {
                        try
                        {
                            tbl.MergeCellsInColumn(j, i, i + nUp);
                        }
                        catch (System.ArgumentOutOfRangeException e)
                        {
                            continue;
                        }
                    }
                    #endregion 合并行

                    if (nLe > 0)
                    {
                        aRowNCol[i] -= nLe;
                        try
                        {
                            tbl.Rows[i].MergeCells(j, j + nLe);
                        }
                        catch (System.ArgumentOutOfRangeException e)
                        {
                            continue;
                        }

                        #region 删除合并单元格中的多余回车
                        //for (int z = 0; z < nLe; z++)
                        //{
                        //    tbl.Rows[i].Cells[j].Paragraphs[tbl.Rows[i].Cells[j].Paragraphs.Count - 1].Remove(false);
                        //}
                        #endregion 删除合并单元格中的多余回车

                        if (nUp > 0)
                        {
                            for (int l = i + 1; l <= i + nUp; l++)
                            {
                                try
                                {
                                    tbl.Rows[l].MergeCells(j, j + nLe);
                                }
                                catch (System.ArgumentOutOfRangeException e)
                                {
                                    continue;
                                }
                            }
                        }
                    }
                    #region 科学计数法 & 上下标
                    int nParagraphs = tbl.Rows[i].Cells[j].Paragraphs.Count;//看这个单元格有多少paragraphs
                    for (int iParagraphs = 0; iParagraphs < nParagraphs; iParagraphs++)
                    {
                        String sComment = tbl.Rows[i].Cells[j].Paragraphs[iParagraphs].Text.ToUpper();

                        if (sComment.IndexOf("E") > 0)
                        {
                            string sPre = sComment.Substring(0, sComment.IndexOf("E"));
                            string sZhishu = sComment.Substring(sComment.IndexOf("E") + 1, sComment.Length - sComment.IndexOf("E") - 1);
                            if ((Regex.IsMatch(sPre, @"^\d+\.\d+$") || Regex.IsMatch(sPre, @"^\d+$") || Regex.IsMatch(sPre, @"^[-]+\d+$") || Regex.IsMatch(sPre, @"^[-]+\d+\.\d+$"))
                                    && (Regex.IsMatch(sZhishu, @"^\d+$") || Regex.IsMatch(sZhishu, @"^[-]+\d+$")))
                            {
                                tbl.Rows[i].Cells[j].Paragraphs[iParagraphs].RemoveText(0, sComment.Length);
                                tbl.Rows[i].Cells[j].Paragraphs[iParagraphs].InsertText(Convert.ToDecimal(sPre).ToString() + "×10{U|" + Int32.Parse(sZhishu).ToString() + "}");
                            }
                        }
                    }
                    for (int iParagraphs = 0; iParagraphs < nParagraphs; iParagraphs++)
                    {
                        #region 上标
                        int iBeginUp = -1;
                        Xceed.Document.NET.Formatting formattingUp = new Xceed.Document.NET.Formatting();
                        formattingUp.Script = Script.superscript;

                        while (tbl.Rows[i].Cells[j].Paragraphs[iParagraphs].Text.ToUpper().IndexOf("{U|") > 0)
                        {
                            iBeginUp = tbl.Rows[i].Cells[j].Paragraphs[iParagraphs].Text.ToUpper().IndexOf("{U|");
                            int iEndUp = -1;
                            for (iEndUp = iBeginUp + 3; iEndUp < tbl.Rows[i].Cells[j].Paragraphs[iParagraphs].Text.Length; iEndUp++)
                            {
                                if (tbl.Rows[i].Cells[j].Paragraphs[iParagraphs].Text.Substring(iEndUp, 1) == "}")
                                {
                                    break;
                                }
                            }
                            if ((iBeginUp + 3) != iEndUp)
                            {
                                string strSub = tbl.Rows[i].Cells[j].Paragraphs[iParagraphs].Text.Substring(iBeginUp + 3, iEndUp - iBeginUp - 3);
                                tbl.Rows[i].Cells[j].Paragraphs[iParagraphs].RemoveText(iBeginUp, iEndUp - iBeginUp + 1);
                                tbl.Rows[i].Cells[j].Paragraphs[iParagraphs].InsertText(iBeginUp, strSub, false, formattingUp);
                            }

                        }
                        #endregion 上标

                        #region 下标
                        int iBeginDown = -1;
                        Xceed.Document.NET.Formatting formattingDown = new Xceed.Document.NET.Formatting();
                        formattingDown.Script = Script.subscript;

                        while (tbl.Rows[i].Cells[j].Paragraphs[iParagraphs].Text.ToUpper().IndexOf("{D|") > 0)
                        {
                            iBeginDown = tbl.Rows[i].Cells[j].Paragraphs[iParagraphs].Text.ToUpper().IndexOf("{D|");
                            int iEndDown = -1;
                            for (iEndDown = iBeginDown + 3; iEndDown < tbl.Rows[i].Cells[j].Paragraphs[iParagraphs].Text.Length; iEndDown++)
                            {
                                if (tbl.Rows[i].Cells[j].Paragraphs[iParagraphs].Text.Substring(iEndDown, 1) == "}")
                                {
                                    break;
                                }
                            }
                            if ((iBeginDown + 3) != iEndDown)
                            {
                                string strSub = tbl.Rows[i].Cells[j].Paragraphs[iParagraphs].Text.Substring(iBeginDown + 3, iEndDown - iBeginDown - 3);
                                tbl.Rows[i].Cells[j].Paragraphs[iParagraphs].RemoveText(iBeginDown, iEndDown - iBeginDown + 1);
                                tbl.Rows[i].Cells[j].Paragraphs[iParagraphs].InsertText(iBeginDown, strSub, false, formattingDown);
                            }
                        }
                        #endregion 下标

                        #region μ转换成Times New Roman
                        int iBegin_u = -1;
                        Xceed.Document.NET.Formatting formatting_u = new Xceed.Document.NET.Formatting();
                        formatting_u.FontFamily = new Xceed.Document.NET.Font("Times New Roman");
                        iBegin_u = tbl.Rows[i].Cells[j].Paragraphs[iParagraphs].Text.ToString().IndexOf("µ");
                        while (iBegin_u >= 0)
                        {
                            tbl.Rows[i].Cells[j].Paragraphs[iParagraphs].RemoveText(iBegin_u, 1);
                            tbl.Rows[i].Cells[j].Paragraphs[iParagraphs].InsertText(iBegin_u, "µ", false, formatting_u);
                            iBegin_u = tbl.Rows[i].Cells[j].Paragraphs[iParagraphs].Text.ToString().IndexOf("µ", iBegin_u + 1);
                        }

                        iBegin_u = tbl.Rows[i].Cells[j].Paragraphs[iParagraphs].Text.ToString().IndexOf("μ");
                        while (iBegin_u >= 0)
                        {
                            tbl.Rows[i].Cells[j].Paragraphs[iParagraphs].RemoveText(iBegin_u, 1);
                            tbl.Rows[i].Cells[j].Paragraphs[iParagraphs].InsertText(iBegin_u, "μ", false, formatting_u);
                            iBegin_u = tbl.Rows[i].Cells[j].Paragraphs[iParagraphs].Text.ToString().IndexOf("μ", iBegin_u + 1);
                        }
                        #endregion μ转换成Times New Roman
                    }

                    #endregion 上下标


                }
            }


            #endregion 合并单元格,从右下角往左上角扫

            for (int i = 0; i < tbl.Rows.Count; i++)
            {
                tbl.Rows[i].BreakAcrossPages = false;
            }

            if (!bEmptyParagraph)
            {
                Paragraph p1 = tbl.InsertParagraphAfterSelf("");

            }

            RemoveParagraphByReplaceFlag(doc, replaceFlag);

            //这里不能保存，保存要到最后一个步骤操作
            //Save(doc);

            return tbl;
        }
    
        public Table GenerateTable(DocX doc, string strXml, string replaceFlag, bool bEmptyParagraph, bool deleteReplaceFlag)
        {
            DataTable dt = XMLDeserialize(strXml);
            int nRow = dt.Rows.Count;
            int nCol = dt.Columns.Count;

            Paragraph p = GetParagraphByReplaceFlag(doc, replaceFlag, null);

            if (p == null)
            {
                return null;
            }

            Table tbl = p.InsertTableBeforeSelf(nRow, nCol);

            int[] aRowNCol = new int[nRow];//每一行有多少列

            //DataRow dr = dt.NewRow();
            //for (int j = 0; j < nCol; j++)
            //{
            //    dr[j] = "";
            //}
            //dt.Rows.Add(dr);

            #region 数组：每行有多少列
            //目的是由于合并情况的出现会导致每行列数减少，但是合并单元格以后行不会减少
            for (int i = 0; i < aRowNCol.Length; i++)
            {
                aRowNCol[i] = nCol;
            }
            #endregion 数组：每行有多少列

            #region 合并单元格,从右下角往左上角扫
            int nLe;
            int nUp;
            string sCellText;

            for (int j = nCol - 1; j >= 0; j--)
            {
                for (int i = nRow - 1; i >= 0; i--)
                {
                    nUp = 0;
                    nLe = 0;

                    if (dt.Rows[i][j].ToString().ToUpper().Trim() == "UP" || dt.Rows[i][j].ToString().ToUpper().Trim() == "LE")
                    {
                        continue;
                    }
                    else
                    {
                        sCellText = dt.Rows[i][j].ToString().Replace("*10e", "×10{U|").Replace("*10E", "×10{U|");
                        if (sCellText.IndexOf("×10{U|") > 0 && sCellText.IndexOf("}") < 0)
                        {
                            sCellText += "}";
                        }
                        tbl.Rows[i].Cells[j].Paragraphs[0].Append(sCellText);
                    }

                    for (int n = i + 1; n < nRow; n++)
                    {
                        if (dt.Rows[n][j].ToString().ToUpper().Trim() == "UP")
                        {
                            aRowNCol[n] -= 1;
                            nUp++;
                        }
                        else
                        {
                            break;
                        }
                    }

                    for (int m = j + 1; m < aRowNCol[i]; m++)
                    {
                        if (dt.Rows[i][m].ToString().ToUpper().Trim() == "LE")
                        {
                            nLe++;
                        }
                        else
                        {
                            break;
                        }
                    }

                    #region 合并行
                    if (nUp > 0)
                    {
                        try
                        {
                            tbl.MergeCellsInColumn(j, i, i + nUp);
                        }
                        catch (System.ArgumentOutOfRangeException e)
                        {
                            continue;
                        }
                    }
                    #endregion 合并行

                    if (nLe > 0)
                    {
                        aRowNCol[i] -= nLe;
                        try
                        {
                            tbl.Rows[i].MergeCells(j, j + nLe);
                        }
                        catch (System.ArgumentOutOfRangeException e)
                        {
                            continue;
                        }

                        #region 删除合并单元格中的多余回车
                        for (int z = 0; z < nLe; z++)
                        {
                            tbl.Rows[i].Cells[j].Paragraphs[tbl.Rows[i].Cells[j].Paragraphs.Count - 1].Remove(false);
                        }
                        #endregion 删除合并单元格中的多余回车

                        if (nUp > 0)
                        {
                            for (int l = i + 1; l <= i + nUp; l++)
                            {
                                try
                                {
                                    tbl.Rows[l].MergeCells(j, j + nLe);
                                }
                                catch (System.ArgumentOutOfRangeException e)
                                {
                                    continue;
                                }
                            }
                        }
                    }

                    #region 科学计数法 & 上下标
                    int nParagraphs = tbl.Rows[i].Cells[j].Paragraphs.Count;//看这个单元格有多少paragraphs
                    for (int iParagraphs = 0; iParagraphs < nParagraphs; iParagraphs++)
                    {
                        String sComment = tbl.Rows[i].Cells[j].Paragraphs[iParagraphs].Text.ToUpper();

                        if (sComment.IndexOf("E") > 0)
                        {
                            string sPre = sComment.Substring(0, sComment.IndexOf("E"));
                            string sZhishu = sComment.Substring(sComment.IndexOf("E") + 1, sComment.Length - sComment.IndexOf("E") - 1);
                            if ((Regex.IsMatch(sPre, @"^\d+\.\d+$") || Regex.IsMatch(sPre, @"^\d+$") || Regex.IsMatch(sPre, @"^[-]+\d+$") || Regex.IsMatch(sPre, @"^[-]+\d+\.\d+$"))
                                    && (Regex.IsMatch(sZhishu, @"^\d+$") || Regex.IsMatch(sZhishu, @"^[-]+\d+$")))
                            {
                                tbl.Rows[i].Cells[j].Paragraphs[iParagraphs].RemoveText(0, sComment.Length);
                                tbl.Rows[i].Cells[j].Paragraphs[iParagraphs].InsertText(Convert.ToDecimal(sPre).ToString() + "×10{U|" + Int32.Parse(sZhishu).ToString() + "}");
                            }
                        }
                    }
                    for (int iParagraphs = 0; iParagraphs < nParagraphs; iParagraphs++)
                    {
                        #region 上标
                        int iBeginUp = -1;
                        Xceed.Document.NET.Formatting formattingUp = new Xceed.Document.NET.Formatting();
                        formattingUp.Script = Script.superscript;

                        while (tbl.Rows[i].Cells[j].Paragraphs[iParagraphs].Text.ToUpper().IndexOf("{U|") > 0)
                        {
                            iBeginUp = tbl.Rows[i].Cells[j].Paragraphs[iParagraphs].Text.ToUpper().IndexOf("{U|");
                            int iEndUp = -1;
                            for (iEndUp = iBeginUp + 3; iEndUp < tbl.Rows[i].Cells[j].Paragraphs[iParagraphs].Text.Length; iEndUp++)
                            {
                                if (tbl.Rows[i].Cells[j].Paragraphs[iParagraphs].Text.Substring(iEndUp, 1) == "}")
                                {
                                    break;
                                }
                            }
                            if ((iBeginUp + 3) != iEndUp)
                            {
                                string strSub = tbl.Rows[i].Cells[j].Paragraphs[iParagraphs].Text.Substring(iBeginUp + 3, iEndUp - iBeginUp - 3);
                                tbl.Rows[i].Cells[j].Paragraphs[iParagraphs].RemoveText(iBeginUp, iEndUp - iBeginUp + 1);
                                tbl.Rows[i].Cells[j].Paragraphs[iParagraphs].InsertText(iBeginUp, strSub, false, formattingUp);
                            }
                        }
                        #endregion 上标

                        #region 下标
                        int iBeginDown = -1;
                        Xceed.Document.NET.Formatting formattingDown = new Xceed.Document.NET.Formatting();
                        formattingDown.Script = Script.subscript;

                        while (tbl.Rows[i].Cells[j].Paragraphs[iParagraphs].Text.ToUpper().IndexOf("{D|") > 0)
                        {
                            iBeginDown = tbl.Rows[i].Cells[j].Paragraphs[iParagraphs].Text.ToUpper().IndexOf("{D|");
                            int iEndDown = -1;
                            for (iEndDown = iBeginDown + 3; iEndDown < tbl.Rows[i].Cells[j].Paragraphs[iParagraphs].Text.Length; iEndDown++)
                            {
                                if (tbl.Rows[i].Cells[j].Paragraphs[iParagraphs].Text.Substring(iEndDown, 1) == "}")
                                {
                                    break;
                                }
                            }
                            if ((iBeginDown + 3) != iEndDown)
                            {
                                string strSub = tbl.Rows[i].Cells[j].Paragraphs[iParagraphs].Text.Substring(iBeginDown + 3, iEndDown - iBeginDown - 3);
                                tbl.Rows[i].Cells[j].Paragraphs[iParagraphs].RemoveText(iBeginDown, iEndDown - iBeginDown + 1);
                                tbl.Rows[i].Cells[j].Paragraphs[iParagraphs].InsertText(iBeginDown, strSub, false, formattingDown);
                            }
                        }
                        #endregion 下标

                        #region μ转换成Times New Roman
                        int iBegin_u = -1;
                        Xceed.Document.NET.Formatting formatting_u = new Xceed.Document.NET.Formatting();
                        formatting_u.FontFamily = new Xceed.Document.NET.Font("Times New Roman");
                        iBegin_u = tbl.Rows[i].Cells[j].Paragraphs[iParagraphs].Text.ToString().IndexOf("µ");
                        while (iBegin_u >= 0)
                        {
                            tbl.Rows[i].Cells[j].Paragraphs[iParagraphs].RemoveText(iBegin_u, 1);
                            tbl.Rows[i].Cells[j].Paragraphs[iParagraphs].InsertText(iBegin_u, "µ", false, formatting_u);
                            iBegin_u = tbl.Rows[i].Cells[j].Paragraphs[iParagraphs].Text.ToString().IndexOf("µ", iBegin_u + 1);
                        }

                        iBegin_u = tbl.Rows[i].Cells[j].Paragraphs[iParagraphs].Text.ToString().IndexOf("μ");
                        while (iBegin_u >= 0)
                        {
                            tbl.Rows[i].Cells[j].Paragraphs[iParagraphs].RemoveText(iBegin_u, 1);
                            tbl.Rows[i].Cells[j].Paragraphs[iParagraphs].InsertText(iBegin_u, "μ", false, formatting_u);
                            iBegin_u = tbl.Rows[i].Cells[j].Paragraphs[iParagraphs].Text.ToString().IndexOf("μ", iBegin_u + 1);
                        }
                        #endregion μ转换成Times New Roman
                    }
                    #endregion 上下标
                }
            }


            #endregion 合并单元格,从右下角往左上角扫

            #region 合并单元格从上往下扫，算法有问题，在合并行的时候，暂废弃不用
            //int nLe;
            //int nUp;
            //int nDataTableCol;

            //for (int i = 0; i < nRow; i++)
            //{
            //    nDataTableCol = 0;
            //    for (int j = 0; j < aRowNCol[i]; j++)
            //    {
            //        nLe = 0;
            //        nUp = 0;
            //        for (int m = j + 1; m < aRowNCol[i]; m++)
            //        {
            //            if (dt.Rows[i][m].ToString().ToUpper().Trim() == "LE")
            //            {
            //                dt.Rows[i][m] = "";
            //                nLe++;
            //            }
            //            else
            //            {
            //                break;
            //            }
            //        }

            //        if (nLe > 0)
            //        {
            //            aRowNCol[i] -= nLe;
            //            tbl.Rows[i].MergeCells(j, j + nLe);
            //        }

            //        for (int n = i + 1; n < nRow; n++)
            //        {
            //            if (dt.Rows[n][j].ToString().ToUpper().Trim() == "UP")
            //            {
            //                dt.Rows[n][j] = "";
            //                nUp++;
            //                if (nLe > 0)
            //                {
            //                    aRowNCol[n] -= nLe;
            //                    tbl.Rows[n].MergeCells(j, j + nLe);
            //                }
            //            }
            //            else
            //            {
            //                break;
            //            }
            //        }

            //        #region 合并行
            //        if (nUp > 0)
            //        {
            //            tbl.MergeCellsInColumn(j, i, i + nUp);
            //        }
            //        #endregion 合并行

            //        #region 删除合并单元格中的多余回车
            //        for (int nDelEnter = 0; nDelEnter < nLe; nDelEnter++)
            //        {
            //            tbl.Rows[i].Cells[j].Paragraphs[nDelEnter].Remove(false);
            //        }
            //        #endregion 删除合并单元格中的多余回车

            //        tbl.Rows[i].Cells[j].Paragraphs[0].Append(dt.Rows[i][nDataTableCol].ToString());

            //        //#region 上下标
            //        //string curData = tbl.Rows[i].Cells[j].Paragraphs[0].Text.ToUpper();
            //        //for (int ibegin = curData.IndexOf("{UP|"); ibegin < curData.Length; ibegin++)
            //        //{
            //        //    curData = curData.Substring(ibegin);
            //        //    int iEnd = curData.IndexOf('}');

            //        //}
            //        //#endregion 上下标


            //        nDataTableCol += nLe + 1;
            //    }
            //}
            #endregion 合并单元格从上往下扫，算法有问题，在合并列的时候，暂废弃不用

            #region 为产生和AutoFit.Window相同效果
            //AutoFit.Window 似乎有BUG，用此循环来解决AutoFit.Window问题，使表格自适应于doc
            //tbl.AutoFit = AutoFit.Contents;

            //int maxRowNCol = 0;
            //int iMaxRowNCol;
            //for (iMaxRowNCol = 0; iMaxRowNCol < nRow; iMaxRowNCol++)
            //{
            //    if (aRowNCol[iMaxRowNCol] > maxRowNCol)
            //    {
            //        maxRowNCol = aRowNCol[iMaxRowNCol];
            //    }
            //}

            //double avg = doc.PageWidth * 1.0 / maxRowNCol;
            //for (int j = 0; j < maxRowNCol; j++)
            //{
            //    tbl.Rows[iMaxRowNCol - 1].Cells[j].Width = avg;
            //}
            //tbl.AutoFit = AutoFit.Window;
            #endregion 为产生和AutoFit.Window相同效果

            //tbl.Rows[nRow - 1].Height = 0.1;
            //string sTableCellStyleString = "BORDER:TOP:BORDERSTYLE_NONE;BORDER:LEFT:BORDERSTYLE_NONE;BORDER:RIGHT:BORDERSTYLE_NONE;BORDER:BOTTOM:BORDERSTYLE_SINGLE;";
            //for (int j = 0; j < nCol; j++)
            //{
            //    SetTableCellStyle(tbl, nRow - 1, j, sTableCellStyleString);
            //}
            for (int i = 0; i < tbl.Rows.Count; i++)
            {
                tbl.Rows[i].BreakAcrossPages = false;
            }
            if (bEmptyParagraph)
            {
                Paragraph p1 = tbl.InsertParagraphAfterSelf("");

            }
            
            if (deleteReplaceFlag)
            {
                RemoveParagraphByReplaceFlag(doc, replaceFlag);
            }

            //Save(doc);

            return tbl;
        }
        /// <summary>
        /// 设置表格中单元格格式
        /// </summary>
        /// <param name="tbl"></param>
        /// <param name="strCellStyleXml"></param>
        /// <returns></returns>
        public Table SetTableCellFont(Table tbl, int x, int y, string sCellFont = "", double nCellSize = 0, bool bBold = false, bool bItalic = false)
        {
            if (x < tbl.Rows.Count && x >= 0)
            {
                if (y < tbl.Rows[x].Cells.Count && y >= 0)
                {
                    for (int i = 0; i < tbl.Rows[x].Cells[y].Paragraphs.Count; i++)
                    {
                        if (sCellFont != "")
                        {

                            tbl.Rows[x].Cells[y].Paragraphs[i].Font(new Xceed.Document.NET.Font(sCellFont));
                            #region μ转换成Times New Roman
                            int iBegin_u = -1;
                            Xceed.Document.NET.Formatting formatting_u = new Xceed.Document.NET.Formatting();
                            formatting_u.FontFamily = new Xceed.Document.NET.Font("Times New Roman");
                            iBegin_u = tbl.Rows[x].Cells[y].Paragraphs[i].Text.ToString().IndexOf("µ");

                            while (iBegin_u >= 0)
                            {
                                tbl.Rows[x].Cells[y].Paragraphs[i].RemoveText(iBegin_u, 1);
                                tbl.Rows[x].Cells[y].Paragraphs[i].InsertText(iBegin_u, "μ", false, formatting_u);
                                iBegin_u = tbl.Rows[x].Cells[y].Paragraphs[i].Text.ToString().IndexOf("μ", iBegin_u + 1);
                            }

                            iBegin_u = tbl.Rows[x].Cells[y].Paragraphs[i].Text.IndexOf("μ");

                            while (iBegin_u >= 0)
                            {
                                tbl.Rows[x].Cells[y].Paragraphs[i].RemoveText(iBegin_u, 1);
                                tbl.Rows[x].Cells[y].Paragraphs[i].InsertText(iBegin_u, "μ", false, formatting_u);
                                iBegin_u = tbl.Rows[x].Cells[y].Paragraphs[i].Text.ToString().IndexOf("μ", iBegin_u + 1);
                            }

                            #endregion μ转换成Times New Roman
                        }
                        if (nCellSize > 0)
                        {
                            tbl.Rows[x].Cells[y].Paragraphs[i].FontSize(nCellSize);
                        }
                        if (bBold)
                        {
                            tbl.Rows[x].Cells[y].Paragraphs[i].Bold();
                        }
                        if (bItalic)
                        {
                            tbl.Rows[x].Cells[y].Paragraphs[i].Italic();
                        }
                    }
                }
            }

            return tbl;
        }
        /// <summary>
        /// 设置表格中单元格格式
        /// </summary>
        /// <param name="tbl"></param>
        /// <param name="strCellStyleXml"></param>
        /// <returns></returns>
        public Table SetTableCellStyle(Table tbl, int x, int y, string sCellStyleAll)
        {
            if (x >= tbl.Rows.Count || x < 0)
            {
                return tbl;
            }
            else
            {
                if (y >= tbl.Rows[x].Cells.Count || (y != -1 && y < 0))
                {
                    return tbl;
                }
            }
            string[] aCellStyle;//由属性串分割成的属性数组
            string sStyleName = "";
            string sStyleValue = "";

            aCellStyle = sCellStyleAll.Split(';');
            for (int m = 0; m < aCellStyle.Length; m++)
            {
                if (aCellStyle[m].IndexOf(':') < 0)
                {
                    continue;
                }

                //sStyleName = aCellStyle[m].Substring(0, aCellStyle[m].IndexOf(':'));
                //sStyleValue = aCellStyle[m].Substring(aCellStyle[m].IndexOf(':') + 1);
                string[] aCellStyleOne = aCellStyle[m].Split(':');
                sStyleName = aCellStyleOne[0];
                sStyleValue = aCellStyleOne[1];

                BorderStyle BorderStyle_Tcbs = BorderStyle.Tcbs_single;
                //BorderStyle BorderStyle_Tcbs = BorderStyle.Tcbs_dotted;
                BorderSize BorderStyle_Size = BorderSize.one;
                int BorderStyle_Space = 0;
                System.Drawing.Color BorderStyle_Color = System.Drawing.Color.Black;
                TableCellBorderType tblCellBorderType = TableCellBorderType.Left;

                if (aCellStyleOne.Length > 2)
                {
                    for (int p = 2; p < aCellStyleOne.Length; p++)
                    {
                        string sBorderProperty = aCellStyleOne[p].Substring(0, aCellStyleOne[p].IndexOf('_'));
                        string sBorderValue = aCellStyleOne[p].Substring(aCellStyleOne[p].IndexOf('_') + 1);
                        if (sBorderProperty.ToUpper() == "BORDERSTYLE")//调边线样式
                        {
                            if (sBorderValue.ToUpper() == "DOTTED")
                            {
                                BorderStyle_Tcbs = BorderStyle.Tcbs_dotted;
                            }
                            else if (sBorderValue.ToUpper() == "NONE")
                            {
                                BorderStyle_Tcbs = BorderStyle.Tcbs_none;
                            }
                            else if (sBorderValue.ToUpper() == "SINGLE")
                            {
                                BorderStyle_Tcbs = BorderStyle.Tcbs_single;
                            }
                        }
                        else if (sBorderProperty.ToUpper() == "BORDERSIZE")//调边线粗细
                        {
                            if (sBorderValue.ToUpper() == "ONE")
                            {
                                BorderStyle_Size = BorderSize.one;
                            }
                            else if (sBorderValue.ToUpper() == "TWO")
                            {
                                BorderStyle_Size = BorderSize.two;
                            }
                            else if (sBorderValue.ToUpper() == "TWO")
                            {
                                BorderStyle_Size = BorderSize.two;
                            }
                            else if (sBorderValue.ToUpper() == "THREE")
                            {
                                BorderStyle_Size = BorderSize.three;
                            }
                            else if (sBorderValue.ToUpper() == "FOUR")
                            {
                                BorderStyle_Size = BorderSize.four;
                            }
                            else if (sBorderValue.ToUpper() == "FIVE")
                            {
                                BorderStyle_Size = BorderSize.five;
                            }
                            else if (sBorderValue.ToUpper() == "SIX")
                            {
                                BorderStyle_Size = BorderSize.six;
                            }
                            else if (sBorderValue.ToUpper() == "SEVEN")
                            {
                                BorderStyle_Size = BorderSize.seven;
                            }
                            else if (sBorderValue.ToUpper() == "EIGHT")
                            {
                                BorderStyle_Size = BorderSize.eight;
                            }
                            else if (sBorderValue.ToUpper() == "NINE")
                            {
                                BorderStyle_Size = BorderSize.nine;
                            }

                        }
                        else if (sBorderProperty.ToUpper() == "BORDERCOLOR")//调边线颜色
                        {
                            if (sBorderValue.ToUpper() == "RED")
                            {
                                BorderStyle_Color = System.Drawing.Color.Red;
                            }
                            else if (sBorderValue.ToUpper() == "WHITE")
                            {
                                BorderStyle_Color = System.Drawing.Color.White;
                            }
                        }
                    }
                }

                if (sStyleName.ToUpper() == "BORDER")
                {
                    if (sStyleValue.ToUpper() == "LEFT")
                    {
                        tblCellBorderType = TableCellBorderType.Left;
                    }
                    else if (sStyleValue.ToUpper() == "TOP")
                    {
                        tblCellBorderType = TableCellBorderType.Top;
                    }
                    else if (sStyleValue.ToUpper() == "RIGHT")
                    {
                        tblCellBorderType = TableCellBorderType.Right;
                    }
                    else if (sStyleValue.ToUpper() == "BOTTOM")
                    {
                        tblCellBorderType = TableCellBorderType.Bottom;
                    }

                    if (y == -1)
                    {
                        for (int z = 0; z < tbl.Rows[x].Cells.Count; z++)
                        {
                            tbl.Rows[x].Cells[z].SetBorder(tblCellBorderType, new Border(BorderStyle_Tcbs, BorderStyle_Size, BorderStyle_Space, BorderStyle_Color));
                        }
                    }
                    else
                    {
                        tbl.Rows[x].Cells[y].SetBorder(tblCellBorderType, new Border(BorderStyle_Tcbs, BorderStyle_Size, BorderStyle_Space, BorderStyle_Color));
                    }
                }

                if (sStyleName.ToUpper() == "PARAGRAPHALIGN")
                {
                    if (sStyleValue.ToUpper() == "LEFT")
                    {
                        if (y == -1)
                        {
                            for (int z = 0; z < tbl.Rows[x].Cells.Count; z++)
                            {
                                tbl.Rows[x].Cells[z].Paragraphs[0].Alignment = Alignment.left;
                            }
                        }
                        else
                        {
                            tbl.Rows[x].Cells[y].Paragraphs[0].Alignment = Alignment.left;
                        }
                    }
                    else if (sStyleValue.ToUpper() == "RIGHT")
                    {
                        if (y == -1)
                        {
                            for (int z = 0; z < tbl.Rows[x].Cells.Count; z++)
                            {
                                tbl.Rows[x].Cells[z].Paragraphs[0].Alignment = Alignment.right;
                            }
                        }
                        else
                        {
                            tbl.Rows[x].Cells[y].Paragraphs[0].Alignment = Alignment.right;
                        }
                    }
                    else if (sStyleValue.ToUpper() == "CENTER")
                    {
                        if (y == -1)
                        {
                            for (int z = 0; z < tbl.Rows[x].Cells.Count; z++)
                            {
                                tbl.Rows[x].Cells[z].Paragraphs[0].Alignment = Alignment.center;
                            }
                        }
                        else
                        {
                            tbl.Rows[x].Cells[y].Paragraphs[0].Alignment = Alignment.center;
                        }
                    }
                    else if (sStyleValue.ToUpper() == "VTOP")
                    {
                        if (y == -1)
                        {
                            for (int z = 0; z < tbl.Rows[x].Cells.Count; z++)
                            {
                                tbl.Rows[x].Cells[z].VerticalAlignment = VerticalAlignment.Top;
                            }
                        }
                        else
                        {
                            tbl.Rows[x].Cells[y].VerticalAlignment = VerticalAlignment.Top;
                        }
                    }
                    else if (sStyleValue.ToUpper() == "VCENTER")
                    {
                        if (y == -1)
                        {
                            for (int z = 0; z < tbl.Rows[x].Cells.Count; z++)
                            {
                                tbl.Rows[x].Cells[z].VerticalAlignment = VerticalAlignment.Center;
                            }
                        }
                        else
                        {
                            tbl.Rows[x].Cells[y].VerticalAlignment = VerticalAlignment.Center;
                        }
                    }
                    else if (sStyleValue.ToUpper() == "VBOTTOM")
                    {
                        if (y == -1)
                        {
                            for (int z = 0; z < tbl.Rows[x].Cells.Count; z++)
                            {
                                tbl.Rows[x].Cells[z].VerticalAlignment = VerticalAlignment.Bottom;
                            }
                        }
                        else
                        {
                            tbl.Rows[x].Cells[y].VerticalAlignment = VerticalAlignment.Bottom;
                        }
                    }
                    else if (sStyleValue.ToUpper() == "BOTH")
                    {
                        if (y == -1)
                        {
                            for (int z = 0; z < tbl.Rows[x].Cells.Count; z++)
                            {
                                tbl.Rows[x].Cells[z].Paragraphs[0].Alignment = Alignment.both;
                            }
                        }
                        else
                        {
                            tbl.Rows[x].Cells[y].Paragraphs[0].Alignment = Alignment.both;
                        }
                    }
                }
            }

            return tbl;
        }

        /// <summary>
        /// 设置表格边框
        /// </summary>
        /// <param name="doc"></param>
        /// <param name="tbl"></param>
        /// <param name="directions"></param>
        /// <returns></returns>
        public Table SetTableBorderLine(Table tbl, string sDirections)
        {
            if (tbl == null)
            {
                return null;
            }

            string[] aDirections = sDirections.Split(',');
            for (int i = 0; i < aDirections.Length; i++)
            {
                if (aDirections[i].ToUpper() == "LEFT")
                {
                    for (int j = 0; j < tbl.RowCount; j++)
                    {
                        SetTableCellStyle(tbl, j, 0, "BORDER:LEFT;");
                    }
                }

                if (aDirections[i].ToUpper() == "TOP")
                {
                    for (int j = 0; j < tbl.Rows[0].Cells.Count; j++)
                    {
                        SetTableCellStyle(tbl, 0, j, "BORDER:TOP;");
                    }
                }

                if (aDirections[i].ToUpper() == "BOTTOM")
                {
                    for (int j = 0; j < tbl.Rows[tbl.RowCount - 1].Cells.Count; j++)
                    {
                        SetTableCellStyle(tbl, tbl.RowCount - 1, j, "BORDER:BOTTOM;");
                    }
                }

                if (aDirections[i].ToUpper() == "RIGHT")
                {
                    for (int j = 0; j < tbl.RowCount; j++)
                    {
                        SetTableCellStyle(tbl, j, tbl.Rows[j].Cells.Count - 1, "BORDER:RIGHT;");
                    }
                }

                if (aDirections[i].ToUpper() == "INSIDEH")
                {
                    for (int j = 0; j < tbl.RowCount - 1; j++)
                    {
                        for (int k = 0; k < tbl.Rows[j].Cells.Count; k++)
                        {
                            SetTableCellStyle(tbl, j, k, "BORDER:BOTTOM;");
                        }
                    }
                }

                if (aDirections[i].ToUpper() == "INSIDEV")
                {
                    for (int j = 0; j < tbl.RowCount; j++)
                    {
                        for (int k = 0; k < tbl.Rows[j].Cells.Count - 1; k++)
                        {
                            SetTableCellStyle(tbl, j, k, "BORDER:RIGHT;");
                        }
                    }
                }
            }
            return tbl;
        }

        /// <summary>
        /// Xml反序列化为DataTable
        /// </summary>
        /// <param name="strXml"></param>
        /// <returns></returns>
        public DataTable XMLDeserialize(string strXml)
        {
            /*例子：
                <?xml version='1.0' ?>
                <complexType length='2'>
	                <complexType length='3'>
		                <string>1</string>
		                <string>2</string>
		                <string>3</string>
	                </complexType>
	                <complexType length='3'>
		                <string>4</string>
		                <string>5</string>
		                <string>6</string>
	                </complexType>
                </complexType>
             */
            System.Xml.XmlDocument xmlDoc = new System.Xml.XmlDocument();
            xmlDoc.LoadXml(strXml);

            XmlNode xn = xmlDoc.SelectSingleNode("complexType");

            XmlNodeList xnl = xn.ChildNodes;
            string sCol = ((XmlElement)xnl[0]).GetAttribute("length");
            int nCol = int.Parse(sCol);

            DataTable dt = new DataTable();

            for (int i = 0; i < nCol; i++)
            {
                dt.Columns.Add();
            }

            int tmpCol = 0;
            foreach (XmlNode xnf in xnl)
            {
                DataRow dr = dt.NewRow();
                tmpCol = 0;

                XmlElement xe = (XmlElement)xnf;

                XmlNodeList xnf1 = xe.ChildNodes;

                foreach (XmlNode xn2 in xnf1)
                {
                    dr[tmpCol++] = xn2.InnerText;
                }
                dt.Rows.Add(dr);
            }

            return dt;
        }

        /// <summary>
        /// 替换字符串（替换标记）
        /// </summary>
        /// <param name="doc"></param>
        /// <param name="replaceFlag"></param>
        /// <param name="newValue"></param>
        /// <returns></returns>
        public Boolean ReplaceFlag(DocX doc, string replaceFlag, string newValue, string alignment)
        {
            
            Paragraph p = GetParagraphByReplaceFlag(doc, replaceFlag, alignment);
            int nIndex;
            string newReplaceFlag;
            while (p != null)
            {
                try
                {
                    if (newValue.ToUpper().IndexOf("{U|") >= 0 || newValue.ToUpper().IndexOf("{D|") >= 0
                        || newValue.IndexOf("µ") >= 0 || newValue.ToString().IndexOf("μ") >= 0)
                    {
                        nIndex = p.Text.ToString().IndexOf(replaceFlag);
                        p.RemoveText(nIndex, 1, false);
                        p.RemoveText(nIndex + replaceFlag.Length - 2, 1, false);
                        newReplaceFlag = replaceFlag.Replace( "】", "").Replace("【", "");
                        //兼容英文中括号
                        newReplaceFlag = replaceFlag.Replace("]", "").Replace("[", "");

                      
                        p.ReplaceText(new StringReplaceTextOptions{SearchValue= newReplaceFlag,NewValue= newValue });



                          
                    }
                    else
                    {
                        p.ReplaceText(new StringReplaceTextOptions { SearchValue = replaceFlag, NewValue = newValue });
                    }
                    #region 上标
                    int iBeginUp = -1;
                    Xceed.Document.NET.Formatting formattingUp = new Xceed.Document.NET.Formatting();
                    formattingUp.Script = Script.superscript; 
                    while (p.Text.ToUpper().IndexOf("{U|") > 0)
                    {
                        iBeginUp = p.Text.ToUpper().IndexOf("{U|");
                        int iEndUp = -1;
                        for (iEndUp = iBeginUp + 3; iEndUp < p.Text.Length; iEndUp++)
                        {
                            if (p.Text.Substring(iEndUp, 1) == "}")
                            {
                                break;
                            }
                        }
                        if ((iBeginUp + 3) != iEndUp)
                        {
                            string strSub = p.Text.Substring(iBeginUp + 3, iEndUp - iBeginUp - 3);
                            p.RemoveText(iBeginUp, iEndUp - iBeginUp + 1);
                            p.InsertText(iBeginUp, strSub, false, formattingUp);

                        }

                    }
                    #endregion 上标

                    #region 下标
                    int iBeginDown = -1;
                    Xceed.Document.NET.Formatting formattingDown = new Xceed.Document.NET.Formatting();
                    formattingDown.Script = Script.subscript;

                    while (p.Text.ToUpper().IndexOf("{D|") > 0)
                    {
                        iBeginDown = p.Text.ToUpper().IndexOf("{D|");
                        int iEndDown = -1;
                        for (iEndDown = iBeginDown + 3; iEndDown < p.Text.Length; iEndDown++)
                        {
                            if (p.Text.Substring(iEndDown, 1) == "}")
                            {
                                break;
                            }
                        }
                        if ((iBeginDown + 3) != iEndDown)
                        {
                            string strSub = p.Text.Substring(iBeginDown + 3, iEndDown - iBeginDown - 3);
                            p.RemoveText(iBeginDown, iEndDown - iBeginDown + 1);
                            p.InsertText(iBeginDown, strSub, false, formattingDown);
                        }
                    }
                    #endregion 下标

                    #region μ转换成Times New Roman
                    int iBegin_u = -1;
                    Xceed.Document.NET.Formatting formatting_u = new Xceed.Document.NET.Formatting();
                    formatting_u.FontFamily = new Xceed.Document.NET.Font("Times New Roman");
                    iBegin_u = p.Text.ToString().IndexOf("µ");
                    while (iBegin_u >= 0)
                    {
                        p.RemoveText(iBegin_u, 1);
                        p.InsertText(iBegin_u, "µ", false, formatting_u);
                        iBegin_u = p.Text.ToString().IndexOf("µ", iBegin_u + 1);
                    }

                    iBegin_u = p.Text.ToString().IndexOf("μ");
                    while (iBegin_u >= 0)
                    {
                        p.RemoveText(iBegin_u, 1);
                        p.InsertText(iBegin_u, "μ", false, formatting_u);
                        iBegin_u = p.Text.ToString().IndexOf("μ", iBegin_u + 1);
                    }
                    #endregion μ转换成Times New Roman
                }
                catch (System.NullReferenceException e)
                {
                    continue;
                }
                p = GetParagraphByReplaceFlag(doc, replaceFlag, alignment);
            }
            return true;
        }

        /// <summary>
        /// 插入图片（用户自定义图片尺寸）
        /// </summary>
        /// <param name="doc"></param>
        /// <param name="replaceFlag"></param>
        /// <param name="imgPath"></param>
        /// <param name="height"></param>
        /// <param name="width"></param>
        public void InsertPicture(DocX doc, string replaceFlag, string imgPath, string alignment, double height, double width)
        {
            height = height / OldVersionFactor;
            width = width / OldVersionFactor;


            Picture pic = InsertPicture(doc, replaceFlag, imgPath, alignment);

            if (pic == null)
            {
                return;
            }

            if (Convert.ToInt32(height) == 0 && Convert.ToInt32(width) == 0)
            {
                return;
            }
            else if (Convert.ToInt32(height) == 0)
            {
                height = Convert.ToDouble(pic.Height) / Convert.ToDouble(pic.Width) * width;
            }
            else if (Convert.ToInt32(width) == 0)
            {
                width = Convert.ToDouble(pic.Width) / Convert.ToDouble(pic.Height) * height;
            }
            pic.Height = Convert.ToInt32(height);
            pic.Width = Convert.ToInt32(width);

        }

        /// <summary>
        /// 插入图片（对图片尺寸没有要求）
        /// </summary>
        /// <param name="doc"></param>
        /// <param name="replaceFlag"></param>
        /// <param name="imgPath"></param>
        public Picture InsertPicture(DocX doc, string replaceFlag, string imgPath, string alignment)
        {
            Paragraph p = GetParagraphByReplaceFlag(doc, replaceFlag, alignment);

            if (p == null)
            {
                return null;
            }

            p.ReplaceText(replaceFlag, "");

            Xceed.Document.NET.Image img = null;
            try
            {
                img = doc.AddImage(imgPath);
            }
            catch (System.InvalidOperationException e)
            {
                return null;
            }

            Picture pic = img.CreatePicture();

            p.AppendPicture(pic);

            pic.Height = Convert.ToInt32(Convert.ToDouble(pic.Height) / Convert.ToDouble(pic.Width) * Convert.ToDouble(doc.PageWidth - doc.MarginLeft - doc.MarginRight));
            pic.Width = Convert.ToInt32(Convert.ToDouble(doc.PageWidth - doc.MarginLeft - doc.MarginRight));
            return pic;
        }

        /// <summary>
        /// 插入图片,图片放在word中，可以在Word自定义图片的格式(图片尺寸，布局等)
        /// </summary>
        /// <param name="doc"></param>
        /// <param name="replaceFlag"></param>
        /// <param name="imgWordPath"></param>
        public void InsertPicture2(DocX doc, string replaceFlag, string imgWordPath, string alignment)
        {
            Picture pic = null;
            int oriPicCount = doc.Pictures.Count;
            try
            {
                DocX imgDoc = DocX.Load(imgWordPath);
                doc.InsertDocument(imgDoc);
            }
            catch (System.InvalidOperationException e)
            {
                pic = null;
            }
            if (doc.Pictures.Count == oriPicCount + 1)
            {
                pic = doc.Pictures[oriPicCount];
            }

            Paragraph p = GetParagraphByReplaceFlag(doc, replaceFlag, alignment);

            while (p != null)
            {
                if (p == null)
                {
                    break;
                }
                p.ReplaceText(replaceFlag, "");
                if (pic != null)
                {
                    p.AppendPicture(pic);
                }

                p = GetParagraphByReplaceFlag(doc, replaceFlag, alignment);
            }

            if (pic != null)
            {
                pic.Remove();
            }
        }

        /// <summary>
        /// 根据字符串获取字符串所在段落
        /// </summary>
        /// <param name="doc"></param>
        /// <param name="replaceFlag"></param>
        /// <returns></returns>
        public Paragraph GetParagraphByReplaceFlag(DocX doc, string replaceFlag, string alignment)
        {
            List<Paragraph> lstParagraphInHeaderFirst = null;
            List<Paragraph> lstParagraphInHeaderOdd = null;
            List<Paragraph> lstParagraphInHeaderEven = null;
            List<Paragraph> lstParagraphInFooterFirst = null;
            List<Paragraph> lstParagraphInFooterOdd = null;
            List<Paragraph> lstParagraphInFooterEven = null;

            if (doc.Headers.First != null)
            {
                lstParagraphInHeaderFirst = doc.Headers.First.Paragraphs.Where(paragraph => paragraph.Text.Trim().Contains(replaceFlag)).ToList<Paragraph>();
            }
            if (doc.Headers.Odd != null)
            {
                lstParagraphInHeaderOdd = doc.Headers.Odd.Paragraphs.Where(paragraph => paragraph.Text.Trim().Contains(replaceFlag)).ToList<Paragraph>();
            }
            if (doc.Headers.Even != null)
            {
                lstParagraphInHeaderEven = doc.Headers.Even.Paragraphs.Where(paragraph => paragraph.Text.Trim().Contains(replaceFlag)).ToList<Paragraph>();
            }
            if (doc.Footers.First != null)
            {
                lstParagraphInFooterFirst = doc.Footers.First.Paragraphs.Where(paragraph => paragraph.Text.Trim().Contains(replaceFlag)).ToList<Paragraph>();
            }
            if (doc.Footers.Odd != null)
            {
                lstParagraphInFooterOdd = doc.Footers.Odd.Paragraphs.Where(paragraph => paragraph.Text.Trim().Contains(replaceFlag)).ToList<Paragraph>();
            }
            if (doc.Footers.Even != null)
            {
                lstParagraphInFooterEven = doc.Footers.Even.Paragraphs.Where(paragraph => paragraph.Text.Trim().Contains(replaceFlag)).ToList<Paragraph>();
            }


            List<Paragraph> lstParagraph = doc.Paragraphs.Where(paragraph => paragraph.Text.Trim().Contains(replaceFlag)).ToList<Paragraph>();

            Paragraph p = null;
            Boolean bBreakOutOfFor = false;

            if (lstParagraphInHeaderFirst != null && lstParagraphInHeaderFirst.Count != 0)
            {
                p = lstParagraphInHeaderFirst[0];
            }
            else if (lstParagraphInHeaderOdd != null && lstParagraphInHeaderOdd.Count != 0)
            {
                p = lstParagraphInHeaderOdd[0];
            }
            else if (lstParagraphInHeaderEven != null && lstParagraphInHeaderEven.Count != 0)
            {
                p = lstParagraphInHeaderEven[0];
            }
            else if (lstParagraphInFooterFirst != null && lstParagraphInFooterFirst.Count != 0)
            {
                p = lstParagraphInFooterFirst[0];
            }
            else if (lstParagraphInFooterOdd != null && lstParagraphInFooterOdd.Count != 0)
            {
                p = lstParagraphInFooterOdd[0];
            }
            else if (lstParagraphInFooterEven != null && lstParagraphInFooterEven.Count != 0)
            {
                p = lstParagraphInFooterEven[0];
            }
            else if (lstParagraph.Count != 0)
            {
                p = lstParagraph[0];
            }
            else if (doc.Headers.First != null || doc.Headers.Odd != null || doc.Headers.Even != null ||
                    doc.Footers.First != null || doc.Footers.Odd != null || doc.Footers.Even != null)
            {
                List<Table> lstTables = null;
                if (doc.Headers.First != null && doc.Headers.First.Tables.Count != 0)
                {
                    lstTables = doc.Headers.First.Tables;
                }
                else if (doc.Headers.Odd != null && doc.Headers.Odd.Tables.Count != 0)
                {
                    lstTables = doc.Headers.Odd.Tables;
                }
                else if (doc.Headers.Even != null && doc.Headers.Even.Tables.Count != 0)
                {
                    lstTables = doc.Headers.Even.Tables;
                }
                else if (doc.Footers.First != null && doc.Footers.First.Tables.Count != 0)
                {
                    lstTables = doc.Footers.First.Tables;
                }
                else if (doc.Footers.Odd != null && doc.Footers.Odd.Tables.Count != 0)
                {
                    lstTables = doc.Footers.Odd.Tables;
                }
                else if (doc.Footers.Even != null && doc.Footers.Even.Tables.Count != 0)
                {
                    lstTables = doc.Footers.Even.Tables;
                }
                else
                {
                    lstTables = doc.Tables;
                }

                int nTablesInDoc = lstTables.Count;

                for (int i = 0; i < nTablesInDoc; i++)
                {
                    for (int m = 0; m < lstTables[i].RowCount; m++)
                    {
                        for (int n = 0; n < lstTables[i].Rows[m].Cells.Count; n++)
                        {
                            List<Paragraph> lstParagraphInCell = lstTables[i].Rows[m].Cells[n].Paragraphs.Where(paragraph => paragraph.Text.Trim().Contains(replaceFlag)).ToList<Paragraph>();

                            if (lstParagraphInCell.Count != 0)
                            {
                                p = lstParagraphInCell[0];
                                bBreakOutOfFor = true;
                            }

                            if (bBreakOutOfFor)
                            {
                                break;
                            }

                            if ((m == lstTables[i].RowCount - 1) &&
                                (n == lstTables[i].Rows[m].Cells.Count - 1) &&
                                (lstParagraphInCell.Count == 0))
                            {
                                return null;
                            }

                        }

                        if (bBreakOutOfFor)
                        {
                            break;
                        }
                    }

                    if (bBreakOutOfFor)
                    {
                        break;
                    }
                }
            }

            if (alignment != null && alignment.Length > 0 && p != null)
            {
                if (alignment.ToUpper() != "LEFT" && alignment.ToUpper() != "RIGHT" && alignment.ToUpper() != "CENTER" && alignment.ToUpper() != "BOTH")
                {
                    p.Alignment = Alignment.left;
                }

                if (alignment.ToUpper() == "LEFT")
                {
                    p.Alignment = Alignment.left;
                }
                else if (alignment.ToUpper() == "RIGHT")
                {
                    p.Alignment = Alignment.right;
                }
                else if (alignment.ToUpper() == "CENTER")
                {
                    p.Alignment = Alignment.center;
                }
                else
                {
                    p.Alignment = Alignment.both;
                }
            }
           
            
            return p;
        }

        /// <summary>
        /// 设置表格列宽（页边距内）
        /// </summary>
        /// <param name="doc"></param>
        /// <param name="tbl"></param>
        /// <param name="sPercent">各列所占表格的比例</param>
        /// <returns></returns>
        public Table SetTableColWidth(DocX doc, Table tbl, string sPercent)
        {
            string[] aPercent = sPercent.Split(',');
            if (tbl.ColumnCount != aPercent.Length)
            {
                return tbl;
            }
            double sum_aPercent = 0.0;
            for (int i = 0; i < aPercent.Length; i++)
            {
                sum_aPercent += Convert.ToDouble(aPercent[i]);
            }
            float[] colWidths = new float[aPercent.Length];
            for (int i = 0; i < aPercent.Length; i++)
            {
                colWidths[i] = (float)(Convert.ToDouble(aPercent[i]) / sum_aPercent * doc.PageWidth);
                //colWidths[i] = (float)(Convert.ToDouble(aPercent[i]) / 100 * doc.PageWidth);

                colWidths[i] = (float)(colWidths[i] / OldVersionFactor);
            }
            tbl.SetWidths(colWidths);

            return tbl;
        }
        /// <summary>
        /// 设置表格列宽（页面比例）
        /// </summary>
        /// <param name="doc"></param>
        /// <param name="tbl"></param>
        /// <param name="sPercent">各列所占表格的百分比</param>
        /// <param name="sPagePercent">表格所占页面的百分比</param>
        /// <returns></returns>
        public Table SetTableColWidth(DocX doc, Table tbl, string sPercent, string sPagePercent)
        {
            string[] aPercent = sPercent.Split(',');
            if (tbl.ColumnCount != aPercent.Length)
            {
                return tbl;
            }
            double nPagePercent = Convert.ToDouble(sPagePercent) / 100;
            double sum_aPercent = 0.0;
            for (int i = 0; i < aPercent.Length; i++)
            {
                sum_aPercent += Convert.ToDouble(aPercent[i]);
            }
            float[] colWidths = new float[aPercent.Length];
            for (int i = 0; i < aPercent.Length; i++)
            {
                colWidths[i] = (float)(Convert.ToDouble(aPercent[i]) / sum_aPercent * nPagePercent * doc.PageWidth);

                colWidths[i] = (float)(colWidths[i] / OldVersionFactor);
            }
            tbl.SetWidths(colWidths);

            tbl.Alignment = Alignment.center;
            tbl.AutoFit = AutoFit.ColumnWidth;

            return tbl;
        }

        /// <summary>
        /// 设置表格行高（cm）
        /// </summary>
        /// <param name="doc"></param>
        /// <param name="tbl"></param>
        /// <param name="RowNum">行号</param>
        /// <param name="Height">行高（cm）</param>
        /// <param name="bOverSet">新值比原来的行高高时才赋值</param>
        /// <returns></returns>
        public Table SetTableRowHeight(DocX doc, Table tbl, int nRowCount, double dHeight, bool bOverSet = false)
        {
            if (nRowCount < -1)
            {
                return tbl;
            }
            if (nRowCount == -1)
            {
                for (int i = 0; i < tbl.Rows.Count; i++)
                {
                    if (!bOverSet)
                    {
                        tbl.Rows[i].MinHeight = dHeight / 0.0265;
                    }
                    else
                    {
                        if (tbl.Rows[i].Height * 0.0265 < dHeight)
                        {
                            tbl.Rows[i].MinHeight = dHeight / 0.0265;
                        }

                    }
                }
                
            }
            else
            {
                if (!bOverSet)
                {
                    tbl.Rows[nRowCount].MinHeight = dHeight / 0.0265;
                }
                else
                {
                    if (tbl.Rows[nRowCount].Height * 0.0265 < dHeight)
                    {
                        tbl.Rows[nRowCount].MinHeight = dHeight / 0.0265;
                    }
                }
            }
            return tbl;
        }
        /// <summary>
        /// document合并
        /// </summary>
        /// <param name="oldDocument"></param>
        /// <param name="newDocument"></param>
        /// <returns></returns>
        public DocX DocUnitAsOne(DocX oldDocument, DocX newDocument)
        {
            oldDocument.InsertDocument(newDocument);
            oldDocument.Save();
            return oldDocument;
        }

        /// <summary>
        /// 在Table前面或者后面插入分页符
        /// </summary>
        /// <param name="doc"></param>
        /// <param name="tbl"></param>
        /// <param name="Location">在table的前面还是后面插入换页符</param>
        public void TableInsertPageBreak(DocX doc, Table tbl, string Location)
        {
            if (Location.ToUpper() == "BEFORE")
            {
                tbl.InsertPageBreakBeforeSelf();
            }
            else if (Location.ToUpper() == "AFTER")
            {
                tbl.InsertPageBreakAfterSelf();
            }
            doc.Save();
        }

        /// <summary>
        /// 设置表格重复标题行
        /// </summary>
        /// <param name="tbl">表格</param>
        /// <param name="nRow">第几行</param>
        public void SetTableHeader(Table tbl, int nRow)
        {
            tbl.Rows[nRow].TableHeader = true;
        }


        /// <summary>
        /// 删除替换标签和所在的行
        /// </summary>
        /// <param name="doc"></param>
        /// <param name="replaceFlag"></param>
        public void RemoveParagraphByReplaceFlag(DocX doc,string replaceFlag)
        {
            Paragraph rp = GetParagraphByReplaceFlag(doc, replaceFlag, "LEFT");
            rp.Remove(false);
        }



        #region DEMO=======================================================================================================================================

        #region Charts

        public class ChartData
        {
            public String Mounth { get; set; }
            public Double Money { get; set; }

            public static List<ChartData> CreateCompanyList1()
            {
                List<ChartData> company1 = new List<ChartData>();
                company1.Add(new ChartData() { Mounth = "January", Money = 100 });
                company1.Add(new ChartData() { Mounth = "February", Money = 120 });
                company1.Add(new ChartData() { Mounth = "March", Money = 140 });
                return company1;
            }

            public static List<ChartData> CreateCompanyList2()
            {
                List<ChartData> company2 = new List<ChartData>();
                company2.Add(new ChartData() { Mounth = "January", Money = 80 });
                company2.Add(new ChartData() { Mounth = "February", Money = 160 });
                company2.Add(new ChartData() { Mounth = "March", Money = 130 });
                return company2;
            }
        }

        public void BarChart(DocX document)
        {
            // Create chart.
            BarChart c = new BarChart();
            c.BarDirection = BarDirection.Column;
            c.BarGrouping = BarGrouping.Standard;
            c.GapWidth = 400;
            c.AddLegend(ChartLegendPosition.Bottom, false);

            // Create data.
            List<ChartData> company1 = ChartData.CreateCompanyList1();
            List<ChartData> company2 = ChartData.CreateCompanyList2();

            // Create and add series
            Series s1 = new Series("Microsoft");
            s1.Color = Color.GreenYellow;
            s1.Bind(company1, "Mounth", "Money");
            c.AddSeries(s1);
            Series s2 = new Series("Apple");
            s2.Bind(company2, "Mounth", "Money");
            c.AddSeries(s2);

            // Insert chart into document
            document.InsertParagraph("Diagram").FontSize(20);
            document.InsertChart(c);
            document.Save();
        }

        public void PieChart(DocX document)
        {
            // Create chart.
            PieChart c = new PieChart();
            c.AddLegend(ChartLegendPosition.Bottom, false);

            // Create data.
            List<ChartData> company2 = ChartData.CreateCompanyList2();

            // Create and add series
            Series s = new Series("Apple");
            s.Bind(company2, "Mounth", "Money");
            c.AddSeries(s);

            // Insert chart into document
            document.InsertParagraph("Diagram").FontSize(20);
            document.InsertChart(c);
            document.Save();
        }

        public void LineChart(DocX document)
        {
            // Create chart.
            LineChart c = new LineChart();
            c.AddLegend(ChartLegendPosition.Bottom, false);

            // Create data.
            List<ChartData> company1 = ChartData.CreateCompanyList1();
            List<ChartData> company2 = ChartData.CreateCompanyList2();

            // Create and add series
            Series s1 = new Series("Microsoft");
            s1.Color = Color.GreenYellow;
            s1.Bind(company1, "Mounth", "Money");
            c.AddSeries(s1);
            Series s2 = new Series("Apple");
            s2.Bind(company2, "Mounth", "Money");
            c.AddSeries(s2);

            // Insert chart into document
            document.InsertParagraph("Diagram").FontSize(20);
            document.InsertChart(c);
            document.Save();
        }

        public void Chart3D(DocX document)
        {
            // Create chart.
            BarChart c = new BarChart();
            c.View3D = true;

            // Create data.
            List<ChartData> company1 = ChartData.CreateCompanyList1();

            // Create and add series
            Series s = new Series("Microsoft");
            s.Color = Color.GreenYellow;
            s.Bind(company1, "Mounth", "Money");
            c.AddSeries(s);

            // Insert chart into document
            document.InsertParagraph("3D Diagram").FontSize(20);
            document.InsertChart(c);
            document.Save();
        }

        #endregion Charts

        /// <summary>
        /// Create a document with two equations.
        /// </summary>
        public void Equations(DocX document)
        {
            // Insert first Equation in this document.
            Paragraph pEquation1 = document.InsertEquation("x = y+z");

            // Insert second Equation in this document and add formatting.
            Paragraph pEquation2 = document.InsertEquation("x = (y+z)/t").FontSize(18).Color(Color.Blue);

            // Save this document to disk.
            document.Save();
        }

        public void DocumentHeading(DocX document)
        {
            foreach (HeadingType heading in (HeadingType[])Enum.GetValues(typeof(HeadingType)))
            {
                string text = string.Format("{0} - The quick brown fox jumps over the lazy dog", heading.EnumDescription());

                Paragraph p = document.InsertParagraph();
                p.AppendLine(text).Heading(heading);
            }

            document.Save();
        }

        public void Bookmarks(DocX document)
        {
            var paragraph = document.InsertBookmark("firstBookmark");

            var paragraph2 = document.InsertParagraph("This is a paragraph which contains a ");
            paragraph2.AppendBookmark("secondBookmark");
            paragraph2.Append("bookmark");

            paragraph2.InsertAtBookmark("handy ", "secondBookmark");

            document.Save();
        }

        /// <summary>
        /// Create a document with a Paragraph whos first line is indented.
        /// </summary>
        public void Indentation(DocX document)
        {
            // Create a new Paragraph.
            Paragraph p = document.InsertParagraph("Line 1\nLine 2\nLine 3");

            // Indent only the first line of the Paragraph.
            p.IndentationFirstLine = 1.0f;

            // Save all changes made to this document.
            document.Save();
        }

        /// <summary>
        /// Create a document that with RightToLeft text flow.
        /// </summary>
        public void RightToLeft(DocX document)
        {
            // Create a new Paragraph with the text "Hello World".
            Paragraph p = document.InsertParagraph("Hello World.");

            // Make this Paragraph flow right to left. Default is left to right.
            p.Direction = Direction.RightToLeft;

            // You don't need to manually set the text direction foreach Paragraph, you can just call this function.
            document.SetDirection(Direction.LeftToRight);

            // Save all changes made to this document.
            document.Save();
        }

        /// <summary>
        /// Creates a document with a Hyperlink, an Image and a Table.
        /// </summary>
        public void HyperlinksImagesTables(DocX document)
        {
            // Add a hyperlink into the document.
            Hyperlink link = document.AddHyperlink("link", new Uri("http://www.google.com"));

            // Add a Table into the document.
            Table table = document.AddTable(2, 2);
            table.Design = TableDesign.ColorfulGridAccent2;
            table.Alignment = Alignment.center;
            table.Rows[0].Cells[0].Paragraphs[0].Append("1");
            table.Rows[0].Cells[1].Paragraphs[0].Append("2");
            table.Rows[1].Cells[0].Paragraphs[0].Append("3");
            table.Rows[1].Cells[1].Paragraphs[0].Append("4");

            Row newRow = table.InsertRow(table.Rows[1]);
            newRow.ReplaceText("4", "5");

            // Add an image into the document.
            Xceed.Document.NET.Image image = document.AddImage("c:\\pig.png");

            //Create a picture (A custom view of an Image).
            Picture picture = image.CreatePicture();
            picture.Rotation = 10;
            picture.SetPictureShape(BasicShapes.cube);

            // Insert a new Paragraph into the document.
            Paragraph title = document.InsertParagraph().Append("Test").FontSize(20).Font(new Xceed.Document.NET.Font("Comic Sans MS"));
            title.Alignment = Alignment.center;

            // Insert a new Paragraph into the document.
            Paragraph p1 = document.InsertParagraph();

            // Append content to the Paragraph
            p1.AppendLine("This line contains a ").Append("bold").Bold().Append(" word.");
            p1.AppendLine("Here is a cool ").AppendHyperlink(link).Append(".");
            p1.AppendLine();
            //p1.AppendLine("Check out this picture ").AppendPicture(picture).Append(" its funky don't you think?");
            p1.AppendLine();
            p1.AppendLine("Can you check this Table of figures for me?");
            p1.AppendLine();

            // Insert the Table after Paragraph 1.
            p1.InsertTableAfterSelf(table);

            // Insert a new Paragraph into the document.
            Paragraph p2 = document.InsertParagraph();

            // Append content to the Paragraph.
            p2.AppendLine("Is it correct?");

            //Save this document.
            document.Save();
        }

        public void HyperlinksImagesTablesWithLists(DocX document)
        {
            // Add a hyperlink into the document.
            Hyperlink link = document.AddHyperlink("link", new Uri("http://www.google.com"));

            // created numbered lists 
            var numberedList = document.AddList("First List Item.", 0, ListItemType.Numbered, 1);
            document.AddListItem(numberedList, "First sub list item", 1);
            document.AddListItem(numberedList, "Second List Item.");
            document.AddListItem(numberedList, "Third list item.");
            document.AddListItem(numberedList, "Nested item.", 1);
            document.AddListItem(numberedList, "Second nested item.", 1);

            // created bulleted lists
            var bulletedList = document.AddList("First Bulleted Item.", 0, ListItemType.Bulleted);
            document.AddListItem(bulletedList, "Second bullet item");
            document.AddListItem(bulletedList, "Sub bullet item", 1);
            document.AddListItem(bulletedList, "Second sub bullet item", 1);
            document.AddListItem(bulletedList, "Third bullet item");


            // Add a Table into the document.
            Table table = document.AddTable(2, 2);
            table.Design = TableDesign.ColorfulGridAccent2;
            table.Alignment = Alignment.center;
            table.Rows[0].Cells[0].Paragraphs[0].Append("1");
            table.Rows[0].Cells[1].Paragraphs[0].Append("2");
            table.Rows[1].Cells[0].Paragraphs[0].Append("3");
            table.Rows[1].Cells[1].Paragraphs[0].Append("4");

            Row newRow = table.InsertRow(table.Rows[1]);
            newRow.ReplaceText("4", "5");

            // Add an image into the document.
            Xceed.Document.NET.Image image = document.AddImage(@"C:\pig.jpeg");

            // Create a picture (A custom view of an Image).
            Picture picture = image.CreatePicture();
            picture.Rotation = 10;
            picture.SetPictureShape(BasicShapes.cube);

            // Insert a new Paragraph into the document.
            Paragraph title = document.InsertParagraph().Append("Test").FontSize(20).Font(new Xceed.Document.NET.Font("Comic Sans MS"));
            title.Alignment = Alignment.center;

            // Insert a new Paragraph into the document.
            Paragraph p1 = document.InsertParagraph();

            // Append content to the Paragraph
            p1.AppendLine("This line contains a ").Append("bold").Bold().Append(" word.");
            p1.AppendLine("Here is a cool ").AppendHyperlink(link).Append(".");
            p1.AppendLine();
            p1.AppendLine("Check out this picture ").AppendPicture(picture).Append(" its funky don't you think?");
            p1.AppendLine();
            p1.AppendLine("Can you check this Table of figures for me?");
            p1.AppendLine();

            // Insert the Table after Paragraph 1.
            p1.InsertTableAfterSelf(table);

            // Insert a new Paragraph into the document.
            Paragraph p2 = document.InsertParagraph();

            // Append content to the Paragraph.
            p2.AppendLine("Is it correct?");
            p2.AppendLine();
            p2.AppendLine("Adding bullet list below: ");

            document.InsertList(bulletedList);

            // Adding another paragraph to add table and bullet list after it
            Paragraph p3 = document.InsertParagraph();
            p3.AppendLine();
            p3.AppendLine("Adding another table...");

            // Adding another table
            Table table1 = document.AddTable(2, 2);
            table1.Design = TableDesign.ColorfulGridAccent2;
            table1.Alignment = Alignment.center;
            table1.Rows[0].Cells[0].Paragraphs[0].Append("1");
            table1.Rows[0].Cells[1].Paragraphs[0].Append("2");
            table1.Rows[1].Cells[0].Paragraphs[0].Append("3");
            table1.Rows[1].Cells[1].Paragraphs[0].Append("4");

            Paragraph p4 = document.InsertParagraph();
            p4.InsertTableBeforeSelf(table1);

            p4.AppendLine();

            // Insert numbered list after table
            Paragraph p5 = document.InsertParagraph();
            p5.AppendLine("Adding numbered list below: ");
            p5.AppendLine();
            document.InsertList(numberedList);

            // Save this document.
            document.Save();
        }

        public void DocumentMargins(DocX document)
        {
            // Create a float var that contains doc Margins properties.
            float leftMargin = document.MarginLeft;
            float rightMargin = document.MarginRight;
            float topMargin = document.MarginTop;
            float bottomMargin = document.MarginBottom;

            // Modify using your own vars.
            leftMargin = 95F;
            rightMargin = 45F;
            topMargin = 50F;
            bottomMargin = 180F;

            // Or simply work the margins by setting the property directly. 
            document.MarginLeft = leftMargin;
            document.MarginRight = rightMargin;
            document.MarginTop = topMargin;
            document.MarginBottom = bottomMargin;

            // created bulleted lists

            var bulletedList = document.AddList("First Bulleted Item.", 0, ListItemType.Bulleted);
            document.AddListItem(bulletedList, "Second bullet item");
            document.AddListItem(bulletedList, "Sub bullet item", 1);
            document.AddListItem(bulletedList, "Second sub bullet item", 1);
            document.AddListItem(bulletedList, "Third bullet item");

            document.InsertList(bulletedList);

            // Save this document.
            document.Save();
        }

        public void DocumentsWithListsFontChange(DocX document)
        {
            foreach (FontFamily oneFontFamily in FontFamily.Families)
            {
                FontFamily fontFamily = oneFontFamily;
                double fontSize = 15;

                // created numbered lists 
                var numberedList = document.AddList("First List Item.", 0, ListItemType.Numbered, 1);
                document.AddListItem(numberedList, "First sub list item", 1);
                document.AddListItem(numberedList, "Second List Item.");
                document.AddListItem(numberedList, "Third list item.");
                document.AddListItem(numberedList, "Nested item.", 1);
                document.AddListItem(numberedList, "Second nested item.", 1);

                // created bulleted lists
                var bulletedList = document.AddList("First Bulleted Item.", 0, ListItemType.Bulleted);
                document.AddListItem(bulletedList, "Second bullet item");
                document.AddListItem(bulletedList, "Sub bullet item", 1);
                document.AddListItem(bulletedList, "Second sub bullet item", 1);
                document.AddListItem(bulletedList, "Third bullet item");

                document.InsertList(bulletedList);


                document.InsertList(numberedList, new Xceed.Document.NET.Font(fontFamily.Name), fontSize);

            }

            // Save this document.
            document.Save();
        }

        public void AddList(DocX document)
        {
            var numberedList = document.AddList("First List Item.", 0, ListItemType.Numbered, 2);
            document.AddListItem(numberedList, "First sub list item", 1);
            document.AddListItem(numberedList, "Second List Item.");
            document.AddListItem(numberedList, "Third list item.");
            document.AddListItem(numberedList, "Nested item.", 1);
            document.AddListItem(numberedList, "Second nested item.", 1);

            var bulletedList = document.AddList("First Bulleted Item.", 0, ListItemType.Bulleted);
            document.AddListItem(bulletedList, "Second bullet item");
            document.AddListItem(bulletedList, "Sub bullet item", 1);
            document.AddListItem(bulletedList, "Second sub bullet item", 1);
            document.AddListItem(bulletedList, "Third bullet item");

            document.InsertList(numberedList);
            document.InsertList(bulletedList);
            document.Save();
        }

        public void HeadersAndFooters(DocX document)
        {
            // Add Headers and Footers to this document.
            document.AddHeaders();
            document.AddFooters();

            // Force the first page to have a different Header and Footer.
            document.DifferentFirstPage = true;

            // Force odd & even pages to have different Headers and Footers.
            document.DifferentOddAndEvenPages = true;

            // Get the first, odd and even Headers for this document.
            Header header_first = document.Headers.First;
            Header header_odd = document.Headers.Odd;
            Header header_even = document.Headers.Even;

            // Get the first, odd and even Footer for this document.
            Footer footer_first = document.Footers.First;
            Footer footer_odd = document.Footers.Odd;
            Footer footer_even = document.Footers.Even;

            // Insert a Paragraph into the first Header.
            Paragraph p0 = header_first.InsertParagraph();
            p0.Append("Hello First Header.").Bold();

            // Insert a Paragraph into the odd Header.
            Paragraph p1 = header_odd.InsertParagraph();
            p1.Append("Hello Odd Header.").Bold();

            // Insert a Paragraph into the even Header.
            Paragraph p2 = header_even.InsertParagraph();
            p2.Append("Hello Even Header.").Bold();

            // Insert a Paragraph into the first Footer.
            Paragraph p3 = footer_first.InsertParagraph();
            p3.Append("Hello First Footer.").Bold();

            // Insert a Paragraph into the odd Footer.
            Paragraph p4 = footer_odd.InsertParagraph();
            p4.Append("Hello Odd Footer.").Bold();

            // Insert a Paragraph into the even Footer.
            Paragraph p5 = footer_even.InsertParagraph();
            p5.Append("Hello Even Footer.").Bold();

            // Insert a Paragraph into the document.
            Paragraph p6 = document.InsertParagraph();
            p6.AppendLine("Hello First page.");

            // Create a second page to show that the first page has its own header and footer.
            p6.InsertPageBreakAfterSelf();

            // Insert a Paragraph after the page break.
            Paragraph p7 = document.InsertParagraph();
            p7.AppendLine("Hello Second page.");

            // Create a third page to show that even and odd pages have different headers and footers.
            p7.InsertPageBreakAfterSelf();

            // Insert a Paragraph after the page break.
            Paragraph p8 = document.InsertParagraph();
            p8.AppendLine("Hello Third page.");

            //Insert a next page break, which is a section break combined with a page break
            document.InsertSectionPageBreak();

            //Insert a paragraph after the "Next" page break
            Paragraph p9 = document.InsertParagraph();
            p9.Append("Next page section break.");

            //Insert a continuous section break
            document.InsertSection();

            //Create a paragraph in the new section
            var p10 = document.InsertParagraph();
            p10.Append("Continuous section paragraph.");

            // Save all changes to this document.
            document.Save();

        }

        public void HeadersAndFootersWithImagesAndTables(DocX document)
        {
            // Add a template logo image to this document.
            Xceed.Document.NET.Image logo = document.AddImage(@"C:\pig.jpeg");
            document.InsertParagraph().AppendPicture(logo.CreatePicture());

            // Add Headers and Footers to this document.
            document.AddHeaders();
            document.AddFooters();

            // Force the first page to have a different Header and Footer.
            document.DifferentFirstPage = true;

            // Force odd & even pages to have different Headers and Footers.
            document.DifferentOddAndEvenPages = true;

            // Get the first, odd and even Headers for this document.
            Header header_first = document.Headers.First;
            Header header_odd = document.Headers.Odd;
            Header header_even = document.Headers.Even;

            // Get the first, odd and even Footer for this document.
            Footer footer_first = document.Footers.First;
            Footer footer_odd = document.Footers.Odd;
            Footer footer_even = document.Footers.Even;

            // Insert a Paragraph into the first Header.
            Paragraph p0 = header_first.InsertParagraph();
            p0.Append("Hello First Header.").Bold();

            // Insert a Paragraph into the odd Header.
            Paragraph p1 = header_odd.InsertParagraph();
            p1.Append("Hello Odd Header.").Bold();

            // Insert a Paragraph into the even Header.
            Paragraph p2 = header_even.InsertParagraph();
            p2.Append("Hello Even Header.").Bold();

            // Insert a Paragraph into the first Footer.
            Paragraph p3 = footer_first.InsertParagraph();
            p3.Append("Hello First Footer.").Bold();

            // Insert a Paragraph into the odd Footer.
            Paragraph p4 = footer_odd.InsertParagraph();
            p4.Append("Hello Odd Footer.").Bold();

            // Insert a Paragraph into the even Footer.
            Paragraph p5 = footer_even.InsertParagraph();
            p5.Append("Hello Even Footer.").Bold();

            // Insert a Paragraph into the document.
            Paragraph p6 = document.InsertParagraph();
            p6.AppendLine("Hello First page.");

            // Create a second page to show that the first page has its own header and footer.
            p6.InsertPageBreakAfterSelf();

            // Insert a Paragraph after the page break.
            Paragraph p7 = document.InsertParagraph();
            p7.AppendLine("Hello Second page.");

            // Create a third page to show that even and odd pages have different headers and footers.
            p7.InsertPageBreakAfterSelf();

            // Insert a Paragraph after the page break.
            Paragraph p8 = document.InsertParagraph();
            p8.AppendLine("Hello Third page.");

            //Insert a next page break, which is a section break combined with a page break
            document.InsertSectionPageBreak();

            //Insert a paragraph after the "Next" page break
            Paragraph p9 = document.InsertParagraph();
            p9.Append("Next page section break.");

            //Insert a continuous section break
            document.InsertSection();

            //Create a paragraph in the new section
            var p10 = document.InsertParagraph();
            p10.Append("Continuous section paragraph.");


            // Inserting logo into footer and header into Tables


            #region Company Logo in Header in Table
            // Insert Table into First Header - Create a new Table with 2 columns and 1 rows.
            Table header_first_table = header_first.InsertTable(1, 3);
            header_first_table.Design = TableDesign.TableGrid;
            header_first_table.AutoFit = AutoFit.Contents;
            header_first_table.AutoFit = AutoFit.Window;
            // Get the upper right Paragraph in the layout_table.
            Paragraph upperRightParagraph = header_first.Tables[0].Rows[0].Cells[1].Paragraphs[0];
            // Insert this template logo into the upper right Paragraph of Table.
            upperRightParagraph.AppendPicture(logo.CreatePicture());
            header_first.Tables[0].Rows[0].Cells[2].Paragraphs[0].AppendPicture(logo.CreatePicture());
            upperRightParagraph.Alignment = Alignment.right;

            // Get the upper left Paragraph in the layout_table.
            Paragraph upperLeftParagraphFirstTable = header_first.Tables[0].Rows[0].Cells[0].Paragraphs[0];
            upperLeftParagraphFirstTable.Append("Company Name - DocX Corporation");
            #endregion


            #region Company Logo in Header in Invisible Table
            // Insert Table into First Header - Create a new Table with 2 columns and 1 rows.
            Table header_second_table = header_odd.InsertTable(1, 2);
            header_second_table.Design = TableDesign.None;
            header_second_table.AutoFit = AutoFit.Window;
            // Get the upper right Paragraph in the layout_table.
            Paragraph upperRightParagraphSecondTable = header_second_table.Rows[0].Cells[1].Paragraphs[0];
            // Insert this template logo into the upper right Paragraph of Table.
            upperRightParagraphSecondTable.AppendPicture(logo.CreatePicture());
            upperRightParagraphSecondTable.Alignment = Alignment.right;

            // Get the upper left Paragraph in the layout_table.
            Paragraph upperLeftParagraphSecondTable = header_second_table.Rows[0].Cells[0].Paragraphs[0];
            upperLeftParagraphSecondTable.Append("Company Name - DocX Corporation");
            #endregion

            #region Company Logo in Footer in Table
            // Insert Table into First Header - Create a new Table with 2 columns and 1 rows.
            Table footer_first_table = footer_first.InsertTable(1, 2);
            footer_first_table.Design = TableDesign.TableGrid;
            footer_first_table.AutoFit = AutoFit.Window;
            // Get the upper right Paragraph in the layout_table.
            Paragraph upperRightParagraphFooterParagraph = footer_first.Tables[0].Rows[0].Cells[1].Paragraphs[0];
            // Insert this template logo into the upper right Paragraph of Table.
            upperRightParagraphFooterParagraph.AppendPicture(logo.CreatePicture());
            upperRightParagraphFooterParagraph.Alignment = Alignment.right;

            // Get the upper left Paragraph in the layout_table.
            Paragraph upperLeftParagraphFirstTableFooter = footer_first.Tables[0].Rows[0].Cells[0].Paragraphs[0];
            upperLeftParagraphFirstTableFooter.Append("Company Name - DocX Corporation");
            #endregion

            #region Company Logo in Header in Invisible Table
            // Insert Table into First Header - Create a new Table with 2 columns and 1 rows.
            Table footer_second_table = footer_odd.InsertTable(1, 2);
            footer_second_table.Design = TableDesign.None;
            footer_second_table.AutoFit = AutoFit.Window;
            // Get the upper right Paragraph in the layout_table.
            Paragraph upperRightParagraphSecondTableFooter = footer_second_table.Rows[0].Cells[1].Paragraphs[0];
            // Insert this template logo into the upper right Paragraph of Table.
            upperRightParagraphSecondTableFooter.AppendPicture(logo.CreatePicture());
            upperRightParagraphSecondTableFooter.Alignment = Alignment.right;

            // Get the upper left Paragraph in the layout_table.
            Paragraph upperLeftParagraphSecondTableFooter = footer_second_table.Rows[0].Cells[0].Paragraphs[0];
            upperLeftParagraphSecondTableFooter.Append("Company Name - DocX Corporation");
            #endregion

            // Save all changes to this document.
            document.Save();
        }

        public void CreateInvoice()
        {
            DocX g_document;

            try
            {
                // Store a global reference to the loaded document.
                g_document = DocX.Load(@"C:\InvoiceTemplate.docx");

                /*
                 * The template 'InvoiceTemplate.docx' does exist, 
                 * so lets use it to create an invoice for a factitious company
                 * called "The Happy Builder" and store a global reference it.
                 */
                g_document = CreateInvoiceFromTemplate(DocX.Load(@"C:\InvoiceTemplate.docx"));

                // Save all changes made to this template as Invoice_The_Happy_Builder.docx (We don't want to replace InvoiceTemplate.docx).
                g_document.SaveAs(@"C:\Invoice_The_Happy_Builder.docx");
            }
            // The template 'InvoiceTemplate.docx' does not exist, so create it.
            catch (FileNotFoundException)
            {
                // Create and store a global reference to the template 'InvoiceTemplate.docx'.
                g_document = CreateInvoiceTemplate();

                // Save the template 'InvoiceTemplate.docx'.
                g_document.Save();

                // The template exists now so re-call CreateInvoice().
                CreateInvoice();
            }
        }

        private static DocX CreateInvoiceTemplate()
        {
            // Create a new document.
            DocX document = DocX.Create(@"C:\InvoiceTemplate.docx");

            // Create a table for layout purposes (This table will be invisible).
            Table layout_table = document.InsertTable(2, 2);
            layout_table.Design = TableDesign.TableNormal;
            layout_table.AutoFit = AutoFit.Window;

            // Dark formatting
            Xceed.Document.NET.Formatting dark_formatting = new Xceed.Document.NET.Formatting();
            dark_formatting.Bold = true;
            dark_formatting.Size = 12;
            dark_formatting.FontColor = Color.FromArgb(31, 73, 125);

            // Light formatting
            Xceed.Document.NET.Formatting light_formatting = new Xceed.Document.NET.Formatting();
            light_formatting.Italic = true;
            light_formatting.Size = 11;
            light_formatting.FontColor = Color.FromArgb(79, 129, 189);

            #region Company Name
            // Get the upper left Paragraph in the layout_table.
            Paragraph upper_left_paragraph = layout_table.Rows[0].Cells[0].Paragraphs[0];

            // Create a custom property called company_name
            CustomProperty company_name = new CustomProperty("company_name", "Company Name");

            // Insert a field of type doc property (This will display the custom property 'company_name')
            layout_table.Rows[0].Cells[0].Paragraphs[0].InsertDocProperty(company_name, f: dark_formatting);

            // Force the next text insert to be on a new line.
            upper_left_paragraph.InsertText("\n", false);
            #endregion

            #region Company Slogan
            // Create a custom property called company_slogan
            CustomProperty company_slogan = new CustomProperty("company_slogan", "Company slogan goes here.");

            // Insert a field of type doc property (This will display the custom property 'company_slogan')
            upper_left_paragraph.InsertDocProperty(company_slogan, f: light_formatting);
            #endregion

            #region Company Logo
            // Get the upper right Paragraph in the layout_table.
            Paragraph upper_right_paragraph = layout_table.Rows[0].Cells[1].Paragraphs[0];

            // Add a template logo image to this document.
            Xceed.Document.NET.Image logo = document.AddImage(@"C:\pig.jpeg");

            // Insert this template logo into the upper right Paragraph.
            upper_right_paragraph.InsertPicture(logo.CreatePicture());

            upper_right_paragraph.Alignment = Alignment.right;
            #endregion

            // Custom properties cannot contain newlines, so the company address must be split into 3 custom properties.
            #region Hired Company Address
            // Create a custom property called company_address_line_one
            CustomProperty hired_company_address_line_one = new CustomProperty("hired_company_address_line_one", "Street Address,");

            // Get the lower left Paragraph in the layout_table. 
            Paragraph lower_left_paragraph = layout_table.Rows[1].Cells[0].Paragraphs[0];
            lower_left_paragraph.InsertText("TO:\n", false, dark_formatting);

            // Insert a field of type doc property (This will display the custom property 'hired_company_address_line_one')
            lower_left_paragraph.InsertDocProperty(hired_company_address_line_one, f: light_formatting);

            // Force the next text insert to be on a new line.
            lower_left_paragraph.InsertText("\n", false);

            // Create a custom property called company_address_line_two
            CustomProperty hired_company_address_line_two = new CustomProperty("hired_company_address_line_two", "City,");

            // Insert a field of type doc property (This will display the custom property 'hired_company_address_line_two')
            lower_left_paragraph.InsertDocProperty(hired_company_address_line_two, f: light_formatting);

            // Force the next text insert to be on a new line.
            lower_left_paragraph.InsertText("\n", false);

            // Create a custom property called company_address_line_two
            CustomProperty hired_company_address_line_three = new CustomProperty("hired_company_address_line_three", "Zip Code");

            // Insert a field of type doc property (This will display the custom property 'hired_company_address_line_three')
            lower_left_paragraph.InsertDocProperty(hired_company_address_line_three, f: light_formatting);
            #endregion

            #region Date & Invoice number
            // Get the lower right Paragraph from the layout table.
            Paragraph lower_right_paragraph = layout_table.Rows[1].Cells[1].Paragraphs[0];

            CustomProperty invoice_date = new CustomProperty("invoice_date", DateTime.Today.Date.ToString("d"));
            lower_right_paragraph.InsertText("Date: ", false, dark_formatting);
            lower_right_paragraph.InsertDocProperty(invoice_date, f: light_formatting);

            CustomProperty invoice_number = new CustomProperty("invoice_number", 1);
            lower_right_paragraph.InsertText("\nInvoice: ", false, dark_formatting);
            lower_right_paragraph.InsertText("#", false, light_formatting);
            lower_right_paragraph.InsertDocProperty(invoice_number, f: light_formatting);

            lower_right_paragraph.Alignment = Alignment.right;
            #endregion

            // Insert an empty Paragraph between two Tables, so that they do not touch.
            document.InsertParagraph(string.Empty, false);

            // This table will hold all of the invoice data.
            Table invoice_table = document.InsertTable(4, 4);
            invoice_table.Design = TableDesign.LightShadingAccent1;
            invoice_table.Alignment = Alignment.center;

            // A nice thank you Paragraph.
            Paragraph thankyou = document.InsertParagraph("\nThank you for your business, we hope to work with you again soon.", false, dark_formatting);
            thankyou.Alignment = Alignment.center;

            #region Hired company details
            CustomProperty hired_company_details_line_one = new CustomProperty("hired_company_details_line_one", "Street Address, City, ZIP Code");
            CustomProperty hired_company_details_line_two = new CustomProperty("hired_company_details_line_two", "Phone: 000-000-0000, Fax: 000-000-0000, e-mail: support@companyname.com");

            Paragraph companyDetails = document.InsertParagraph(string.Empty, false);
            companyDetails.InsertDocProperty(hired_company_details_line_one, f: light_formatting);
            companyDetails.InsertText("\n", false);
            companyDetails.InsertDocProperty(hired_company_details_line_two, f: light_formatting);
            companyDetails.Alignment = Alignment.center;
            #endregion

            // Return the document now that it has been created.
            return document;
        }

        // Create an invoice for a factitious company called "The Happy Builder".
        private static DocX CreateInvoiceFromTemplate(DocX template)
        {
            #region Logo
            // A quick glance at the template shows us that the logo Paragraph is in row zero cell 1.
            Paragraph logo_paragraph = template.Tables[0].Rows[0].Cells[1].Paragraphs[0];
            // Remove the template Picture that is in this Paragraph.
            logo_paragraph.Pictures[0].Remove();

            // Add the Happy Builders logo to this document.
            Xceed.Document.NET.Image logo = template.AddImage(@"C:\pig.jpeg");

            // Insert the Happy Builders logo into this Paragraph.
            logo_paragraph.InsertPicture(logo.CreatePicture());
            #endregion

            #region Set CustomProperty values
            // Set the value of the custom property 'company_name'.
            template.AddCustomProperty(new CustomProperty("company_name", "The Happy Builder"));

            // Set the value of the custom property 'company_slogan'.
            template.AddCustomProperty(new CustomProperty("company_slogan", "No job too small"));

            // Set the value of the custom properties 'hired_company_address_line_one', 'hired_company_address_line_two' and 'hired_company_address_line_three'.
            template.AddCustomProperty(new CustomProperty("hired_company_address_line_one", "The Crooked House,"));
            template.AddCustomProperty(new CustomProperty("hired_company_address_line_two", "Dublin,"));
            template.AddCustomProperty(new CustomProperty("hired_company_address_line_three", "12345"));

            // Set the value of the custom property 'invoice_date'.
            template.AddCustomProperty(new CustomProperty("invoice_date", DateTime.Today.Date.ToString("d")));

            // Set the value of the custom property 'invoice_number'.
            template.AddCustomProperty(new CustomProperty("invoice_number", 1));

            // Set the value of the custom property 'hired_company_details_line_one' and 'hired_company_details_line_two'.
            template.AddCustomProperty(new CustomProperty("hired_company_details_line_one", "Business Street, Dublin, 12345"));
            template.AddCustomProperty(new CustomProperty("hired_company_details_line_two", "Phone: 012-345-6789, Fax: 012-345-6789, e-mail: support@thehappybuilder.com"));
            #endregion

            /* 
             * InvoiceTemplate.docx contains a blank Table, 
             * we want to replace this with a new Table that
             * contains all of our invoice data.
             */
            Table t = template.Tables[1];
            Table invoice_table = CreateAndInsertInvoiceTableAfter(t, ref template);
            t.Remove();

            // Return the template now that it has been modified to hold all of our custom data.
            return template;
        }

        private static Table CreateAndInsertInvoiceTableAfter(Table t, ref DocX document)
        {
            // Grab data from somewhere (Most likely a database)
            DataTable data = GetDataFromDatabase();

            /* 
             * The trick to replacing one Table with another,
             * is to insert the new Table after the old one, 
             * and then remove the old one.
             */
            Table invoice_table = t.InsertTableAfterSelf(data.Rows.Count + 1, data.Columns.Count);
            invoice_table.Design = TableDesign.LightShadingAccent1;

            #region Table title
            Xceed.Document.NET.Formatting table_title = new Xceed.Document.NET.Formatting();
            table_title.Bold = true;

            invoice_table.Rows[0].Cells[0].Paragraphs[0].InsertText("Description", false, table_title);
            invoice_table.Rows[0].Cells[0].Paragraphs[0].Alignment = Alignment.center;
            invoice_table.Rows[0].Cells[1].Paragraphs[0].InsertText("Hours", false, table_title);
            invoice_table.Rows[0].Cells[1].Paragraphs[0].Alignment = Alignment.center;
            invoice_table.Rows[0].Cells[2].Paragraphs[0].InsertText("Rate", false, table_title);
            invoice_table.Rows[0].Cells[2].Paragraphs[0].Alignment = Alignment.center;
            invoice_table.Rows[0].Cells[3].Paragraphs[0].InsertText("Amount", false, table_title);
            invoice_table.Rows[0].Cells[3].Paragraphs[0].Alignment = Alignment.center;
            #endregion

            // Loop through the rows in the Table and insert data from the data source.
            for (int row = 1; row < invoice_table.RowCount; row++)
            {
                for (int cell = 0; cell < invoice_table.Rows[row].Cells.Count; cell++)
                {
                    Paragraph cell_paragraph = invoice_table.Rows[row].Cells[cell].Paragraphs[0];
                    cell_paragraph.InsertText(data.Rows[row - 1].ItemArray[cell].ToString(), false);
                }
            }

            // We want to fill in the total by suming the values from the amount column.
            Row total = invoice_table.InsertRow();
            total.Cells[0].Paragraphs[0].InsertText("Total:", false);
            Paragraph total_paragraph = total.Cells[invoice_table.ColumnCount - 1].Paragraphs[0];

            /* 
             * Lots of people are scared of LINQ,
             * so I will walk you through this line by line.
             * 
             * invoice_table.Rows is an IEnumerable<Row> (i.e a collection of rows), with LINQ you can query collections.
             * .Where(condition) is a filter that you want to apply to the items of this collection. 
             * My condition is that the index of the row must be greater than 0 and less than RowCount.
             * .Select(something) lets you select something from each item in the filtered collection.
             * I am selecting the Text value from each row, for example €100, then I am remove the €, 
             * and then I am parsing the remaining string as a double. This will return a collection of doubles,
             * the final thing I do is call .Sum() on this collection which return one double the sum of all the doubles,
             * this is the total.
             */
            double totalCost =
            (
                invoice_table.Rows
                .Where((row, index) => index > 0 && index < invoice_table.RowCount - 1)
                .Select(row => double.Parse(row.Cells[row.Cells.Count() - 1].Paragraphs[0].Text.Remove(0, 1)))
            ).Sum();

            // Insert the total calculated above using LINQ into the total Paragraph.
            total_paragraph.InsertText(string.Format("€{0}", totalCost), false);

            // Let the tables columns expand to fit its contents.
            invoice_table.AutoFit = AutoFit.Contents;

            // Center the Table
            invoice_table.Alignment = Alignment.center;

            // Return the invloce table now that it has been created.
            return invoice_table;
        }

        // You need to rewrite this function to grab data from your data source.
        private static DataTable GetDataFromDatabase()
        {
            DataTable table = new DataTable();
            table.Columns.AddRange(new DataColumn[] { new DataColumn("Description"), new DataColumn("Hours"), new DataColumn("Rate"), new DataColumn("Amount") });

            table.Rows.Add
            (
                "Install wooden doors (Kitchen, Sitting room, Dining room & Bedrooms)",
                "5",
                "€25",
                string.Format("€{0}", 5 * 25)
            );

            table.Rows.Add
            (
                "Fit stairs",
                "20",
                "€30",
                string.Format("€{0}", 20 * 30)
            );

            table.Rows.Add
            (
                "Replace Sitting room window",
                "6",
                "€50",
                string.Format("€{0}", 6 * 50)
            );

            table.Rows.Add
            (
                "Build garden shed",
                "10",
                "€10",
                string.Format("€{0}", 10 * 10)
            );

            table.Rows.Add
             (
                 "Fit new lock on back door",
                 "0.5",
                 "€30",
                 string.Format("€{0}", 0.5 * 30)
             );

            table.Rows.Add
             (
                 "Tile Kitchen floor",
                 "24",
                 "€25",
                 string.Format("€{0}", 24 * 25)
             );

            return table;
        }

        public void HelloWorldProtectedDocument(DocX document)
        {
            // Insert a Paragraph into this document.
            Paragraph p = document.InsertParagraph();

            // Append some text and add formatting.
            p.Append("Hello World!^011Hello World!")
            .Font(new Xceed.Document.NET.Font("Times New Roman"))
            .FontSize(32)
            .Color(Color.Blue)
            .Bold();

            // Save this document to disk with different options
            // Protected with password for Read Only
            EditRestrictions erReadOnly = EditRestrictions.readOnly;
            document.AddProtection(erReadOnly);
            document.SaveAs(@"C:\HelloWorldPasswordProtectedReadOnly.docx");

            // Protected with password for Comments
            EditRestrictions erComments = EditRestrictions.comments;
            document.AddProtection(erComments);
            document.SaveAs(@"C:\HelloWorldPasswordProtectedCommentsOnly.docx");

            // Protected with password for Forms
            EditRestrictions erForms = EditRestrictions.forms;
            document.AddProtection(erForms);
            document.SaveAs(@"C:\HelloWorldPasswordProtectedFormsOnly.docx");

            // Protected with password for Tracked Changes
            EditRestrictions erTrackedChanges = EditRestrictions.trackedChanges;
            document.AddProtection(erTrackedChanges);
            document.SaveAs(@"C:\HelloWorldPasswordProtectedTrackedChangesOnly.docx");

            // But it's also possible to add restrictions without protecting it with password.

            // Protected with password for Read Only
            document.AddProtection(erReadOnly);
            document.SaveAs(@"C:\HelloWorldWithoutPasswordReadOnly.docx");

            // Protected with password for Comments
            document.AddProtection(erComments);
            document.SaveAs(@"C:\HelloWorldWithoutPasswordCommentsOnly.docx");

            // Protected with password for Forms
            document.AddProtection(erForms);
            document.SaveAs(@"C:\HelloWorldWithoutPasswordFormsOnly.docx");

            // Protected with password for Tracked Changes
            document.AddProtection(erTrackedChanges);
            document.SaveAs(@"C:\HelloWorldWithoutPasswordTrackedChangesOnly.docx");
        }

        /// <summary>
        /// Loads a document 'Input.docx' and writes the text 'Hello World' into the first imbedded Image.
        /// This code creates the file 'Output.docx'.
        /// </summary>
        public void ProgrammaticallyManipulateImbeddedImage(DocX document)
        {
            const string str = "Hello World";

            // Make sure this document has at least one Image.
            if (document.Images.Count() > 0)
            {
                Xceed.Document.NET.Image img = document.Images[0];

                // Write "Hello World" into this Image.
                Bitmap b = new Bitmap(img.GetStream(FileMode.Open, FileAccess.ReadWrite));

                /*
                * Get the Graphics object for this Bitmap.
                * The Graphics object provides functions for drawing.
                */
                Graphics g = Graphics.FromImage(b);

                // Draw the string "Hello World".
                g.DrawString
                (
                    str,
                    new System.Drawing.Font("Tahoma", 20),
                    Brushes.Blue,
                    new PointF(0, 0)
                );

                // Save this Bitmap back into the document using a Create\Write stream.
                b.Save(img.GetStream(FileMode.Create, FileAccess.Write), ImageFormat.Png);
            }
            else
            {
                Console.WriteLine("The provided document contains no Images.");
            }

            // Save this document as Output.docx.
            document.SaveAs(@"C:\Output.docx");
        }

        public void AddToc(DocX document)
        {
            document.InsertTableOfContents("I can haz table of contentz", TableOfContentsSwitches.O | TableOfContentsSwitches.U | TableOfContentsSwitches.Z | TableOfContentsSwitches.H, "Heading2");
            var h1 = document.InsertParagraph("Heading 1");
            h1.StyleName = "Heading1";
            document.InsertParagraph("Some very interesting content here");
            var h2 = document.InsertParagraph("Heading 2");
            document.InsertSectionPageBreak();
            h2.StyleName = "Heading1";
            document.InsertParagraph("Some very interesting content here as well");
            var h3 = document.InsertParagraph("Heading 2.1");
            h3.StyleName = "Heading2";
            document.InsertParagraph("Not so very interesting....");

            document.Save();
        }

        public void HelloWorldKeepWithNext(DocX document)
        {
            // Create a new Paragraph with the text "Hello World".
            Paragraph p = document.InsertParagraph("Hello World.");
            //p.KeepWithNext();
            p.KeepWithNextParagraph();
            document.InsertParagraph("Previous paragraph will appear on the same page as this paragraph");

            // Save all changes made to this document.
            document.Save();
        }

        public void HelloWorldKeepLinesTogether(DocX document)
        {
            // Create a new Paragraph with the text "Hello World".
            Paragraph p = document.InsertParagraph("All lines of this paragraph will appear on the same page...\nLine 2\nLine 3\nLine 4\nLine 5\nLine 6...");
            p.KeepLinesTogether();
            // Save all changes made to this document.
            document.Save();
        }

        private static Border BlankBorder = new Border(BorderStyle.Tcbs_dashed, 0, 0, Color.Red);

        public void LargeTable(DocX doc)
        {
            var tbl = doc.InsertTable(1, 18);

            var wholeWidth = doc.PageWidth - doc.MarginLeft - doc.MarginRight;
            var colWidth = wholeWidth / tbl.ColumnCount;
            var colWidths = new int[tbl.ColumnCount];
            tbl.AutoFit = AutoFit.Contents;
            var r = tbl.Rows[0];
            var cx = 0;
            foreach (var cell in r.Cells)
            {
                cell.Paragraphs.First().Append("Col " + cx);
                cell.MarginBottom = 0;
                cell.MarginLeft = 0;
                cell.MarginRight = 0;
                cell.MarginTop = 0;

                cx++;
            }
            tbl.SetBorder(TableBorderType.Bottom, BlankBorder);
            tbl.SetBorder(TableBorderType.Left, BlankBorder);
            tbl.SetBorder(TableBorderType.Right, BlankBorder);
            tbl.SetBorder(TableBorderType.Top, BlankBorder);
            tbl.SetBorder(TableBorderType.InsideV, BlankBorder);
            tbl.SetBorder(TableBorderType.InsideH, BlankBorder);

            doc.Save();
        }

        public void TableWithSpecifiedWidths(DocX doc)
        {
            var widths = new float[] { 600f, 500f, 400f };
            var tbl = doc.InsertTable(1, widths.Length);
            tbl.SetWidths(widths);
            var wholeWidth = doc.PageWidth - doc.MarginLeft - doc.MarginRight;
            tbl.AutoFit = AutoFit.Contents;
            var r = tbl.Rows[0];
            var cx = 0;
            foreach (var cell in r.Cells)
            {
                cell.Paragraphs.First().Append("Col " + cx);
                //cell.Width = colWidth;
                cell.MarginBottom = 0;
                cell.MarginLeft = 0;
                cell.MarginRight = 0;
                cell.MarginTop = 0;

                cx++;
            }
            //add new rows 
            for (var x = 0; x < 5; x++)
            {
                r = tbl.InsertRow();
                cx = 0;
                foreach (var cell in r.Cells)
                {
                    cell.Paragraphs.First().Append("Col " + cx);
                    //cell.Width = colWidth;
                    cell.MarginBottom = 0;
                    cell.MarginLeft = 0;
                    cell.MarginRight = 0;
                    cell.MarginTop = 0;

                    cx++;
                }
            }
            tbl.SetBorder(TableBorderType.Bottom, BlankBorder);
            tbl.SetBorder(TableBorderType.Left, BlankBorder);
            tbl.SetBorder(TableBorderType.Right, BlankBorder);
            tbl.SetBorder(TableBorderType.Top, BlankBorder);
            tbl.SetBorder(TableBorderType.InsideV, BlankBorder);
            tbl.SetBorder(TableBorderType.InsideH, BlankBorder);

            doc.Save();
        }
        public void InsertTableRowsAndCopyFlag(Table tbl, int nRow)
        {
            for (int i = 0; i < nRow; i++)
            {
                Row r0 = tbl.Rows[tbl.RowCount - 1];
                tbl.InsertRow(r0);
                for (int j = 0; j < tbl.Rows[tbl.RowCount - 1].ColumnCount; j++)
                {
                    string s1 = tbl.Rows[tbl.RowCount - i - 2].Cells[j].Paragraphs[0].Text.ToString();
                    if (s1 != "" && s1.IndexOf("【#") >= 0 && s1.IndexOf("】") > 0)
                    {
                        tbl.Rows[tbl.RowCount - 1].Cells[j].Paragraphs[0].RemoveText(0, false);
                        tbl.Rows[tbl.RowCount - 1].Cells[j].Paragraphs[0].InsertText(0, s1 + "[" + i.ToString() + "]", false);
                    }
                }

            }
        }


        #endregion DEMO=======================================================================================================================================

    }
}