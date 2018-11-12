using NPOI.OpenXmlFormats.Dml.WordProcessing;
using NPOI.XWPF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Reflection;

namespace Ade.OfficeService.Word
{
    public static class WordHelper
    {
        /// <summary>
        /// 替换Word中的占位符
        /// </summary>
        /// <param name="doc">word</param>
        /// <param name="placeHolderAndValueDict">占位符：值 字典</param>
        /// <param name="hook">钩子委托</param>
        public static void ReplacePlaceHolderInWord(XWPFDocument doc, Dictionary<string, string> placeHolderAndValueDict, List<AddPictureOptions> listAddPictureOptions, List<string> listAllPlaceHolder)
        {
            IEnumerator<XWPFParagraph> paragraphEnumerator = doc.GetParagraphsEnumerator();
            XWPFParagraph paragraph;
            while (paragraphEnumerator.MoveNext())
            {
                paragraph = paragraphEnumerator.Current;
                ReplacePlaceHolderInParagraph(paragraph, placeHolderAndValueDict);
                AddPicture(paragraph, listAddPictureOptions);
            }

            IEnumerator<XWPFTable> tableEnumerator = doc.GetTablesEnumerator();
            XWPFTable table;
            while (tableEnumerator.MoveNext())
            {
                table = tableEnumerator.Current;
                ReplacePlaceHolderInTable(table, placeHolderAndValueDict);
                AddPictureInTable(table, listAddPictureOptions);
            }

            ClearNotUsedPlaceHolder(listAllPlaceHolder, doc); 
        }

        /// <summary>
        /// 清空所有未使用的占位符用空字符串
        /// </summary>
        /// <param name="doc"></param>
        /// <param name="listAllPlaceHolder"></param>
        private static void ClearNotUsedPlaceHolder(List<string> listAllPlaceHolder, XWPFDocument doc)
        {
            Dictionary<string, string> placeHolderAndValueDict = new Dictionary<string, string>();
            foreach (var placeHolder in listAllPlaceHolder)
            {
                placeHolderAndValueDict.Add(placeHolder, string.Empty);
            }

            IEnumerator<XWPFParagraph> paragraphEnumerator = doc.GetParagraphsEnumerator();
            XWPFParagraph paragraph;
            while (paragraphEnumerator.MoveNext())
            {
                paragraph = paragraphEnumerator.Current;
                ReplacePlaceHolderInParagraph(paragraph, placeHolderAndValueDict);
            }

            IEnumerator<XWPFTable> tableEnumerator = doc.GetTablesEnumerator();
            XWPFTable table;
            while (tableEnumerator.MoveNext())
            {
                table = tableEnumerator.Current;
                ReplacePlaceHolderInTable(table, placeHolderAndValueDict);
            }
        }

        private static void DoAddPicture(XWPFParagraph paragraph, List<AddPictureOptions> listAddPictureOptions)
        {
            XWPFRun textRun, imgRun, placeHolderRun;
            IList<XWPFRun> listRun = paragraph.Runs;
            for (int i = 0; i < listRun.Count; i++)
            {
                placeHolderRun = paragraph.Runs[i];

                for (int j = 0; j < listAddPictureOptions.Count; j++)
                {
                    var addPictureOptions = listAddPictureOptions[j];
                    if (addPictureOptions.PlaceHolder.ToString() == placeHolderRun.Text)
                    {
                        if (j == 0)
                        {
                            paragraph.RemoveRun(i);
                        }

                        textRun = paragraph.CreateRun();

                        string newText = addPictureOptions.PictureName + Environment.NewLine;
                        if (j > 0)
                        {
                            newText = Environment.NewLine + "    " + newText;
                        }

                        //textRun.SetText(newText);
                        textRun.SetText(string.Empty);

                        CopyRunStyle(placeHolderRun, textRun);

                        imgRun = paragraph.CreateRun();

                        if (File.Exists(addPictureOptions.LocalPictureUrl))
                        {
                            using (FileStream fs = File.OpenRead(addPictureOptions.LocalPictureUrl))
                            {
                                imgRun.AddPicture(
                                    fs,
                                    (int)GetPictureTypeFromUrl(addPictureOptions.LocalPictureUrl),
                                    addPictureOptions.ImageType.ToString() + Path.GetExtension(addPictureOptions.LocalPictureUrl),
                                    5000000,
                                    3000000);
                                CT_Inline inline = imgRun.GetCTR().GetDrawingList()[0].inline[0];
                                inline.docPr.id = addPictureOptions.PicId; //id必须从1开始
                            }
                        }
                    }
                }
            }
        }

        private static PictureType GetPictureTypeFromUrl(string localFileUrl)
        {
            PictureType picType = PictureType.JPEG;

            string extension = Path.GetExtension(localFileUrl);
            switch (extension.ToLower())
            {
                case "png":
                    picType = PictureType.PNG;
                    break;
                case "gif":
                    picType = PictureType.GIF;
                    break;
                case "bmp":
                    picType = PictureType.BMP;
                    break;
                default:
                    break;
            }

            return picType;
        }

        private static void AddPicture(XWPFParagraph paragraph, List<AddPictureOptions> listAddPictureOptions)
        {
            foreach (var groupedList in listAddPictureOptions.GroupBy(e => e.ImageType))
            {
                DoAddPicture(paragraph, groupedList.ToList());
            }
        }

        private static void AddPictureInTable(XWPFTable table, List<AddPictureOptions> listAddPictureOptions)
        {
            foreach (XWPFTableRow row in table.Rows)
            {
                foreach (XWPFTableCell cell in row.GetTableCells())
                {
                    XWPFParagraph paragraph;
                    for (int i = 0; i < cell.Paragraphs.Count; i++)
                    {
                        paragraph = cell.Paragraphs[i];
                        AddPicture(paragraph, listAddPictureOptions);
                    }
                }
            }
        }

        /// <summary>
        /// 替换表格中的文本
        /// </summary>
        /// <param name="table"></param>
        /// <param name="placeHolder"></param>
        /// <param name="replaceText"></param>
        /// <param name="setRunStyleDelegate"></param>
        private static void ReplacePlaceHolderInTable(XWPFTable table, Dictionary<string, string> oldAndNewStringDict)
        {
            foreach (XWPFTableRow row in table.Rows)
            {
                foreach (XWPFTableCell cell in row.GetTableCells())
                {
                    XWPFParagraph paragraph;
                    for (int i = 0; i < cell.Paragraphs.Count; i++)
                    {
                        paragraph = cell.Paragraphs[i];
                        ReplacePlaceHolderInParagraph(paragraph, oldAndNewStringDict);
                    }
                }
            }
        }

        /// <summary>
        /// 替换段落中的文本
        /// </summary>
        /// <param name="paragraph">段落</param>
        /// <param name="placeHolder">占位符(不要包含特殊字符，建议全英文)</param>
        /// <param name="replaceText">替换后字符</param>
        /// <param name="setRunStyle">设置新文本格式</param>
        private static void ReplacePlaceHolderInParagraph(XWPFParagraph paragraph, Dictionary<string, string> oldAndNewStringDict)
        {
            string runText, newRunText;
            XWPFRun newRun;
            IList<XWPFRun> listRun = paragraph.Runs;

            for (int i = 0; i < listRun.Count; i++)
            {
                XWPFRun run = listRun[i];
                runText = run.Text;
                foreach (var kvp in oldAndNewStringDict)
                {
                    if (run.Text == kvp.Key)
                    {
                        newRunText = runText.Replace(kvp.Key, kvp.Value);

                        paragraph.RemoveRun(i);
                        newRun = paragraph.InsertNewRun(i);

                        if (newRunText.Contains('\n'))
                        {
                            string[] arrayStringLines = newRunText.Split('\n');
                            newRun.SetText(arrayStringLines[0], 0);
                            for (int j = 1; j < arrayStringLines.Length; j++)
                            {
                                newRun.AddBreak(BreakType.TEXTWRAPPING);
                                newRun.SetText(arrayStringLines[j], j);
                            }
                        }
                        else
                        {
                            newRun.SetText(newRunText);
                        }

                        CopyRunStyle(run, newRun);
                        if (newRun.FontSize <= 0)
                        {
                            newRun.FontSize = 10;
                        }
                    }
                }
            }
        }

        /// <summary>
        /// 将旧文本域的样式拷贝至新文本域
        /// </summary>
        /// <param name="fromRun"></param>
        /// <param name="toRun"></param>
        private static void CopyRunStyle(XWPFRun fromRun, XWPFRun toRun)
        {
            Type type = typeof(XWPFRun);
            PropertyInfo[] props = type.GetProperties();
            foreach (PropertyInfo prop in props)
            {
                if (prop.CanWrite)
                {
                    prop.SetValue(toRun, prop.GetValue(fromRun));
                }
            }
        }

        public static string GetRemoteFileExtention(string srcFileUrl)
        {
            string ext = string.Empty;
            if (string.IsNullOrWhiteSpace(srcFileUrl) || !srcFileUrl.Contains("."))
            {
                return ext;
            }

            string[] arrayUrls = srcFileUrl.Split('.');
            ext = arrayUrls[arrayUrls.Length - 1];
            return ext;
        }
    }
}
