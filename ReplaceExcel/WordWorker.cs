using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Text.RegularExpressions;

namespace ReplaceExcel
{
    public class WordWorker
    {
        protected Application wApp;
        protected Document wDoc;

        public void Destroy()
        {
            if (wApp != null) {
                wApp.Quit();
                wApp = null;
            }
        }

        #region 打开Word文档
        public bool OpenWord(string fileName)
        {
            Object FileName = fileName; // 文档的名称。默认值是当前文件夹名和文件名。如果文档在以前没有保存过，则使用默认名称（例如，Doc1.doc）。如果已经存在具有指定文件名的文档，则会在不先提示用户的情况下改写文档。  
            if (wApp == null) {
                wApp = new Microsoft.Office.Interop.Word.Application();
            }
            wApp.Visible = false;
            object isread = false;
            object isvisible = true;
            object miss = System.Reflection.Missing.Value;

            try {
                wDoc = wApp.Documents.Open(ref FileName, ref miss, ref isread, ref miss, ref miss, ref miss, ref miss, ref miss,
                                  ref miss, ref miss, ref miss, ref isvisible, ref miss, ref miss, ref miss, ref miss);
                return true;
            }
            catch (Exception ex) {
                string err = string.Format("另存文件出错，错误原因：{0}", ex.Message);
                throw new Exception(err, ex);
            }
        }
        #endregion


        #region 文档另存为其他文件名
        /// <summary>  
        /// 文档另存为其他文件名  
        /// </summary>  
        /// <param name="fileName">文件名</param>  
        /// <param name="wDoc">Document对象</param>  
        public bool SaveAs(string fileName)
        {
            try {
                return SaveAs(fileName, wDoc);
            }
            catch (Exception ex) {
                throw ex;
            }
        }
        #endregion


        #region 文档另存为其他文件名
        /// <summary>  
        /// 文档另存为其他文件名  
        /// </summary>  
        /// <param name="fileName">文件名</param>  
        /// <param name="wDoc">Document对象</param>  
        public static bool SaveAs(string fileName, Document wDoc)
        {
            Object FileName = fileName; // 文档的名称。默认值是当前文件夹名和文件名。如果文档在以前没有保存过，则使用默认名称（例如，Doc1.doc）。如果已经存在具有指定文件名的文档，则会在不先提示用户的情况下改写文档。  
            Object FileFormat = WdSaveFormat.wdFormatDocument; // 文档的保存格式。可以是任何 WdSaveFormat 值。要以另一种格式保存文档，请为 SaveFormat 属性指定适当的值。  
            Object LockComments = false; // 如果为 true，则锁定文档以进行注释。默认值为 false。  
            Object Password = System.Type.Missing; // 用来打开文档的密码字符串。（请参见下面的备注。）  
            Object AddToRecentFiles = false; // 如果为 true，则将该文档添加到“文件”菜单上最近使用的文件列表中。默认值为 true。  
            Object WritePassword = System.Type.Missing; // 用来保存对文件所做更改的密码字符串。（请参见下面的备注。）  
            Object ReadOnlyRecommended = false; // 如果为 true，则让 Microsoft Office Word 在打开文档时建议只读状态。默认值为 false。  
            Object EmbedTrueTypeFonts = false; //如果为 true，则将 TrueType 字体随文档一起保存。如果省略的话，则 EmbedTrueTypeFonts 参数假定 EmbedTrueTypeFonts 属性的值。  
            Object SaveNativePictureFormat = true; // 如果图形是从另一个平台（例如，Macintosh）导入的，则 true 表示仅保存导入图形的 Windows 版本。  
            Object SaveFormsData = false; // 如果为 true，则将用户在窗体中输入的数据另存为数据记录。  
            Object SaveAsAOCELetter = false; // 如果文档附加了邮件程序，则 true 表示会将文档另存为 AOCE 信函（邮件程序会进行保存）。  
            Object Encoding = System.Type.Missing; // MsoEncoding。要用于另存为编码文本文件的文档的代码页或字符集。默认值是系统代码页。  
            Object InsertLineBreaks = true; // 如果文档另存为文本文件，则 true 表示在每行文本末尾插入分行符。  
            Object AllowSubstitutions = false; //如果文档另存为文本文件，则 true 允许 Word 将某些符号替换为外观与之类似的文本。例如，将版权符号显示为 (c)。默认值为 false。  
            Object LineEnding = WdLineEndingType.wdCRLF;// Word 在另存为文本文件的文档中标记分行符和换段符。可以是任何 WdLineEndingType 值。  
            Object AddBiDiMarks = true;//如果为 true，则向输出文件添加控制字符，以便保留原始文档中文本的双向布局。  
            try {
                wDoc.SaveAs(ref FileName, ref FileFormat, ref LockComments, ref Password, ref AddToRecentFiles, ref WritePassword
                        , ref ReadOnlyRecommended, ref EmbedTrueTypeFonts, ref SaveNativePictureFormat
                        , ref SaveFormsData, ref SaveAsAOCELetter, ref Encoding, ref InsertLineBreaks, ref AllowSubstitutions
                        , ref LineEnding, ref AddBiDiMarks);
                return true;
            }
            catch (Exception ex) {
                string err = string.Format("另存文件出错，错误原因：{0}", ex.Message);
                throw new Exception(err, ex);
            }
        }
        #endregion


        #region 关闭文档
        /// <summary>  
        /// 关闭文档  
        /// </summary>  
        public void Close()
        {
            Close(wDoc, wApp);
            wDoc = null;
        }
        #endregion


        #region 关闭文档
        /// <summary>  
        /// 关闭文档  
        /// </summary>  
        /// <param name="wDoc">Document对象</param>  
        /// <param name="WApp">Application对象</param>  
        public static void Close(Document wDoc, Application WApp)
        {
            try {
                Object SaveChanges = WdSaveOptions.wdSaveChanges;// 指定文档的保存操作。可以是下列 WdSaveOptions 值之一：wdDoNotSaveChanges、wdPromptToSaveChanges 或 wdSaveChanges。  
                Object OriginalFormat = WdOriginalFormat.wdOriginalDocumentFormat;// 指定文档的保存格式。可以是下列 WdOriginalFormat 值之一：wdOriginalDocumentFormat、wdPromptUser 或 wdWordDocument。  
                Object RouteDocument = false;// 如果为 true，则将文档传送给下一个收件人。如果没有为文档附加传送名单，则忽略此参数。  
                if (wDoc != null) {
                    wDoc.Close(ref SaveChanges, ref OriginalFormat, ref RouteDocument);
                    wDoc = null;
                }
                //if (WApp != null) WApp.Quit(ref SaveChanges, ref OriginalFormat, ref RouteDocument);
            }
            catch (Exception ex) {
                throw ex;
            }
        }
        #endregion
 
        #region 找到表格
        public bool FindTable(string bookmarkTable)
        {
            try {
                object bkObj = bookmarkTable;
                if (wApp.ActiveDocument.Bookmarks.Exists(bookmarkTable) == true) {
                    wApp.ActiveDocument.Bookmarks.get_Item(ref bkObj).Select();
                    return true;
                }
                else
                    return false;
            }
            catch (Exception ex) {
                throw ex;
            }
        }
        #endregion

        #region 移动到下一单元格
        public void MoveNextCell()
        {
            try {
                Object unit = WdUnits.wdCell;
                Object count = 1;
                wApp.Selection.Move(ref unit, ref count);
            }
            catch (Exception ex) {
                throw ex;
            }
        }
        #endregion

        #region 是文件名有效
        public string MakeFilenameValid(string filename)
        {
            if (filename == null)
                throw new ArgumentNullException();


            if (filename.EndsWith("."))
                filename = Regex.Replace(filename, @"\.+$", "");


            if (filename.Length == 0)
                throw new ArgumentException();


            if (filename.Length > 245)
                throw new PathTooLongException();


            foreach (char c in System.IO.Path.GetInvalidFileNameChars()) {
                filename = filename.Replace(c, '_');
            }


            return filename;
        }
        #endregion

        //用模板字符替换原字符
        public void ReplaceString(string origialString, string destinationString)
        {
            #region 如果目标字符串长度大于255，则分成小于255的若干多分别进行替换
            /*
             * 例如：destinationString的长度是300；
             * 第一步：将origialString替换为“$$$1$”和“$$$2$”;
             * 第二步：将destinationString分解为两个字符串，第一个字符串长度255，第二个字符串长度300-255=45；
             * 第三步：用第一个字符串替换“$$$1$”，用第二个字符串替换“$$$2$”
             */
            if (destinationString.Length > 255) {
                int count = destinationString.Length / 255 + ((destinationString.Length % 255) == 0 ? 0 : 1);
                List<string> origianlStringList = new List<string>();
                List<string> destinationStringList = new List<string>();
                for (int i = 0; i < count; i++) {
                    origianlStringList.Add("$$$" + i.ToString() + "$");

                    int length;//每小段的长度
                    if (i != count - 1) {
                        length = 255;
                    }
                    else {
                        length = destinationString.Length % 255;
                    }
                    destinationStringList.Add(destinationString.Substring(i * 255, length));
                }
                string origianlStringListTotalString = string.Empty;
                for (int i = 0; i < count; i++) {
                    origianlStringListTotalString += origianlStringList[i];
                }
                this.ReplaceString(origialString, origianlStringListTotalString);
                for (int i = 0; i < count; i++) {
                    this.ReplaceString(origianlStringList[i], destinationStringList[i]);
                }
                return;
            }

            #endregion

            Object missing = Missing.Value;
            object replaceAll = WdReplace.wdReplaceAll;

            wApp.Selection.Find.ClearFormatting();
            wApp.Selection.Find.Text = origialString;

            wApp.Selection.Find.Replacement.ClearFormatting();
            wApp.Selection.Find.Replacement.Text = destinationString;

            wApp.Selection.Find.Execute(
                ref missing, ref missing, ref missing, ref missing, ref missing,
                ref missing, ref missing, ref missing, ref missing, ref missing,
                ref replaceAll, ref missing, ref missing, ref missing, ref missing);
            ////
            ////替换在文本框中的文字
            ////
            //StoryRanges ranges = wApp.ActiveDocument.StoryRanges;
            //foreach (Range item in ranges) {
            //    if (WdStoryType.wdTextFrameStory == item.StoryType) {
            //        Range range = item;
            //        while (range != null) {
            //            range.Find.ClearFormatting();
            //            range.Find.Text = origialString;
            //            range.Find.Replacement.Text = destinationString;
            //            range.Find.Execute(
            //        ref missing, ref missing, ref missing, ref missing, ref missing,
            //        ref missing, ref missing, ref missing, ref missing, ref missing,
            //        ref replaceAll, ref missing, ref missing, ref missing, ref missing);
            //            range = range.NextStoryRange;
            //        }
            //        break;
            //    }
            //}
        }
    }
}
