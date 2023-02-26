using Aspose.Words;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;

namespace DOCGenerator
{
    /// <summary>
    /// 文档生成器
    /// </summary>
    public class DocumentGenerator
    {
        public static DocumentGenerator<DT> Generator<DT>(string templateFile) where DT : DocumentEntity
        {
            var dg = new DocumentGenerator<DT>(templateFile);
            return dg;
        }

        /// <summary>
        /// 生成文档
        /// </summary>
        /// <param name="templateFile">模板文件</param>
        /// <param name="data">生成文档的数据实体</param>
        /// <param name="action">自定义回调处理函数</param>
        public static DocumentGenerator<DT> Generator<DT>(string templateFile, DT data, Action<Document> action = null) where DT : DocumentEntity
        {
            var dg = new DocumentGenerator<DT>(templateFile);
            dg.Process(data, action);
            return dg;
        }
    }
    /// <summary>
    /// 文档生成器
    /// </summary>
    /// <typeparam name="T"></typeparam>
    public class DocumentGenerator<T> : DocumentGenerator
        where T : DocumentEntity
    {

        /// <summary>
        /// Word
        /// </summary>
        public Document _doc;


        /// <summary>
        /// 初始化插件
        /// </summary>
        static DocumentGenerator()
        {
            string key = "";
            if (System.IO.File.Exists(@"./key.lic"))
            {
                key = File.ReadAllText(@"./key.lic");
            }
            new Aspose.Words.License().SetLicense(new MemoryStream(Convert.FromBase64String(key)));
        }


        /// <summary>
        /// 构造函数
        /// </summary>
        /// <param name="templateFile">模板文件</param>
        public DocumentGenerator(string templateFile)
        {
            TemplateFile = templateFile;
            OpenTempelte(TemplateFile);
        }

        /// <summary>
        /// 当前生成的文档对象
        /// </summary>
        public Document Document { get => _doc; }

        /// <summary>
        /// 模板文件地址
        /// </summary>
        public virtual String TemplateFile { get; }

        /// <summary>
        /// 处理文档
        /// </summary>
        /// <param name="data">需要绑定的数据</param>
        /// <param name="action">处理过程中回调函数，用于定制其他生成逻辑</param>
        /// <returns></returns>
        public DocumentGenerator<T> Process(T data, Action<Document> action = null)
        {
            _doc.MailMerge.Execute(data.Fields.Keys.ToArray(), data.Fields.Values.ToArray());
            _doc.MailMerge.ExecuteWithRegions(data.GetDataSet());
            /// 最后一个结
            var sec = Document.LastSection;

            action?.Invoke(Document);

            OnProcess(data);

            #region 合并结/更新域
            var idx = Document.Sections.IndexOf(sec);

            if (idx != Document.Sections.Count - 1)
            {
                List<Section> removeSections = new List<Section>();

                for (int i = idx + 1; i < Document.Sections.Count; i++)
                {
                    var csec = Document.Sections[i];
                    sec.AppendContent(csec);
                    csec.ClearContent();
                    removeSections.Add(csec);
                }

                foreach (var rsec in removeSections)
                {
                    rsec.Remove();
                }
            }


            // 更新域
            Document.UpdateFields();
            Document.UpdateListLabels();
            Document.UpdateTableLayout();
            Document.UpdateThumbnail();
            Document.UpdatePageLayout();
            Document.UpdateWordCount();
            #endregion

            return this;
        }

        /// <summary>
        /// 处理文档
        /// </summary>
        /// <param name="data">需要绑定的数据</param>
        public virtual void OnProcess(T data)
        {
        }

        /// <summary>
        /// 合并文档
        /// </summary>
        /// <param name="document">追加到最后的文档</param>
        public void AppendDocument(Document document)
        {
            Document.AppendDocument(document, ImportFormatMode.UseDestinationStyles, new ImportFormatOptions
            {
                KeepSourceNumbering = false,
                SmartStyleBehavior = true,
            });
        }

        /// <summary>
        /// 保持DOC文件
        /// </summary>
        /// <param name="filename">文件名</param>
        /// <returns>返回当前对象</returns>
        public DocumentGenerator<T> SaveDoc(string filename)
        {
            _doc.Save(filename, SaveFormat.Doc);
            return this;
        }

        /// <summary>
        /// 保持PDF文件
        /// </summary>
        /// <param name="filename">文件名</param>
        /// <returns>返回当前对象</returns>
        public DocumentGenerator<T> SavePdf(string filename)
        {
            _doc.Save(filename, SaveFormat.Pdf);
            return this;
        }


        /// <summary>
        /// 基于模版新建Word文件
        /// </summary>
        /// <param name="path">模板路径</param>
        public void OpenTempelte(string path)
        {
            _doc = new Document(path);
        }

        /// <summary>
        /// 书签赋值用法
        /// </summary>
        /// <param name="LabelId">书签名</param>
        /// <param name="Content">内容</param>
        public void WriteBookMark(string LabelId, string Content)
        {
            if (_doc.Range.Bookmarks[LabelId] != null)
            {
                _doc.Range.Bookmarks[LabelId].Text = Content;
            }
        }
        /// <summary>
        /// 不可编辑受保护，需输入密码
        /// </summary>
        /// <param name="pwd">密码</param>
        public void NoEdit(string pwd)
        {
            _doc.Protect(ProtectionType.ReadOnly, pwd);
        }

        /// <summary>
        /// 设置只读
        /// </summary>
        public void ReadOnly()
        {
            _doc.Protect(ProtectionType.ReadOnly);
        }
    }
}
