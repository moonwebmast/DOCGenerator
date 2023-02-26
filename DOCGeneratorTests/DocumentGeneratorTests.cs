using Microsoft.VisualStudio.TestTools.UnitTesting;
using DOCGenerator;
using System;
using System.Collections.Generic;
using System.Text;
using Aspose.Words.Tables;
using Aspose.Words;
using System.Drawing;
using System.IO;
using System.Drawing.Imaging;
using System.Drawing.Drawing2D;
using System.Linq;

namespace DOCGenerator.Tests
{
    [TestClass()]
    public class DocumentGeneratorTests
    {
        public class RowItem : DocumentEntity
        {
            public string Id { get; set; }

            public string Name { get; set; }

            public string Desc { get; set; }
        }

        public class TestEntity : DocumentEntity
        {
            public List<RowItem> Items { get; } = new List<RowItem>();
        }

        [TestMethod()]
        public void GeneratorTest()
        {
            TestEntity data = new TestEntity()
            {
                // 直接初始化字段
                Fields = {
                    {"DocumentTitle","测试项目名称" },
                    {"DocumentVersion","100101" }
                },
                // 初始化对象属性
                Items = {
                    new RowItem{ Id = "1",Name = "Name1",Desc = "Desc-1" },
                    new RowItem{ Id = "2",Name = "Name2",Desc = "Desc-2" },
                    new RowItem{ Id = "3",Name = "Name3",Desc = "Desc-3" },
                }
            };

            DocumentGenerator
                .Generator<TestEntity>(@"./template/TestTemplate.docx")
                .Process(data, (doc) =>
                {
                    // 获得文档
                    DocumentBuilder b = new DocumentBuilder(doc);
                    // 跳转到书签位置
                    b.MoveToBookmark("flow");
                    // 画简易流程图
                    var imgData = FlowChart.Draw(data.Items.Select(x => x.Name).ToArray());
                    // 插入图片
                    b.InsertImage(imgData);

                    data.Items.ForEach(x =>
                    {
                        doc.AppendDocument(DocumentGenerator.Generator<RowItem>(@"./template/SubTemplate.docx")
                            .Process(x, (subDoc) =>
                            {
                                // 获得文档
                                DocumentBuilder subBuilder = new DocumentBuilder(subDoc);
                                // 跳转到书签位置
                                subBuilder.MoveToBookmark("subflow");
                                // 画简易流程图
                                var imgData = FlowChart.Draw(new string[] { x.Name });
                                // 插入图片
                                subBuilder.InsertImage(imgData);
                            }).Document, ImportFormatMode.UseDestinationStyles);
                    });
                })
                .SaveDoc("./output/Test.doc")
                .SavePdf("./output/Test.pdf");
        }



    }
}