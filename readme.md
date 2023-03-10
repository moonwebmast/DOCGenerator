# 文档自动生成组件

## 模板制作
1.	新建Doc文件
2.	插入->文档部件->域
3.	域名选择MergeField, 
4.	域名输入属性名称 DocumentTitle（定义的实体类中属性名称或字典中的关键字）

## 代码编写
~~~
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
~~~
