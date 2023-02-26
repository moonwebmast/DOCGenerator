using Aspose.Words;
using Aspose.Words.Tables;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;

namespace DOCGenerator
{
    /// <summary>
    /// 文档扩展操作
    /// </summary>
    public static class DocExtension
    {
        /// <summary>
        /// 复制表格最后一行并添加到表格最后
        /// </summary>
        /// <param name="self">表格</param>
        /// <returns>添加的行记录</returns>
        public static Row AddRow(this Table self)
        {
            Row row = (Row)self.LastRow.Clone(true);
            //Row newRow = new Row(self.Document);
            self.Rows.Add(row);
            return row;
        }

        /// <summary>
        /// 复制表格制定的行并在行号位置增加行
        /// </summary>
        /// <param name="self">表格</param>
        /// <param name="idx">行索引值</param>
        /// <returns>添加的行记录</returns>
        public static Row InsertRow(this Table self, int idx)
        {
            Row row = (Row)self.Rows[idx].Clone(true);
            //Row newRow = new Row(self.Document);
            self.Rows.Insert(idx, row);
            return row;
        }

        /// <summary>
        /// 添加单元格
        /// </summary>
        /// <param name="self">行</param>
        /// <returns>添加的单元格</returns>
        public static Cell AddCell(this Row self)
        {

            Cell newCell = new Cell(self.Document);
            self.Cells.Add(newCell);
            return newCell;
        }

        /// <summary>
        /// 设置单元格文本内容
        /// </summary>
        /// <param name="self">单元格</param>
        /// <param name="text">需要设置的文本</param>
        public static void SetText(this Cell self, string text)
        {
            if (self.Paragraphs.Count == 0)
            {
                self.Paragraphs.Add(new Paragraph(self.Document));
            }

            self.FirstParagraph.Runs.Clear();
            self.FirstParagraph.Runs.Add(new Run(self.Document, text));
        }
    }
}
