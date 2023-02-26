using System;
using System.Collections.Generic;
using System.Data;
using System.Reflection;

namespace DOCGenerator
{
    public class DocumentEntity
    {

        Dictionary<String, String> _fields = null;
        /// <summary>
        /// 所有直接替换的属性
        /// </summary>
        public Dictionary<String, String> Fields
        {
            get
            {
                if (_fields != null) {
                    return _fields;
                }                
                this.InitFields();
                return _fields;
            }
        }

        /// <summary>
        /// 根据对象的属性、集合来生成文档数据实体
        /// </summary>
        /// <returns>读取对象的集合</returns>
        public virtual DataSet GetDataSet()
        {
            Type t = this.GetType();
            PropertyInfo[] pi = t.GetProperties();
            DataSet entity = new DataSet();

            foreach (PropertyInfo p in pi)
            {
                // 查找对象属性的集合实现
                var collectionType = p.PropertyType.GetInterface(typeof(IEnumerable<>).FullName);

                if (p.Name == "Fields") { continue; }

                if (p.PropertyType.Name.ToLower() == "string" && !Fields.ContainsKey(p.Name))
                {
                    Fields.Add(p.Name, p.GetValue(this, null).ToString());
                }
                else if (collectionType != null && p.PropertyType.IsGenericType)
                {
                    var itemType = p.PropertyType.GetGenericArguments()[0];
                    // 每个集合转换成一张表
                    var table = entity.Tables.Add(p.Name);

                    //初始化表头
                    table.Columns.Add("RowNo");
                    foreach (var itemP in itemType.GetProperties())
                    {
                        // 文档生成文档生成只用字符类型的属性
                        if (itemP.PropertyType.Name.ToLower() == "string")
                        {
                            table.Columns.Add(itemP.Name, typeof(string));
                        }
                    }
                    // 读取对象属性值
                    var items = (IEnumerable<object>)p.GetValue(this, null);
                    int i = 1;
                    foreach (var item in items)
                    {
                        // 遍历添加行
                        DataRow row = table.NewRow();
                        row["RowNo"] = i.ToString();

                        // 所有字符型属性添加到表中
                        foreach (var itemP in itemType.GetProperties())
                        {
                            // 文档生成只用字符类型的属性
                            if (itemP.PropertyType.Name.ToLower() == "string")
                            {
                                row[itemP.Name] = itemP.GetValue(item).ToString();
                            }
                        }
                        table.Rows.Add(row);
                        i++;
                    }
                }
            }

            return entity;
        }

        private void InitFields()
        {
            _fields = new Dictionary<string, string>();
            Type t = this.GetType();
            PropertyInfo[] pi = t.GetProperties();
            DataSet entity = new DataSet();

            foreach (PropertyInfo p in pi)
            {
                if (p.PropertyType.Name.ToLower() == "string" && !Fields.ContainsKey(p.Name))
                {
                    var val = p.GetValue(this, null);
                    if (val == null)
                    {
                        val = "";
                    }
                    Fields.Add(p.Name, val.ToString());
                }
            }
        }
    }
}
