using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Reflection.Emit;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using ExcelDataReader;

namespace ExcelFormat
{
    public class JsonFormater : Formater
    {
        public override void Format(string[] filePaths,int codepage = 65001)
        {
            foreach (var path in filePaths)
            {
                StringBuilder stringBuilder = new StringBuilder();
                byte[] jsonBytes = null;
                using (var rawStream = File.Open(path, FileMode.Open, FileAccess.Read))
                {
                    ExcelReaderConfiguration excelReaderConfiguration = new ExcelReaderConfiguration();
                    excelReaderConfiguration.FallbackEncoding = Encoding.GetEncoding(codepage);

                    using (var excelStream = path.EndsWith("cvs") ? ExcelReaderFactory.CreateCsvReader(rawStream) : ExcelReaderFactory.CreateReader(rawStream))
                    {
                        var dataSet = excelStream.AsDataSet();
                        var table = dataSet.Tables[0];

                        AssemblyName assemblyName = new AssemblyName();
                        assemblyName.Name = "TempAssembly";
                        AssemblyBuilder assemblyBuilder = AssemblyBuilder.DefineDynamicAssembly(assemblyName, AssemblyBuilderAccess.Run);
                        ModuleBuilder moduleBuilder = assemblyBuilder.DefineDynamicModule("TempMoudle");
                        TypeBuilder typeBuilder = moduleBuilder.DefineType("TempClass", TypeAttributes.Public);

                        Dictionary<int, Type> keyValuePairs = new Dictionary<int, Type>();
                        for (int j = 0; j < table.Columns.Count; j++)
                        {
                            var varName = table.Rows[1][j];
                            var varType = table.Rows[2][j];
  
                            try
                            {
                                Type buildType = Translate.FieldTranslate(typeBuilder, varName.ToString(), varType.ToString().Trim());
                                if(buildType != null)
                                {
                                    keyValuePairs.Add(j, buildType);
                                }
                            }
                            catch (Exception e)
                            {
                                throw new Exception("在文件" + path + "中，第" + (j + 1) + "列" + e.Message);
                            }


                        }

                        Type type = typeBuilder.CreateType();

                        List<Object> list = new List<Object>();

                        for (int i = 3; i < table.Rows.Count; i++)
                        {
                            Object tempClass = Activator.CreateInstance(type);

                            FieldInfo[] infos = tempClass.GetType().GetFields();

                            for (int j = 0; j < infos.Length; j++)
                            {
                                try
                                {
                                    if(keyValuePairs.ContainsKey(j))
                                    {
                                        Translate.FieldInfoTranslate(tempClass, infos[j], table.Rows[i][j],keyValuePairs[j]);
                                    }
                                    else
                                    {
                                        Translate.FieldInfoTranslate(tempClass, infos[j], table.Rows[i][j], null);
                                    }
                                    
        
                                }
                                catch (Exception e)
                                {
                                    throw new Exception("在文件" + path + "中，第" + (i + 1) + "行,第" + (j + 1) + "列出现数据错误，" + e.Message);
                                }
                            }

                            list.Add(tempClass);
                        }

                        string jsonContent = LitJson.JsonMapper.ToJson(list);
                        #region
                        //stringBuilder.Append("[");
                        //for (int k = 0; k < list.Count; k++)
                        //{
                        //    var item = list[k];

                        //    stringBuilder.Append("{");

                        //    var fields = item.GetType().GetFields();
                        //    for (int i = 0; i < fields.Length; i++)
                        //    {
                        //        if (fields[i].FieldType.ToString().Equals(typeof(string).FullName))
                        //        {
                        //            stringBuilder.Append("\"" + fields[i].Name + "\"" + ":" + "\"" + fields[i].GetValue(item) + "\"");
                        //        }
                        //        else if (fields[i].FieldType.ToString().Equals(typeof(IDAndValue).FullName))
                        //        {
                        //            if (fields[i].GetValue(item) == null)
                        //            {
                        //                stringBuilder.Append("\"" + fields[i].Name + "\"" + ":" + "null");
                        //            }
                        //            else
                        //            {
                        //                stringBuilder.Append("\"" + fields[i].Name + "\"" + ":" + fields[i].GetValue(item).ToString());
                        //            }
                        //        }
                        //        else if (fields[i].FieldType.FullName.ToString().Equals(typeof(List<int>).FullName))
                        //        {
                        //            if(fields[i].GetValue(item) == null)
                        //            {
                        //                stringBuilder.Append("\"" + fields[i].Name + "\"" + ":" + "null");
                        //            }
                        //            else
                        //            {
                        //                stringBuilder.Append("\"" + fields[i].Name + "\"" + ":" + "[");
                        //                List<int> l = (List<int>)fields[i].GetValue(item);
                        //                for (int li = 0; li < l.Count; li++)
                        //                {
                        //                    stringBuilder.Append(l[li]);
                        //                    if (li != l.Count - 1)
                        //                    {
                        //                        stringBuilder.Append(",");
                        //                    }
                        //                }
                        //                stringBuilder.Append("]");
                        //            }

                        //        }
                        //        else if(fields[i].FieldType.FullName.ToString().Equals(typeof(List<string>).FullName))
                        //        {
                        //            if(fields[i].GetValue(item) == null)
                        //            {
                        //                stringBuilder.Append("\"" + fields[i].Name + "\"" + ":" + "null");
                        //            }
                        //            else
                        //            {
                        //                stringBuilder.Append("\"" + fields[i].Name + "\"" + ":" + "[");
                        //                List<string> l = (List<string>)fields[i].GetValue(item);
                        //                for (int li = 0; li < l.Count; li++)
                        //                {
                        //                    stringBuilder.Append("\"" + l[li] + "\"");
                        //                    if (li != l.Count - 1)
                        //                    {
                        //                        stringBuilder.Append(",");
                        //                    }
                        //                }
                        //                stringBuilder.Append("]");
                        //            }

                        //        }
                        //        else if(fields[i].FieldType.FullName.ToString().Equals(typeof(List<IDAndValue>).FullName))
                        //        {
                        //            if (fields[i].GetValue(item) == null)
                        //            {
                        //                stringBuilder.Append("\"" + fields[i].Name + "\"" + ":" + "null");
                        //            }
                        //            else
                        //            {
                        //                stringBuilder.Append("\"" + fields[i].Name + "\"" + ":" + "[");
                        //                List<IDAndValue> l = (List<IDAndValue>)fields[i].GetValue(item);
                        //                for (int li = 0; li < l.Count; li++)
                        //                {
                        //                    stringBuilder.Append(l[li].ToString());
                        //                    if (li != l.Count - 1)
                        //                    {
                        //                        stringBuilder.Append(",");
                        //                    }
                        //                }
                        //                stringBuilder.Append("]");
                        //            }
                        //        }
                        //        else if(fields[i].FieldType.ToString().Equals(typeof(bool).FullName))
                        //        {
                        //            stringBuilder.Append("\"" + fields[i].Name + "\"" + ":" + fields[i].GetValue(item).ToString().ToLower());
                        //        }
                        //        else if(fields[i].FieldType.ToString().Equals("Obj"))
                        //        {
                        //            //stringBuilder.Append(fields[i].GetValue(item).ToString());
                        //            if (fields[i].GetValue(item) == null)
                        //            {
                        //                stringBuilder.Append("\"" + fields[i].Name + "\"" + ":" + "null");
                        //            }
                        //            else
                        //            {
                        //                FieldInfo[] infos = fields[i].GetValue(item).GetType().GetFields(BindingFlags.Instance | BindingFlags.Public);
                        //                stringBuilder.Append("{");
                        //                for (int findex = 0; findex < infos.Length; findex++)
                        //                {

                        //                    stringBuilder.Append("\"" + fields[i].Name + "\"" + ":" + fields[i].GetValue(item));
                        //                    if(findex != findex -1)
                        //                    {
                        //                        stringBuilder.Append(",");
                        //                    }
                        //                }
                        //                stringBuilder.Append("}");
                        //                Console.WriteLine(fields[i].GetValue(item));
                        //                stringBuilder.Append("\"" + fields[i].Name + "\"" + ":" + fields[i].GetValue(item).ToString());
                        //            }

                        //        }

                        //        else
                        //        {
                        //            stringBuilder.Append("\"" + fields[i].Name + "\"" + ":" + fields[i].GetValue(item));
                        //        }

                        //        if (i != fields.Length - 1 && fields.Length > 1)
                        //        {
                        //            stringBuilder.Append(",");
                        //        }
                        //    }
                        //    stringBuilder.Append("}");

                        //    if (k != list.Count - 1 && list.Count > 1)
                        //    {
                        //        stringBuilder.Append(",");
                        //    }
                        //}
                        //stringBuilder.Append("]");
                        #endregion
                        jsonContent = Regex.Unescape(jsonContent);
                        jsonBytes = Encoding.UTF8.GetBytes(jsonContent);
                    }
                }


                var rawPath = path.Split('/', '\\');
                string[] absPath = new string[rawPath.Length - 1];
                Array.Copy(rawPath, 0, absPath, 0, rawPath.Length - 1);
                string jsonPath = null;
                foreach (var item in absPath)
                {
                    jsonPath += (item + "\\");
                }
                jsonPath = jsonPath + rawPath[rawPath.Length - 1].Split('.')[0] + ".json";

                using (var jsonStream = File.Open(jsonPath, FileMode.Create, FileAccess.Write))
                {
                    jsonStream.Write(jsonBytes, 0, jsonBytes.Length);
                }
            }
        }
    }
}
