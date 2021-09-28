using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Reflection.Emit;
using System.Text;
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
                        for (int j = 0; j < table.Columns.Count; j++)
                        {
                            var varName = table.Rows[1][j];
                            var varType = table.Rows[2][j];
                            try
                            {
                                Translate.FieldTranslate(typeBuilder, varName.ToString(), varType.ToString().Trim());
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
                                    Translate.FieldInfoTranslate(tempClass, infos[j], table.Rows[i][j]);
                                }
                                catch (Exception e)
                                {
                                    throw new Exception("在文件" + path + "中，第" + (i + 1) + "行,第" + (j + 1) + "列出现数据错误，" + e.Message);
                                }
                            }

                            list.Add(tempClass);
                        }

                        stringBuilder.Append("[");
                        for (int k = 0; k < list.Count; k++)
                        {
                            var item = list[k];

                            stringBuilder.Append("{");

                            var fields = item.GetType().GetFields();
                            for (int i = 0; i < fields.Length; i++)
                            {
                                if (fields[i].FieldType.ToString().Equals(typeof(string).FullName))
                                {
                                    stringBuilder.Append("\"" + fields[i].Name + "\"" + ":" + "\"" + fields[i].GetValue(item) + "\"");
                                }
                                else
                                {
                                    stringBuilder.Append("\"" + fields[i].Name + "\"" + ":" + fields[i].GetValue(item));
                                }


                                if (i != fields.Length - 1 && fields.Length > 1)
                                {
                                    stringBuilder.Append(",");
                                }
                            }
                            stringBuilder.Append("}");

                            if (k != list.Count - 1 && list.Count > 1)
                            {
                                stringBuilder.Append(",");
                            }
                        }
                        stringBuilder.Append("]");

                        jsonBytes = Encoding.UTF8.GetBytes(stringBuilder.ToString());
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

                using (var jsonStream = File.Open(jsonPath, FileMode.OpenOrCreate, FileAccess.Write))
                {
                    jsonStream.Write(jsonBytes, 0, jsonBytes.Length);
                }
            }
            //}
            //catch(FileNotFoundException e)
            //{

            //}
            //catch(IndexOutOfRangeException e)
            //{
            //    Console.WriteLine("此表不符合我们达成的约定，请检查你的表是否定义第二行和第三行。错误信息： " + e.Message);
            //}
            //catch(Exception e)
            //{
            //    Console.WriteLine("未知异常。异常信息：" + e.Message);
            //}
        }
    }
}
