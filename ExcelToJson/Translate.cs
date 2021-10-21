using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Reflection.Emit;
using System.Text;
using System.Threading.Tasks;

namespace ExcelFormat
{
    public class Translate
    {
        public static void  FieldTranslate(TypeBuilder typeBuilder,string varName,string varType )
        {        
            if(varType.Trim().Equals("bool"))
            {
                FieldBuilder fieldBuilder = typeBuilder.DefineField(varName.Trim().ToString(), typeof(bool), FieldAttributes.Public);
            }
            else if (varType.Trim().Equals("sbyte"))
            {
                FieldBuilder fieldBuilder = typeBuilder.DefineField(varName.Trim().ToString(), typeof(sbyte), FieldAttributes.Public);
            }
            else if  (varType.Trim().Equals("byte"))
            {
                FieldBuilder fieldBuilder = typeBuilder.DefineField(varName.Trim().ToString(), typeof(byte), FieldAttributes.Public);
            }
            else if  (varType.Trim().Equals("short"))
            {
                FieldBuilder fieldBuilder = typeBuilder.DefineField(varName.Trim().ToString(), typeof(short), FieldAttributes.Public);
            }
            else if (varType.Trim().Equals("ushort"))
            {
                FieldBuilder fieldBuilder = typeBuilder.DefineField(varName.Trim().ToString(), typeof(ushort), FieldAttributes.Public);
            }
            else if(varType.Trim().Equals("int"))
            {
                FieldBuilder fieldBuilder = typeBuilder.DefineField(varName.Trim().ToString(), typeof(int), FieldAttributes.Public);
            }
            else if (varType.Trim().Equals("uint"))
            {
                FieldBuilder fieldBuilder = typeBuilder.DefineField(varName.Trim().ToString(), typeof(uint), FieldAttributes.Public);
            }
            else if (varType.Trim().Equals("long"))
            {
                FieldBuilder fieldBuilder = typeBuilder.DefineField(varName.Trim().ToString(), typeof(long), FieldAttributes.Public);
            }
            else if (varType.Trim().Equals("ulong"))
            {
                FieldBuilder fieldBuilder = typeBuilder.DefineField(varName.Trim().ToString(), typeof(ulong), FieldAttributes.Public);
            }
            else if (varType.Trim().Equals("float"))
            {
                FieldBuilder fieldBuilder = typeBuilder.DefineField(varName.Trim().ToString(), typeof(float), FieldAttributes.Public);
            }
            else if (varType.Trim().Equals("double"))
            {
                FieldBuilder fieldBuilder = typeBuilder.DefineField(varName.Trim().ToString(), typeof(double), FieldAttributes.Public);
            }
            else if (varType.Trim().Equals("decimal"))
            {
                FieldBuilder fieldBuilder = typeBuilder.DefineField(varName.Trim().ToString(), typeof(decimal), FieldAttributes.Public);
            }
            else if (varType.Trim().Equals("string"))
            {
                FieldBuilder fieldBuilder = typeBuilder.DefineField(varName.Trim().ToString(), typeof(string), FieldAttributes.Public);
            }
            else if(varType.Trim().ToLower().Equals("idandvalue"))
            {
                FieldBuilder fieldBuilder = typeBuilder.DefineField(varName.Trim().ToString(), typeof(IDAndValue), FieldAttributes.Public);
            }
            else if(varType.Trim().ToLower().Equals("list<int>"))
            {
                FieldBuilder fieldBuilder = typeBuilder.DefineField(varName.Trim().ToString(), typeof(List<int>), FieldAttributes.Public);
            }
            else if(varType.Trim().ToLower().Equals("list<string>"))
            {
                FieldBuilder fieldBuilder = typeBuilder.DefineField(varName.Trim().ToString(), typeof(List<string>), FieldAttributes.Public);
            }
            else if (varType.Trim().ToLower().Equals("list<idandvalue>"))
            {
                FieldBuilder fieldBuilder = typeBuilder.DefineField(varName.Trim().ToString(), typeof(List<IDAndValue>), FieldAttributes.Public);
            }
            else
            {
                throw new Exception("检测到不支持的数据类型，请检查您定义的数据类型是否错误");
            }

        }

        public static void FieldInfoTranslate(object instance,FieldInfo info,object value)
        {
            try
            {
                var v = value.ToString();
                if (info.FieldType.ToString().Equals(typeof(bool).FullName))
                {
                    if (string.IsNullOrEmpty(v))
                        info.SetValue(instance, false);
                    else
                        info.SetValue(instance, bool.Parse(v));
                }
                if (info.FieldType.ToString().Equals(typeof(sbyte).FullName))
                {
                    if (string.IsNullOrEmpty(v))
                        info.SetValue(instance, 0);
                    else
                        info.SetValue(instance, sbyte.Parse(v));
                }
                if (info.FieldType.ToString().Equals(typeof(byte).FullName))
                {
                    if (string.IsNullOrEmpty(v))
                        info.SetValue(instance, 0);
                    else
                        info.SetValue(instance, byte.Parse(v));
                }
                if (info.FieldType.ToString().Equals(typeof(short).FullName))
                {
                    if (string.IsNullOrEmpty(v))
                        info.SetValue(instance, 0);
                    else
                        info.SetValue(instance, short.Parse(v));
                }
                if (info.FieldType.ToString().Equals((typeof(ushort).FullName)))
                {
                    if (string.IsNullOrEmpty(v))
                        info.SetValue(instance, 0);
                    else
                        info.SetValue(instance, ushort.Parse(v));
                }
                if (info.FieldType.ToString().Equals((typeof(int).FullName)))
                {
                    if (string.IsNullOrEmpty(v))
                        info.SetValue(instance, 0);
                    else
                        info.SetValue(instance, int.Parse(v));
                }
                if (info.FieldType.ToString().Equals((typeof(uint).FullName)))
                {
                    if (string.IsNullOrEmpty(v))
                        info.SetValue(instance, 0);
                    else
                        info.SetValue(instance, uint.Parse(v));
                }
                if (info.FieldType.ToString().Equals((typeof(long).FullName)))
                {
                    if (string.IsNullOrEmpty(v))
                        info.SetValue(instance, 0);
                    else
                        info.SetValue(instance, long.Parse(v));
                }
                if (info.FieldType.ToString().Equals((typeof(ulong).FullName)))
                {
                    if (string.IsNullOrEmpty(v))
                        info.SetValue(instance, 0);
                    else
                        info.SetValue(instance, ulong.Parse(v));
                }
                if (info.FieldType.ToString().Equals((typeof(float).FullName)))
                {
                    if (string.IsNullOrEmpty(v))
                        info.SetValue(instance, 0f);
                    else
                        info.SetValue(instance, float.Parse(v));
                }
                if (info.FieldType.ToString().Equals((typeof(double).FullName)))
                {
                    if (string.IsNullOrEmpty(v))
                        info.SetValue(instance, 0);
                    else
                        info.SetValue(instance, double.Parse(v));
                }
                if (info.FieldType.ToString().Equals((typeof(decimal).FullName)))
                {
                    if (string.IsNullOrEmpty(v))
                        info.SetValue(instance, 0);
                    else
                        info.SetValue(instance, decimal.Parse(v));
                }
                if (info.FieldType.ToString().Equals((typeof(string).FullName)))
                {
                    info.SetValue(instance, v);
                }
                if (info.FieldType.ToString().Equals(typeof(IDAndValue).FullName))
                {
                    if(string.IsNullOrEmpty(v))
                    {

                        info.SetValue(instance, null);
                    }
                    else
                    {
                        IDAndValue iDAndValue = new IDAndValue();
                        string[] args = v.Split('/', '_', ',');
                        iDAndValue.id = int.Parse(args[0]);
                        iDAndValue.value = int.Parse(args[1]);
                        info.SetValue(instance, iDAndValue);
                    }

                }
                if (info.FieldType.FullName.ToString().Equals(typeof(List<int>).FullName))
                {
                    List<int> li = new List<int>();
                    string[] args = v.Split('/', ',', '_');
                    if (args.Length > 0)
                    {
                        foreach (var item in args)
                        {
                            li.Add(int.Parse(item.Trim()));
                        }
                    }

                    info.SetValue(instance, li);
                }
                if(info.FieldType.FullName.ToString().Equals(typeof(List<string>).FullName))
                {
                    List<string> ls = new List<string>();
                    string[] args = v.Split('/', ',');
                    if(args.Length > 0)
                    {
                        foreach (var item in args)
                        {
                            ls.Add(item.Trim());
                        }
                    }
                    info.SetValue(instance, ls);
                }

                if(info.FieldType.FullName.ToString().Equals(typeof(List<IDAndValue>).FullName))
                {
                    List<IDAndValue> iDAndValues = new List<IDAndValue>();
                    string[] args = v.Split('/',';');
                    if(args.Length > 0)
                    {
                        foreach (var arg in args)
                        {
                            string[] iavs = arg.Split('_',',');
                            IDAndValue iDAndValue = new IDAndValue();
                            iDAndValue.id = int.Parse(iavs[0]);
                            iDAndValue.value = int.Parse(iavs[1]);
                            iDAndValues.Add(iDAndValue);
                        }
                    }
                    info.SetValue(instance, iDAndValues);
                }

            }
            catch(Exception)
            {
                throw new Exception("请确认您填的数据是否符合表头所定义的数据类型");
            }
        }
    }
}
