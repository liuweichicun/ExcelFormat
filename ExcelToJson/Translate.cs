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
        public static Type  FieldTranslate(TypeBuilder typeBuilder,string varName,string varType )
        {

            if(varType.StartsWith("list"))
            {
                FieldBuilder fieldBuilder = typeBuilder.DefineField(varName.Trim().ToString(), typeof(List<object>), FieldAttributes.Public);
            }
            else
            {
                FieldBuilder fieldBuilder = typeBuilder.DefineField(varName.Trim().ToString(), typeof(object), FieldAttributes.Public);
            }


            //if (varType.Trim().Equals("bool"))
            //{
            //    FieldBuilder fieldBuilder = typeBuilder.DefineField(varName.Trim().ToString(), typeof(bool), FieldAttributes.Public);
            //}
            //else if (varType.Trim().Equals("sbyte"))
            //{
            //    FieldBuilder fieldBuilder = typeBuilder.DefineField(varName.Trim().ToString(), typeof(sbyte), FieldAttributes.Public);
            //}
            //else if  (varType.Trim().Equals("byte"))
            //{
            //    FieldBuilder fieldBuilder = typeBuilder.DefineField(varName.Trim().ToString(), typeof(byte), FieldAttributes.Public);
            //}
            //else if  (varType.Trim().Equals("short"))
            //{
            //    FieldBuilder fieldBuilder = typeBuilder.DefineField(varName.Trim().ToString(), typeof(short), FieldAttributes.Public);
            //}
            //else if (varType.Trim().Equals("ushort"))
            //{
            //    FieldBuilder fieldBuilder = typeBuilder.DefineField(varName.Trim().ToString(), typeof(ushort), FieldAttributes.Public);
            //}
            //else if(varType.Trim().Equals("int"))
            //{
            //    FieldBuilder fieldBuilder = typeBuilder.DefineField(varName.Trim().ToString(), typeof(int), FieldAttributes.Public);
            //}
            //else if (varType.Trim().Equals("uint"))
            //{
            //    FieldBuilder fieldBuilder = typeBuilder.DefineField(varName.Trim().ToString(), typeof(uint), FieldAttributes.Public);
            //}
            //else if (varType.Trim().Equals("long"))
            //{
            //    FieldBuilder fieldBuilder = typeBuilder.DefineField(varName.Trim().ToString(), typeof(long), FieldAttributes.Public);
            //}
            //else if (varType.Trim().Equals("ulong"))
            //{
            //    FieldBuilder fieldBuilder = typeBuilder.DefineField(varName.Trim().ToString(), typeof(ulong), FieldAttributes.Public);
            //}
            //else if (varType.Trim().Equals("float"))
            //{
            //    FieldBuilder fieldBuilder = typeBuilder.DefineField(varName.Trim().ToString(), typeof(float), FieldAttributes.Public);
            //}
            //else if (varType.Trim().Equals("double"))
            //{
            //    FieldBuilder fieldBuilder = typeBuilder.DefineField(varName.Trim().ToString(), typeof(double), FieldAttributes.Public);
            //}
            //else if (varType.Trim().Equals("decimal"))
            //{
            //    FieldBuilder fieldBuilder = typeBuilder.DefineField(varName.Trim().ToString(), typeof(decimal), FieldAttributes.Public);
            //}
            //else if (varType.Trim().Equals("string"))
            //{
            //    FieldBuilder fieldBuilder = typeBuilder.DefineField(varName.Trim().ToString(), typeof(string), FieldAttributes.Public);
            //}
            //else if(varType.Trim().ToLower().Equals("idandvalue"))
            //{
            //    FieldBuilder fieldBuilder = typeBuilder.DefineField(varName.Trim().ToString(), typeof(IDAndValue), FieldAttributes.Public);
            //}
            //else if(varType.Trim().ToLower().Equals("list<int>"))
            //{
            //    FieldBuilder fieldBuilder = typeBuilder.DefineField(varName.Trim().ToString(), typeof(List<int>), FieldAttributes.Public);
            //}
            //else if(varType.Trim().ToLower().Equals("list<string>"))
            //{
            //    FieldBuilder fieldBuilder = typeBuilder.DefineField(varName.Trim().ToString(), typeof(List<string>), FieldAttributes.Public);
            //}
            //else if (varType.Trim().ToLower().Equals("list<idandvalue>"))
            //{
            //    FieldBuilder fieldBuilder = typeBuilder.DefineField(varName.Trim().ToString(), typeof(List<IDAndValue>), FieldAttributes.Public);
            //}
            //else if(varType.StartsWith("obj"))
            //{
            //    AssemblyName assemblyName = new AssemblyName();
            //    assemblyName.Name = "TempAssembly";
            //    AssemblyBuilder assemblyBuilder = AssemblyBuilder.DefineDynamicAssembly(assemblyName, AssemblyBuilderAccess.Run);
            //    ModuleBuilder moduleBuilder = assemblyBuilder.DefineDynamicModule("TempMoudle");
            //    TypeBuilder tb = moduleBuilder.DefineType("Obj", TypeAttributes.Public);
                
            //    string[] typeMarks = varType.Split(':');
            //    string[] types = typeMarks[1].Split('_');
            //    string[] varNameMark = varName.Split(':');
            //    string[] vn = varNameMark[1].Split('_');

            //    for (int i = 0; i < vn.Length; i++)
            //    {
            //        FieldTranslate(tb, vn[i], types[i]);
            //    }

            //    FieldBuilder fieldBuilder = typeBuilder.DefineField(varNameMark[0].Trim().ToString(),tb.CreateType(), FieldAttributes.Public);
            //}
            //else if(varType.StartsWith("list<obj>"))
            //{
            //    string[] varNameMark = varName.Split(':');
            //    FieldBuilder fieldBuilder = typeBuilder.DefineField(varNameMark[0].Trim(), typeof(List<Object>), FieldAttributes.Public);

            //    AssemblyName assemblyName = new AssemblyName();
            //    assemblyName.Name = "TempAssembly";
            //    AssemblyBuilder assemblyBuilder = AssemblyBuilder.DefineDynamicAssembly(assemblyName, AssemblyBuilderAccess.Run);
            //    ModuleBuilder moduleBuilder = assemblyBuilder.DefineDynamicModule("TempMoudle");
            //    TypeBuilder tb = moduleBuilder.DefineType("TObj", TypeAttributes.Public);

            //    string[] typeMarks = varType.Split(':');
            //    string[] types = typeMarks[1].Split('_');
            //    string[] vn = varNameMark[1].Split('_');

            //    for (int i = 0; i < vn.Length; i++)
            //    {
            //        FieldTranslate(tb, vn[i], types[i]);
            //    }
            //    return tb.CreateType();
            //}
            //else
            //{
            //    throw new Exception("检测到不支持的数据类型，请检查您定义的数据类型是否错误");
            //}
            return null;
        }

        public static void FieldInfoTranslate(object instance,FieldInfo info,object value,Type type)
        {
            try
            {
                string v = value.ToString();

                Console.WriteLine(info.FieldType.FullName + "====");
                if (info.FieldType.ToString().Equals(typeof(object).FullName))
                {

                   
                    info.SetValue(instance, value);
                }
                /*if (info.FieldType.ToString().Equals(typeof(bool).FullName))
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
                }*/
                //if (info.FieldType.ToString().Equals(typeof(IDAndValue).FullName))
                //{
                //    if (string.IsNullOrEmpty(v))
                //    {
                //        info.SetValue(instance, null);
                //    }
                //    else
                //    {
                //        IDAndValue iDAndValue = new IDAndValue();
                //        string[] args = v.Split('/', '_', ',');
                //        iDAndValue.id = int.Parse(args[0]);
                //        iDAndValue.value = int.Parse(args[1]);
                //        info.SetValue(instance, iDAndValue);
                //    }
                //}


                if (info.FieldType.FullName.ToString().Equals(typeof(List<object>).FullName))
                {



                    Console.WriteLine("Excute");
                    if (string.IsNullOrEmpty(v))
                    {
                        info.SetValue(instance, null);
                    }
                    else
                    {
                        List<object> li = new List<object>();
                        string[] args = v.Split('/', ',', '_');
                        if (args.Length > 0)
                        {
                            foreach (var item in args)
                            {
                                Console.WriteLine(item);
                                li.Add(item);
                            }
                        }
                        info.SetValue(instance, li);
                    }

                }


                //if (info.FieldType.FullName.ToString().Equals(typeof(List<string>).FullName))
                //{
                //    if (string.IsNullOrEmpty(v))
                //    {
                //        info.SetValue(instance, null);
                //    }
                //    else
                //    {
                //        List<string> ls = new List<string>();
                //        string[] args = v.Split('/', ',');
                //        if (args.Length > 0)
                //        {
                //            foreach (var item in args)
                //            {
                //                ls.Add(item.Trim());
                //            }
                //        }
                //        info.SetValue(instance, ls);
                //    }
                //}

                //if (info.FieldType.FullName.ToString().Equals(typeof(List<IDAndValue>).FullName))
                //{
                //    if (string.IsNullOrEmpty(v))
                //    {
                //        info.SetValue(instance, null);
                //    }
                //    else
                //    {
                //        List<IDAndValue> iDAndValues = new List<IDAndValue>();
                //        string[] args = v.Split('/', ';');
                //        if (args.Length > 0)
                //        {
                //            foreach (var arg in args)
                //            {
                //                string[] iavs = arg.Split('_', ',');
                //                IDAndValue iDAndValue = new IDAndValue();
                //                iDAndValue.id = int.Parse(iavs[0]);
                //                iDAndValue.value = int.Parse(iavs[1]);
                //                iDAndValues.Add(iDAndValue);
                //            }
                //        }
                //        info.SetValue(instance, iDAndValues);
                //    }
                //}

                //if (info.FieldType.FullName.ToString().Equals("Obj"))
                //{
                //    var obj = Activator.CreateInstance(info.FieldType);
                //    object[] valueMark = v.Split('_');

                //    var fieldInfo = obj.GetType().GetFields(BindingFlags.Public | BindingFlags.Instance);
                //    for (int i = 0; i < fieldInfo.Length; i++)
                //    {
                //        FieldInfoTranslate(obj, fieldInfo[i], valueMark[i], null);
                //    }
                //    info.SetValue(instance, obj);
                //}

                //if (info.FieldType.FullName.ToString().Equals(typeof(List<Object>).FullName))
                //{

                //    string[] args = v.Split('/', ';');
                //    List<Object> objs = new List<object>();
                //    if (args.Length > 0)
                //    {
                //        foreach (var arg in args)
                //        {
                //            string[] iavs = arg.Split('_', ',');
                //            var obj = Activator.CreateInstance(type);
                //            FieldInfo[] infos = obj.GetType().GetFields(BindingFlags.Public | BindingFlags.Instance);
                //            for (int i = 0; i < iavs.Length; i++)
                //            {
                //                FieldInfoTranslate(obj, infos[i], iavs[i], null);
                //            }
                //            objs.Add(obj);
                //        }
                //    }
                //    info.SetValue(instance, objs);
                //}

            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                throw new Exception("请确认您填的数据是否符合表头所定义的数据类型");
            }
        }


        static void Covert(object instance,FieldInfo info,object value,Type type)
        {
            string v = value.ToString();
            Dictionary<string, Type> d = new Dictionary<string, Type>();
            if (info.FieldType.ToString().Equals((typeof(long).FullName)))
            {
                string aaa = info.FieldType.ToString();
                if (string.IsNullOrEmpty(v))
                    info.SetValue(instance, 0);
                else
                   info.SetValue(instance, v);
            }
        }
    }
}
