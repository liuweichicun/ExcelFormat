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
                FieldBuilder fieldBuilder = typeBuilder.DefineField(varName.ToString(), typeof(bool), FieldAttributes.Public);
            }
            else if (varType.Trim().Equals("sbyte"))
            {
                FieldBuilder fieldBuilder = typeBuilder.DefineField(varName.ToString(), typeof(sbyte), FieldAttributes.Public);
            }
            else if  (varType.Trim().Equals("byte"))
            {
                FieldBuilder fieldBuilder = typeBuilder.DefineField(varName.ToString(), typeof(byte), FieldAttributes.Public);
            }
            else if  (varType.Trim().Equals("short"))
            {
                FieldBuilder fieldBuilder = typeBuilder.DefineField(varName.ToString(), typeof(short), FieldAttributes.Public);
            }
            else if (varType.Trim().Equals("ushort"))
            {
                FieldBuilder fieldBuilder = typeBuilder.DefineField(varName.ToString(), typeof(ushort), FieldAttributes.Public);
            }
            else if(varType.Trim().Equals("int"))
            {
                FieldBuilder fieldBuilder = typeBuilder.DefineField(varName.ToString(), typeof(int), FieldAttributes.Public);
            }
            else if (varType.Trim().Equals("uint"))
            {
                FieldBuilder fieldBuilder = typeBuilder.DefineField(varName.ToString(), typeof(uint), FieldAttributes.Public);
            }
            else if (varType.Trim().Equals("long"))
            {
                FieldBuilder fieldBuilder = typeBuilder.DefineField(varName.ToString(), typeof(long), FieldAttributes.Public);
            }
            else if (varType.Trim().Equals("ulong"))
            {
                FieldBuilder fieldBuilder = typeBuilder.DefineField(varName.ToString(), typeof(ulong), FieldAttributes.Public);
            }
            else if (varType.Trim().Equals("float"))
            {
                FieldBuilder fieldBuilder = typeBuilder.DefineField(varName.ToString(), typeof(float), FieldAttributes.Public);
            }
            else if (varType.Trim().Equals("double"))
            {
                FieldBuilder fieldBuilder = typeBuilder.DefineField(varName.ToString(), typeof(double), FieldAttributes.Public);
            }
            else if (varType.Trim().Equals("decimal"))
            {
                FieldBuilder fieldBuilder = typeBuilder.DefineField(varName.ToString(), typeof(decimal), FieldAttributes.Public);
            }
            else if (varType.Trim().Equals("string"))
            {
                FieldBuilder fieldBuilder = typeBuilder.DefineField(varName.ToString(), typeof(string), FieldAttributes.Public);
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
                    info.SetValue(instance, bool.Parse(v));
                }
                if (info.FieldType.ToString().Equals(typeof(sbyte).FullName))
                {
                    info.SetValue(instance, sbyte.Parse(v));
                }
                if (info.FieldType.ToString().Equals(typeof(byte).FullName))
                {
                    info.SetValue(instance, byte.Parse(v));
                }
                if (info.FieldType.ToString().Equals(typeof(short).FullName))
                {
                    info.SetValue(instance, short.Parse(v));
                }
                if (info.FieldType.ToString().Equals((typeof(ushort).FullName)))
                {
                    info.SetValue(instance, ushort.Parse(v));
                }
                if (info.FieldType.ToString().Equals((typeof(int).FullName)))
                {
                    info.SetValue(instance, int.Parse(v));
                }
                if (info.FieldType.ToString().Equals((typeof(uint).FullName)))
                {
                    info.SetValue(instance, uint.Parse(v));
                }
                if (info.FieldType.ToString().Equals((typeof(long).FullName)))
                {
                    info.SetValue(instance, long.Parse(v));
                }
                if (info.FieldType.ToString().Equals((typeof(ulong).FullName)))
                {
                    info.SetValue(instance, ulong.Parse(v));
                }
                if (info.FieldType.ToString().Equals((typeof(float).FullName)))
                {
                    info.SetValue(instance, float.Parse(v));
                }
                if (info.FieldType.ToString().Equals((typeof(double).FullName)))
                {
                    info.SetValue(instance, double.Parse(v));
                }
                if (info.FieldType.ToString().Equals((typeof(decimal).FullName)))
                {
                    info.SetValue(instance, decimal.Parse(v));
                }
                if (info.FieldType.ToString().Equals((typeof(string).FullName)))
                {
                    info.SetValue(instance, v);
                }
            }
            catch(Exception)
            {
                throw new Exception("请确认您填的数据是否符合表头所定义的数据类型");
            }
        }
    }
}
