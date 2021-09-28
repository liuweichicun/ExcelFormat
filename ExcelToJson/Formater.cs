namespace ExcelFormat
{
    /// <summary>
    /// 格式转换的父类
    /// </summary>
    public abstract class Formater
    {
        /// <summary>
        /// 调用执行format操作
        /// </summary>
        /// <param name="filePaths">文件的绝对路径</param>
        public abstract void Format(string[] filePaths,int codepage = 65001);
    }
}
