using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Runtime.InteropServices;

namespace Tools
{
    public class INI
    {
        [DllImport("kernel32")]
        private static extern int GetPrivateProfileString(string section, string key, string defVal, StringBuilder retVal, int size, string filePath);
        //section：要读取的段落名
        //key: 要读取的键
        //defVal: 读取异常的情况下的缺省值
        //retVal: key所对应的值，如果该key不存在则返回空值
        //size: 值允许的大小
        //filePath: INI文件的完整路径和文件名

        [DllImport("kernel32")]
        private static extern long WritePrivateProfileString(string section, string key, string val, string filePath);
        //section: 要写入的段落名
        //key: 要写入的键，如果该key存在则覆盖写入
        //val: key所对应的值
        //filePath: INI文件的完整路径和文件名

        string _path;
        public string Path
        {
            get { return _path; }
            set { _path = value; }
        }

        int _size = 255;
        public int Size
        {
            get { return _size; }
            set { _size = value; }
        }

        public INI(string path)
        {
            _path = path;
        }

        public string GetValue(string key, string section = "main")
        {
            StringBuilder value = new StringBuilder(255);
            GetPrivateProfileString(section, key, string.Empty, value, _size, _path);
            //string ss = value.ToString();
            if (value.ToString() == string.Empty)
                throw new Exception(string.Format("参数 {0} 不存在!", key));
            return value.ToString();
        }

        public void SetValue(string key, string value, string section = "main")
        {
            WritePrivateProfileString(section, key, value, _path);
        }

    }
}
