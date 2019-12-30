using System.Collections.Generic;

namespace ETModel
{
    public class ExcelMD5Info
    {
        public Dictionary<string, string> FileMD5 = new Dictionary<string, string>();

        public string Get(string fileName)
        {
            string md5 = string.Empty;
            this.FileMD5.TryGetValue(fileName, out md5);
            return md5;
        }

        public void Add(string fileName, string md5)
        {
            this.FileMD5[fileName] = md5;
        }
    }
}