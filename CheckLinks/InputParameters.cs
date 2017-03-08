using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CheckLinks
{
    class InputParameters
    {
        private String fileName;
        private String tableName;
        private String linkField;
        private String okField;

        public InputParameters(String fileName, String tableName, String linkField, String okField)
        {
            this.fileName = fileName;
            this.tableName = tableName;
            this.linkField = linkField;
            this.okField = okField;
        }

        public static InputParameters FromArray(String[] array)
        {
            if (array == null) throw new ArgumentException("Array must not be null");
            if (array.Length != 4) throw new ArgumentException("Array must have length 4");

            return new InputParameters(array[0], array[1], array[2], array[3]);
        }

        public String FileName
        {
            get { return fileName; }
        }

        public String TableName
        {
            get { return tableName; }
        }

        public String LinkField
        {
            get { return linkField; }
        }

        public String OkField
        {
            get { return okField; }
        }
    }
}
