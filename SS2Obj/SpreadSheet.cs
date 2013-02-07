using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data.OleDb;
using System.Data;

namespace SS2Obj
{
    public class SpreadSheet
    {
        private string _path = string.Empty;
        public SpreadSheet(string path)
        {
            this._path = path;
        }

        public DataTable GetDataTable()
        {
            string ConnectionString = string.Format(@"Provider=Microsoft.ACE.OLEDB.12.0;
Data Source={0};
Extended Properties=Excel 5.0",_path);

            StringBuilder stbQuery = new StringBuilder();
            stbQuery.Append("SELECT * FROM [A1:ZZ118097]");
            OleDbDataAdapter adp = new OleDbDataAdapter(stbQuery.ToString(), ConnectionString);

            DataTable dtSchools = new DataTable();
            adp.Fill(dtSchools);

            return dtSchools;
        }
    }
}
