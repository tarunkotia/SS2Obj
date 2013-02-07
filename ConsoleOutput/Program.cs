using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using SS2Obj;
using System.Data;
using System.Data.SqlClient;

namespace ConsoleOutput
{
    class Program
    {
        static void Main(string[] args)
        {            
            Console.Write("Please enter the file path you want to parse: ");
            string path = @"C:\Users\tkotia\Data\SAM_Exclusions_Public_Extract_13036.xlsx";//Console.ReadLine();
            //path = @"\\frxntnyc.frx2.com\UserDataDFS\MRL03\tkotia\My Documents\U_DriveData\FRI\Polaris\SAM_Exclusions_Public_Extract_13036\SAM_Exclusions_Public_Extract_13036.xlsx";
            SpreadSheet ss = new SpreadSheet(path);
            DataTable dt = ss.GetDataTable();

            Console.WriteLine(string.Format("Total rows selected:{0}",dt.Rows.Count));
            Console.WriteLine(string.Format("Total columns selected:{0}", dt.Columns.Count));


            Console.WriteLine("Dumping Data into database");
            DumpToDB(dt);
            Console.WriteLine("Finished dumping data");
            Console.ReadLine();
        }

        static SqlConnection Connection()
        {
            SqlConnection conn =
             new System.Data.SqlClient.SqlConnection("Data Source=frxsqldev;Initial Catalog=MuseUAT;Integrated Security=True;");

            conn.Open();

            //try
            //{
            //    //your Stuff                    
            //}
            //catch (SqlException)
            //{
            //    throw;
            //}
            //finally
            //{
            //    if (conn.State == ConnectionState.Open) conn.Close();
            //}
            
            return conn;
        }
        public static void DumpToDB(DataTable tbl)
        {
            System.Data.SqlClient.SqlCommand cmd = new System.Data.SqlClient.SqlCommand();

            SqlConnection conn = Connection();
            //SqlDataAdapter adpt = new SqlDataAdapter("Select * From SAM_Exclusions_Public_Extract;", conn);           

            //DataSet ds = new DataSet();
            //adpt.Fill(ds);
            tbl.TableName = "SAM_Exclusions_Public_Extract";

            SqlBulkCopy bulkCopy = new SqlBulkCopy(conn);           


            bulkCopy.DestinationTableName = tbl.TableName;
            try
            {
                bulkCopy.WriteToServer(tbl);
            }
            catch (Exception e) { Console.WriteLine(string.Format("Error dumping data:{0}", e.Message)); }
            
            conn.Close();
        }
    }
}
