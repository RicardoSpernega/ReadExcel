using ExcelDataReader;
using System;
using System.Data;
using System.Data.OleDb;
using System.IO;

namespace AppCid
{
    class Program
    {
        static void Main(string[] args)
        {
            string con = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Doc\Patologias_Gestão_Saúde-Analise.xlsx;Extended Properties='Excel 8.0;HDR=Yes;'";
            using (OleDbConnection connection = new OleDbConnection(con))
            {
                connection.Open();
                OleDbCommand command = new OleDbCommand("select * from [Patologias_Gestão_Saúde-Analise$]", connection);
                using (OleDbDataReader dr = command.ExecuteReader())
                {
                    while (dr.Read())
                    {
                        var row1Col0 = dr[0];
                        var row1Col1 = dr[1];
                        var row1Col2 = dr[2];

                        Console.WriteLine(row1Col0 +"\n" + row1Col1 + "\n" + row1Col2);
                    }
                }
            }
        }

    }

}
