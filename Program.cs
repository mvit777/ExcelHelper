using Microsoft.Extensions.Configuration;
using ranxlib;
using System;
using System.Collections;
using System.Data;
using System.Data.SqlClient;

namespace ranx
{
    class Program
    {
        static void Main(string[] args)
        {
            //var env = Environment.GetEnvironmentVariable("ASPNETCORE_ENVIRONMENT");
            var builder = new ConfigurationBuilder()
                .AddJsonFile($"appsettings.json", true, true)
                //.AddJsonFile($"appsettings.{env}.json", true, true)
                .AddEnvironmentVariables();

            var config = builder.Build();

            Console.WriteLine("Test addRowTotal");
            TestAddRowTotal(config);
            
            Console.ReadLine();

        }

        public static void TestAddRowTotal(IConfigurationRoot config)
        {     
            string sql = @"SELECT TOP 10
                        ProductName
                          ,sum(quantity) as Bottles_sold
	                      ,sum(ExtendedPrice) as Subtotal_per_Wine
                      FROM [dbo].[Order Details Extended] 
                      group by ProductName ORDER BY sum(ExtendedPrice) DESC";
            var results = GetDataTable(config, sql, null);
            
            var file = AppDomain.CurrentDomain.BaseDirectory + "/output/rowtotal.xlsx";
            ExcelHelper eh = ExcelHelper.Create(file);
            eh.Dump(results, "SHEET1");
            eh.AddRowTotal("SHEET1");
            //eh.DoPivotTable("SHEET1", "SHEET2", new string[]{"ProductName"}, new string[] {"Subtotal_per_Wine" });
            eh.Save();
            eh.Dispose();
            Console.WriteLine("file created in " + file);
            Console.WriteLine("End testAddRowTotal");
        }

        public static DataTable GetDataTable(IConfigurationRoot config, string sql, Hashtable _params = null)
        {
            //Console.WriteLine(config["ConnectionString"]);
            SqlConnection conn = new SqlConnection(config["ConnectionString"]);
            DataTable result = new DataTable();
            SqlCommand cmd = null;
            try
            {
                cmd = new SqlCommand(sql, conn);
                cmd.Connection.Open();
                if (_params != null)
                {

                }
                SqlDataReader reader = cmd.ExecuteReader();
                result.Load(reader);
                reader.Close();
                cmd.Connection.Close();
            }
            catch(Exception ex)
            {
                Console.WriteLine(ex.Message);
                cmd.Connection.Close();
            }
            

            return result;
        }

    }
}
