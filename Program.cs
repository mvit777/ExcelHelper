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

            BuildMenu();
            //TestBuildDataForPivot(config);
            //TestAddRowTotal(config);
            //TestBuildPivot(config);
            RunAll(config);
            
            Console.ReadLine();

        }

        public static void BuildMenu()
        {
           
        }

        public static void RunAll(IConfigurationRoot config)
        {
            TestAddRowTotal(config);
            //TestBuildPivot(config);
            TestPivotAtOnce(config);
        }

        public static void TestAddRowTotal(IConfigurationRoot config)
        {
            Console.WriteLine("Test addRowTotal");
            string sql = @"SELECT TOP 10
                        ProductName
                          ,sum(quantity) as Bottles_sold
	                      ,sum(ExtendedPrice) as Subtotal_per_Wine
                      FROM [dbo].[Order Details Extended] 
                      group by ProductName ORDER BY sum(ExtendedPrice) DESC";
            var results = GetDataTable(config, sql, null);
            
            var file = AppDomain.CurrentDomain.BaseDirectory + "/output/rowtotal.xlsx";
            if (System.IO.File.Exists(file))
            {
                System.IO.File.Delete(file);
            }
            ExcelHelper eh = ExcelHelper.Create(file)
                        .Dump(results, "SHEET1")
                        //.DoPivotTable("SHEET1","SHEET2", new string[]{"ProductName"}, new string[]{"Bottles_sold"}, new string[]{"Subtotal_per_Wine" })
                        .AddRowTotal("SHEET1")
                        .Save();
            eh.Dispose();
            Console.WriteLine("file created in " + file);
            Console.WriteLine("End testAddRowTotal");
        }

        public static void TestBuildDataForPivot(IConfigurationRoot config)
        {
            string sql = "SELECT * FROM Invoices ORDER BY OrderDate";
            var results = GetDataTable(config, sql, null);
            var file = AppDomain.CurrentDomain.BaseDirectory + "/output/pivot.xlsx";
            ExcelHelper eh = ExcelHelper.Create(file)
                        .Dump(results, "SHEET1")
                        .Save();
            eh.Dispose();
        }

        public static void TestBuildPivot(IConfigurationRoot config)
        {
            Console.WriteLine("Begin TestBuildPivot");
            var file = AppDomain.CurrentDomain.BaseDirectory + "/output/pivot.xlsx";
            if (System.IO.File.Exists(file) != true)
            {
                TestBuildDataForPivot(config);
            }
            ExcelHelper eh = ExcelHelper.Create(file)
                                        .DoPivotTable("SHEET1", 
                                                        "SHEET2", 
                                                        new string[] { "ShipCountry" },
                                                        new string[] { "ProductName" },
                                                        new string[] { "ExtendedPrice" }
                                                     )
                                        .Save();
            eh.Dispose();
            Console.WriteLine("file created in " + file);
            Console.WriteLine("End testBuildPivot");

        }

        public static void TestPivotAtOnce(IConfigurationRoot config)
        {
            Console.WriteLine("Begin TestPivotAtOnce");
            string sql = "SELECT * FROM Invoices ORDER BY OrderDate";
            var results = GetDataTable(config, sql, null);
            var file = AppDomain.CurrentDomain.BaseDirectory + "/output/pivot_at_once.xlsx";
            if (System.IO.File.Exists(file))
            {
                System.IO.File.Delete(file);
            }
            ExcelHelper eh = ExcelHelper.Create(file)
                        .Dump(results, "SHEET1")
                        .DoPivotTable("SHEET1",
                                                        "SHEET2",
                                                        new string[] { "ShipCountry" },
                                                        new string[] { "ProductName" },
                                                        new string[] { "ExtendedPrice" }
                                                     )
                        .Save();
            eh.Dispose();
            Console.WriteLine("file created in " + file);
            Console.WriteLine("End testPivotAtOnce");
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
