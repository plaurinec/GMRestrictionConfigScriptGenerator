using Datapac.Posybe.ERPInterface;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.ExtendedProperties;
using GMRestrictionConfigScriptGenerator.Helpers;
using GMRestrictionConfigScriptGenerator.Models;
using Microsoft.EntityFrameworkCore;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Serilog;
using StoreMovementsScriptor.Helpers;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Threading;

namespace StoreMovementsScriptor
{
    class Program
    {
        static void Main(string[] args)
        {
            // check if already running
            if (IsAppRunning()) return;
            // load configuration
            IConfiguration config = new ConfigurationBuilder()
            .AddJsonFile("appsettings.json")
            .Build();

            var optionsBuilder = new DbContextOptionsBuilder<PosybeContext>();
            optionsBuilder.UseSqlServer(Datapac.CryptUnit.KX_Decrypt(config.GetConnectionString("PosybeConnection"), "PosybeConnection"));
            // setup logger
            Log.Logger = new LoggerConfiguration()
            .ReadFrom.Configuration(config)
            .CreateLogger();

            Log.Information($"App started ..");
            Environment.ExitCode = 0;
            bool generateEnabled = false, exportTemplateEnabled = false;
            string inputFileName = string.Empty, outputFileName = string.Empty, mappingFile = "Posybe2ERPMapping.xml";

            var o = args.ToList();
            var p = o.Select(p => p.ToLower()).ToList();

            if (o.Count > 0)
            {
                generateEnabled = p[0] == "/generatesql";
                exportTemplateEnabled = p[0] == "/exporttemplate";
            }
            //exportTemplateEnabled = j != -1;

            if (o.Count > 1) inputFileName = o[1];
            if (o.Count > 2) outputFileName = o[2];            
            if(o.Count>3) mappingFile = o[3];
            if (!generateEnabled && !exportTemplateEnabled || (generateEnabled && (string.IsNullOrEmpty(inputFileName) || string.IsNullOrEmpty(outputFileName))) || (exportTemplateEnabled && string.IsNullOrEmpty(inputFileName)))
            {
                Log.Information($"Wrong input parameters defined. ");
                Console.WriteLine("No arguments specified. !!");
                Console.WriteLine();
                Console.WriteLine("Available arguments:");
                Console.WriteLine("Commands:");
                Console.WriteLine("   /generatesql       - Generate SQL script from excel, input parameters [excel file path] [output file path] [mapping file- default name Posybe2ERPMapping.xml]");
                Console.WriteLine("   /exporttemplate    - Generate Excel file from database, input parameters [excel file path to create] (ConnectionString PosybeConnection must be defined in appsettings.json)");
                Environment.ExitCode = 1;
            }

            try
            {
                if (generateEnabled)
                {
                    Log.Information($"Generating SQL file {outputFileName}");
                    if (File.Exists(outputFileName)) File.Delete(outputFileName);
                    string result = string.Empty;
                    var ex = new Export(null, mappingFile, null);
                    var ei = new ExcelImport();
                    int dc = 0;
                    int errors = 0;
                    int[] sposob = new int[2];
                    ei.ImportMovementsFromExcel(inputFileName, (a) =>
                     {
                         try
                         {
                             dc++;
                             var name = a.Operation == 1 ? "Issue" : "Receipt";
                             Mapping mapping = null;
                             if (a.IsExternalCode)
                             {
                                 mapping = ex.GoodsMovementConversionBack(a.Type, a.Operation, a.ExternalCode, 0);
                                 //if (mapping == null) mapping = new Mapping() { PosybeCode = "0" };
                                 if (mapping == null) throw new Exception($"Cannot map code from external code {a.ExternalCode} for {name} item name {a.Name}");
                             }
                             else if (!string.IsNullOrEmpty(a.SubCode))
                             {
                                 mapping = new Mapping() { PosybeCode = a.SubCode };
                             }
                             else
                             {
                                 throw new Exception($"No mapping code defined for {name} item name {a.Name}");
                             }
                             string sql = $"If Not Exists( Select * From [Posybe].[dbo].[F_SKLAD_POHYBY_SPOSOBY] Where OPERACIA={a.Operation} And SPOSOB={sposob[a.Operation]} And TYP={a.Type}) \r\n" +
                                 $"Insert[Posybe].[dbo].[F_SKLAD_POHYBY_SPOSOBY] (OPERACIA,SPOSOB,POVOLENY,SKRATKA,NAZOV,MAPOVANIE_POHYB,SQL_FILTER,SQL_FILTER_PARTNER,TYP) Values({a.Operation}, {sposob[a.Operation]}, 0, '', N'', 0, Null, Null, {a.Type})";

                             string enabled = string.Join("'',''", a.Items.Where(p => p.State).Select(p => p.Prefix));
                             string disabled = string.Join("'',''", a.Items.Where(p => !p.State).Select(p => p.Prefix));
                             int povoleny = 1;
                             string filter = $"\r\nSQL_FILTER = 'And s.TOVAR In ( Select us.TOVAR From UPL_SKLAD us Left Outer Join UPL_SUBCATEGORIES usc on us.SUBCATEGORY_ID = usc.ID Where usc.NOTES ##In## (\r\n''";

                             if (enabled.Length == 0)
                             {
                                 povoleny = 0;
                                 filter = string.Empty;
                             }
                             else if (disabled.Length == 0)
                             {
                                 povoleny = 1;
                                 filter = string.Empty;
                             }

                             if (enabled.Length < disabled.Length)
                             {
                                 povoleny = 1;
                                 filter = filter.Replace("##In##", "In");
                                 filter += enabled + "'') )";
                             }
                             else
                             {
                                 povoleny = 1;
                                 filter = filter.Replace("##In##", "Not In");
                                 filter += disabled + "'') )";
                             }

                             sql += $"\r\n\r\nUpdate[Posybe].[dbo].[F_SKLAD_POHYBY_SPOSOBY] Set POVOLENY = {povoleny}, SKRATKA = '{a.ShortName}', NAZOV = N'{a.Name}', MAPOVANIE_POHYB = {mapping.PosybeCode},";


                             if (a.Limits[0].HasValue && a.Limits[0].Value && a.Limits[1].HasValue && a.Limits[1].Value)
                             {
                                 throw new Exception($"Cannot define both item limits for {name} item name {a.Name}");
                             }
                             if (a.Limits[2].HasValue && a.Limits[2].Value && a.Limits[3].HasValue && a.Limits[3].Value)
                             {
                                 throw new Exception($"Cannot define both partner limits for {name} item name {a.Name}");
                             }
                             if (a.Limits[4].HasValue && a.Limits[4].Value && a.Limits[5].HasValue && a.Limits[5].Value)
                             {
                                 throw new Exception($"Cannot define both partner ses limits for {name} item name {a.Name}");
                             }
                             string itemLimits = string.Empty;
                             if (a.Limits[0].HasValue && a.Limits[0].Value) itemLimits += " And IsNull(ctd.DPH_T, -1) <> -1";
                             if (a.Limits[1].HasValue && a.Limits[1].Value) itemLimits += "  And IsNull(ctd.DPH_T, -1) = -1";

                             if (string.IsNullOrEmpty(itemLimits))
                             {
                                 if (string.IsNullOrEmpty(filter)) filter = "\r\nSQL_FILTER = Null";
                                 else filter += "'";
                             }
                             else
                             {
                                 filter += itemLimits + "'";
                             }
                             sql += filter + ",";

                             filter = "\r\nSQL_FILTER_PARTNER = ";

                             string PartnerLimits = string.Empty;
                             if (a.Limits[2].HasValue && a.Limits[2].Value) PartnerLimits += " And IsNull(spe.PRISTUP_K_TOVAROM, 0) <> 2";
                             if (a.Limits[3].HasValue && a.Limits[3].Value) PartnerLimits += " And IsNull(spe.PRISTUP_K_TOVAROM, 0) = 2";

                             string PartnerLimitsSes = string.Empty;
                             if (a.Limits[4].HasValue && a.Limits[4].Value) PartnerLimitsSes += " And sp.PARTNER Between 10000 And 99999";
                             if (a.Limits[5].HasValue && a.Limits[5].Value) PartnerLimitsSes += " And sp.PARTNER In (Select CISLO_CS From C_ID_CS)";

                             if (string.IsNullOrEmpty(PartnerLimits) && string.IsNullOrEmpty(PartnerLimitsSes))
                             {
                                 filter += "Null";
                             }
                             else
                             {
                                 filter += "'";
                                 if (!string.IsNullOrEmpty(PartnerLimits)) filter += PartnerLimits;
                                 if (!string.IsNullOrEmpty(PartnerLimitsSes)) filter += PartnerLimitsSes;
                                 filter += "'";
                             }
                             sql += filter;
                             sql += $"\r\nWhere OPERACIA = {a.Operation} And SPOSOB = {sposob[a.Operation]} And TYP = {a.Type}";

                             sposob[a.Operation]++;
                             result += "\r\n\r\n" + sql;
                         }
                         catch(Exception ex)
                         {
                             errors++;
                             Console.WriteLine($"Error: {ex.Message}");
                             Log.Error(ex, $"Unexpected error while reading excel file {inputFileName}");
                         }
                     });
                    if (dc == 0)
                    {                        
                        throw new Exception($"No data found in excel file {inputFileName}");                        
                    }
                    else if (errors == 0)
                    {
                        Environment.ExitCode = 1;                        
                        File.AppendAllText(outputFileName, result);
                        Console.WriteLine($"SQL script {outputFileName} generated. Records count: {dc}");
                        Log.Information($"SQL script {outputFileName} generated. Records count: {dc}");
                    }                    
                }
                else if(exportTemplateEnabled)
                {
                    try
                    {
                        using (var db = new PosybeContext(optionsBuilder.Options))
                        {
                            var k = db.FSkladPohybySposoby.AsNoTracking().Where(p => p.Operacia == 1 && p.Typ == 0);
                            var rows = new List<SheetRow>();
                            var r = new SheetRow(); r.Columns.Add(new SheetCellValue()); r.Columns.Add(new SheetCellValue() { Value = "External system movement", Bold = true });
                            rows.Add(r);
                            r = new SheetRow(); r.Columns.Add(new SheetCellValue()); r.Columns.Add(new SheetCellValue() { Value = "Posybe movement", Bold = true });
                            rows.Add(r);
                            foreach (var cc in k)
                            {
                                r.Columns.Add(new SheetCellValue() { Value = cc.MapovaniePohyb.ToString() });
                            }
                            r = new SheetRow(); r.Columns.Add(new SheetCellValue()); r.Columns.Add(new SheetCellValue() { Value = "Posybe movement operation", Bold = true });
                            rows.Add(r);
                            foreach (var cc in k)
                            {
                                r.Columns.Add(new SheetCellValue() { Value = cc.Skratka });
                            }
                            r = new SheetRow(); r.Columns.Add(new SheetCellValue()); r.Columns.Add(new SheetCellValue() { Value = "Item limitation - COCA", Bold = true });
                            rows.Add(r);
                            r = new SheetRow(); r.Columns.Add(new SheetCellValue()); r.Columns.Add(new SheetCellValue() { Value = "Item limitation - CODO", Bold = true });
                            rows.Add(r);
                            r = new SheetRow(); r.Columns.Add(new SheetCellValue()); r.Columns.Add(new SheetCellValue() { Value = "Partner limitation - COCA", Bold = true });
                            rows.Add(r);
                            r = new SheetRow(); r.Columns.Add(new SheetCellValue()); r.Columns.Add(new SheetCellValue() { Value = "Partner limitation - CODO", Bold = true });
                            rows.Add(r);
                            r = new SheetRow(); r.Columns.Add(new SheetCellValue()); r.Columns.Add(new SheetCellValue() { Value = "Partner limitation - only SeS", Bold = true });
                            rows.Add(r);
                            r = new SheetRow(); r.Columns.Add(new SheetCellValue()); r.Columns.Add(new SheetCellValue() { Value = "Partner limitation - same SeS", Bold = true });
                            rows.Add(r);
                            r = new SheetRow(); r.Columns.Add(new SheetCellValue() { Value = "MC Prefix", Bold = true }); r.Columns.Add(new SheetCellValue() { Value = "Movement/MC name", Bold = true });
                            rows.Add(r);
                            foreach (var cc in k)
                            {
                                r.Columns.Add(new SheetCellValue() { Value = cc.Nazov });
                            }

                            var m = new List<string>();
                            foreach (var cc in k)
                            {
                                if (!string.IsNullOrEmpty(cc.SqlFilter))
                                {
                                    if (cc.SqlFilter.Contains("UPL_SUBCATEGORIES"))
                                    {
                                        int i = cc.SqlFilter.IndexOf("'");
                                        int j = cc.SqlFilter.LastIndexOf("'");
                                        if (j >= 0 && i >= 0)
                                        {
                                            foreach (var ca in cc.SqlFilter.Substring(i, j - i).Replace("'", "").Split(','))
                                            {
                                                var ko = ca.Trim();
                                                if (!m.Contains(ko)) m.Add(ko);
                                            };
                                        }
                                    }
                                }
                            }
                            foreach (var sku in m)
                            {
                                foreach (var boi in db.UplSubcategories.AsNoTracking().Where(p => p.Notes.Contains(sku)))
                                {
                                    r = new SheetRow(); r.Columns.Add(new SheetCellValue() { Value = sku }); r.Columns.Add(new SheetCellValue() { Value = boi.Title.Trim() });
                                    rows.Add(r);
                                }
                            }
                            //issue
                            var sheets = new List<SheetDefinition<SheetRow>>();
                            var c = new SheetDefinition<SheetRow>() { Name = "Issue", Fields = new List<SpreadsheetField>(), Objects = rows };
                            sheets.Add(c);


                            k = db.FSkladPohybySposoby.AsNoTracking().Where(p => p.Operacia == 0 && p.Typ == 0);
                            rows = new List<SheetRow>();
                            r = new SheetRow(); r.Columns.Add(new SheetCellValue()); r.Columns.Add(new SheetCellValue() { Value = "External system movement", Bold = true });
                            rows.Add(r);
                            r = new SheetRow(); r.Columns.Add(new SheetCellValue()); r.Columns.Add(new SheetCellValue() { Value = "Posybe movement", Bold = true });
                            rows.Add(r);
                            foreach (var cc in k)
                            {
                                r.Columns.Add(new SheetCellValue() { Value = cc.MapovaniePohyb.ToString() });
                            }
                            r = new SheetRow(); r.Columns.Add(new SheetCellValue()); r.Columns.Add(new SheetCellValue() { Value = "Posybe movement operation", Bold = true });
                            rows.Add(r);
                            foreach (var cc in k)
                            {
                                r.Columns.Add(new SheetCellValue() { Value = cc.Skratka });
                            }
                            r = new SheetRow(); r.Columns.Add(new SheetCellValue()); r.Columns.Add(new SheetCellValue() { Value = "Item limitation - COCA", Bold = true });
                            rows.Add(r);
                            r = new SheetRow(); r.Columns.Add(new SheetCellValue()); r.Columns.Add(new SheetCellValue() { Value = "Item limitation - CODO", Bold = true });
                            rows.Add(r);
                            r = new SheetRow(); r.Columns.Add(new SheetCellValue()); r.Columns.Add(new SheetCellValue() { Value = "Partner limitation - COCA", Bold = true });
                            rows.Add(r);
                            r = new SheetRow(); r.Columns.Add(new SheetCellValue()); r.Columns.Add(new SheetCellValue() { Value = "Partner limitation - CODO", Bold = true });
                            rows.Add(r);
                            r = new SheetRow(); r.Columns.Add(new SheetCellValue()); r.Columns.Add(new SheetCellValue() { Value = "Partner limitation - only SeS", Bold = true });
                            rows.Add(r);
                            r = new SheetRow(); r.Columns.Add(new SheetCellValue()); r.Columns.Add(new SheetCellValue() { Value = "Partner limitation - same SeS", Bold = true });
                            rows.Add(r);
                            r = new SheetRow(); r.Columns.Add(new SheetCellValue() { Value = "MC Prefix", Bold = true }); r.Columns.Add(new SheetCellValue() { Value = "Movement/MC name", Bold = true });
                            rows.Add(r);
                            foreach (var cc in k)
                            {
                                r.Columns.Add(new SheetCellValue() { Value = cc.Nazov });
                            }

                            m = new List<string>();
                            foreach (var cc in k)
                            {
                                if (!string.IsNullOrEmpty(cc.SqlFilter))
                                {
                                    if (cc.SqlFilter.Contains("UPL_SUBCATEGORIES"))
                                    {
                                        int i = cc.SqlFilter.IndexOf("'");
                                        int j = cc.SqlFilter.LastIndexOf("'");
                                        if (j >= 0 && i >= 0)
                                        {
                                            foreach (var ca in cc.SqlFilter.Substring(i, j - i).Replace("'", "").Split(','))
                                            {
                                                var ko = ca.Trim();
                                                if (!m.Contains(ko)) m.Add(ko);
                                            };
                                        }
                                    }
                                }
                            }
                            foreach (var sku in m)
                            {
                                foreach (var boi in db.UplSubcategories.AsNoTracking().Where(p => p.Notes.Contains(sku)))
                                {
                                    r = new SheetRow(); r.Columns.Add(new SheetCellValue() { Value = sku }); r.Columns.Add(new SheetCellValue() { Value = boi.Title.Trim() });
                                    rows.Add(r);
                                }
                            }

                            c = new SheetDefinition<SheetRow>() { Name = "Receipt", Fields = new List<SpreadsheetField>(), Objects = rows };
                            sheets.Add(c);

                            // genrate excel file from table
                            Spreadsheet.Create<SheetRow>(inputFileName, sheets.ToArray());
                            Console.WriteLine($"Excel file {inputFileName} generated.");
                            Log.Information($"Excel file {inputFileName} generated.");
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Error: {ex.Message}");
                        Log.Error(ex, $"Unexpected error while generating excel file {inputFileName}");
                    }
                }
            }
            catch(Exception ex)
            {
                Environment.ExitCode = 1;
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        private static Mutex m_appRunnigMutex;
        /// <summary>
        /// check if app with key appName is running, global means check RDP sessions
        /// </summary>
        /// <param name="appName"></param>
        /// <param name="global"></param>
        /// <returns></returns>
        public static bool IsAppRunning(string appName = null, bool global = true)
        {
            if (string.IsNullOrEmpty(appName)) appName = Assembly.GetExecutingAssembly().GetName().Name;
            if (m_appRunnigMutex != null) return false;
            bool cn;
            m_appRunnigMutex = new Mutex(true, global ? @"Global\" + appName : appName, out cn);
            return !cn;
        }
    }    
}
