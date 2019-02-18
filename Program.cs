using Microsoft.Deployment.WindowsInstaller;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace VerifyMsi
{
    class Program
    {
        static Dictionary<string, string> baseline = new Dictionary<string, string>
        {
            { "F_Microsoft.Data.Edm52", "5.2.0.51212" },
            { "F_Microsoft.Data.Edm560", "5.6.0.61587" },
            { "F_Microsoft.Data.Edm562", "5.6.2.61936" },
            { "F_Microsoft.Data.Edm564", "5.6.4.62175" },
            { "F_Microsoft.Data.OData52", "5.2.0.51212" },
            { "F_Microsoft.Data.OData560", "5.6.0.61587" },
            { "F_Microsoft.Data.OData562", "5.6.2.61936" },
            { "F_Microsoft.Data.OData564", "5.6.4.62175" },
            { "F_Microsoft.Data.Services.Client52", "5.2.0.51212" },
            { "F_Microsoft.Data.Services.Client560", "5.6.0.61587" },
            { "F_Microsoft.Data.Services.Client562", "5.6.2.61936" },
            { "F_Microsoft.Data.Services.Client564", "5.6.4.62175" },
            { "Microsoft.Data.Edm", "5.6.2.61936" },
            { "Microsoft.Data.OData", "5.6.2.61936" },
            { "Microsoft.Data.Services", "5.6.2.61936" },
            { "F_System.Spatial52", "5.2.0.51212" },
            { "F_System.Spatial560", "5.6.0.61587" },
            { "F_System.Spatial562", "5.6.2.61936" },
            { "F_System.Spatial564", "5.6.4.62175" },
            { "System.Spatial.dll", "5.6.2.61936" },
        };

        static void Main(string[] args)
        {
            try
            {
                if (args.Length == 0)
                {
                    Console.WriteLine(@"Usage: VerifyMsi.exe \\reddog\Builds\branches\git_aapt_antares_websites_master\80.0.7.48");
                    return;
                }

                VerifyMsi(Path.Combine(args[0], @"retail-amd64\Hosting\WebHosting.msi"));
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
            }
        }

        static void VerifyMsi(string msiFile)
        {
            var files = new SortedList<string, string>();
            using (var database = new Database(msiFile, DatabaseOpenMode.ReadOnly))
            {
                using (var view = database.OpenView(database.Tables["File"].SqlSelectString))
                {
                    view.Execute();
                    foreach (var rec in view)
                    {
                        using (rec)
                        {
                            var file = rec.GetString("File");
                            if (baseline.ContainsKey(file))
                            {
                                files.Add(rec.GetString("File"), rec.GetString("Version"));
                            }
                        }
                    }
                }
            }

            foreach (var file in files)
            {
                var result = "mismatched!";
                ConsoleColor color = ConsoleColor.Red;
                if (baseline[file.Key] == file.Value)
                {
                    color = ConsoleColor.Green;
                    result = "matched";
                }

                using (new ApplyFGColor(color))
                {
                    Console.WriteLine("{0} = {1} ... {2}", file.Key, file.Value, result);
                }
            }
        }

        class ApplyFGColor : IDisposable
        {
            ConsoleColor _prev;

            public ApplyFGColor(ConsoleColor color)
            {
                _prev = Console.ForegroundColor;
                Console.ForegroundColor = color;
            }

            public void Dispose()
            {
                Console.ForegroundColor = _prev;
            }
        }
    }
}
