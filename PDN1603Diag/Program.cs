using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.IO.Compression;
using System.Windows.Forms;

namespace PDN1603Diag
{
    class Program
    {
        [STAThread]
        static void Main(string[] args)
        {
            string appTitle = "toe_head2001's diagnostic tool for Error 1603 in the paint.net installation";
            Console.Title = appTitle + " - v" + Application.ProductVersion;

            Console.ForegroundColor = ConsoleColor.Cyan;
            Console.WriteLine(appTitle);
            Console.ForegroundColor = ConsoleColor.White;
            Console.WriteLine(string.Empty);

            string relPath = Directory.GetCurrentDirectory();
            string searchPattern = "paint.net.*.install.*";
            string[] installFiles = Directory.GetFiles(relPath, searchPattern);
            string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);
            string msiPath = $@"\PaintDotNetMsi\PaintDotNet_x{(Environment.Is64BitOperatingSystem ? "64" : "86")}.msi";

            if (installFiles.Length <= 0)
            {
                NoFile("installation");
                return;
            }

            bool exePresent = false;
            foreach (string file in installFiles)
            {
                if (Path.GetExtension(file).Equals(".exe", StringComparison.OrdinalIgnoreCase))
                {
                    exePresent = true;
                    Console.WriteLine("Creating a MSI file to work with...");
                    Process.Start(file, "/CreateMsi /auto").WaitForExit();
                    if (File.Exists(desktopPath + msiPath))
                    {
                        Console.WriteLine("Done.");
                    }
                    else
                    {
                        NoFile("MSI");
                        return;
                    }
                    Console.WriteLine(string.Empty);
                    break;
                }
            }

            if (!exePresent)
            {
                bool zipPresent = false;
                foreach (string file in installFiles)
                {
                    if (Path.GetExtension(file).Equals(".zip", StringComparison.OrdinalIgnoreCase))
                    {
                        zipPresent = true;
                        ZipFile.ExtractToDirectory(file, relPath);
                        break;
                    }
                }

                if (!zipPresent)
                {
                    NoFile("installation");
                    return;
                }

                installFiles = Directory.GetFiles(relPath, searchPattern);

                foreach (string file in installFiles)
                {
                    if (Path.GetExtension(file).Equals(".exe", StringComparison.OrdinalIgnoreCase))
                    {
                        exePresent = true;
                        Console.WriteLine("Creating a MSI file to work with...");
                        Process.Start(file, "/CreateMsi /auto").WaitForExit();
                        if (File.Exists(desktopPath + msiPath))
                        {
                            Console.WriteLine("Done.");
                        }
                        else
                        {
                            NoFile("MSI");
                            return;
                        }
                        Console.WriteLine(string.Empty);
                        if (File.Exists(file))
                        {
                            try
                            {
                                File.Delete(file);
                            }
                            catch
                            {
                            }
                        }
                        break;
                    }
                }

                if (!exePresent)
                {
                    NoFile("installation");
                    return;
                }

            }

            Console.WriteLine("Running the MSI with logging enabled...");

            ProcessStartInfo startInfo = new ProcessStartInfo();
            startInfo.FileName = "msiexec.exe";
            startInfo.WorkingDirectory = desktopPath;
            startInfo.Arguments = $"/i {msiPath} /L*V pdn.log";
            Process process = new Process();
            process.StartInfo = startInfo;
            process.Start();
            process.WaitForExit();

            Console.WriteLine("Done.");
            Console.WriteLine(string.Empty);

            string msiFolder = desktopPath + @"\PaintDotNetMsi";
            if (Directory.Exists(msiFolder))
            {
                try
                {
                    Directory.Delete(msiFolder, true);
                }
                catch
                {
                }
            }

            string logPath = desktopPath + @"\pdn.log";
            if (!File.Exists(logPath))
            {
                NoFile("log");
                return;
            }

            Console.WriteLine("---------- Below are the relevant lines from the installation log ----------");
            Console.WriteLine(string.Empty);
            Console.WriteLine(string.Empty);

            bool found1603 = false;
            List<string> errorLines = new List<string>();
            string lineOutput;
            string[] logText = File.ReadAllLines(logPath);
            for (int line = 0; line < logText.Length; line++)
            {
                if (logText[line].Contains("1603") || logText[line].Contains("Return value 3."))
                {
                    found1603 = true;
                    try
                    {
                        for (int beforeLine = line - 4; beforeLine < line; beforeLine++)
                        {
                            lineOutput = $"> Line {beforeLine + 1}: {logText[beforeLine]}";
                            errorLines.Add(lineOutput);
                            Console.WriteLine(lineOutput);
                        }
                    }
                    catch
                    {
                    }

                    lineOutput = $"> Line {line + 1}: {logText[line]}";
                    errorLines.Add(lineOutput);
                    Console.ForegroundColor = ConsoleColor.Yellow;
                    Console.WriteLine(lineOutput);
                    Console.ForegroundColor = ConsoleColor.White;

                    try
                    {
                        for (int afterLine = line + 1; afterLine <= line + 4; afterLine++)
                        {
                            lineOutput = $"> Line {afterLine + 1}: {logText[afterLine]}";
                            errorLines.Add(lineOutput);
                            Console.WriteLine(lineOutput);
                        }
                    }
                    catch
                    {
                        lineOutput = "> -EOF-";
                        errorLines.Add(lineOutput);
                        Console.WriteLine(lineOutput);
                    }

                    errorLines.Add(string.Empty);
                    Console.WriteLine(string.Empty);
                    Console.WriteLine(string.Empty);
                    Console.WriteLine(string.Empty);
                }
            }

            if (!found1603)
            {
                bool pdnSuccess = false;
                for (int line = 0; line < logText.Length; line++)
                {
                    if (logText[line].Contains("paint.net -- Configuration completed successfully.") || 
                        logText[line].Contains("paint.net -- Installation completed successfully."))
                    {
                        lineOutput = "There where no 1603 errors, and it appears paint.net was successfully installed.";
                        errorLines.Add(lineOutput);
                        Console.ForegroundColor = ConsoleColor.Green;
                        Console.WriteLine(lineOutput);
                        Console.ForegroundColor = ConsoleColor.White;
                        pdnSuccess = true;
                        break;
                    }
                }

                if (!pdnSuccess)
                {
                    lineOutput = "There where no 1603 errors, but something else went wrong in the installation.";
                    errorLines.Add(lineOutput);
                    Console.ForegroundColor = ConsoleColor.Yellow;
                    Console.WriteLine(lineOutput);
                    Console.ForegroundColor = ConsoleColor.White;
                }

                Console.WriteLine(string.Empty);
                Console.WriteLine(string.Empty);
            }

            Console.WriteLine("---------- Above are the relevant lines from the installation log----------");
            Console.WriteLine(string.Empty);
            Console.WriteLine(string.Empty);

            Console.WriteLine("The full installation log has been saved to your desktop as 'pdn.log'.");

            if (found1603)
            {
                Clipboard.SetText(string.Join("\r\n", errorLines.ToArray()));
                Console.WriteLine("The relevant lines from above have also been copied to the clipboard.");
                Console.WriteLine(string.Empty);
            }

            Console.WriteLine(string.Empty);
            Console.WriteLine("Press any key to close.");
            Console.ReadKey();
        }


        static void NoFile(string type)
        {
            Console.ForegroundColor = ConsoleColor.Red;
            Console.WriteLine($"The {type} file was not found!");
            Console.ForegroundColor = ConsoleColor.White;
            Console.WriteLine(string.Empty);
            Console.WriteLine("Press any key to close.");
            Console.ReadKey();
        }

    }
}
