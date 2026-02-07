
using System;
using System.Diagnostics;
using System.IO;
using System.Text;
using System.Web.Script.Serialization; // Needs System.Web.Extensions reference
using System.Collections.Generic;

namespace InvoiceInspector
{
    public static class PythonBridge
    {
        // Simple JSON Parser helper since we can't easily reference Newtonsoft.JSON without NuGet
        // We will use System.Web.Script.Serialization (built-in .NET 4.0+)
        // OR we can use simple string parsing if the JSON is simple enough.
        // Actually, System.Web.Extensions might not be referenced by default in csc.
        // Let's try to stick to simple string operations or Regex for dependencies-free experience
        // Or better: Use System.Runtime.Serialization.Json (DataContractJsonSerializer)
        
        // For simplicity given the requirement "don't install too many things", 
        // I will write a tiny JSON parser or use Regex for the specific fields we need.
        
        public static string RunCommand(string args)
        {
            ProcessStartInfo start = new ProcessStartInfo();
            
            // Portable Backend Check
            string exePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "dist", "inspector_cli.exe");
            if (!File.Exists(exePath))
            {
                // Fallback: Try root folder (if moved out of dist)
                exePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "inspector_cli.exe");
            }

            if (File.Exists(exePath))
            {
                // Use compiled backend
                start.FileName = exePath;
                start.Arguments = args;
            }
            else
            {
                // Dev Mode: Use python cli.py
                start.FileName = "python"; 
                start.Arguments = "cli.py " + args;
            }

            start.UseShellExecute = false;
            start.RedirectStandardOutput = true;
            start.RedirectStandardError = true;
            start.CreateNoWindow = true;
            start.StandardOutputEncoding = Encoding.UTF8;

            using (Process process = Process.Start(start))
            {
                string result = process.StandardOutput.ReadToEnd();
                // string stderr = process.StandardError.ReadToEnd(); // Can block logic if both redirect? 
                // Best practice is read one async or use OutputDataReceived. 
                // But for small output, sequential read is okay IF buffer doesn't fill.
                // StandardError read after StandardOutput read *might* deadlock if stdout is huge.
                // But typically okay here.
                string stderr = process.StandardError.ReadToEnd();
                
                process.WaitForExit();

                if (process.ExitCode != 0)
                {
                    throw new Exception("Backend Error: " + stderr);
                }
                return result;
            }
        }
    }
    
    // Minimal Data Structures for JSON Deserialization (Manual or via built-in)
}
