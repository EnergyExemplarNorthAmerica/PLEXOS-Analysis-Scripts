 // C# program to print Hello World!
using System;
using System.IO;
using System.Text.RegularExpressions;
using System.Collections.Generic;

// namespace declaration
namespace CSVCollector {  
    // Class declaration


    class CSVCollector {
        private static StreamWriter OpenProjectCSV(string folder, string project, string timestamp) 
        {
            string filename = Path.Combine(folder.Length == 0 ? "." : folder, string.Format("{0}-{1}.csv", project, timestamp));
            // Console.WriteLine(filename);
            if (!File.Exists(filename))
            {
                StreamWriter fout = File.CreateText(filename);
                fout.WriteLine("Model,Period,Phase,Class,Name,Property,Timestamp,Value");
                fout.Close();
            }
            return new StreamWriter(new FileStream(filename, FileMode.Append, FileAccess.Write));
        }
    
        private static string[] UnpivotHeader(string header) {
            return header.Split(','); // has an extra column at the front
        }

        private static int UnpivotValues(StreamWriter outfile, string[] objects, string line, Match m) {
            string[] fields = line.Split(',');
            for (int idx = 1; idx < fields.Length; idx++) {
                outfile.WriteLine(
                    "{0},{1},{2},{3},{4},{5},{6},{7}",
                    m.Groups["model"].Value, m.Groups["period"].Value, m.Groups["phase"].Value,
                    m.Groups["class"].Value, objects[idx], m.Groups["prop"], fields[0], fields[idx]
                );
            }
            return fields.Length - 1;
        }

        private static void UnpivotFile(string folder, string filepath, string timestamp, int progress = 500) {
            Match m = Regex.Match(filepath, @"Project (?<project>.+?) Solution\\Model (?<model>.+?) Solution\\(?<period>\w+)\\(?<phase>\w+)\s(?<class>[\w\s]+)\.(?<prop>.+?)\.csv");
            // Console.WriteLine(filepath);

            // Skip no match or Interval .csv files
            if (m.Length == 0 || m.Groups["period"].Value == "Interval") return; // skip the interval files
            Console.WriteLine(
                "Model->{0}, Phase-->{2}, Period->{1}, Property-->{3}.{4}", 
                m.Groups["model"].Value, m.Groups["period"].Value, m.Groups["phase"], 
                m.Groups["class"].Value, m.Groups["prop"].Value
            );
            
            // open the .csv file and a target file for aggregation
            StreamReader infile = File.OpenText(filepath);
            StreamWriter outfile = OpenProjectCSV(folder, m.Groups["project"].Value, timestamp);

            string[] headers = UnpivotHeader(infile.ReadLine());
            int values = 0, newrows;

            while (!infile.EndOfStream) { 
                newrows = UnpivotValues(outfile, headers, infile.ReadLine(), m);
                Console.Write(new string('.',((newrows + (values % progress))/progress)));
                values += newrows;
            }

            // close both files
            outfile.Close();
            infile.Close();
        }

        // Main Method
        static void Main(string[] args)
        {
            string[] fld = args[0].Split(',');
            UnpivotFile(fld[0], fld[1], fld[2]);
        }
    }
}