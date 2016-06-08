using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using MinimalisticTelnet;

namespace UpdatedMBM
{
    class Program
    {
        
        static void Main(string[] args)
        {
            
            try
            {
                if (File.Exists("output.xls"))
                {
                    File.Delete("output.xls");
                }
                BSCDataOperations aOperation = new BSCDataOperations();
                aOperation.LoadInputFile("input.txt");
                aOperation.ExecuteCommandInBSC();
            }
            catch (Exception exception)
            {
                Console.WriteLine("Error: " + exception.Message);
            }
            Console.ReadKey();
        }

       
       
    }
}
