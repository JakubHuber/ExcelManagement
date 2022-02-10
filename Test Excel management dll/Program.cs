using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelManagement;

namespace Test_Excel_management_dll
{
    internal class Program
    {
        static void Main()
        {
            ManageExcelTasks oMET = new ExcelManagement.ManageExcelTasks(); 
            Console.WriteLine(oMET.CloseAllExcelInstances());
            Console.ReadLine(); 
        }
    }
}
