using System;
using System.IO;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;
using System.Reflection;
using Get_a_collection_of_all_running_Excel_instances;
using System.Activities;
using System.ComponentModel;

namespace ExcelManagement
{
    public class ManageExcelTasks
    {
        /// <summary>
        /// Method for Closing all Excel instances
        /// </summary>
        /// <param name="saveWorkbooks">Save opened workbooks. Default is true</param>
        /// <returns>return 1 if everything is OK</returns>
        public int CloseAllExcelInstances(bool saveWorkbooks = true)
        {
            ExcelAppCollection myApps = new ExcelAppCollection();
            List<Process> ExcelProcesses = (List<Process>)myApps.GetProcesses();

            Application ExcelAppication;

            foreach (Process process in ExcelProcesses)
            {

                ExcelAppication = myApps.FromProcess(process);

                if (ExcelAppication != null)
                {

                    foreach (Workbook oWorkbook in ExcelAppication.Workbooks)
                    {
                        //Check if workbook was never saved
                        if (oWorkbook.Path != string.Empty)
                        {
                            oWorkbook.Close(saveWorkbooks, Missing.Value, Missing.Value);
                        }
                        else
                        {

                            string folderPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                            string savePath = Path.Combine(folderPath, oWorkbook.Name + ".xlsx");
                            Console.WriteLine("Workbook first time opened - savied {0}", savePath);

                            if (File.Exists(savePath))
                            {
                                savePath = Path.Combine(folderPath, oWorkbook.Name + DateTime.Now.ToString("ssmmHHddMMyyyy") + ".xlsx");
                            }

                            oWorkbook.SaveAs(savePath, XlFileFormat.xlOpenXMLWorkbook, Missing.Value, Missing.Value, Missing.Value, Missing.Value, XlSaveAsAccessMode.xlExclusive, XlSaveConflictResolution.xlLocalSessionChanges , Missing.Value, Missing.Value, Missing.Value, Missing.Value);
                            oWorkbook.Close(saveWorkbooks, Missing.Value, Missing.Value);
                        }

                        ReleaseAll(oWorkbook);

                    }

                    ExcelAppication.Quit();
                    ReleaseAll(ExcelAppication);

                }
                else
                {
                    //Excel is in task manager but not visible. Kill it with fire!
                    process.Kill();

                }

            }

            return 1;

        }

        private void ReleaseAll(object obj)
        {
            try
            {
                Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch
            {
                obj = null;
            }
            finally
            {
                GC.Collect();
            }
        }
    }
}
