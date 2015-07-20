using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Net;
using System.IO;
using System.Data;
using System.Threading;
using System.Threading.Tasks;
using System.Data.SqlClient;
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;
using System.Web;
using System.Web.Script.Serialization;
using System.Drawing;

namespace WRPRmigrate {
  class Program {
    static void Main(string[] args) {
      string fileA = "";
      string fileB = "";
      bool found = false;
      if (args.Length >= 2) {
        //dowork
        fileA = args[0];
        fileB = args[1];
        found = true;
      } else {
        Console.WriteLine("Please enter two spreadsheets to merge from console."+Environment.NewLine+"WRPRmigrate.exe  <c:\\folder\\from.xlsx> <c:\\folder\\into.xlsx> [c (cert-only?)]");
        Console.ReadKey();
      }

      if (found) {
        Excel.Application excelApp = null;
        Excel.Workbook OldBook = null;
        Excel.Workbook NewBook = null;
        Excel.Worksheet OldSheet = null;
        Excel.Worksheet NewSheet = null;
        Excel.Worksheet dtSheet = null;
        Excel.Range visibleCells = null;
        Excel.Range R2 = null;
        try {
          excelApp = new Excel.Application(); ;
          excelApp.DisplayAlerts = false;
          OldBook = excelApp.Workbooks.Open(fileA, false, false, Type.Missing, Type.Missing, Type.Missing, true, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
          NewBook = excelApp.Workbooks.Open(fileB, false, false, Type.Missing, Type.Missing, Type.Missing, true, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
          OldSheet = OldBook.Sheets[1];
          NewSheet = NewBook.Sheets[1];
          int NewMR = NewSheet.UsedRange.Rows.Count;
          dtSheet = (Worksheet)NewBook.Sheets.Add(Type.Missing, NewBook.Sheets[1], Type.Missing, Type.Missing);
          if (NewBook.Sheets.Count > 1) {
            dtSheet.Name = "Merge "+(NewBook.Sheets.Count - 1).ToString();
          } else {
            dtSheet.Name = "Merge Info";
          }

          #region Formatted Cells Parser

          visibleCells = OldSheet.UsedRange.SpecialCells(Excel.XlCellType.xlCellTypeVisible, Type.Missing);
          int OldMR = visibleCells.Rows.Count;
          object[,] OldData = visibleCells.get_Range("A2:V" + OldMR).Value;
          object[,] NewData = NewSheet.get_Range("C2:E" + NewMR).Value;//get full new WR list
          List<string> OldWRs = new List<string>();
          int te = visibleCells.Areas.Count;
          int tb = visibleCells.Columns.Count;
          foreach (Excel.Range area in visibleCells.Areas) {
            object[,] areaData = area.Value;
            OldWRs.AddRange(getColumn(areaData, 5));
          }
          List<string> NewWRs = getColumn(NewData, 3);
          List<string> NewPID = getColumn(NewData, 1);
          List<string> NewerWRs = NewWRs;
          int itemF = 0;
          int rowCount = 0;
          foreach (Excel.Range area in visibleCells.Areas) {
            rowCount += area.Rows.Count;
          }

          #endregion
          
          double modifier = 100.0 / rowCount;
          double progress = 0.0;
          foreach (Excel.Range area in visibleCells.Areas) {
            foreach (Excel.Range xlRow in area.Rows) {
              object[,] rowData = xlRow.Value;
              for (int i = 0; i < NewerWRs.Count; i++) {
                if (NewerWRs[i] != "DUPE"&&NewerWRs[i]==rowData[1,5].ToString()&&NewPID[i]==rowData[1,3].ToString()) {
                  Excel.Range y1 = null;
                  Excel.Range y2 = null;
                  if (args.Length == 2) {
                    //copy data for A, B, S, and T
                    NewSheet.Cells[i + 2, 1] = rowData[1, 1];//A
                    NewSheet.Cells[i + 2, 2] = rowData[1, 2];//B
                    NewSheet.Cells[i + 2, 19] = rowData[1, 19];//S
                    NewSheet.Cells[i + 2, 20] = rowData[1, 20];//T
                    //copy formatting from row on columns 1-20
                    y1 = NewSheet.Cells[i + 2, 1];
                    y2 = NewSheet.Cells[i + 2, 20];
                    xlRow.Copy(Type.Missing);
                    R2 = (Excel.Range)NewSheet.get_Range(y1, y2);
                    R2.PasteSpecial(Excel.XlPasteType.xlPasteFormats, Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, false, false);
                    
                  } else {
                    //SCT copy data
                    NewSheet.Cells[i + 2, 21] = rowData[1, 21];//U
                    NewSheet.Cells[i + 2, 22] = rowData[1, 22];//V
                    //SCT copy formatting
                    y1 = NewSheet.Cells[i + 2, 21];
                    y2 = NewSheet.Cells[i + 2, 22];
                    xlRow.Columns["U:V"].Copy(Type.Missing);
                    R2 = (Excel.Range)NewSheet.get_Range(y1, y2);
                    R2.PasteSpecial(Excel.XlPasteType.xlPasteFormats, Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, false, false);

                  }
                  releaseObject(y1);
                  releaseObject(y2);
                  itemF++;
                }
              }
              releaseObject(xlRow);
              progress += modifier;
              Console.WriteLine(Math.Round(progress)+"%");
            }
            releaseObject(area);            
          }
          Console.WriteLine("Finishing the report...");
          NewSheet.get_Range("R1", "R1").EntireColumn.NumberFormat = "m/d/yyyy";
          NewSheet.get_Range("H1", "J1").EntireColumn.NumberFormat = "m/d/yyyy";
          int DupeWR = NewSheet.UsedRange.Rows.Count;
          object[,] DupeData = NewSheet.get_Range("C2:E" + DupeWR).Value;
          List<string> SecondWR = getColumn(DupeData, 3);
          List<string>SecondPID = getColumn(DupeData,1);
          List<string> duplicatedWRs = SecondWR.GroupBy(x => x)
                .Where(group => group.Count() > 1)
                .Select(group => group.Key).ToList();
          dtSheet.Cells[1, 3] = "Duplicated"+Environment.NewLine+"WRs Found:";
          dtSheet.Cells[1, 4] = "Projects Linked";
          int rowToPrint = 2;
          foreach (string dWRx in duplicatedWRs) {
            string ProjectsFound = "";
            for (int i = 0; i < SecondWR.Count; i++) {
              if (dWRx == SecondWR[i]) {
                if (ProjectsFound == "") {
                  ProjectsFound += SecondPID[i];
                } else {
                  ProjectsFound += Environment.NewLine + SecondPID[i];
                }
              }
            }
            dtSheet.Cells[rowToPrint, 3] = dWRx;
            dtSheet.Cells[rowToPrint, 4] = ProjectsFound;
            rowToPrint++;
          }
          dtSheet.Cells[2, 1] = "Items merged: " + itemF;
          dtSheet.Cells[3, 1] = "Items closed: " + ((OldWRs.Count()-1) - itemF);
          dtSheet.Cells[4, 1] = "New Items: " + (((from x in NewerWRs select x).Distinct().Count() - 1) - itemF);
          NewBook.SaveAs(Directory.GetCurrentDirectory()+"\\output", XlFileFormat.xlOpenXMLWorkbook, Type.Missing, Type.Missing, false, false, XlSaveAsAccessMode.xlNoChange, XlSaveConflictResolution.xlLocalSessionChanges, Type.Missing, Type.Missing, Type.Missing);
          NewBook.Close(true, Type.Missing, Type.Missing);
          OldBook.Close(false, Type.Missing, Type.Missing);
          excelApp.Quit();
        } catch (Exception e){
          NewBook.Close(false, Type.Missing, Type.Missing);
          OldBook.Close(false, Type.Missing, Type.Missing);
          excelApp.Quit();
          Console.WriteLine("Danger to manifold! Check for format issues.");
          Console.WriteLine(e.Message);
          Console.ReadKey();
        } finally {
          releaseObject(R2);
          releaseObject(NewSheet);
          releaseObject(OldSheet);
          releaseObject(NewBook);
          releaseObject(OldBook);
          releaseObject(excelApp);
        }
      }
    }
    static List<string> getColumn(object[,] dataTable, int column) {
      List<string> data = new List<string>();
      int maxR = dataTable.GetLength(0);
      for (int i = 1; i <= maxR; i++) {
        data.Add(dataTable[i, column].ToString());
      }
      return data;
    }
    static void releaseObject(object obj) {
      try {
        System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
        obj = null;
      } catch {
        obj = null;
      } finally {
        GC.Collect();
        GC.WaitForPendingFinalizers();
      }
    }
    static int getIndex(List<string> values, string value) {
      int index = -1;
      for (int i = 0; i < values.Count; i++) {
        if (values[i] == value) {
          index = i;
          break;
        }
      }
      return index;
    }
    static int getSIndex(List<string> values1, List<string> values2, string value1, string value2) {
      int index = -1;
      for (int i = 0; i < values1.Count; i++) {
        if (values1[i] == value1 && values2[i] == value2) {
          index = i;
          break;
        }
      }
      return index;
    }
    static List<string> markDuplicates(List<string> values) {
      List<string> duplicates = values.GroupBy(x => x)
                .Where(group => group.Count() > 1)
                .Select(group => group.Key).ToList();
      for (int i = 0; i < values.Count; i++) {
        int k = getIndex(duplicates, values[i]);
        if (k >= 0) {
          values.RemoveAt(i);
          values.Insert(i, "DUPE");
        }
      }
      return values;
    }

  }
}
