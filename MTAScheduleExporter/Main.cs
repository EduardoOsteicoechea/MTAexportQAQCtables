using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace MTAScheduleExporter
{
   [Transaction(TransactionMode.Manual)]   
    public class Main: IExternalCommand
   {
      public ExternalCommandData ExternalCommandDataInstance { get; set; } = null;
      public static Autodesk.Revit.DB.Document DatabaseDocument { get; set; }
      public FilteredElementCollector ScheduleCollector { get; set; } = null;
      public List<ViewSchedule> DocumentSchedules { get; set; } = null;
      public List<ViewSchedule> TargetedSchedules { get; set; } = null;
      public StringBuilder ScheduleSection0Data { get; set; } = new StringBuilder();
      public StringBuilder ScheduleSection1Data { get; set; } = new StringBuilder();
      public StringBuilder ScheduleSection2Data { get; set; } = new StringBuilder();
      public StringBuilder ScheduleFullData { get; set; } = new StringBuilder();
      public string CSVFilePath { get; set; } = null;
      public Result Execute
      (
         ExternalCommandData commandData,
         ref string message,
         ElementSet elements
      )
      {
         ExternalCommandDataInstance = commandData;
         DatabaseDocument = ExternalCommandDataInstance.Application.ActiveUIDocument.Document;
         CollectDocumentSchedules();
         CastScheduleCollectorElementsToViewScheduleTypes();
         FilterTargetedSchedules();
         if(ThereAreTargetedSchedules())
         {
            GenereateCSVForEachSchedule();
         };

         return Result.Succeeded;
      }

      public void PrepareDataContainers()
      {
         ScheduleSection0Data.Clear();
         ScheduleSection1Data.Clear();
         ScheduleSection2Data.Clear();
         ScheduleFullData.Clear();
         CSVFilePath = "";
      }

      public void GenereateCSVForEachSchedule() 
      {
         foreach(ViewSchedule schedule in TargetedSchedules)
         {
            PrepareDataContainers();
            GetScheduleSection0Data(schedule, ScheduleSection0Data, 0);
            GetScheduleSection1Data(schedule, ScheduleSection1Data, 1);
            //GetScheduleSection2Data(schedule, ScheduleSection2Data, 2);
            GenerateFullDataString();
            GenerateCSVFilePath(schedule);
            GenerateCSVFile(schedule);
         };
      }

      public void CollectDocumentSchedules()
      {
         ScheduleCollector = new FilteredElementCollector(DatabaseDocument);
         ScheduleCollector.OfClass(typeof(ViewSchedule));
      }

      public void CastScheduleCollectorElementsToViewScheduleTypes()
      {
         DocumentSchedules = ScheduleCollector.ToElements().Cast<ViewSchedule>().ToList();
      }

      public void FilterTargetedSchedules()
      {
         TargetedSchedules = new List<ViewSchedule>();

         if(DocumentSchedules.Count > 0)
         {
            foreach(ViewSchedule schedule in DocumentSchedules)
            {
               //if(schedule != null && schedule.Name.Contains("WS-WSS-693-3700-SMA-DDE-AR-A00-0100001_Area Schedule") == true)
               //if(schedule != null && schedule.Name.Contains("WS-WSS-693-3700-SMA-DDE-AR-A00-0100001_Level Schedule") == true)
               //if(schedule != null && schedule.Name.Contains("WS-WSS-693-3700-SMA-DDE-AR-A00-0100001_Multi-Category Schedule") == true)
               //if(schedule != null && schedule.Name.Contains("WS-WSS-693-3700-SMA-DDE-AR-A00-0100001_Project Schedule") == true)
               //if(schedule != null && schedule.Name.Contains("WS-WSS-693-3700-SMA-DDE-AR-A00-0100001_Room Schedule") == true)
               //if(schedule != null && schedule.Name.Contains("WS-WSS-693-3700-SMA-DDE-AR-A00-0100001") == true)
                  //if(schedule != null && schedule.Name.Contains("WS-WSS-693-3700-SMA-DDE-AR") == true)
                  if(schedule != null && schedule.Name.Contains(DatabaseDocument.Title) == true)
                  {
                  //DataBlockVisualizer dataBlockVisualizer = new DataBlockVisualizer($"Proccessing schedule {schedule.Name}");
                  //dataBlockVisualizer.Show();
                  TargetedSchedules.Add(schedule);
               }
               //else
               //{
               //   DataBlockVisualizer dataBlockVisualizer = new DataBlockVisualizer($"No schedules contained the Document title");
               //   dataBlockVisualizer.Show();
               //};
            };
         }
         else
         {
            DataBlockVisualizer dataBlockVisualizer = new DataBlockVisualizer($"No schedules in the Document");
            dataBlockVisualizer.Show();
         }
      }

      public bool ThereAreTargetedSchedules()
      {
         return TargetedSchedules.Count > 0;
      }

      public void GetScheduleSection0Data(ViewSchedule schedule, StringBuilder sectionDataBuilder, int sectionDataIndex)
      {
         TableData scheduleTableData = schedule.GetTableData();

         if(scheduleTableData.GetSectionData(sectionDataIndex) != null)
         {
            sectionDataBuilder.Append(schedule.Name);
         };
      }

      public void GetScheduleSection1Data(ViewSchedule schedule, StringBuilder sectionDataBuilder, int sectionDataIndex)
      {
         TableData scheduleTableData = schedule.GetTableData();

         if(scheduleTableData.GetSectionData(sectionDataIndex) != null)
         {
            TableSectionData tableSectionData = scheduleTableData.GetSectionData(sectionDataIndex);
            ProcessTableSectionData(schedule, sectionDataBuilder, tableSectionData);
         };
      }

      public void GetScheduleSection2Data(ViewSchedule schedule, StringBuilder sectionDataBuilder, int sectionDataIndex)
      {
         TableData scheduleTableData = schedule.GetTableData();

         if(scheduleTableData.GetSectionData(sectionDataIndex) != null)
         {
            TableSectionData tableSectionData = scheduleTableData.GetSectionData(sectionDataIndex);
            ProcessTableSectionData(schedule, sectionDataBuilder, tableSectionData);
         };
      }

      public void ProcessTableSectionData(ViewSchedule schedule, StringBuilder sectionDataBuilder, TableSectionData tableSectionData)
      {
         for(int j = 0; j < tableSectionData.NumberOfRows; j++)
         {
            StringBuilder columnsDataBuilder = new StringBuilder();

            for(int k = 0; k < tableSectionData.NumberOfColumns; k++)
            {
               int lastColumnIndex = tableSectionData.NumberOfColumns - 1;

               if(k != lastColumnIndex)
               {
                  columnsDataBuilder.Append(schedule.GetCellText(SectionType.Body, j, k) + "\t");
               }
               else
               {
                  columnsDataBuilder.Append(schedule.GetCellText(SectionType.Body, j, k));
               };
            };

            sectionDataBuilder.AppendLine(columnsDataBuilder.ToString());
         };
      }

      public void GenerateFullDataString()
      {  
         ScheduleFullData.Clear();
         ScheduleFullData.AppendLine(ScheduleSection0Data.ToString());
         ScheduleFullData.AppendLine(ScheduleSection1Data.ToString());
         ScheduleFullData.AppendLine(ScheduleSection2Data.ToString());
      }

      public void GenerateCSVFilePath(ViewSchedule schedule)
      {
         CSVFilePath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), schedule.Name) + ".csv";
      }

      public void GenerateCSVFile(ViewSchedule schedule)
      {
         Encoding encoding = Encoding.UTF8;

         using(StreamWriter writer = new StreamWriter(CSVFilePath, false, encoding))
         {
            writer.Write(ScheduleFullData.ToString());
         };         
         
         //DataBlockVisualizer dataBlockVisualizer = new DataBlockVisualizer($"Exported {schedule.Name} Schedule to:\n\n{CSVFilePath}");
         //dataBlockVisualizer.ShowDialog();
      }
   }
}