using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using System;
using System.Collections.Generic;
using System.Text;
using System.IO;

namespace SelectCount;

public class BlockCounter
{
    [CommandMethod("BlockCounter")]
    public static void SelectObjectsOnscreen()
    {
        Document doc = Application.DocumentManager.MdiActiveDocument;
        Editor ed = doc.Editor;
        Dictionary<string, int> blockCounts = new Dictionary<string, int>();

        try
        {
            using (Transaction tr = doc.Database.TransactionManager.StartTransaction())
            {
                PromptSelectionResult selectionRes = ed.GetSelection();

                if (selectionRes.Status == PromptStatus.OK)
                {
                    SelectionSet selectionSet = selectionRes.Value;

                    foreach (SelectedObject selectedObj in selectionSet)
                    {
                        if (selectedObj == null) continue;

                        BlockReference blockRef = tr.GetObject(selectedObj.ObjectId, OpenMode.ForRead) as BlockReference;
                        if (blockRef != null)
                        {
                            // Using DynamicBlockTableRecord to get the proper name even for dynamic blocks
                            ObjectId blockId = blockRef.DynamicBlockTableRecord;
                            BlockTableRecord btr = tr.GetObject(blockId, OpenMode.ForRead) as BlockTableRecord;
                            
                            string blockName = btr.Name;
                            
                            if (blockCounts.ContainsKey(blockName))
                                blockCounts[blockName]++;
                            else
                                blockCounts.Add(blockName, 1);
                        }
                    }
                }
                tr.Commit();
            }

            if (blockCounts.Count > 0)
            {
                SaveToCsv(blockCounts);
                ed.WriteMessage("\n[Success]: Data exported to BOM.csv on your Desktop.");
            }
            else
            {
                ed.WriteMessage("\n[Info]: No blocks were selected.");
            }
        }
        catch (System.Exception ex)
        {
            ed.WriteMessage("\n[Error]: " + ex.Message);
        }
    }

    private static void SaveToCsv(Dictionary<string, int> data)
    {
        var csv = new StringBuilder();
        csv.AppendLine("Block Name;Quantity");
        foreach (var item in data)
        {
            csv.AppendLine($"{item.Key};{item.Value}");
        }

        string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
        string filePath = Path.Combine(desktopPath, "BOM.csv");
        
        File.WriteAllText(filePath, csv.ToString(), Encoding.UTF8);
    }
}