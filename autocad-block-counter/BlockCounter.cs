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
        Dictionary<string, int> licznikBlokow = new Dictionary<string, int>();

        try
        {
            using (Transaction tr = doc.Database.TransactionManager.StartTransaction())
            {
                PromptSelectionResult acSSPrompt = ed.GetSelection();

                if (acSSPrompt.Status == PromptStatus.OK)
                {
                    SelectionSet acSSet = acSSPrompt.Value;

                    foreach (SelectedObject acSSObj in acSSet)
                    {
                        if (acSSObj == null) continue;

                        BlockReference acBlkRef = tr.GetObject(acSSObj.ObjectId, OpenMode.ForRead) as BlockReference;
                        if (acBlkRef != null)
                        {
                            ObjectId idBlock = acBlkRef.DynamicBlockTableRecord;
                            BlockTableRecord btr = tr.GetObject(idBlock, OpenMode.ForRead) as BlockTableRecord;
                            
                            string blockName = btr.Name;
                            
                            if (licznikBlokow.ContainsKey(blockName))
                                licznikBlokow[blockName]++;
                            else
                                licznikBlokow.Add(blockName, 1);
                        }
                    }
                }
                tr.Commit();
            }

            if (licznikBlokow.Count > 0)
            {
                SaveToCsv(licznikBlokow);
                ed.WriteMessage("\n[Sukces]: Dane wyeksportowane do BOM.csv na pulpicie.");
            }
            else
            {
                ed.WriteMessage("\n[Info]: Nie zaznaczono żadnych bloków.");
            }
        }
        catch (System.Exception ex)
        {
            ed.WriteMessage("\n[Błąd]: " + ex.Message);
        }
    }

    private static void SaveToCsv(Dictionary<string, int> dane)
    {
        var csv = new StringBuilder();
        csv.AppendLine("Nazwa bloku;Ilosc");
        foreach (var item in dane)
        {
            csv.AppendLine($"{item.Key};{item.Value}");
        }

        string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
        string filePath = Path.Combine(desktopPath, "BOM.csv");
        
        File.WriteAllText(filePath, csv.ToString(), Encoding.UTF8);
    }
}