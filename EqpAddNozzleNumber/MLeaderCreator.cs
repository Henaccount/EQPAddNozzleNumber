using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;
using Autodesk.ProcessPower.PnP3dObjects;
using Autodesk.ProcessPower.PnP3dEquipment;
using System;
using System.Collections.Generic;

namespace EqpAddNozzleNumber
{
    public class MLeaderCreator
    {
        [CommandMethod("EqpAddNozzleNumber")]
        public void CreateMLeaders()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;

            PromptEntityResult equipResult = ed.GetEntity("\nSelect an equipment object: ");
            if (equipResult.Status != PromptStatus.OK || !(equipResult.ObjectId.IsValid))
            {
                ed.WriteMessage("\nInvalid selection. Please select an equipment object.");
                return;
            }

            using (Transaction tr = doc.TransactionManager.StartTransaction())
            {
                List<ObjectId> mTextObjects = new List<ObjectId>();

                if (equipResult.ObjectId.ObjectClass.Name.Equals("AcPpDb3dEquipment"))
                {

                    Equipment Eqp = tr.GetObject(equipResult.ObjectId, OpenMode.ForRead) as Equipment;


                    if (Eqp != null)
                    {
                        foreach (SubPart sp in Eqp.AllSubParts)
                        {
                            PartSizeProperties spprops = sp.PartSizeProperties;

                            try
                            {
                                if (sp.Id.ObjectClass.Name != null)
                                {
                                    BlockTable table = tr.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
                                    BlockTableRecord model = tr.GetObject(table[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;

                                    MText mText = new MText();

                                    //mText.Width = 100;
                                    //mText.Height = 50;
                                    //mText.TextHeight = 5;

                                    mText.Location = new Point3d(Convert.ToDouble(spprops.PropValue("Position X")), Convert.ToDouble(spprops.PropValue("Position Y")), Convert.ToDouble(spprops.PropValue("Position Z")));

                                    mText.SetContentsRtf(spprops.PropValue("Tag").ToString());
           
                                    model.AppendEntity(mText);
                                    tr.AddNewlyCreatedDBObject(mText, true);

                                    mTextObjects.Add(mText.Id);

                                }
                            }
                            catch (System.Exception e) { ed.WriteMessage("\n" + e.Message); }

                        }
                    }
                    else
                    {
                        ed.WriteMessage("\nSelected object is not a valid equipment. Please select an equipment object.");
                        return;
                    }


                }
                ObjectId[] mTextObjectsArray = mTextObjects.ToArray();
                new EquipmentHelper().AttachGraphics(equipResult.ObjectId, mTextObjectsArray);
                tr.Commit();
            }
        }
    }
}
