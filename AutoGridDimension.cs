using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.Attributes;

namespace NSA.Revit.RibbonButton
{
    [TransactionAttribute(TransactionMode.Manual)]
    public class AutoGridDimension : IExternalCommand
    {
        double min_ptX = new double();
        double max_ptX = new double();
        double min_ptY = new double();
        double max_ptY = new double();

        IList<ViewPlan> collectionViews = new List<ViewPlan>();
        IList<IList<ReferenceArray>> collectionGrids = new List<IList<ReferenceArray>>();
        IList<IList<Line>> collectionLine = new List<IList<Line>>();
        IList<ElementId> collectionViewID = new List<ElementId>();

        private void setBoundingBoxXY(View currentView)
        {
            BoundingBoxXYZ cropBB = currentView.CropBox;

            min_ptX = cropBB.Min.X;
            max_ptX = cropBB.Max.X;
            min_ptY = cropBB.Min.Y;
            max_ptY = cropBB.Max.Y;
        }

        private Dictionary<Grid, Double> gridToAngle(IList<Grid> grids)
        {
            Dictionary<Grid, Double> temp = new Dictionary<Grid, double>();

            foreach (Grid g in grids)
            {
                Curve grd_c = g.Curve;
                XYZ grd_Dir = grd_c.GetEndPoint(1) - grd_c.GetEndPoint(0);

                double angle = grd_Dir.AngleTo(XYZ.BasisX) * 180 / Math.PI;
                angle = Math.Round(angle, 2, MidpointRounding.AwayFromZero);

                temp.Add(g, angle);
            }

            return temp;
        }

        private static void VectorOffsetParrallely(ref XYZ pt1, ref XYZ pt2, double dist)
        {
            double mag = Math.Sqrt(Math.Pow(pt2.X - pt1.X, 2) + Math.Pow(pt2.Y - pt1.Y, 2));
            double x_Dist = (dist * (pt2.Y - pt1.Y)) / mag;
            double y_Dist = (dist * (pt1.X - pt2.X)) / mag;
            pt1 = new XYZ(pt1.X + x_Dist, pt1.Y + y_Dist, pt1.Z);
            pt2 = new XYZ(pt2.X + x_Dist, pt2.Y + y_Dist, pt2.Z);
        }

        private (IList<ReferenceArray>, IList<Line>) setGridnRefLine(Dictionary<double, List<Grid>> grd_Group)
        {
            IList<ReferenceArray> temp_grid_Collection = new List<ReferenceArray>();
            IList<Line> temp_ref_Line = new List<Line>();

            for (int j = 0; j < grd_Group.Count; j++)
            {
                ReferenceArray tempArray = new ReferenceArray();

                XYZ firstPt = grd_Group.ElementAt(j).Value.First().Curve.GetEndPoint(0);
                XYZ lastPt = grd_Group.ElementAt(j).Value.Last().Curve.GetEndPoint(0);

                if (grd_Group.ElementAt(j).Value.Count != 1)
                {
                    switch (grd_Group.ElementAt(j).Key)
                    {
                        case var _ when (grd_Group.ElementAt(j).Key == 90.0 && lastPt.Y > max_ptY):
                            firstPt = new XYZ(firstPt.X, max_ptY, firstPt.Z);
                            lastPt = new XYZ(lastPt.X, max_ptY, firstPt.Z);
                            break;
                        case var _ when (grd_Group.ElementAt(j).Key == 180.0 && lastPt.X > max_ptX):
                            firstPt = new XYZ(max_ptX, firstPt.Y, firstPt.Z);
                            lastPt = new XYZ(max_ptX, lastPt.Y, firstPt.Z);
                            break;
                        case var _ when (grd_Group.ElementAt(j).Key > 90.0 && firstPt.X > max_ptX ||
                        grd_Group.ElementAt(j).Key > 90.0 && lastPt.X > max_ptX):
                            if (firstPt.X >= lastPt.X)
                            {
                                VectorOffsetParrallely(ref firstPt, ref lastPt,
                                (max_ptX - firstPt.X) / Math.Cos(grd_Group.ElementAt(j).Key * Math.PI / 180.0));
                            }
                            else
                            {
                                VectorOffsetParrallely(ref firstPt, ref lastPt,
                                (max_ptX - lastPt.X) / Math.Cos(grd_Group.ElementAt(j).Key * Math.PI / 180.0));
                            }
                            break;
                        case var _ when (grd_Group.ElementAt(j).Key < 90.0 && firstPt.X < min_ptX ||
                        grd_Group.ElementAt(j).Key < 90.0 && lastPt.X < min_ptX):
                            if (firstPt.X <= lastPt.X)
                            {
                                VectorOffsetParrallely(ref firstPt, ref lastPt,
                                (min_ptX - firstPt.X) / Math.Cos(grd_Group.ElementAt(j).Key * Math.PI / 180.0));
                            }
                            else
                            {
                                VectorOffsetParrallely(ref firstPt, ref lastPt,
                                (min_ptX - lastPt.X) / Math.Cos(grd_Group.ElementAt(j).Key * Math.PI / 180.0));
                            }
                            break;
                    }

                    Line tempLine = Line.CreateBound(firstPt, lastPt);

                    for (int k = 0; k < grd_Group.ElementAt(j).Value.Count; k++)
                    {
                        var gRef = new Reference(grd_Group.ElementAt(j).Value[k] as Element);

                        tempArray.Append(gRef);
                    }
                    temp_grid_Collection.Add(tempArray);
                    temp_ref_Line.Add(tempLine);
                }
            }

            return (temp_grid_Collection, temp_ref_Line);
        }

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc.Document;
            View currentView = doc.ActiveView;

            var sheetCollector = new FilteredElementCollector(doc).OfClass(typeof(ViewSheet));
            List<ViewSheet> sheets = sheetCollector.Cast<ViewSheet>().ToList();

            var viewCollector = new FilteredElementCollector(doc).OfClass(typeof(ViewPlan));
            List<ViewPlan> ViewPlans = viewCollector.Cast<ViewPlan>().ToList();

            using (DataGridSelectSheet form = new DataGridSelectSheet(sheets))
            {
                var dr = form.ShowDialog();
                if (dr == System.Windows.Forms.DialogResult.OK)
                {
                    if (form.selectedSheets != null)
                    {
                        sheets = form.selectedSheets;
                        foreach (var s in sheets)
                        {
                            //msg += "\n" + s.Name;
                            var viewsID = s.GetAllPlacedViews();
                            if (viewsID != null)
                            {
                                foreach (var id in viewsID)
                                {
                                    View v = doc.GetElement(id) as View;
                                    //msg += "\n" + v;
                                    
                                    if (v.ViewType == ViewType.FloorPlan || v.ViewType == ViewType.CeilingPlan || 
                                        v.ViewType == ViewType.AreaPlan /*|| v.ViewType == ViewType.Elevation ||
                                        v.ViewType == ViewType.Section*/)
                                    {
                                        //collectionViews.Add((ViewPlan)v);
                                        collectionViewID.Add(id);
                                    }
                                }
                            }
                        }
                    }
                    else { return Result.Failed; }
                }
            }

            if (collectionViewID.Count > 0)
            {
                int i = 0;
                while (i != collectionViewID.Count)
                {
                    //ElementId viewID = collectionViews[i].Id;
                    ElementId viewID = collectionViewID[i];
                    View view = doc.GetElement(collectionViewID[i]) as View;

                    //setBoundingBoxXY(collectionViews[i]);
                    setBoundingBoxXY(view);

                    FilteredElementCollector collector = new FilteredElementCollector(doc, viewID);
                    ElementQuickFilter filter = new ElementCategoryFilter(BuiltInCategory.OST_Grids);

                    IList<Element> grids_Elem = collector.WherePasses(filter).WhereElementIsNotElementType().ToElements();
                    IList<Grid> grids = grids_Elem.Cast<Grid>().ToList();

                    Dictionary<Grid, double> grid_Dict = gridToAngle(grids);

                    Dictionary<double, List<Grid>> grd_Group = grid_Dict.GroupBy(r => r.Value)
                        .ToDictionary(t => t.Key, t => t.Select(r => r.Key).ToList());

                    (IList<ReferenceArray> grid_Collection, IList<Line> ref_Line) = setGridnRefLine(grd_Group);

                    collectionGrids.Add(grid_Collection);
                    collectionLine.Add(ref_Line);

                    i++;
                }
            }
            else
            {
                return Result.Failed;
            }

            try
            {
                using (Transaction trans = new Transaction(doc, "Grid Dimension"))
                {
                    trans.Start();
                    //for (int i = 0; i < collectionViews.Count; i++)
                    for (int i = 0; i < collectionViewID.Count; i++)
                    {
                        if (collectionGrids[i] != null)
                        {
                            var gridCollectionNrefLine =
                            collectionGrids[i].Zip(collectionLine[i], (gr, lr) => new { gridRef = gr, lineRef = lr });

                            foreach (var gridLine in gridCollectionNrefLine)
                            {
                                {
                                    //doc.Create.NewDimension(collectionViews[i], gridLine.lineRef, gridLine.gridRef);
                                    doc.Create.NewDimension((View)doc.GetElement(collectionViewID[i]), gridLine.lineRef, gridLine.gridRef);
                                }
                            }
                        } 
                    }

                    trans.Commit();
                }
                return Result.Succeeded;
            }
            catch (Exception e)
            {
                message = e.Message;
                return Result.Failed;
            }

        }
    }
}
