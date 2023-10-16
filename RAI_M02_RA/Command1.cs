#region Namespaces
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Reflection;

#endregion

namespace RAI_M02_RA
{
    [Transaction(TransactionMode.Manual)]
    public class Command1 : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            // this is a variable for the Revit application
            UIApplication uiapp = commandData.Application;

            // this is a variable for the current Revit model
            Document doc = uiapp.ActiveUIDocument.Document;

            // 1a. Filtered Element Collector by view
            View curView = doc.ActiveView;
            FilteredElementCollector collector = new FilteredElementCollector(doc, curView.Id);

            // 1b. ElementMultiCategoryFilter
            List<BuiltInCategory> catList = new List<BuiltInCategory>();
            catList.Add(BuiltInCategory.OST_Areas);
            catList.Add(BuiltInCategory.OST_Walls);
            catList.Add(BuiltInCategory.OST_Doors);
            catList.Add(BuiltInCategory.OST_Furniture);
            catList.Add(BuiltInCategory.OST_LightingFixtures);
            catList.Add(BuiltInCategory.OST_Rooms);
            catList.Add(BuiltInCategory.OST_Windows);

            ElementMulticategoryFilter catFilter = new ElementMulticategoryFilter(catList);
            collector.WherePasses(catFilter).WhereElementIsNotElementType();

            // 1c. Use LINQ to get family symbol by name
            FamilySymbol curAreaTag = new FilteredElementCollector(doc)
                .OfClass(typeof(FamilySymbol))
                .Cast<FamilySymbol>()
                .Where(x => x.FamilyName.Equals("M_Area Tag"))
                .First();

            FamilySymbol curCurtainWallTag = new FilteredElementCollector(doc)
                .OfClass(typeof(FamilySymbol))
                .Cast<FamilySymbol>()
                .Where(x => x.FamilyName.Equals("M_Curtain Wall Tag"))
                .First();

            FamilySymbol curDoorTag = new FilteredElementCollector(doc)
                .OfClass(typeof(FamilySymbol))
                .Cast<FamilySymbol>()
                .Where(x => x.FamilyName.Equals("M_Door Tag"))
                .First();

            FamilySymbol curFurnitureTag = new FilteredElementCollector(doc)
                .OfClass(typeof(FamilySymbol))
                .Cast<FamilySymbol>()
                .Where(x => x.FamilyName.Equals("M_Furniture Tag"))
                .First();

            FamilySymbol curLightingFixtureTag = new FilteredElementCollector(doc)
                .OfClass(typeof(FamilySymbol))
                .Cast<FamilySymbol>()
                .Where(x => x.FamilyName.Equals("M_Lighting Fixture Tag"))
                .First();

            FamilySymbol curRoomTag = new FilteredElementCollector(doc)
                .OfClass(typeof(FamilySymbol))
                .Cast<FamilySymbol>()
                .Where(x => x.FamilyName.Equals("M_Room Tag"))
                .First();

            FamilySymbol curWallTag = new FilteredElementCollector(doc)
                .OfClass(typeof(FamilySymbol))
                .Cast<FamilySymbol>()
                .Where(x => x.FamilyName.Equals("M_Wall Tag"))
                .First();

            FamilySymbol curWindowTag = new FilteredElementCollector(doc)
                .OfClass(typeof(FamilySymbol))
                .Cast<FamilySymbol>()
                .Where(x => x.FamilyName.Equals("M_Window Tag"))
                .First();

            // 2. create dictionary for tags
            Dictionary<string, FamilySymbol> tags = new Dictionary<string, FamilySymbol>();
            tags.Add("Areas", curAreaTag);
            tags.Add("Curtain Walls", curCurtainWallTag);
            tags.Add("Doors", curDoorTag);
            tags.Add("Furniture", curFurnitureTag);
            tags.Add("Lighting Fixtures", curLightingFixtureTag);
            tags.Add("Rooms", curRoomTag);
            tags.Add("Walls", curWallTag);
            tags.Add("Windows", curWindowTag);

            using (Transaction t = new Transaction(doc))
            {
                t.Start("Insert Tags");
                foreach (Element curElem in collector)
                {
                    // 3. get point from location
                    XYZ insPoint;
                    LocationPoint locPoint;
                    LocationCurve locCurve;

                    Location curLoc = curElem.Location;

                    if (curLoc == null)
                        continue;

                    locPoint = curLoc as LocationPoint;
                    if (locPoint != null)
                    {
                        // is a location point
                        insPoint = locPoint.Point;
                    }
                    else
                    {
                        // is a location curve
                        locCurve = curLoc as LocationCurve;
                        Curve curCurve = locCurve.Curve;
                        //insPoint = curCurve.GetEndPoint(1);
                        insPoint = GetMidPointBetweenTwoPoints(curCurve.GetEndPoint(0), curCurve.GetEndPoint(1));
                    }
                    ViewType curViewType = curView.ViewType;

                    // 4. create reference to element
                    Reference curRef = new Reference(curElem);

                    // 5. tag area plan
                    if (curViewType == ViewType.AreaPlan)
                    {
                        if (curElem.Category.Name == "Areas")
                        {
                            ViewPlan curAreaPlan = curView as ViewPlan;
                            Area curArea = curElem as Area;

                            // 5b. place area tag
                            AreaTag curAreaTag2 = doc.Create.NewAreaTag(curAreaPlan, curArea, new UV(insPoint.X, insPoint.Y));
                            curAreaTag2.TagHeadPosition = new XYZ(insPoint.X, insPoint.Y, 0);
                            curAreaTag2.HasLeader = false;
                        }
                    }

                    // 6. tag ceiling plan

                    if (curViewType == ViewType.CeilingPlan)
                    {
                        if (curElem.Category.Name == "Lighting Fixtures")
                        {
                            IndependentTag newLightingFixtureTag = IndependentTag.Create(doc, curLightingFixtureTag.Id, curView.Id, curRef, false, TagOrientation.Horizontal, insPoint);
                        }
                        if (curElem.Category.Name == "Rooms")
                        {
                            IndependentTag newRoomTag = IndependentTag.Create(doc, curRoomTag.Id, curView.Id, curRef, false, TagOrientation.Horizontal, insPoint);
                        }
                    }

                    // 7. tag floor plan

                    if (curViewType == ViewType.FloorPlan)
                    {
                        if (curElem.Category.Name == "Walls")
                        {
                            Wall curWall = curElem as Wall;
                            WallType curWallType = curWall.WallType;


                            if (curWallType.Kind == WallKind.Curtain)
                            {
                                IndependentTag newCurtainWallTag = IndependentTag.Create(doc, curCurtainWallTag.Id, curView.Id, curRef, false, TagOrientation.Horizontal, insPoint);
                            }
                            else
                            {
                                IndependentTag newWallTag = IndependentTag.Create(doc, curWallTag.Id, curView.Id, curRef, false, TagOrientation.Horizontal, insPoint);
                            }
                        }

                        if (curElem.Category.Name == "Doors")
                        {
                            IndependentTag newDoorTag = IndependentTag.Create(doc, curDoorTag.Id, curView.Id, curRef, false, TagOrientation.Horizontal, insPoint);
                        }
                        if (curElem.Category.Name == "Furniture")
                        {
                            IndependentTag newFurnitureTag = IndependentTag.Create(doc, curFurnitureTag.Id, curView.Id, curRef, false, TagOrientation.Horizontal, insPoint);
                        }
                        if (curElem.Category.Name == "Rooms")
                        {
                            IndependentTag newRoomTag = IndependentTag.Create(doc, curRoomTag.Id, curView.Id, curRef, false, TagOrientation.Horizontal, insPoint);
                        }
                        if (curElem.Category.Name == "Windows")
                        {
                            XYZ newInsPoint = new XYZ(insPoint.X, (insPoint.Y+3), insPoint.Z);
                            IndependentTag newWindowTag = IndependentTag.Create(doc, curWindowTag.Id, curView.Id, curRef, false, TagOrientation.Horizontal, newInsPoint);
                        }
                    }

                    // 8. tag section

                    if (curViewType == ViewType.Section)
                    {
                        if (curElem.Category.Name == "Rooms")
                        {

                            ElementId curLevelId = curElem.LevelId as ElementId;
                            Level curLevel = doc.GetElement(curLevelId) as Level;
                            double curLevelElev = curLevel.Elevation;
                            IndependentTag newRoomTag = IndependentTag.Create(doc, curRoomTag.Id, curView.Id, curRef, false, TagOrientation.Horizontal, insPoint);
                            XYZ moveVector = new XYZ (0, 0, 3);
                            ElementTransformUtils.MoveElement(doc, newRoomTag.Id, moveVector);
                        }
                    }


                    //FamilySymbol curTagType = tags[curElem.Category.Name];

                    // 5a. place tags
                    //IndependentTag newTag = IndependentTag.Create(doc, curTagType.Id, curView.Id, curRef, false, TagOrientation.Horizontal, insPoint);


                }
                t.Commit();

            }

            return Result.Succeeded;
        }

        private XYZ GetMidPointBetweenTwoPoints(XYZ point1, XYZ point2)
        {
            XYZ midPoint = new XYZ(
                (point1.X + point2.X) / 2,
                (point1.Y + point2.Y) / 2,
                (point1.Z + point2.Z) / 2);

            return midPoint;
        }

        internal static PushButtonData GetButtonData()
        {
            // use this method to define the properties for this command in the Revit ribbon
            string buttonInternalName = "btnCommand1";
            string buttonTitle = "Button 1";

            ButtonDataClass myButtonData1 = new ButtonDataClass(
                buttonInternalName,
                buttonTitle,
                MethodBase.GetCurrentMethod().DeclaringType?.FullName,
                Properties.Resources.Blue_32,
                Properties.Resources.Blue_16,
                "This is a tooltip for Button 1");

            return myButtonData1.Data;
        }
    }
}
