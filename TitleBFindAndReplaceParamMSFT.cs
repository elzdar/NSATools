#region Header
//
//Picks out the Workset parameter value and places it in the @Work Package parameter 
//Filter selected chilled beams by FamilyName = "AUS_Arup_Chilled Beam_Passive-Adjustable" and FamilyType = "CHB - Adjustable";
//Adds all air terminals to the selection
//
//
//
//
//

#endregion
#region Namespaces
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using ExcelDataReader;
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.SqlClient;
using Parameter = Autodesk.Revit.DB.Parameter;
using DataTable = System.Data.DataTable;
using System.Data.OleDb;
using System.Text;
using System.Runtime.InteropServices;
using Newtonsoft.Json;
using System.Net.Http;
#endregion // Namespaces

namespace MEPExtensions
{
    [Transaction(TransactionMode.Manual)]
    public class TitleBFindAndReplaceParamMSFT : IExternalCommand
    {
        private Stream stream;
        private UIDocument UIdoc;

        public object ConfigurationManager { get; private set; }

        //public XYZ LeaderEnd { get; set; }
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc.Document;
            Autodesk.Revit.DB.View activeView = doc.ActiveView;
            Autodesk.Revit.DB.View view = doc.ActiveView;
            ProjectInfo projectinfo = doc.ProjectInformation;


            IEnumerable<ViewSheet> sheetInst = new FilteredElementCollector(doc)
                .WhereElementIsNotElementType()
                .OfClass(typeof(ViewSheet))
                .OfCategory(BuiltInCategory.OST_Sheets)
                .Cast<ViewSheet>()
                .OrderBy(e => e.SheetNumber);

            IEnumerable<FamilyInstance> titleBlocks = new FilteredElementCollector(doc)
                .WhereElementIsNotElementType()
                .OfClass(typeof(FamilyInstance))
                .OfCategory(BuiltInCategory.OST_TitleBlocks)
                .Cast<FamilyInstance>()
                .OrderBy(e => e.Name);

            #region MyRegion


            //FilteredElementCollector sheetInst = new FilteredElementCollector(doc)
            //    .OfCategory(BuiltInCategory.OST_Sheets).WhereElementIsNotElementType();



            //IEnumerable<FamilySymbol> familyList = from elem in new FilteredElementCollector(doc)
            //                                       .OfClass(typeof(FamilySymbol))
            //                                       .OfCategory(BuiltInCategory.OST_TitleBlocks)
            //                                       let type = elem as FamilySymbol
            //                                       where type.Name.Contains("E1")
            //                                       select type;


            //DataTable tbTable = new DataTable();
            //tbTable.Columns.Add("Sheet Number", typeof(string));
            //tbTable.Columns.Add("Sheet Name", typeof(string));
            //tbTable.Columns.Add("DIP Stamp", typeof(int));
            //tbTable.Columns.Add("Rev Box 2", typeof(int));
            //tbTable.Columns.Add("Rev Box 3", typeof(int));
            //tbTable.Columns.Add("Key Plan", typeof(int)); 
            #endregion



            string sheetName = "";
            string sheetNumber = "";
            string sheetNumberWithPrefix = "";
            //string docType = "";
            double countOftB = 0;
            int dipStamp = 0;
            //int revBox2 = 0;
            int revBox3 = 0;
            int keyPlan = 0;


            //Parameters for Post to HTTP
            //the strings can be put into the function call directly
            UIdoc = commandData.Application.ActiveUIDocument;
            string userId = UIdoc.Application.Application.Username;
            string software = UIdoc.Application.Application.SubVersionNumber;
            var activeDocTitle = commandData.Application.ActiveUIDocument.Document.Title;
            //JsonAndPostToHttp("TBF", "TitleBlockFindnReplace", userId, software, activeDocTitle);
            //

            #region MyRegion
            //foreach (FamilyInstance tB in titleBlocks)
            //{
            //    if (tB.Name.Contains("ISO"))
            //    {
            //        countOftB++;
            //        sheetName = tB.get_Parameter(BuiltInParameter.SHEET_NUMBER).AsString();
            //        sheetNumber = tB.get_Parameter(BuiltInParameter.SHEET_NAME).AsString();
            //        dipStamp = tB.LookupParameter("Arup DIP Stamp").AsInteger();
            //        revBox2 = tB.LookupParameter("Rev Box 2").AsInteger();
            //        revBox3 = tB.LookupParameter("Rev Box 3").AsInteger();
            //        keyPlan = tB.LookupParameter("Key Plan").AsInteger();

            //        tbTable.Rows.Add(sheetName, sheetNumber, dipStamp, revBox2, revBox3, keyPlan);

            //    }

            //}
            //Application.Run(new SelectViewsFrPBPanel(tbTable)); 
            #endregion







            DataTable dt = new DataTable();
            string excelFileName = @"C:\Users\vince.bombardiere\OneDrive - Arup\280842 - SYD04\280842 - MSFT SYD04 Lane Cove - VAB Drawing list - Copy.xlsm";
            Workbook xlWorkBook;
            Excel.Application xlApp = new Excel.Application();
            Worksheet xlWorkSheet;
            int excelTabNumber = 9;//12=MasterList_Mech, 9=MEPFullSet
            string excelTabName = "MEPFullSet";
            int rangeRows, rangeColumns;
            GetExcelSheetNameAndRange(excelFileName, xlApp,
                //excelTabNumber, 
                excelTabName,
                out xlWorkBook, out xlWorkSheet, out rangeRows, out rangeColumns);
            //string stng = StringOfTextFromExcelSheet(xlWorkSheet, rangeRows, rangeColumns);
            //Clipboard.SetDataObject(stng);




            #region MyRegion
            //object[] mainBodyArray = new object[c];
            //string rowData = "";
            //for (int i = 1; i <= r; i++)//for each row
            //{
            //    DataRow dr = dt.NewRow();//set a new row item
            //    for (int j = 1; j <= c; j++)//for columns get the name allocate the value
            //    {
            //        string colName = Convert.ToString((xlWorkSheet.Cells[1, i] as Range).Value2);//getting the header name
            //        rowData = Convert.ToString((xlWorkSheet.Cells[j, i] as Range).Value2);
            //        mainBodyArray[j] = rowData;
            //    }
            //    dt.Rows.Add(mainBodyArray);//apply 
            //} 
            #endregion

            //Creating the table columns
            ConvertWorksheetToDataTable(dt, xlWorkSheet, rangeRows, rangeColumns);


            #region MyRegion
            //for (int columnCount = 0; columnCount < nColumns; columnCount++)//foreach column in each row
            //{
            //    rowData = ViewSchedule.GetCellText(SectionType.Body, rowCount, columnCount);
            //    mainBodyArray[columnCount] = rowData;
            //}
            //currentTable.Rows.Add(mainBodyArray);
            //string data = "";
            //for (int i = 0; i < getRowColumnCount(viewSchedule, sectionType, true); i++)
            //{
            //    for (int j = 0; j < getRowColumnCount(viewSchedule, sectionType, false); j++)
            //    {
            //        data += viewSchedule.GetCellText(sectionType, i, j);
            //        xlWorkSheet.Cells[i + 1, j + 1] = data;// Moving to first cell in Excel sheet
            //        data = "";
            //    }

            //} 
            #endregion
            #region MyRegion
            //Worksheet
            //DataSet ds = new DataSet();


            //////Working do not delete//////
            //dt.Columns.Add("Sheet Number");
            //dt.Columns.Add("ARUP_BDR_TITLE1");
            //dt.Columns.Add("ARUP_BDR_TITLE2");
            //dt.Columns.Add("ARUP_BDR_TITLE3");
            //dt.Columns.Add("ARUP_BDR_TITLE4");
            //dt.Columns.Add("IFT");


            //for (int i = 1; i <= r; i++)
            //{
            //    DataRow dr = dt.NewRow();
            //    dr["Sheet Number"] = Convert.ToString((xlWorkSheet.Cells[i, 1] as Range).Value2);
            //    dr["ARUP_BDR_TITLE1"] = Convert.ToString((xlWorkSheet.Cells[i, 2] as Range).Value2);
            //    dr["ARUP_BDR_TITLE2"] = Convert.ToString((xlWorkSheet.Cells[i, 3] as Range).Value2);
            //    dr["ARUP_BDR_TITLE3"] = Convert.ToString((xlWorkSheet.Cells[i, 4] as Range).Value2);
            //    dr["ARUP_BDR_TITLE4"] = Convert.ToString((xlWorkSheet.Cells[i, 5] as Range).Value2);
            //    dr["IFT"] = Convert.ToString((xlWorkSheet.Cells[i, 9] as Range).Value2);
            //    dt.Rows.Add(dr);
            //}
            //////Working do not delete////// 
            #endregion

            //close the workbook
            object misValue = System.Reflection.Missing.Value;
            xlWorkBook.Close(true, misValue, misValue);
            xlApp.Quit();

            Marshal.ReleaseComObject(xlWorkSheet);
            Marshal.ReleaseComObject(xlWorkBook);
            Marshal.ReleaseComObject(xlApp);

            #region MyRegion
            //StringBuilder sb2 = CreateStringBuilder(dt);
            //Clipboard.Clear();
            //string stng = sb2.ToString();
            //Clipboard.SetDataObject(stng);

            //Range r = ws.Range["A12"];
            //MessageBox.Show(r.Value); //value for special cell
            //MessageBox.Show(rSum);//value for sum a range 

            //double rangeTotal = rangeRows*rangeColumns;
            //var w = new System.Windows.Forms.Form();
            //Task.Delay(TimeSpan.FromSeconds(2.0))
            //    .ContinueWith((t) => w.Close(), TaskScheduler.FromCurrentSynchronizationContext());
            //MessageBox.Show(w, "Total Count  :-..." + rangeTotal.ToString() + "\nexiting function now", "Excel List total count");

            #endregion


            using (Transaction trans = new Transaction(doc, "Set parameter Value"))
            {
                trans.Start();


                #region MyRegion
                string sheetCurrentRevDate = "";
                Parameter sheetIssueDateParam = null;
                string sheetCurrentRevision = "";
                string sheetCurrentRevisionDescription = "";
                string sheetIssueTickBox = "";

                Parameter admin_Doc_Suitability = null;
                Parameter admin_Doc_Revision = null;
                Parameter admin_Doc_Role = null;
                string admin_Doc_RoleString = "";

                Parameter eleIdData = null;


                Parameter tblockPrefix = null;
                Parameter arup_bdr_title1 = null;
                Parameter arup_bdr_title2 = null;
                Parameter arup_bdr_title3 = null;
                Parameter arup_bdr_title4 = null;

                Parameter arup_bdr_role = null;
                Parameter arup_bdr_suitability = null;//not required

                Parameter arup_bdr_drawing_type = null;

                Parameter issueSixtyPc = null;
                Parameter issueDraftIFT = null;
                Parameter issueIFT = null;

                Parameter issue4Param = null;//Issue-4-Desc - XX/XX/2022
                Parameter issue5Param = null;//Issue-5-Desc - XX/XX/2022
                Parameter issue6Param = null;//Issue-6-Desc - XX/XX/2022
                Parameter issue7Param = null;//Issue-7-Desc - XX/XX/2022
                Parameter issue8Param = null;//Issue-8-Desc - XX/XX/2022
                Parameter issue9Param = null;//Issue-9-Desc - XX/XX/2022


                Parameter pkg01 = null;
                Parameter pkg02 = null;
                Parameter pkg03 = null;

                Parameter outputFromRevit = null;




                #endregion

                #region MyRegion
                //foreach title block
                //add datatable items

                //string testSt = titleBlocks.First().get_Parameter;

                //var w = new System.Windows.Forms.Form();
                //Task.Delay(TimeSpan.FromSeconds(2.0))
                //    .ContinueWith((t) => w.Close(), TaskScheduler.FromCurrentSynchronizationContext());
                //MessageBox.Show(w, "You Selected  :-..."+ testSt +"\nexiting function now", "Title block Replace Param");

                #endregion

                foreach (DataRow row in dt.Rows)
                {
                   // ... Write value of first field as integer.
                   //Console.WriteLine(row.Field<int>(0));

                    //var w = new System.Windows.Forms.Form();
                    //Task.Delay(TimeSpan.FromSeconds(0.3))
                    //    .ContinueWith((t) => w.Close(), TaskScheduler.FromCurrentSynchronizationContext());
                    //MessageBox.Show(w, "Sheet Name  :-..." +
                    //    "Name      - " + row.Field<string>(0) + "\n," +
                    //    "Draft IFT - " + row.Field<string>(8) + "\n," +
                    //    "IFT       - " + row.Field<string>(9) + "\n," +
                    //    "PKG01     - " + row.Field<string>(10) + "\n," +
                    //    "PKG02     - " + row.Field<string>(11) + "\n," +
                    //    "PKG03     - " + row.Field<string>(12) + "\n," +
                    //    "From Revit- " + row.Field<string>(13) + "\n," +
                    //    //row.Field<string>(14) + ", " +
                    //    "\nexiting function now", "Excel List Sheet No");
                }




                foreach (ViewSheet vs in sheetInst)
                {
                    tblockPrefix = vs.LookupParameter("Tblock_Prefix");
                    sheetNumber = vs.get_Parameter(BuiltInParameter.SHEET_NUMBER).AsString();
                    sheetNumberWithPrefix = vs.LookupParameter("Tblock_Prefix").AsString() + sheetNumber;
                    sheetCurrentRevision = vs.get_Parameter(BuiltInParameter.SHEET_CURRENT_REVISION).AsString();
                    sheetCurrentRevDate = vs.get_Parameter(BuiltInParameter.SHEET_CURRENT_REVISION_DATE).AsString();
                    sheetIssueDateParam = vs.get_Parameter(BuiltInParameter.SHEET_ISSUE_DATE);


                    //var w = new System.Windows.Forms.Form();
                    //Task.Delay(TimeSpan.FromSeconds(0.4))
                    //    .ContinueWith((t) => w.Close(), TaskScheduler.FromCurrentSynchronizationContext());
                    //MessageBox.Show(w, "Sheet Name  :-..." + sheetNumberWithPrefix +"\nexiting function now", "Excel List Sheet No");

                    

                    arup_bdr_title1 = vs.LookupParameter("ARUP_BDR_TITLE1");
                    arup_bdr_title2 = vs.LookupParameter("ARUP_BDR_TITLE2");
                    arup_bdr_title3 = vs.LookupParameter("ARUP_BDR_TITLE3");
                    arup_bdr_title4 = vs.LookupParameter("ARUP_BDR_TITLE4");

                    arup_bdr_role = vs.LookupParameter("ARUP_BDR_ROLE");
                    arup_bdr_suitability = vs.LookupParameter("ARUP_BDR_SUITABILITY");

                    arup_bdr_drawing_type = vs.LookupParameter("ARUP_BDR_DRAWING TYPE");

                    issueSixtyPc = vs.LookupParameter("Issue-1-60pcDD - 08/06/2022");
                    issueDraftIFT = vs.LookupParameter("Issue-2-Draft IFT - 14/07/2022");
                    issueIFT = vs.LookupParameter("Issue-3-Final IFT - 28/07/2022");

                    issue4Param = vs.LookupParameter("Issue-4-Desc - XX/XX/2022");
                    issue5Param = vs.LookupParameter("Issue-5-Desc - XX/XX/2022");
                    issue6Param = vs.LookupParameter("Issue-6-Desc - XX/XX/2022");
                    issue7Param = vs.LookupParameter("Issue-7-Desc - XX/XX/2022");
                    issue8Param = vs.LookupParameter("Issue-8-Desc - XX/XX/2022");
                    issue9Param = vs.LookupParameter("Issue-9-Desc - XX/XX/2022");

                    pkg01 = vs.LookupParameter("PKG01");
                    pkg02 = vs.LookupParameter("PKG02");
                    pkg03 = vs.LookupParameter("PKG03");

                    foreach (DataRow row in dt.Rows)
                    {
                        if (row.Field<string>(0) == sheetNumberWithPrefix)//if datatable sheet number equals revit sheet number
                        {
                            //var w = new System.Windows.Forms.Form();
                            //Task.Delay(TimeSpan.FromSeconds(0.3))
                            //    .ContinueWith((t) => w.Close(), TaskScheduler.FromCurrentSynchronizationContext());
                            //MessageBox.Show(w, "Sheet No  :-...\n"+ 
                            //    row.Field<string>(0) + " - "+sheetNumberWithPrefix + "\nexiting function now", "Excel List Sheet No");

                            string title1 = row[1].ToString(); arup_bdr_title1.Set(title1);
                            string title2 = row[2].ToString(); arup_bdr_title2.Set(title2);
                            string title3 = row[3].ToString(); arup_bdr_title3.Set(title3);
                            string title4 = row[4].ToString(); arup_bdr_title4.Set(title4);

                            string role = row[5].ToString(); arup_bdr_role.Set(role);
                            string type = row[6].ToString(); arup_bdr_drawing_type.Set(type);


                            //string issue1 = row[7].ToString(); if (issue1 == "Yes") { issueSixtyPc.Set(1); }
                            string issue2 = row[8].ToString(); if (issue2 == "Yes") { issueDraftIFT.Set(1); } else { issueDraftIFT.Set(0); }
                            string issue3 = row[9].ToString(); if (issue3 == "Yes") { issueIFT.Set(1); } else { issueIFT.Set(0); }
                            issue4Param.Set(0);
                            issue5Param.Set(0);
                            issue6Param.Set(0);
                            issue7Param.Set(0);
                            issue8Param.Set(0);
                            issue9Param.Set(0);


                            string stringPkg01 = row[10].ToString(); if (stringPkg01 == "Yes") { pkg01.Set(1); } else { pkg01.Set(1); }
                            string stringPkg02 = row[11].ToString(); if (stringPkg02 == "Yes") { pkg02.Set(1); } else { pkg02.Set(0); }
                            string stringPkg03 = row[12].ToString(); if (stringPkg03 == "Yes") { pkg03.Set(1); } else { pkg03.Set(0); }


                            sheetIssueDateParam.Set(sheetCurrentRevDate);//set the sheet issue date to the current revision date



                        }
                    }

                    #region MyRegion
                    //admin_Doc_Revision = vs.LookupParameter("Admin_Document_Revision");
                    //sheetIssueDateParam.Set(sheetCurrentRevDate);
                    //admin_Doc_Revision.Set(sheetCurrentRevision);

                    // docType = vs.LookupParameter("Aconex - Type");

                    //if (sheetCurrentRevDate == "13/05/22")
                    //{
                    //    aconexPhase = vs.LookupParameter("Aconex - Phase");
                    //    aconexPhase.Set("Phase 4, Phase 5");

                    //}

                    //if (sheetCurrentRevDate == "24/03/22")
                    //{
                    //    //eleIdData = vs.LookupParameter("Phase"); //this will change the drawing number
                    //    //eleIdData.Set("Phase 4, Phase 5");

                    //    aconexPhase = vs.LookupParameter("Aconex - Phase");
                    //    //aconexPhase.Set("Phase 2, Phase 3, Phase 4, Phase 5");
                    //    //aconexPhase.Set("Phase 4, Phase 5");
                    //} 



                    //aconexTypeParam = vs.LookupParameter("Aconex - Type");
                    //aconexTypeParam.Set("Drawing");
                    #endregion
                    #region Older code
                    //set Aconex - Status to tender
                    //sheetIssueTickBox = vs.LookupParameter("Issue - Tender Addendum 1 - Phase 2+ 3  16/02/2022").AsValueString();
                    //sheetIssueTickBox = vs.LookupParameter("Issue - Final FOH - 04/03/2022").AsValueString();

                    //aconexStatusParam = vs.LookupParameter("Aconex - Status");
                    //admin_Doc_Suitability = vs.LookupParameter("Admin_Document_Suitability");
                    //if (sheetIssueTickBox == "Yes") //param for this issue tick box selected
                    //{
                    //    aconexStatusParam.Set("For Tender");
                    //    admin_Doc_Suitability.Set("For Tender");
                    //}05 Detailed Design & Documentation

                    //sheetIssueTickBox = vs.LookupParameter("Issue - Prelim Ph 4+5 - 23/02/2022").AsValueString(); 


                    //sheetIssueTickBox = vs.LookupParameter("Issue-2-Final IFC - 14/07/2022").AsValueString();
                    //aconexStatusParam = vs.LookupParameter("Aconex - Status");
                    //admin_Doc_Suitability = vs.LookupParameter("Admin_Document_Suitability");

                    //aconexStage = vs.LookupParameter("Aconex - Stage");

                    //if (sheetIssueTickBox == "Yes") //param for this issue tick box selected
                    //{
                    //    //aconexStatusParam.Set("For Information");
                    //    //admin_Doc_Suitability.Set("70% Design Development");
                    //    //aconexStage.Set("05 Detailed Design & Documentation");

                    //}
                    #endregion
                    #region Older code
                    //admin_Doc_Role = vs.LookupParameter("Admin_Document_Role");
                    //admin_Doc_RoleString = vs.LookupParameter("Admin_Document_Role").AsString();
                    //aconexDiscipline = vs.LookupParameter("Aconex - Discipline");
                    //if (admin_Doc_RoleString?.Length == 0) { aconexDiscipline.Set("Not Listed"); }
                    //if (admin_Doc_RoleString == "Combined Services") { aconexDiscipline.Set("Combined Services"); }
                    //if (admin_Doc_RoleString == "Communications") { aconexDiscipline.Set("Telecommunications"); }
                    //if (admin_Doc_RoleString == "Electrical Services") { aconexDiscipline.Set("Electrical"); }
                    //if (admin_Doc_RoleString == "Facility Control and Monitoring System") { aconexDiscipline.Set("Controls"); }
                    //if (admin_Doc_RoleString == "Fire Services") { aconexDiscipline.Set("Fire Services"); }
                    //if (admin_Doc_RoleString == "Hydraulic") { aconexDiscipline.Set("Hydraulic"); }
                    //if (admin_Doc_RoleString == "Mechanical Services") { aconexDiscipline.Set("Mechanical"); }
                    //if (admin_Doc_RoleString == "Security") { aconexDiscipline.Set("Security"); }
                    //aconexFileParam = vs.LookupParameter("Aconex - File");
                    //string aconexFile = "SY6-2-DRG-" + sheetNumber + "_" + sheetCurrentRevision + ".pdf";
                    //aconexFileParam.Set(aconexFile); 


                    //aconexCategory = vs.LookupParameter("Aconex - Auth Orgntn");
                    //aconexCategory.Set("25043 Architectural");


                    //aconexAuthoringOrganisation = vs.LookupParameter("Aconex - Category");
                    //aconexAuthoringOrganisation.Set("Arup");

                    //aconexPrintSizeParam = vs.LookupParameter("Aconex - Print Size");
                    //aconexPrintSizeParam.Set("A0");
                    #endregion
                    #region Older code
                    //aconexRevDateParam = vs.LookupParameter("Aconex - Revision Date");
                    //aconexRevDateParam.Set(sheetCurrentRevDate);

                    //aconexCreatedBy = vs.LookupParameter("Aconex - Created By");
                    //aconexCreatedBy.Set("Arup");

                    //if tblock name equals vs name then set the dip to true
                    //and if current rev date is set to the next issue
                    //if (sheetCurrentRevDate == "30/03/22")
                    //{
                    //    foreach (FamilyInstance tB in titleBlocks)
                    //    {
                    //        string tBlockNumber = tB.get_Parameter(BuiltInParameter.SHEET_NUMBER).AsString();
                    //        string tBlockName = tB.get_Parameter(BuiltInParameter.ELEM_FAMILY_PARAM).AsValueString();//filter out the cover sheet title blocks

                    //        if (tBlockNumber == sheetNumber && tBlockName.Contains("ISO"))
                    //        {
                    //            Parameter dIPStamp = tB.LookupParameter("Arup DIP Stamp"); dIPStamp.Set(0);
                    //            Parameter rev1 = tB.LookupParameter("Rev 1"); rev1.Set(1);
                    //            Parameter rev2 = tB.LookupParameter("Rev 2"); rev2.Set(1);
                    //            Parameter rev3 = tB.LookupParameter("Rev 3"); rev3.Set(1);
                    //            Parameter rev4 = tB.LookupParameter("Rev 4"); rev4.Set(1);
                    //            Parameter rev5 = tB.LookupParameter("Rev 5"); rev5.Set(1);
                    //            Parameter rev6 = tB.LookupParameter("Rev 6"); rev6.Set(1);
                    //            Parameter revBox2 = tB.LookupParameter("Rev Box 2"); revBox2.Set(1);
                    //        }

                    //    }
                    //}
                    //if (sheetCurrentRevDate == "04/03/22")
                    //{
                    //    foreach (FamilyInstance tB in titleBlocks)
                    //    {
                    //        string tBlockNumber = tB.get_Parameter(BuiltInParameter.SHEET_NUMBER).AsString();
                    //        string tBlockName = tB.get_Parameter(BuiltInParameter.ELEM_FAMILY_PARAM).AsValueString();//filter out the cover sheet title blocks

                    //        if (tBlockNumber == sheetNumber && tBlockName.Contains("ISO"))
                    //        {
                    //            Parameter dIPStamp = tB.LookupParameter("Arup DIP Stamp"); dIPStamp.Set(0);
                    //            Parameter rev1 = tB.LookupParameter("Rev 1"); rev1.Set(1);
                    //            Parameter rev2 = tB.LookupParameter("Rev 2"); rev2.Set(1);
                    //            Parameter rev3 = tB.LookupParameter("Rev 3"); rev3.Set(1);
                    //            Parameter rev4 = tB.LookupParameter("Rev 4"); rev4.Set(1);
                    //            Parameter rev5 = tB.LookupParameter("Rev 5"); rev5.Set(1);
                    //            Parameter rev6 = tB.LookupParameter("Rev 6"); rev6.Set(1);
                    //            Parameter revBox2 = tB.LookupParameter("Rev Box 2"); revBox2.Set(1);
                    //        }

                    //    }
                    //}






                    #endregion

                }



                trans.Commit();
                return Result.Succeeded;
            }
        }

        private void ConvertWorksheetToDataTable(DataTable dt, Worksheet xlWorkSheet, int rangeRows, int rangeColumns)
        {
            //for (int i = 1; i < rangeColumns; i++)//Create column names
            //{
            //    string cName = Convert.ToString((xlWorkSheet.Cells[1, i] as Range).Value2);
            //    dt.Columns.Add(cName);
            //}

            dt.Columns.Add("Sheet Number");
            dt.Columns.Add("ARUP_BDR_TITLE1");
            dt.Columns.Add("ARUP_BDR_TITLE2");
            dt.Columns.Add("ARUP_BDR_TITLE3");
            dt.Columns.Add("ARUP_BDR_TITLE4");
            dt.Columns.Add("ARUP_BDR_ROLE");
            dt.Columns.Add("ARUP_BDR_DRAWING TYPE");
            dt.Columns.Add("60%");
            dt.Columns.Add("Draft IFT");
            dt.Columns.Add("IFT");
            dt.Columns.Add("SYD04 Colo 1");
            dt.Columns.Add("SYD04 Colo 2");
            dt.Columns.Add("SYD04 Colo 3");
            dt.Columns.Add("Output from Revit");

            StringBuilder bx = new StringBuilder();
            //setting up parameters for data table
            string sheetNumber = "";
            string arup_bdr_title1 = ""; string arup_bdr_title2 = ""; string arup_bdr_title3 = ""; string arup_bdr_title4 = "";
            string arup_bdr_role = "";
            string arup_bdr_drawing_type = "";
            string issueSixtyPc = ""; string issueDraftIFT = ""; string issueIFT = "";
            string package01 = ""; string package02 = ""; string package03 = "";
            string outputFromRevit = "";
            for (int i = 2; i < rangeRows; i++)
            {
                int count = 1;
                sheetNumber = Convert.ToString((xlWorkSheet.Cells[i, count] as Range).Value2); count++;
                arup_bdr_title1 = Convert.ToString((xlWorkSheet.Cells[i, count] as Range).Value2); count++;
                arup_bdr_title2 = Convert.ToString((xlWorkSheet.Cells[i, count] as Range).Value2); count++;
                arup_bdr_title3 = Convert.ToString((xlWorkSheet.Cells[i, count] as Range).Value2); count++;
                arup_bdr_title4 = Convert.ToString((xlWorkSheet.Cells[i, count] as Range).Value2); count++;
                arup_bdr_role = Convert.ToString((xlWorkSheet.Cells[i, count] as Range).Value2); count++;
                arup_bdr_drawing_type = Convert.ToString((xlWorkSheet.Cells[i, count] as Range).Value2); count++;
                issueSixtyPc = Convert.ToString((xlWorkSheet.Cells[i, count] as Range).Value2); count++;
                issueDraftIFT = Convert.ToString((xlWorkSheet.Cells[i, count] as Range).Value2); count++;
                issueIFT = Convert.ToString((xlWorkSheet.Cells[i, count] as Range).Value2); count++;
                package01 = Convert.ToString((xlWorkSheet.Cells[i, count] as Range).Value2); count++;
                package02 = Convert.ToString((xlWorkSheet.Cells[i, count] as Range).Value2); count++;
                package03 = Convert.ToString((xlWorkSheet.Cells[i, count] as Range).Value2); count++;
                outputFromRevit = Convert.ToString((xlWorkSheet.Cells[i, count] as Range).Value2);

                dt.Rows.Add(sheetNumber, arup_bdr_title1, arup_bdr_title2, arup_bdr_title3, arup_bdr_title4, arup_bdr_role, arup_bdr_drawing_type, issueSixtyPc, issueDraftIFT, issueIFT, package01, package02, package03, outputFromRevit);

            }

            StringBuilder sb2 = CreateStringBuilder(dt);
            Clipboard.Clear();
            string stng1 = sb2.ToString();
            Clipboard.SetDataObject(stng1);
        }

        private static void GetExcelSheetNameAndRange(string excelFileName, Excel.Application xlApp, 
            //int excelTabNumber,
            string excelTabName,
            out Workbook xlWorkBook, out Worksheet xlWorkSheet, out int rangeRows, out int rangeColumns)
        {
            xlWorkBook = xlApp.Workbooks.Open(excelFileName);
            //xlWorkSheet = (Worksheet)xlWorkBook.Worksheets.get_Item(excelTabNumber);
            xlWorkSheet = (Worksheet)xlWorkBook.Worksheets.get_Item(excelTabName);
            Range range = xlWorkSheet.UsedRange;
            rangeRows = range.Rows.Count;
            rangeColumns = range.Columns.Count;

            //double rangeTotal = rangeRows*rangeColumns;
            //var w = new System.Windows.Forms.Form();
            //Task.Delay(TimeSpan.FromSeconds(2.0))
            //    .ContinueWith((t) => w.Close(), TaskScheduler.FromCurrentSynchronizationContext());
            //MessageBox.Show(w, "Total Count  :-..." + rangeTotal.ToString() + "\nexiting function now", "Excel List total count");
        }

        private static string StringOfTextFromExcelSheet(Worksheet xlWorkSheet, int r, int c)
        {
            StringBuilder bx = new StringBuilder();
            for (int i = 1; i < r; i++)
            {
                string rowData = "";
                for (int j = 1; j < c; j++)
                {
                    rowData = Convert.ToString((xlWorkSheet.Cells[i, j] as Range).Value2);
                    bx.Append(rowData).Append("\t");
                }
                bx.Append("\n");
            }


            string stng = bx.ToString();
            return stng;
        }

        StringBuilder CreateStringBuilder(DataTable datTbl)
        {
            StringBuilder bx = new StringBuilder();
            foreach (DataRow r in datTbl.Rows)
            {
                foreach (DataColumn c in datTbl.Columns)
                { bx.Append(r[c.ColumnName].ToString()).Append("\t"); }
                bx.Append("\n");
            }

            return bx;
        }


        private static void JsonAndPostToHttp(string toolCode, string toolName, string userId, string software, string activeDocTitle)
        {
            var TableDic = new Dictionary<string, object>
                {
                { "toolCode", toolCode },
                { "toolName", toolName },
                { "userId", userId },
                { "software", software }
                };

            var pdList = new List<Dictionary<string, object>> { TableDic };
            string jsonInfo = JsonConvert.SerializeObject(pdList[0], Formatting.Indented);
            


#if DEBUG
            
            string Today = DateTime.Now.ToString("yyyyMMdd");
            string fPath = @"C:\Temp\ToolUsage1.json";

            //if (File.Exists(fPath)) { File.Delete(fPath); }
            //File.WriteAllText(fPath, jsonInfo);
            System.IO.File.AppendAllText(fPath, jsonInfo);



            //var w = new System.Windows.Forms.Form();
            //Task.Delay(TimeSpan.FromSeconds(8.0))
            //    .ContinueWith((t) => w.Close(), TaskScheduler.FromCurrentSynchronizationContext());
            //MessageBox.Show(w, "Doc title  :-..." +
            //    jsonInfo +
            //    fPath +
            //    "\nexiting function now", "Active doc title");




#endif
            //post to http in this loop 
            HttpClient client = new HttpClient();
            var url = "https://uam8f4ihb6.execute-api.ap-southeast-2.amazonaws.com/dev/log";
            StringContent httpContent1 = new StringContent(jsonInfo, System.Text.Encoding.UTF8, "application/json");
            var response1 = client.PostAsync(url, httpContent1);

        }



    }
}

