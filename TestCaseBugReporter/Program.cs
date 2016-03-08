using System;
using System.Net;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;

using Microsoft.TeamFoundation.Client;
using Microsoft.TeamFoundation.TestManagement.Client;
using Microsoft.TeamFoundation.TestManagement.Common;
using Microsoft.TeamFoundation.WorkItemTracking.Client;



namespace TestCaseBugReporter
{
    class Program
    {
        static void AddPBIData(CreateExcelDoc excell_app, WorkItemStore myWorkItemStore, string iterationPath)
        {
            int row = 1, col = 1;
            excell_app.addData(row, col++, "WorkItem Type", "A1", "B" + col, "#,##0");
            excell_app.addData(row, col++, "WorkItem ID", "A1", "B" + col, "#,##0");
            excell_app.addData(row, col++, "Title", "A1", "B" + col, "#,##0");
            excell_app.addData(row, col++, "TestCase", "A1", "B" + col, "#,##0");
            excell_app.addData(row, col++, "TestCase Title", "A1", "B" + col, "#,##0");

            row++;
            col = 1;

            WorkItemCollection workItemCollection = myWorkItemStore.Query(" SELECT [System.Id]" + " FROM WorkItems " + " WHERE [System.IterationPath] = '" + iterationPath + "' " + "ORDER BY [System.Id]");
            foreach (WorkItem wi in workItemCollection)
            {
                if (wi.Type.Name.Equals("Product Backlog Item") || wi.Type.Name.Equals("Bug"))
                {
                    col = 1;
                    bool foundTest = false;
                    WorkItemLinkCollection wilinks = wi.WorkItemLinks;
                    foreach (WorkItemLink link in wilinks)
                    {
                        col = 4;
                        RegisteredLinkType t = link.ArtifactLinkType;
                        WorkItemLinkTypeEnd typeend = link.LinkTypeEnd;
                        if (!typeend.ImmutableName.Contains("TestedBy"))
                        {
                            continue;
                        }
                        if (!foundTest)
                        {
                            col = 1;
                            excell_app.addData(row, col, wi.Type.Name, "A" + row, "B" + col, "#,##0"); col++;
                            excell_app.addData(row, col, wi.Id.ToString(), "A" + row, "B" + col, "#,##0"); col++;
                            excell_app.addData(row, col, wi.Title, "A" + row, "B" + col, "#,##0"); col++;
                            foundTest = true;
                            col = 4;
                        }
                        excell_app.addData(row, col, link.TargetId.ToString(), "A" + row, "B" + col, "#,##0"); col++;
                        WorkItem linkedTest = myWorkItemStore.GetWorkItem(link.TargetId);
                        excell_app.addData(row, col, linkedTest.Title, "A" + row, "B" + col, "#,##0"); col++;
                        //Console.WriteLine("PBI :" + link.SourceId + ": TestCase" + link.TargetId);
                        row++;

                    }
                    if (!foundTest)
                    {
                        col = 1;
                        excell_app.addData(row, col, wi.Type.Name, "A" + row, "B" + col, "#,##0"); col++;
                        excell_app.addData(row, col, wi.Id.ToString(), "A" + row, "B" + col, "#,##0"); col++;
                        excell_app.addData(row, col, wi.Title, "A" + row, "B" + col, "#,##0"); col++;
                        excell_app.addData(row, col, "NOT FOUND", "A" + row, "B" + col, "#,##0"); col++;
                        row++;
                        //Console.WriteLine("No test found for WI:" + wi.Id);
                    }
                }
            }
        }

        static void AddBugData(CreateExcelDoc excell_app, WorkItemStore myWorkItemStore, WorkItemCollection workItemCollection)
        {
            int row = 1, col = 1;
            excell_app.addData(row, col++, "ID", "A1", "B" + col, "#,##0");
            excell_app.addData(row, col++, "WorkItem Type", "A1", "B" + col, "#,##0");
            excell_app.addData(row, col++, "Title", "A1", "B" + col, "#,##0");
            excell_app.addData(row, col++, "State", "A1", "B" + col, "#,##0");
            excell_app.addData(row, col++, "Created By", "A1", "B" + col, "#,##0");
            excell_app.addData(row, col++, "Iteration Path", "A1", "B" + col, "#,##0");
            excell_app.addData(row, col++, "Area Path", "A1", "B" + col, "#,##0");
            excell_app.addData(row, col++, "Created Date", "A1", "B" + col, "#,##0");

            row++;
            col = 1;

            foreach (WorkItem wi in workItemCollection)
            {
                col = 1;
                excell_app.addData(row, col, wi.Id.ToString(), "A" + row, "B" + col, "#,##0"); col++;
                excell_app.addData(row, col, wi.Type.Name, "A" + row, "B" + col, "#,##0"); col++;
                excell_app.addData(row, col, wi.Title, "A" + row, "B" + col, "#,##0"); col++;
                excell_app.addData(row, col, wi.State, "A" + row, "B" + col, "#,##0"); col++;
                excell_app.addData(row, col, wi.CreatedBy, "A" + row, "B" + col, "#,##0"); col++;
                excell_app.addData(row, col, wi.IterationPath, "A" + row, "B" + col, "#,##0"); col++;
                excell_app.addData(row, col, wi.AreaPath, "A" + row, "B" + col, "#,##0"); col++;
                excell_app.addData(row, col, wi.CreatedDate.ToString(), "A" + row, "B" + col, "#,##0"); col++;
                row++;
            }
        }
        static void Main(string[] args)
        {
            int testPlanNum;
            if (args.Length != 7)
            {
                System.Console.WriteLine("Please enter the 5 arguments");
                System.Console.WriteLine("Usage: TestCaseBugReporter.exe <TFS Server> <Project Name> <Test Plan Number> <Full Path to result spreadsheet> <Iteration Path> <sprint Start Date> <sprint End date>");
                return;
            }
            try
            {               
                testPlanNum = int.Parse(args[2]);                
            }
            catch (System.FormatException)
            {
                System.Console.WriteLine("Please enter a numeric argument.");
                System.Console.WriteLine("Usage: Usage: TestCaseBugReporter.exe <TFS Server> <Project Name> <Test Plan Number> <Full Path to result spreadsheet> <Iteration Path>");
                return;
            }
         
            string serverurl = args[0];
            string workBookName = args[3]; 
            string iterationPath = args[4];            
            string sprintStartDate = args[5];
            string sprintEndDate =  args[6];

            string project = args[1]; 
            TfsTeamProjectCollection tfsCollection = new TfsTeamProjectCollection(TfsTeamProjectCollection.GetFullyQualifiedUriForName(serverurl));
            WorkItemStore myWorkItemStore = (WorkItemStore)tfsCollection.GetService(typeof(WorkItemStore));
            ITestManagementService tms = (ITestManagementService)tfsCollection.GetService(typeof(ITestManagementService));
            ITestManagementTeamProject proj = null;
            
            proj = tms.GetTeamProject(project);
            ITestPlanHelper planHelper = proj.TestPlans;
            ITestPlan foundPlan = null;
            try
            {
                foundPlan = proj.TestPlans.Find(testPlanNum); 
                Console.WriteLine("Got Plan {0} with Id {1}", foundPlan.Name, foundPlan.Id);
            }
            catch(System.NullReferenceException)
            {
                Console.WriteLine("No Plan with ID: {0} found!", testPlanNum);
                return;
            }
            
            int col = 1;
            CreateExcelDoc excell_app = new CreateExcelDoc();

            WorkItemCollection workItemCollection = myWorkItemStore.Query(" SELECT [System.Id]" + " FROM WorkItems " + " WHERE [System.WorkItemType] = 'Bug' AND [System.State] <> 'Done' AND [System.State] <> 'Removed'" + "ORDER BY [System.Id]");
            AddBugData(excell_app, myWorkItemStore, workItemCollection);

            excell_app.createAndMoveToNextWS("Bugs Fixed during the Sprint");
            string queryBuilder = "SELECT [System.Id]" + " FROM WorkItems " + " WHERE [System.WorkItemType] = 'Bug' AND [System.AreaPath] not under 'Glasswall\\Glasswall QFE Team' AND [Microsoft.VSTS.Common.ClosedDate] >= '" + sprintStartDate + "'" + " AND  [Microsoft.VSTS.Common.ClosedDate] <= '" + sprintEndDate + "' " + "ORDER BY [System.Id]";
            workItemCollection = myWorkItemStore.Query(queryBuilder);
            AddBugData(excell_app, myWorkItemStore, workItemCollection); 
            
            excell_app.createAndMoveToNextWS("Bugs Created during the Sprint");
            workItemCollection = myWorkItemStore.Query(" SELECT [System.Id]" + " FROM WorkItems " + " WHERE [System.WorkItemType] = 'Bug' AND [System.AreaPath] not under 'Glasswall\\Glasswall QFE Team' AND [System.CreatedDate] >= '" + sprintStartDate + "' " + " AND  [System.CreatedDate] <= '" + sprintEndDate + "' " + "ORDER BY [System.Id]");
            AddBugData(excell_app, myWorkItemStore, workItemCollection);           

            excell_app.createAndMoveToNextWS("PBI Linked Testcases");
            AddPBIData(excell_app, myWorkItemStore,iterationPath);

            excell_app.createAndMoveToNextWS("Detailed Test results");
            excell_app.addData(1, col++, "TestCase ID", "A1", "B" + col, "#,##0");
            excell_app.addData(1, col++, "Title", "A1", "B" + col, "#,##0");
            excell_app.addData(1, col++, "State", "A1", "B" + col, "#,##0");
            excell_app.addData(1, col++, "Result", "A1", "B" + col, "#,##0"); 
            excell_app.addData(1, col++, "Area Path", "A1", "B" + col, "#,##0");
            excell_app.addData(1, col++, "Iteration Path", "A1", "B" + col, "#,##0");
            excell_app.addData(1, col++, "Tester", "A1", "B" + col, "#,##0");
            excell_app.addData(1, col++, "Last Run", "A1", "B" + col, "#,##0");
            excell_app.addData(1, col++, "Linked Bugs", "A1", "B" + col, "#,##0");
            
            int num = 0, row = 2;
            col = 1;
            ITestPointCollection testPoints = foundPlan.QueryTestPoints(string.Format("SELECT * FROM TestPoint"));
            foreach( ITestPoint testpoint in testPoints)
            {
                col = 1;
                num++;              
                ITestCase testcase = testpoint.TestCaseWorkItem;
                LinkCollection links = testcase.Links;
                               
                string id = testcase.Id.ToString();
                excell_app.addData(row, col,id , "A" + row, "B" + col, "#,##0");
                col++;
                string title = testcase.Title;
                excell_app.addData(row, col, title, "A" + row, "B" + col, "#,##0");
                col++;
                excell_app.addData(row, col, testpoint.TestCaseWorkItem.State.ToString(), "A" + row, "B" + col, "#,##0");                
                col++;
                string result = testpoint.MostRecentResultOutcome.ToString();
                excell_app.addData(row, col, result, "A" + row, "B" + col, "#,##0");
                col++;
                excell_app.addData(row, col, testpoint.TestCaseWorkItem.Area, "A" + row, "B" + col, "#,##0");
                col++;
                excell_app.addData(row, col, testpoint.TestCaseWorkItem.WorkItem.IterationPath, "A" + row, "B" + col, "#,##0");
                col++;
                string tester = testpoint.AssignedTo == null?  "" : testpoint.AssignedTo.DisplayName;
                excell_app.addData(row, col, tester, "A" + row, "B" + col, "#,##0");
                col++;
                excell_app.addData(row, col, testpoint.LastUpdated.ToString(), "A" + row, "B" + col, "#,##0");

                if (testpoint.MostRecentResultOutcome == TestOutcome.Passed)
                {
                    row++;
                    continue;
                }
                   
                String bugList = "";

                int[] workItemList = testpoint.QueryAssociatedWorkItemsFromResults();
              
                if (null == workItemList && null == links)
                {
                    row++;
                    continue;
                }

                if (null != workItemList)
                {
                    foreach (int workItemId in workItemList)
                    {
                        WorkItem currentWorkItem = myWorkItemStore.GetWorkItem(workItemId);
                        if (currentWorkItem.Type.Name.ToUpper() == "BUG" && (!currentWorkItem.State.Equals("Closed") || currentWorkItem.Reason.Equals("Deferred")))
                        {
                            string seperator;
                            if ("" == bugList) seperator = ""; else seperator = "  , ";
                            bugList = bugList + seperator + workItemId + " : " + currentWorkItem.State + " - " + currentWorkItem.Title;
                        }

                    }
                }

                if (null != links)
                {
                    foreach (Link link in links)
                    {
                        if (link.BaseType != BaseLinkType.RelatedLink)
                            continue;
                        
                        var wi = myWorkItemStore.GetWorkItem(((Microsoft.TeamFoundation.WorkItemTracking.Client.RelatedLink)(link)).RelatedWorkItemId);
                        
                        WorkItemType witype = wi.Type;
                        string wiNumber = ((Microsoft.TeamFoundation.WorkItemTracking.Client.RelatedLink)(link)).RelatedWorkItemId.ToString();
                        if (witype.Name.Equals("Bug") && !bugList.Contains(wiNumber) && (!wi.State.Equals("Done") || wi.Reason.Equals("Deferred")))
                        {
                            string seperator;
                            if ("" == bugList) seperator = ""; else seperator = "  , ";
                            bugList = bugList + seperator + wiNumber + " :: " + wi.State + " - " + wi.Title;
                        }
                    }
                }
                col++;
                if(null != bugList) excell_app.addData(row, col, bugList, "A" + row, "B" + col, "#,##0");
                row++;
            }

            excell_app.save(workBookName);
            Console.WriteLine("\nTest Bug Reporter : Job Complete\n");
            return;
        }

       

    }
}
 

