using System;
using System.Windows;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

//word API
using Microsoft.Office.Interop.Word;

//MTM/TFS API
using Microsoft.TeamFoundation;
using Microsoft.TeamFoundation.Client;
using Microsoft.TeamFoundation.TestManagement.Client;
using Microsoft.TeamFoundation.WorkItemTracking;




//Test Management API
//for web tfs api:
//https://docs.microsoft.com/en-us/dotnet/api/microsoft.teamfoundation.testmanagement.webapi?view=azure-devops-dotnet

//for client application:
//https://docs.microsoft.com/en-us/previous-versions/dd998375%28v%3dvs.140%29
//Classes: HierarchyEntry and TestActionCollection

namespace WordReader
{

    public struct TestCase
    {

        public string nLot,
                      nDocument,
                      champ,
                      test,
                      expected,
                      result1,
                      result2;

                      
    }
    class Program
    {
        //LOL

        static ITestManagementTeamProject GetProject(string serverUrl, string project)
        {
            TfsTeamProjectCollection tfs = new TfsTeamProjectCollection(TfsTeamProjectCollection.GetFullyQualifiedUriForName(serverUrl));
            ITestManagementService tms = tfs.GetService<ITestManagementService>();

            return tms.GetTeamProject(project);
        }


        static void Main(string[] args)
        {

            Application word = new Application();
            Document doc = new Document();

            // Define an object to pass to the API for missing parameters
            object missing = System.Type.Missing;
            doc = word.Documents.Open(@"D:\RQemploi\DA0_Devis_RAC_PAL18.docx",
                    ref missing, ref missing, ref missing, ref missing,
                    ref missing, ref missing, ref missing, ref missing,
                    ref missing, ref missing, ref missing, ref missing,
                    ref missing, ref missing, ref missing);

            List<TestCase> data = new List<TestCase>();

            foreach(Table table in doc.Tables)
            {
                for(int i = 1; i <= table.Rows.Count; i++)
                {
                    bool isHeader = false;
                    TestCase temp = new TestCase();
                    int cellNum = 0;
                    for(int j = 1; j <= table.Columns.Count; j++)
                    {
                        
                        Cell cell = null;
                        try
                        {
                            cell = table.Cell(i, j);
                        }
                        catch (System.Runtime.InteropServices.COMException e)
                        {
                            isHeader = true;
                        }
                        if (!isHeader)
                        {
                            switch (cellNum)
                            {
                                case 0:
                                    temp.nLot = cell.Range.Text;
                                    break;
                                case 1:
                                    temp.nDocument = cell.Range.Text;
                                    break;
                                case 2:
                                    temp.champ = cell.Range.Text;
                                    break;
                                case 3:
                                    temp.test = cell.Range.Text;
                                    break;
                                case 4:
                                    temp.expected = cell.Range.Text;
                                    break;
                                case 5:
                                    temp.result1 = cell.Range.Text;
                                    break;
                                case 6:
                                    temp.result2 = cell.Range.Text;
                                    break;
                            }
                        }
                        cellNum++;
                    }
                    if(!isHeader)
                        data.Add(temp);
                }
            }

            foreach(TestCase tc in data)
            {
                Console.WriteLine(tc.nLot);
                Console.WriteLine(tc.nDocument);
                Console.WriteLine(tc.champ);
                Console.WriteLine(tc.test);
                Console.WriteLine(tc.expected);
                Console.WriteLine(tc.result1);
                Console.WriteLine(tc.result2);
            }
            

            Console.ReadLine();

            ((_Document)doc).Close();
            ((_Application)word).Quit();

            /*
            //MTM

            string serverurl = "http://localhost:8080/tfs";
            string project = "project name on tfs server";
            ITestManagementTeamProject proj = GetProject(serverurl, project);

            //create plan
            ITestPlan plan = proj.TestPlans.Create();
            plan.Name = "sprint name";
            plan.StartDate = DateTime.Now;
            plan.EndDate = DateTime.Now.AddMonths(1);
            plan.Save();

            //create suite for plan
            IStaticTestSuite suite = proj.TestSuites.CreateStatic();
            suite.Title = "Types of test cases";

            plan.RootSuite.Entries.Add(suite);
            plan.Save();

            //Config for adding test cases to suite
            ITestConfiguration defaultConfig = null;

            foreach (ITestConfiguration config in proj.TestConfigurations.Query(
                "Select * from TestConfiguration"))
            {
                defaultConfig = config;
                break;
            }

            //creating the test case to the suite
            ITestCase testCase = proj.TestCases.Create();
            testCase.Title = "Verify X";

            //adding a test step to the test case
            ITestStep testStep = testCase.CreateTestStep();
            testStep.Title = "Test step title";
            testStep.ExpectedResult = "Test step expected result";
            testCase.Actions.Add(testStep);

            //saving the test case to the project
            testCase.Save();

            //setting configs i guess?
            IdAndName defaultConfigidAndName = new IdAndName(defaultConfig.Id, defaultConfig.Name);
            suite.SetDefaultConfigurations(new IdAndName[] { defaultConfigidAndName });

            //adding test cases to the new suite
            suite.Entries.Add(testCase);

            //saving the test plan
            plan.Save();

            */
        }

    }
}
