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

    public struct TestStep
    {

        public string nLot,
                      nDocument,
                      champ,
                      test,
                      expected,
                      result1,
                      result2;

                      
    }

    public struct TestCase
    {
        public List<TestStep> testSteps;
    }
   
    

    
    class Program
    {
        

        static ITestManagementTeamProject GetProject(string serverUrl, string project)
        {
            TfsTeamProjectCollection tfs = new TfsTeamProjectCollection(TfsTeamProjectCollection.GetFullyQualifiedUriForName(serverUrl));
            ITestManagementService tms = tfs.GetService<ITestManagementService>();

            return tms.GetTeamProject(project);
        }

        static IStaticTestSuite Traversal(ITestSuiteCollection suitesCollection, string desiredSuiteTitle)
        {
            foreach (IStaticTestSuite childSuite in suitesCollection)
            {
                if (childSuite.Title == desiredSuiteTitle)
                    return childSuite;
            }

            return null;
        }


        static void Main(string[] args)
        {

            Application word = new Application();
            Document doc = new Document();

            // Define an object to pass to the API for missing parameters
            object missing = System.Type.Missing;
            doc = word.Documents.Open(@"D:\RQemploi\DA0_Devis_RAC_PAL19.docx",
                    ref missing, ref missing, ref missing, ref missing,
                    ref missing, ref missing, ref missing, ref missing,
                    ref missing, ref missing, ref missing, ref missing,
                    ref missing, ref missing, ref missing);

            List<TestStep> data = new List<TestStep>();

            foreach(Table table in doc.Tables)
            {
                for(int i = 1; i <= table.Rows.Count; i++)
                {
                    bool isHeader = false;
                    TestStep temp = new TestStep();
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

            foreach(TestStep ts in data)
            {
                Console.WriteLine(ts.nLot);
                Console.WriteLine(ts.nDocument);
                Console.WriteLine(ts.champ);
                Console.WriteLine(ts.test);
                Console.WriteLine(ts.expected);
                Console.WriteLine(ts.result1);
                Console.WriteLine(ts.result2);
            }
            

            Console.ReadLine();

            ((_Document)doc).Close();
            ((_Application)word).Quit();


            //MTM

            string prodName = "TP01";
            string prodYear = "2018-11";
            string testPlanName = $"Liv 2019 - DAO UT {prodName} ({prodYear})";
            int idTestPlan = 0;
            string serverurl = "http://gestsource.services.mrq:8080/tfs";
            string project = @"gestsource.services.mrq\RQ\R4-CAB2D-RAC";
            ITestManagementTeamProject proj = GetProject(serverurl, project);

            // Get a Test Plan by its Id
            int myPlansId = idTestPlan;
            ITestPlan plan = proj.TestPlans.Find(myPlansId);


            IStaticTestSuite foundSuite = Traversal(plan.RootSuite.SubSuites, testPlanName);
            foundSuite = Traversal(foundSuite.SubSuites, "Changement annuel et non régression");
            foundSuite = Traversal(foundSuite.SubSuites, "B- AD'DOC");
            foundSuite = Traversal(foundSuite.SubSuites, "5- Rejets d'OCR");



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

            
        }

    }
}
