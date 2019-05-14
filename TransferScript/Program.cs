using System;
using System.Windows;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

//word API
using Word = Microsoft.Office.Interop.Word;

//MTM/TFS API
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

    struct TestCase
    {
        public int nLot,
                   nDocument;

        public string champ,
                      test,
                      expected;

        public bool result;
                      
    }
    class Program
    {


        static ITestManagementTeamProject GetProject(string serverUrl, string project)
        {
            TfsTeamProjectCollection tfs = new TfsTeamProjectCollection(TfsTeamProjectCollection.GetFullyQualifiedUriForName(serverUrl));
            ITestManagementService tms = tfs.GetService<ITestManagementService>();

            return tms.GetTeamProject(project);
        }


        static void Main(string[] args)
        {

            Word.Application word = new Word.Application();
            Word.Document doc = new Word.Document();

            // Define an object to pass to the API for missing parameters
            object missing = System.Type.Missing;
            doc = word.Documents.Open(@"D:\RQemploi\TestDoc.docx",
                    ref missing, ref missing, ref missing, ref missing,
                    ref missing, ref missing, ref missing, ref missing,
                    ref missing, ref missing, ref missing, ref missing,
                    ref missing, ref missing, ref missing);

            List<string> data = new List<string>();
            for (int i = 0; i < doc.Paragraphs.Count; i++)
            {
                string temp = doc.Paragraphs[i + 1].Range.Text.Trim();
                if (temp != string.Empty)
                    data.Add(temp);
            }

            for (int i = 0; i < doc.Paragraphs.Count; i++)
            {
                Console.WriteLine(data[i]);
            }

            Console.ReadLine();

            ((Word._Document)doc).Close();
            ((Word._Application)word).Quit();


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

            //adding the test case to the suite
            ITestCase testCase = proj.TestCases.Create();
            testCase.Title = "Verify X";
            testCase.Save();

            IdAndName defaultConfigIdAndName = new IdAndName(defaultConfig.Id, defaultConfig.Name);

            suite.Entries.Add(testCase);
            plan.Save();


        }

    }
}
