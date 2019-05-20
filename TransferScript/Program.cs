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


            //MTM

            string prodName = "TP01";
            string prodYear = "2018-11";
            string testPlanName = $"Liv 2019 - DAO UT {prodName} ({prodYear})";
            int myPlansId = 0;
            string serverurl = "http://gestsource.services.mrq:8080/tfs";
            string project = @"gestsource.services.mrq\RQ\R4-CAB2D-RAC";
            ITestManagementTeamProject proj = GetProject(serverurl, project);

            // Get a Test Plan by its Id
            ITestPlan plan = proj.TestPlans.Find(myPlansId);


            IStaticTestSuite foundSuite = Traversal(plan.RootSuite.SubSuites, testPlanName);
            foundSuite = Traversal(foundSuite.SubSuites, "Changement annuel et non régression");
            foundSuite = Traversal(foundSuite.SubSuites, "B- AD'DOC");
            foundSuite = Traversal(foundSuite.SubSuites, "5- Rejets d'OCR");



            //creating the test case
            ITestCase testCase = proj.TestCases.Create();
            testCase.Title = "TEST";

            //adding a test step to the test case
            ITestStep testStep = testCase.CreateTestStep();
            testStep.Title = "Test step title";
            testStep.ExpectedResult = "Test step expected result";
            testCase.Actions.Add(testStep);

            //saving the test case to the project
            testCase.Save();

            //setting configs i guess?
            ITestConfiguration defaultConfig = null;
            IdAndName defaultConfigidAndName = new IdAndName(defaultConfig.Id, defaultConfig.Name);
            foundSuite.SetDefaultConfigurations(new IdAndName[] { defaultConfigidAndName });

            //adding test case to the new suite
            foundSuite.Entries.Add(testCase);

            //saving the test plan
            plan.Save();

            
        }

    }
}
