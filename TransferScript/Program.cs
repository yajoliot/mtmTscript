using System;
using System.Windows;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

//word API
using Microsoft.Office.Interop.Word;

//MTM API
using Microsoft.TeamFoundation;
using Microsoft.TeamFoundation.Client;
using Microsoft.TeamFoundation.TestManagement.Client;
using Microsoft.TeamFoundation.WorkItemTracking;

//TFS API
using Microsoft.TeamFoundationServer.Client;
using Microsoft.TeamFoundationServer.TestManagement.WebApi;
using Microsoft.VisualStudio.Services.Common;
using Microsoft.VisualStudio.Services.Client;
using Microsoft.TeamFoundation.SourceControl.WebApi;
using Microsoft.VisualStudio.Services.WebApi;



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

            //SERVERACCESS

            string prodName = "TP01";
            string prodYear = "2018-11";
            string testPlanName = $"Liv 2019 - DAO UT {prodName} ({prodYear})";
            int myPlansId = 0;
            string serverUrl = "http://gestsource.services.mrq:8080/tfs";
            string project = @"R4-CAB2D-RAC / équipe Caribou"; // é doit etre majuscule

            string personalAccessToken = "";
            var base64Token = Convert.ToBase64String(Encoding.ASCII.GetBytes($":{personalAccessToken}"));

            using (HttpClient client = new HttpClient())
            {
                client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Basic", base64Token);

                var requestMessage = new HttpRequestMessage(HttpMthod.Post, $"https://{serverUrl}/DefaultCollection/{projetc}/_testManagement");
                requestMessage.Content = new StringContent { /* test Title etc...*/, Encoding.UTF8, "application/json" };
                using (HttpResponseMessage response = client.sendAsync(requestMessage).Result)
                {
                    response.EnsureSuccessStatusCode();
                }
            }


            //MTM

            
        }

    }
}
