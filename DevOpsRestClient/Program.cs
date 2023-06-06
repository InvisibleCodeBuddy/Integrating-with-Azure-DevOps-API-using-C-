using Microsoft.TeamFoundation.WorkItemTracking.WebApi;
using Microsoft.TeamFoundation.WorkItemTracking.WebApi.Models;
using Microsoft.VisualStudio.Services.Common;
using Microsoft.VisualStudio.Services.WebApi;
using Microsoft.VisualStudio.Services.WebApi.Patch;
using Microsoft.VisualStudio.Services.WebApi.Patch.Json;
using System.Net;

namespace DevOpsRestClient
{
	internal class Program
	{
		static readonly string TFUrl = "https://dev.azure.com/<your team>";
		static readonly string UserAccount = "";
		static readonly string UserPassword = "";
        //You can get user PAT from here : https://dev.azure.com/<your team>/_usersSettings/tokens
        static readonly string UserPAT = "<your personal access token>";
		static readonly string ProjectName = "<your project name>";
		static readonly int MaxNumberWorkItems = 5;

		static WorkItemTrackingHttpClient WitClient;

		static async Task Main(string[] args)
		{
			try
			{
				//Initialize connection using User Personal Access Token
				await ConnectWithPAT(TFUrl, UserPAT);

				//Initialize connection using User Name Password
				//await ConnectWithCustomCreds(TFUrl, UserAccount,UserPassword );

				//Get all work items
				await DisplayAllWorkItems();

                //Create WorkItem and get generated ID
                int wiId = await CreateWorkItem();

				//Display common fields
				await DisplayWorkItem(wiId);

				Console.WriteLine("\r\nWorkItem created, press any key to continue..\r\n");
				Console.ReadLine();

				//Edit and Update the WorkItem
				wiId = await EditWorkItem(wiId);
				await DisplayWorkItem(wiId);

				Console.WriteLine("\r\nWorkItem updated, press any key to continue..\r\n");
				Console.ReadLine();

				//Delete the WorkItem by ID
				var result = await  DeleteWorkItem (wiId);
				Console.WriteLine($"\r\nThe WorkItem has been deleted on {result.DeletedDate}"); 





			}
			catch (Exception ex)
			{
				Console.WriteLine(ex.Message);
				//Console.WriteLine(ex.StackTrace);
			}
			Console.WriteLine("Press any key to exit..");
			Console.Read();
		}

        async static Task DisplayAllWorkItems()
        {
			// To query all items you want to use Wiql
			Wiql wiql = new Wiql()
			{
				Query = "Select [Title] from WorkItems Where [System.TeamProject] = '" + ProjectName + "'"
			};

			var wis = await WitClient.QueryByWiqlAsync(wiql);
            Console.WriteLine("__________________________________________");
            Console.WriteLine("                   Top " + MaxNumberWorkItems + " Work Items");
            Console.WriteLine("__________________________________________");

			var workItems = wis.WorkItems.ToList();

			for (int i = 0; i < MaxNumberWorkItems; i++)
			{
                var item = await GetWorkItem(workItems[i].Id);
                Console.WriteLine($"{item.Id}");
                if (item.Fields.ContainsKey("System.Title"))
                    Console.WriteLine($"{item.Fields["System.Title"]}");
                if (item.Fields.ContainsKey("System.Description"))
					Console.WriteLine($"{item.Fields["System.Description"]}");
                Console.WriteLine($"{item.Url}");
				Console.WriteLine("");
            }
        }

        async static Task DisplayWorkItem(int wiId)
		{
			var wi = await GetWorkItem(wiId);
			Console.WriteLine("__________________________________________");
			Console.WriteLine("                   Work Item");
			Console.WriteLine("__________________________________________");
			var wiTitle = CheckFieldAndGetFieldValue(wi, "System.Title"); // or just: var fieldValue = GetFieldValue(wi, "System.Title");

			var wiType = CheckFieldAndGetFieldValue(wi, "System.WorkItemType");

			var wiCDate = CheckFieldAndGetFieldValue(wi, "System.CreatedDate");

			var wiMDate = CheckFieldAndGetFieldValue(wi, "System.ChangedDate");

			Console.WriteLine($"Title:{wiTitle}\r\nType:{wiType}\r\nCreated On:{wiCDate}\r\nModified On:{wiMDate}");
		}

		async static Task<int> CreateWorkItem()
		{
			Dictionary<string, object> fields = new Dictionary<string, object>();
			//You can add more fields acording to your set up
			fields.Add("Title", "Task from app");
			fields.Add("Repro Steps", "<ol><li>Run app</li><li>Crash</li></ol>");
			fields.Add("Priority", 1);

			var newBug = await CreateWorkItem(ProjectName , "Task", fields);

			return newBug.Id.GetValueOrDefault();
		}
		async static Task<int> EditWorkItem(int WIId)
		{
			Dictionary<string, object> fields = new Dictionary<string, object>();

			fields.Add("Title", "Task from app updated");
			fields.Add("Repro Steps", "<ol><li>Run app</li><li>Crash</li><li>Updated step</li></ol>");
			fields.Add("History", "Comment from app");
			//You can assign the item to specific user like this
			//fields.Add("System.AssignedTo", "john.doe");

			var editedBug = await UpdateWorkItem(WIId, fields); 


			return editedBug.Id.GetValueOrDefault();
		}

		async static Task<WorkItem> UpdateWorkItem(int WIId, Dictionary<string, object> Fields)
		{
			JsonPatchDocument patchDocument = new JsonPatchDocument();

			foreach (var key in Fields.Keys)
				patchDocument.Add(new JsonPatchOperation()
				{
					Operation = Operation.Add,
					Path = "/fields/" + key,
					Value = Fields[key]
				});

			return await WitClient.UpdateWorkItemAsync(patchDocument, WIId);
		}
		async static Task<WorkItem> CreateWorkItem(string ProjectName, string WorkItemTypeName, Dictionary<string, object> Fields)
		{
			JsonPatchDocument patchDocument = new JsonPatchDocument();

			foreach (var key in Fields.Keys)
				patchDocument.Add(new JsonPatchOperation()
				{
					Operation = Operation.Add,
					Path = "/fields/" + key,
					Value = Fields[key]
				});

			return await WitClient.CreateWorkItemAsync(patchDocument, ProjectName, WorkItemTypeName);
		}

		async static Task<WorkItemDelete> DeleteWorkItem(int Id)
		{
			return await WitClient.DeleteWorkItemAsync(Id);
		}
		
		async static Task<WorkItem> GetWorkItem(int Id)
		{
			return await  WitClient.GetWorkItemAsync(Id, expand: WorkItemExpand.Relations);

			
		}
		static string GetFieldValue(WorkItem WI, string FieldName)
		{
			if (!WI.Fields.Keys.Contains(FieldName)) return string.Empty;

			if (WI.Fields[FieldName] != null)
				return WI.Fields[FieldName].ToString();
			else return string.Empty;
		}
		static string CheckFieldAndGetFieldValue(WorkItem WI, string FieldName)
		{
			WorkItemType wiType = GetWorkItemType(WI);

			var fields = from field in wiType.Fields where field.Name == FieldName || field.ReferenceName == FieldName select field;

			if (fields.Count() < 1) throw new ArgumentException("Work Item Type " + wiType.Name + " does not contain the field " + FieldName, "CheckFieldAndGetFieldValue");

			return GetFieldValue(WI, FieldName);
		}
		static WorkItemType GetWorkItemType(WorkItem WI)
		{
			if (!WI.Fields.Keys.Contains("System.WorkItemType")) throw new ArgumentException("There is no WorkItemType field in the workitem", "GetWorkItemType");
			if (!WI.Fields.Keys.Contains("System.TeamProject")) throw new ArgumentException("There is no TeamProject field in the workitem", "GetWorkItemType");

			return WitClient.GetWorkItemTypeAsync((string)WI.Fields["System.TeamProject"], (string)WI.Fields["System.WorkItemType"]).Result;
		}
				
		async static Task InitClient(VssConnection Connection)
		{
			WitClient = await Connection.GetClientAsync<WorkItemTrackingHttpClient>();
		
		}
		async static Task  ConnectWithCustomCreds(string ServiceURL, string User, string Password)
		{
			VssConnection connection = new VssConnection(new Uri(ServiceURL), new WindowsCredential(new NetworkCredential(User, Password)));
			await InitClient(connection);
		}

		async static Task ConnectWithPAT(string ServiceURL, string PAT)
		{
			VssConnection connection = new VssConnection(new Uri(ServiceURL), new VssBasicCredential(string.Empty, PAT));
			await InitClient(connection);
		}
	}
}