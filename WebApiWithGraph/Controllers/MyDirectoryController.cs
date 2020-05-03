using System.Collections.Generic;
using System.Net;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Graph;
using WebApiWithGraph.Models;
using WebApiWithGraph.Services;

namespace WebApiWithGraph.Controllers
{
    [Produces("application/json")]
    [Route("directory")]
    public class MyDirectoryController : Controller
    {
        internal static class RouteNames
        {
            public const string Users = nameof(Users);
            public const string UserById = nameof(UserById);
            public const string Groups = nameof(Groups);
            public const string GroupById = nameof(GroupById);
        }

        [HttpGet("users/{id}", Name = RouteNames.UserById)]
        public async Task<IActionResult> GetUser(string id)
        {
            Models.User objUser = new Models.User();
            try
            {
                if (string.IsNullOrEmpty(id) || string.IsNullOrWhiteSpace(id))
                {
                    return BadRequest();
                }


                // Initialize the GraphServiceClient.
                GraphServiceClient client = await MicrosoftGraphClient.GetGraphServiceClient();

                // Load user profile.
                var user = await client.Users[id].Request().GetAsync();

                // Copy Microsoft-Graph User to DTO User
                objUser = CopyHandler.UserProperty(user);

                return Ok(objUser);
            }
            catch (ServiceException ex)
            {
                if (ex.StatusCode == HttpStatusCode.BadRequest)
                {
                    return BadRequest();
                }
                else
                {
                    return NotFound();
                }
            }
        }


        [HttpGet("users/")]
        public async Task<IActionResult> GetUsers()
        {
            Users users = new Users();
            try
            {
                users.resources = new List<Models.User>();

                // Initialize the GraphServiceClient.
                GraphServiceClient client = await MicrosoftGraphClient.GetGraphServiceClient();

                // Load users profiles.
                var userList = await client.Users.Request().GetAsync();

                // Copy Microsoft User to DTO User
                foreach (var user in userList)
                {
                    var objUser = CopyHandler.UserProperty(user);
                    users.resources.Add(objUser);
                }
                users.totalResults = users.resources.Count;

                return Ok(users);
            }
            catch (ServiceException ex)
            {
                if (ex.StatusCode == HttpStatusCode.BadRequest)
                {
                    return BadRequest();
                }
                else
                {
                    return NotFound();
                }
            }
        }


        [HttpGet("groups/{id}", Name = RouteNames.GroupById)]
        public async Task<IActionResult> GetGroup(string id)
        {
            Models.Group objGroup = new Models.Group();
            try
            {

                // Initialize the GraphServiceClient.
                GraphServiceClient client = await MicrosoftGraphClient.GetGraphServiceClient();

                // Load group profile.
                var group = await client.Groups[id].Request().GetAsync();

                // Copy Microsoft-Graph Group to DTO Group
                objGroup = CopyHandler.GroupProperty(group);

                return Ok(objGroup);
            }
            catch (ServiceException ex)
            {
                if (ex.StatusCode == HttpStatusCode.BadRequest)
                {
                    return BadRequest();
                }
                else
                {
                    return NotFound();
                }
            }
        }


        [HttpGet("groups/")]
        public async Task<IActionResult> GetGroups()
        {
            Groups groups = new Groups();
            try
            {
                groups.resources = new List<Models.Group>();

                // Initialize the GraphServiceClient.
                GraphServiceClient client = await MicrosoftGraphClient.GetGraphServiceClient();

                // Load groups profiles.
                var groupList = await client.Groups.Request().GetAsync();

                // Copy Microsoft-Graph Group to DTO Group
                foreach (var group in groupList)
                {
                    var objGroup = CopyHandler.GroupProperty(group);
                    groups.resources.Add(objGroup);
                }
                groups.totalResults = groups.resources.Count;

                return Ok(groups);
            }
            catch (ServiceException ex)
            {
                if (ex.StatusCode == HttpStatusCode.BadRequest)
                {
                    return BadRequest();
                }
                else
                {
                    return NotFound();
                }
            }
        }

    }
}