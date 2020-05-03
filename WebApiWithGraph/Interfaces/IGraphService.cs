using Microsoft.Extensions.Logging;
using System.Threading.Tasks;
using WebApiWithGraph.Models;
using System.Collections.Generic;

namespace WebApiWithGraph.Interfaces
{
    public interface IGraphService
    {
        Task<User> GetUser(string id);
        Task<Users> GetUsers(string filter, int? startIndex, int? count, string sortBy);
        Task<Group> GetGroup(string id);
        Task<Groups> GetGroups(string filter, int? startIndex, int? count, string sortBy);
    }

    
}
