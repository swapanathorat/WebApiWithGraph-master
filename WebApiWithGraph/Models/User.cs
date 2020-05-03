using System.Collections.Generic;

namespace WebApiWithGraph.Models
{

    public class User
    {
        public string id { get; set; }
        public string givenName { get; set; }
        public string surname { get; set; }
        public string userPrincipalName { get; set; }
        public string email { get; set; }
    }

    public class Users
    {
        public int itemsPerPage { get; set; }
        public int startIndex { get; set; }
        public int totalResults { get; set; }
        public List<User> resources { get; set; }
    }
}
