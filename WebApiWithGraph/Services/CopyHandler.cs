using WebApiWithGraph.Models;

namespace WebApiWithGraph.Services
{
    public class CopyHandler
    {
        public static User UserProperty(Microsoft.Graph.User graphUser)
        {
            User user = new User();
            user.id = graphUser.Id;
            user.givenName = graphUser.GivenName;
            user.surname = graphUser.Surname;
            user.userPrincipalName = graphUser.UserPrincipalName;
            user.email = graphUser.Mail;

            return user;
        }

        public static Group GroupProperty(Microsoft.Graph.Group graphGroup)
        {
            Group group = new Group();
            group.id = graphGroup.Id;
            group.displayName = graphGroup.DisplayName;

            return group;
        }
    }
}
