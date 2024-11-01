using System.Collections.Generic;

namespace Test.Models
{
    public class ManageUserRolesViewModel
    {
        public string UserId { get; set; }
        public string UserName { get; set; }
        public List<UserRoleViewModel> Roles { get; set; }
    }

    public class UserRoleViewModel
    {
        public string RoleName { get; set; }
        public bool IsSelected { get; set; }
    }
}