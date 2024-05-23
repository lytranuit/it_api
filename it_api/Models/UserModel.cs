using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

using Microsoft.AspNetCore.Identity;

namespace Vue.Models
{

    [Table("AspNetUsers")]
    public class UserModel : IdentityUser
    {
        public string FullName { get; set; }
        public string? image_url { get; set; }
        public string? image_sign { get; set; }
        public string? signature { get; set; }
        public DateTime? last_login { get; set; }

        public DateTime? created_at { get; set; }

        public DateTime? updated_at { get; set; }

        public DateTime? deleted_at { get; set; }
        [NotMapped]
        public List<UserManagerModel> list_users { get; set; }
        [NotMapped]
        public UserManagerModel userreport { get; set; }


    }
}
