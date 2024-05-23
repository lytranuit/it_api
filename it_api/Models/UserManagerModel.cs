using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace Vue.Models
{
    [Table("user_manager")]
    public class UserManagerModel
    {
        [Key]
        public int id { get; set; }
        public string userId { get; set; }
        public string userManagerId { get; set; }

        [NotMapped]
        [ForeignKey("userId")]
        public UserModel user { get; set; }

        [NotMapped]
        [ForeignKey("userManagerId")]
        public UserModel userManager { get; set; }

    }
}
