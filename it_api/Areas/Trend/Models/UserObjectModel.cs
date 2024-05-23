using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace Vue.Models
{
    [Table("user_object")]
    public class UserObjectModel
    {
        [Key]
        public int id { get; set; }
        public string user_id { get; set; }
        public int object_id { get; set; }

        [NotMapped]
        [ForeignKey("user_id")]
        public UserModel user { get; set; }

        [NotMapped]
        [ForeignKey("object_id")]
        public ObjectModel obj { get; set; }

    }
}
