using NuGet.Configuration;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using workflow.Models;
namespace Vue.Models
{
    [Table("department")]
    public class DepartmentModel
    {
        [Key]
        public int id { get; set; }

        public string? color { get; set; }
        public string name { get; set; }
        public int? parent { get; set; } = 0;
        public int? count_child { get; set; } = 0;
        public int? stt { get; set; } = 0;
        public DateTime? created_at { get; set; }

        public DateTime? updated_at { get; set; }

        public DateTime? deleted_at { get; set; }

        public virtual List<UserDepartmentModel>? list_users { get; set; }
        [NotMapped]
        public List<DepartmentModel>? children { get; set; }
        [NotMapped]
        public List<string> list_users_id
        {
            get
            {
                return list_users != null ? list_users.Select(x => x.user_id).ToList() : new List<string>();
            }
        }
    }
}
