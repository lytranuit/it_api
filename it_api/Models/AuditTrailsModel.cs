using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
namespace Vue.Models
{
    [Table("AuditTrails")]
    public class AuditTrailsModel
	{
		[Key]
		public int Id { get; set; }
        public string? UserId { get; set; }

        [ForeignKey("UserId")]
        public virtual UserModel? user { get; set; }
        public string? Type { get; set; }
        public string? TableName { get; set; }
        public string? description { get; set; }
        public DateTime DateTime { get; set; }
        public string? OldValues { get; set; }
        public string? NewValues { get; set; }
        public string? AffectedColumns { get; set; }
        public string? PrimaryKey { get; set; }
    }
}
