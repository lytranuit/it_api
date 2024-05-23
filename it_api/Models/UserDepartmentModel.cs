using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using Vue.Models;

namespace workflow.Models
{
	[Table("user_department")]
	public class UserDepartmentModel
	{
		[Key]
		public int id { get; set; }

		public string user_id { get; set; }
		public int department_id { get; set; }


		[ForeignKey("user_id")]
		public UserModel user { get; set; }

		[ForeignKey("department_id")]
		public DepartmentModel department { get; set; }
	}
}
