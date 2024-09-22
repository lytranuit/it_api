using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using Vue.Models;

namespace Info.Models
{
	[Table("a_salary")]
	public class SalaryModel
	{
		[Key]
		public string id { get; set; }
		public string name { get; set; }
		public string? description { get; set; }
		public string? file_url { get; set; }

		public int? nam { get; set; }
		public int? thang { get; set; }

		public DateTime? date_from { get; set; }
		public DateTime? date_to { get; set; }

		public string? created_by { get; set; }
		public string? status { get; set; }


		public DateTime? deleted_at { get; set; }
		public DateTime? updated_at { get; set; }
		public DateTime? created_at { get; set; }

		[NotMapped]
		public DateTime? date { get; set; }


	}
}
