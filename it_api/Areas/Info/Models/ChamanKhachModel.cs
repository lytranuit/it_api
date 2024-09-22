using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using Vue.Models;

namespace Info.Models
{
	[Table("a_chaman_khach")]
	public class ChamanKhachModel
	{
		[Key]
		public string id { get; set; }
		public string title { get; set; }
		public int? soluong { get; set; }

		public DateTime? date { get; set; }
		public string created_by { get; set; }
		public DateTime? created_at { get; set; }
		public string calendar_id { get; set; }

		[NotMapped]
		public bool? ignore { get; set; }
	}
}
