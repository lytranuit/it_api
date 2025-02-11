using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace Info.Models
{

    [Table("a_meeting")]
    public class MeetingModel
    {
        [Key]
        public int id { get; set; }

        public string name { get; set; }
        public string? ghichu { get; set; }
        public string? list_notify { get; set; }

        [NotMapped]
        public List<string> list_notify_id
        {
            get
            {
                return list_notify != null ? list_notify.Split(",").ToList() : new List<string>();
            }
            set
            {
                list_notify = value != null ? string.Join(",", value) : null;
            }
        }

        public string? phong_hop { get; set; }
        public int? phong_hop_id { get; set; }

        public DateTime? date_from { get; set; }
        public DateTime? date_to { get; set; }


        public DateTime? created_at { get; set; }
        public DateTime? deleted_at { get; set; }
        public string? created_by { get; set; }
    }
}