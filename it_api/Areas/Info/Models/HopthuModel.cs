using it_api.Extensions;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using Vue.Models;

namespace Info.Models
{
    [Table("a_hopthu")]
    public class HopthuModel
    {
        [Key]
        public string id { get; set; }
        public string content { get; set; }
        public LoaiPhanHoi loai { get; set; }
        [NotMapped]
        public string loai_label => loai.GetDescription();

        public TinhtrangHopthu tinhtrang { get; set; }
        [NotMapped]
        public string tinhtrang_label => tinhtrang.GetDescription();

        public bool? is_andanh { get; set; }
        public string? created_by { get; set; }


        [ForeignKey("created_by")]
        public virtual UserModel? user_created_by { get; set; }

        public DateTime? deleted_at { get; set; }
        public DateTime? created_at { get; set; }


        public string? dinhkem { get; set; }

        [NotMapped]
        public virtual List<string>? list_dinhkem
        {
            get
            {
                return string.IsNullOrEmpty(dinhkem) ? new List<string>() : dinhkem.Split(",").ToList();
            }
            set
            {
                dinhkem = string.Join(",", value);
            }
        }
    }
    public enum LoaiPhanHoi
    {
        [Description("Phản ánh")]
        PhanAnh = 1,

        [Description("Ý kiến / Góp ý")]
        GopY = 2,

        [Description("Sáng kiến / Cải tiến")]
        SangKien = 3,
    }

    public enum TinhtrangHopthu
    {
        [Description("Mới")]
        New = 1,

        [Description("Đang xử lý")]
        Pedding = 2,

        [Description("Đã xử lý")]
        Success = 3,
    }
}
