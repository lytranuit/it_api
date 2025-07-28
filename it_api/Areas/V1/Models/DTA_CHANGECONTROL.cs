using System;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace Vue.Areas.V1.Models
{
    [Table("DTA_CHANGECONTROL")]
    public class DTA_CHANGECONTROL
    {
        public string sochange { get; set; } = null!;

        public DateTime ngaydenghi { get; set; }

        public string? mabp { get; set; }

        public string? tenbp { get; set; }

        public string? doituong { get; set; }

        public string? doituong_2 { get; set; }

        public string? tinhtrang { get; set; }

        public string? mota { get; set; }

        public string? mota_2 { get; set; }

        public string? a_facility { get; set; }

        public string? b_equipment { get; set; }

        public string? c_utilities { get; set; }

        public string? d_computerized { get; set; }

        public string? e_raw { get; set; }

        public string? f_packaging { get; set; }

        public string? g_consumables { get; set; }

        public string? h_product { get; set; }

        public string? i_batch { get; set; }

        public string? j_production { get; set; }

        public string? k_specifications { get; set; }

        public string? l_other { get; set; }

        public string? other_detail { get; set; }

        public string? lydo { get; set; }

        public string? lydo_2 { get; set; }

        public DateTime? dukien { get; set; }

        public string? thaydoitamthoi { get; set; }

        public string? ngungsanxuat { get; set; }

        public string? lygiai { get; set; }

        public string? thucthi { get; set; }

        public string? chapthuan_qa { get; set; }

        public string? chapthuan_partner { get; set; }

        public string? nguoidung { get; set; }

        public string? nguoinhan { get; set; }

        [Key]

        public int id { get; set; }
    }
}
