using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace Vue.Models
{
    public class DIEUTRASUCO
    {
        [Key]
        public int id { get; set; }
        public string so { get; set; }
        public string? bienphap { get; set; }
        public bool? laplai { get; set; }


        public string? sothamkhao { get; set; }
        public string? mucdo_mota { get; set; }
        public int? mucdo_diem { get; set; }
        public string? tansuat_mota { get; set; }
        public int? tansuat_diem { get; set; }


        public bool? capa { get; set; }

        public bool? suco { get; set; }
        public string? pheduyet_danhgia { get; set; }
        public string? pheduyet_quyetdinh { get; set; }
        public bool? baocaocaptren { get; set; }
        public bool? baocaonsx { get; set; }


    }
}
