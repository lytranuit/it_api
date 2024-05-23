using it_template.Areas.Trend.Controllers;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Diagnostics;
using System.Text.Json;
using System.Text.Json.Serialization;

namespace Vue.Models
{
    [Table("chart_data")]
    public class ChartDataModel
    {
        [Key]
        public int id { get; set; }
        public string? key { get; set; }
        public string? data { get; set; }
        public long timestamp { get; set; }

        [NotMapped]
        public virtual Chart1 chart
        {
            get
            {
                return JsonSerializer.Deserialize<Chart1>(string.IsNullOrEmpty(data) ? "{}" : data);
            }
            set
            {
                data = JsonSerializer.Serialize(value, new JsonSerializerOptions()
                {
                    ReferenceHandler = ReferenceHandler.IgnoreCycles
                });
            }
        }
    }
}
