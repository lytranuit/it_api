using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Reflection;
using static iText.StyledXmlParser.Jsoup.Select.Evaluator;

namespace Vue.Models
{
    [Table("Point")]
    public class PointModel
    {
        [Key]
        public int id { get; set; }
        public string? code { get; set; }
        public string? name { get; set; }
        public string? name_en { get; set; }
        public string? color { get; set; }
        public int? target_id { get; set; }
        public int? location_id { get; set; }
        public int? object_id { get; set; }
        public int? frequency_id { get; set; }
        public DateTime? deleted_at { get; set; }
        public DateTime? updated_at { get; set; }
        public DateTime? created_at { get; set; }

        [ForeignKey("object_id")]
        public virtual ObjectModel obj { get; set; }

        [ForeignKey("target_id")]
        public virtual TargetModel target { get; set; }
        [ForeignKey("location_id")]
        public virtual LocationModel location { get; set; }

        [NotMapped]
        public string frequency
        {
            get
            {
                var statusEnum = (Frequency)frequency_id;
                var memberInfo = typeof(Frequency).GetMember(statusEnum.ToString());
                var displayName = memberInfo.First().GetCustomAttribute<DisplayAttribute>().GetName();

                //enumValue.GetType()
                //        .GetMember(enumValue.ToString())
                //        .First()
                //        .GetCustomAttribute<DisplayAttribute>()
                //        .GetName();
                return displayName;

            }
        }
        [NotMapped]
        public string frequency_en
        {
            get
            {
                var statusEnum = (Frequency_en)frequency_id;
                var memberInfo = typeof(Frequency_en).GetMember(statusEnum.ToString());
                var displayName = memberInfo.First().GetCustomAttribute<DisplayAttribute>().GetName();

                //enumValue.GetType()
                //        .GetMember(enumValue.ToString())
                //        .First()
                //        .GetCustomAttribute<DisplayAttribute>()
                //        .GetName();
                return displayName;

            }
        }
    }
    enum Frequency
    {
        [Display(Name = "Hàng ngày")]
        Daily = 1,
        [Display(Name = "2 tuần / lần")]
        Two_Week = 2,
        [Display(Name = "Hàng tháng")]
        Monthly = 3,
        [Display(Name = "3 tháng / lần")]
        Three_Monthly = 4,
        [Display(Name = "6 tháng / lần")]
        Six_Monthly = 5,
        [Display(Name = "Hàng năm")]
        Yearly = 6
    }
    enum Frequency_en
    {
        [Display(Name = "Daily")]
        Daily = 1,
        [Display(Name = "2 tuần / lần")]
        Two_Week = 2,
        [Display(Name = "Monthly")]
        Monthly = 3,
        [Display(Name = "3 tháng / lần")]
        Three_Monthly = 4,
        [Display(Name = "6 tháng / lần")]
        Six_Monthly = 5,
        [Display(Name = "Yearly")]
        Yearly = 6
    }
}
