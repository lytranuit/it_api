
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using Microsoft.AspNetCore.Identity;
using Vue.Data;
using Vue.Models;
using System.Collections;
using workflow.Models;
using System.Text.Json.Serialization;
using System.Data;

namespace it_template.Areas.Trend.Controllers
{
    public class AdminController : BaseController
    {
        private UserManager<UserModel> UserManager;
        public AdminController(ItContext context, UserManager<UserModel> UserMgr) : base(context)
        {
            UserManager = UserMgr;
        }

        public async Task<JsonResult> chart(List<int> list_point, DateTime? tungay, DateTime? denngay)
        {
            var cusdata = _context.ResultModel.Where(d => d.deleted_at == null && list_point.Contains(d.point_id.Value));
            if (tungay != null && tungay.HasValue)
            {
                if (tungay.Value.Kind == DateTimeKind.Utc)
                {
                    tungay = tungay.Value.ToLocalTime();
                }
                cusdata = cusdata.Where(d => d.date >= tungay.Value);
            }
            if (denngay != null && denngay.HasValue)
            {
                if (denngay.Value.Kind == DateTimeKind.Utc)
                {
                    denngay = denngay.Value.ToLocalTime();
                }
                cusdata = cusdata.Where(d => d.date <= denngay.Value);
            }
            var results = cusdata.Include(d => d.target).Include(d => d.point).ThenInclude(d => d.location).Where(h => h.target.value_type == "float").ToList();

            var groups = results
                        //.GroupBy(d => d.point.location).Select(d => new Label()
                        //{
                        //    key = d.Key.id,
                        //    label = d.Key.name,
                        //    type = "location",
                        //    timestamp = new DateTimeOffset(DateTime.UtcNow).ToUnixTimeSeconds(),
                        //    children = d.GroupBy(e => new { e.point.frequency_id, e.point.frequency }).Select(e => new Label()
                        //    {
                        //        key = e.Key.frequency_id.Value,
                        //        label = e.Key.frequency,
                        //        type = "frequency",
                        //        timestamp = new DateTimeOffset(DateTime.UtcNow).ToUnixTimeSeconds(),
                        //        children = e
                        .GroupBy(f => f.target).Select(f => new Label()
                        {
                            key = f.Key.id,
                            label = f.Key.name,
                            timestamp = new DateTimeOffset(DateTime.UtcNow).ToUnixTimeSeconds(),
                            data = f.ToList(),
                        }).ToList();
            //    }).ToList()
            //}).ToList();
            var list_pointStyle = new List<string>()
            {
                "circle",
                "cross",
                "crossRot",
                "dash",
                "line",
                "rect",
                "rectRounded",
                "rectRot",
                "star",
                "triangle"
            };
            //foreach (var location in groups)
            //{
            //    foreach (var frequency in location.children)
            //    {
            foreach (var target in groups)
            {
                //target.chart = target.data.Select(d => d.point.code).ToList();
                //var limit_all = _context.LimitModel.Where(d => d.target_id == target.key && d.date_effect <= denngay && d.deleted_at == null).Include(d => d.points).ToList();
                var data = target.data;
                var list_point1 = data.Select(d => d.point_id.Value).Distinct().ToList();
                //limit_all = limit_all.Where(d => d.list_point.Intersect(list_point1).Any()).ToList();
                var labels = target.data.GroupBy(g => g.date).Select(g => g.Key.Value).OrderBy(g => g).ToList();


                var datasets = target.data.GroupBy(g => g.point).Select(g => new Dataset()
                {
                    label = g.Key.code,
                    type = "line",
                    borderWidth = 1,
                    spanGaps = true,
                    borderColor = g.Key.color,
                    backgroundColor = g.Key.color,
                    results = g.ToList(),
                    data = new List<decimal?>()
                }).ToList();
                var annotations = new Dictionary<string, Annotations>();
                var name_line_ghcb = "";
                var name_line_ghhd = "";
                var stt_action = 0;
                var stt_alert = 0;
                decimal? limit_action_prev = 0;
                decimal? limit_alert_prev = 0;
                var row = 0;
                var max = data.Select(g => g.value).Max();
                var suggestedMax = max + (max * 20 / 100);
                foreach (var label in labels)
                {

                    var stt = 0;
                    foreach (var d in datasets)
                    {
                        d.pointStyle = list_pointStyle[stt];
                        var finddata = d.results.Where(d => d.date == label).FirstOrDefault();
                        //var point_id = finddata.point_id;

                        var value = finddata != null ? finddata.value : null;
                        d.data.Add(value);
                        //var date = d.date.Value.ToString("yyyy-MM-dd");
                        //var point_code = d.point.code;
                        stt = list_pointStyle.Count - 1 > stt ? stt + 1 : 0;
                    }
                    var data_date = data.Where(d => d.date == label).ToList();
                    var limit_action = data_date.Select(g => g.limit_action).Min();
                    if (limit_action != null)
                    {
                        if (limit_action == limit_action_prev)
                        {
                            annotations[name_line_ghhd].xMax++;
                        }
                        else
                        {
                            limit_action_prev = limit_action;
                            name_line_ghhd = "line_ghhd_" + stt_action;
                            annotations.Add(name_line_ghhd, new Annotations()
                            {
                                xMin = row,
                                xMax = row + 1,
                                yMin = limit_action.Value,
                                yMax = limit_action.Value,
                                type = "line",
                                borderWidth = 3,
                                borderColor = "red"
                            });
                            annotations.Add("callout_ghhd_" + stt_action, new Annotations()
                            {
                                type = "label",
                                xValue = row,
                                yValue = limit_action.Value,
                                xAdjust = 50,
                                yAdjust = -20,
                                content = new List<string>() { "Action Limit", limit_action.Value.ToString("#,##0.##") },
                                callout = new Callout()
                                {
                                    display = true,
                                    side = 5
                                }
                            });

                            ///
                            if (limit_action + (limit_action * 20 / 100) > suggestedMax)
                            {
                                suggestedMax = limit_action + (limit_action * 20 / 100);
                            }
                            stt_action++;
                        }
                    }
                    var limit_alert = data.Where(d => d.date == label).Select(g => g.limit_alert).Min();

                    if (limit_alert != null)
                    {
                        if (limit_alert == limit_alert_prev)
                        {
                            annotations[name_line_ghcb].xMax++;
                        }
                        else
                        {
                            limit_alert_prev = limit_alert;
                            name_line_ghcb = "line_ghcb_" + stt_alert;
                            annotations.Add(name_line_ghcb, new Annotations()
                            {
                                xMin = row,
                                xMax = row + 1,
                                yMin = limit_alert.Value,
                                yMax = limit_alert.Value,
                                type = "line",
                                borderWidth = 3,
                                borderColor = "orange"
                            });

                            annotations.Add("callout_ghcb_" + stt_alert, new Annotations()
                            {
                                type = "label",
                                xValue = row,
                                yValue = limit_alert.Value,
                                xAdjust = 50,
                                yAdjust = -20,
                                content = new List<string>() { "Alert Limit", limit_alert.Value.ToString("#,##0.##") },
                                callout = new Callout()
                                {
                                    display = true,
                                    side = 5
                                }
                            });
                            ///
                            if (limit_action + (limit_action * 20 / 100) > suggestedMax)
                            {
                                suggestedMax = limit_action + (limit_action * 20 / 100);
                            }
                            stt_alert++;
                        }
                    }

                    row++;
                }
                //if (ghcb.data.Count > 0)
                //{
                //    datasets.Add(ghcb);
                //    datasets.Add(ghhd);
                //}
                target.chart = new Chart1()
                {
                    labels = labels.Select(d => d.ToString("yyyy-MM-dd")).ToList(),
                    datasets = datasets,
                    yTitle = data.FirstOrDefault().target.unit,
                    suggestedMax = suggestedMax,
                    annotations = annotations
                };
                //    }
                //}
            }
            return Json(new { groups }, new System.Text.Json.JsonSerializerOptions()
            {
                DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull
            });
        }
        public async Task<JsonResult> alerts(List<int> list_point, DateTime? tungay, DateTime? denngay)
        {
            var cusdata = _context.ResultModel.Where(d => d.deleted_at == null);
            if (tungay != null && tungay.HasValue)
            {
                if (tungay.Value.Kind == DateTimeKind.Utc)
                {
                    tungay = tungay.Value.ToLocalTime();
                }
                cusdata = cusdata.Where(d => d.date >= tungay.Value);
            }
            if (denngay != null && denngay.HasValue)
            {
                if (denngay.Value.Kind == DateTimeKind.Utc)
                {
                    denngay = denngay.Value.ToLocalTime();
                }
                cusdata = cusdata.Where(d => d.date <= denngay.Value);
            }
            if (list_point != null && list_point.Count() > 0)
            {
                cusdata = cusdata.Where(d => list_point.Contains(d.point_id.Value));
            }
            var results = cusdata
                .Include(d => d.point)
                .Include(d => d.target).Where(d => d.limit_id != null && d.value > d.limit_alert).OrderByDescending(d => d.date).ThenByDescending(d => d.id).ToList();

            return Json(results);
        }
        public async Task<JsonResult> actions(List<int> list_point, DateTime? tungay, DateTime? denngay)
        {
            var cusdata = _context.ResultModel.Where(d => d.deleted_at == null);
            if (tungay != null && tungay.HasValue)
            {
                if (tungay.Value.Kind == DateTimeKind.Utc)
                {
                    tungay = tungay.Value.ToLocalTime();
                }
                cusdata = cusdata.Where(d => d.date >= tungay.Value);
            }
            if (denngay != null && denngay.HasValue)
            {
                if (denngay.Value.Kind == DateTimeKind.Utc)
                {
                    denngay = denngay.Value.ToLocalTime();
                }
                cusdata = cusdata.Where(d => d.date <= denngay.Value);
            }
            if (list_point != null && list_point.Count() > 0)
            {
                cusdata = cusdata.Where(d => list_point.Contains(d.point_id.Value));
            }
            var results = cusdata
                .Include(d => d.point)
                .Include(d => d.target).Where(d => d.limit_id != null && d.value > d.limit_action).OrderByDescending(d => d.date).ThenByDescending(d => d.id).ToList();
            return Json(results);
        }

    }

    public class Annotations
    {
        public string type { get; set; } = "line";

        public decimal? xValue { get; set; }
        public decimal? yValue { get; set; }
        public decimal? xAdjust { get; set; }
        public decimal? yAdjust { get; set; }
        public decimal? xMin { get; set; }
        public decimal? yMin { get; set; }
        public decimal? xMax { get; set; }
        public decimal? yMax { get; set; }
        public string? borderColor { get; set; }
        public int? borderWidth { get; set; }
        public List<string>? content { get; set; }
        public Font font { get; set; }
        public Callout? callout { get; set; }

    }
    public class Font
    {
        public int? size { get; set; }
        public string? style { get; set; }
        public string? weight { get; set; }
    }
    public class Callout
    {
        public int side { get; set; }
        public bool display { get; set; }
    }
    public class Dataset
    {
        public string type { get; set; }
        public string label { get; set; }
        public List<decimal?> data { get; set; }
        public bool fill { get; set; }
        public int borderWidth { get; set; }
        public string borderColor { get; set; }
        public string backgroundColor { get; set; }
        public bool spanGaps { get; set; } = false;
        public string pointStyle { get; set; }
        public List<ResultModel>? results { get; set; }
    }
    public class Chart1
    {
        public List<string> labels { get; set; }
        public List<Dataset> datasets { get; set; }

        public decimal? suggestedMax { get; set; }
        public decimal? suggestedMin { get; set; }
        public string? yTitle { get; set; }
        public Dictionary<string, Annotations> annotations { get; set; }
    }

    public class Label
    {
        public int? key { set; get; }
        public string label { set; get; }
        public string type { set; get; }
        public long timestamp { get; set; }
        public List<Label> children { get; set; }

        public Chart1? chart { get; set; }
        public List<ResultModel>? data { get; set; }



    }
}
