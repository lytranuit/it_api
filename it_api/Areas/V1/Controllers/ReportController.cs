

using iText.Commons.Bouncycastle.Asn1.X509;
using Microsoft.AspNetCore.Identity;
using Microsoft.AspNetCore.Mvc;
using Microsoft.CodeAnalysis;
using Microsoft.EntityFrameworkCore;
using Microsoft.EntityFrameworkCore.Metadata.Internal;
using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Doc.Fields;
using System.Data;
using System.Reflection;
using Vue.Data;
using Vue.Models;
using static iText.Svg.SvgConstants;

namespace it_template.Areas.V1.Controllers
{

    [Area("V1")]
    public class ReportController : Controller
    {
        private UserManager<UserModel> UserManager;
        private readonly IConfiguration _configuration;
        private QLSXContext _qlsxContext;
        public ReportController(QLSXContext QLSXContext, IConfiguration configuration, UserManager<UserModel> UserMgr) : base()
        {
            UserManager = UserMgr;
            _configuration = configuration;
            _qlsxContext = QLSXContext;
        }
        public async Task<JsonResult> lenhxuatxuong(string solenh)
        {
            if (solenh == null)
            {
                return Json(new { success = false });
            }

            var viewPath = _configuration["Source:Path_Private"] + "\\qa\\templates\\mau lenhxuatxuong.docx";
            var documentPath = "/temp/lenhxuatxuong_" + DateTime.Now.ToFileTimeUtc() + ".docx";
            string Domain = (HttpContext.Request.IsHttps ? "https://" : "http://") + HttpContext.Request.Host.Value;


            var report = "";


            ///XUẤT PDF

            //Creates Document instance
            Spire.Doc.Document document = new Spire.Doc.Document();
            Section section = document.AddSection();

            document.LoadFromFile(viewPath, Spire.Doc.FileFormat.Docx);



            //document.MailMerge.ExecuteWidthRegion(datatable_details);

            ////Lấy 
            ///

            Dictionary<string, Spire.Doc.Table> replacements_table = new Dictionary<string, Spire.Doc.Table>(StringComparer.OrdinalIgnoreCase) { };

            var sqla = $"EXECUTE [IN_LENHXUATXUONG_2] '{solenh}' ";

            var list1 = _qlsxContext.LenhXuatXuongModel.FromSqlRaw($"{sqla}").ToList();


            ///FIELD OTHER
            var first = list1.FirstOrDefault();
            var raw = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
                {
                    {"solenh",first.solenh },
                    {"mahh",first.mahh },
                    {"tenhh",first.tenhh },
                    {"mahh_goc",first.mahh_goc },
                    {"tenhh_goc",first.tenhh_goc },
                    {"tenhoatchat",first.tenhoatchat },
                    {"dangbaoche",first.dangbaoche },
                    {"ngaysx",first.ngaysx.Value.ToString("dd/MM/yy") },
                    {"quicachdonggoi",first.quicachdonggoi },
                    {"ngaylenh",first.ngaylenh.ToString("dd/MM/yyyy") },
                    {"malo_goc",first.malo_goc },
                    {"mapl",first.mapl },
                    {"tenpl",first.tenpl },
                    {"malo",first.malo },
                    {"colo",first.colo.Replace(",",".") },
                    {"sodk",first.sodk },
                    {"handung",first.handung },
                    {"tieuchuan",first.tieuchuan },
                    {"soluongdonggoi",first.soluongdonggoi.Value.ToString("#,##0").Replace(",",".") + " " +first.dvt },
                    {"coa",first.coa },
                    {"thung1",first.thung1.Value.ToString("#,##0").Replace(",",".") },
                    {"thung2",first.thung2.Value.ToString("#,##0").Replace(",",".") },
                    {"hop1",first.hop1.Value.ToString("#,##0").Replace(",",".") },
                    {"hop2",first.hop2.Value.ToString("#,##0").Replace(",",".") },
                    {"tong1",first.tong1.Value.ToString("#,##0").Replace(",",".") },
                    {"tong2",first.tong2.Value.ToString("#,##0").Replace(",",".") },
                    {"vi2",first.vi2 != null && first.vi2 != 0 ? first.vi2.Value.ToString("#,##0").Replace(",",".") : "" },
                    {"c",first.vi2 != null && first.vi2 != 0 ? "+" : "" },
                    {"donvi",first.donvi },
                    {"donvi_en",first.donvi_en },
                    {"donvi_thung",first.donvi_thung },
                    {"donvi_thung_en",first.donvi_thung_en },
                    {"donvi_thungle",first.donvi_thungle },
                    {"donvi_thungle_en",first.donvi_thungle_en },
                    {"sop",first.sop },
                    {"ngaysop",first.ngaysop },
                };




            string[] fieldName = raw.Keys.ToArray();
            string[] fieldValue = raw.Values.ToArray();


            document.MailMerge.Execute(fieldName, fieldValue);





            document.SaveToFile("./wwwroot" + documentPath, Spire.Doc.FileFormat.Docx);

            return Json(new { success = true, link = Domain + documentPath, });


        }
        public async Task<JsonResult> phieuCOA(int id)
        {
            if (id == null)
            {
                return Json(new { success = false });
            }

            var viewPath = _configuration["Source:Path_Private"] + "\\qc\\templates\\mau coa.docx";
            var documentPath = "/temp/coa_" + DateTime.Now.ToFileTimeUtc() + ".docx";
            string Domain = (HttpContext.Request.IsHttps ? "https://" : "http://") + HttpContext.Request.Host.Value;


            var report = "";


            ///XUẤT PDF

            //Creates Document instance
            Spire.Doc.Document document = new Spire.Doc.Document();
            Section section = document.AddSection();

            document.LoadFromFile(viewPath, Spire.Doc.FileFormat.Docx);



            //document.MailMerge.ExecuteWidthRegion(datatable_details);

            ////Lấy 
            ///

            Dictionary<string, Spire.Doc.Table> replacements_table = new Dictionary<string, Spire.Doc.Table>(StringComparer.OrdinalIgnoreCase) { };

            var sqla = $"EXECUTE [QC_COA_IN_2] '{id}' ";

            var list1 = _qlsxContext.COAModel.FromSqlRaw($"{sqla}").ToList();

            var parent = list1.Where(d => d.id_parent == null).ToList();

            var count_list2 = list1.Where(d => d.id_parent != null && (d.tieuchuan.Trim() == "" || d.tieuchuan == null)).Count();
            ///FIELD OTHER
            var first = list1.FirstOrDefault();
            var raw = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
                {
                    {"mahh",first.mahh },
                    {"tenhh",first.tenhh },
                    {"tenhoatchat",first.tenhoatchat },
                    {"dangbaoche",first.dangbaoche },
                    {"ngaysx",first.ngaysx.Value.ToString("dd/MM/yy") },
                    {"quicachdonggoi",first.quicachdonggoi },
                    {"sophieu",first.solenh },
                    {"malo",first.malo },
                    {"sodk",first.sodk },
                    {"handung",first.handung },
                    {"theotieuchuan",first.theotieuchuan },
                    {"ghichu",first.ghichu },
                    {"ketluan",first.ketluan },
                    {"sop",first.sop },
                    {"ngaysop",first.ngaysop },
                };

            ///////TABLE
            var text = "Table_COA";
            Spire.Doc.Table table1 = section.AddTable(true);
            PreferredWidth width = new PreferredWidth(WidthType.Percentage, 100);
            table1.PreferredWidth = width;
            replacements_table.Add(text, table1);
            table1.ResetCells(list1.Count() + 1 - count_list2, 4);
            table1.SetColumnWidth(0, 40, CellWidthType.Point);
            table1.SetColumnWidth(1, 121, CellWidthType.Point);
            table1.SetColumnWidth(2, 226, CellWidthType.Point);
            table1.SetColumnWidth(3, 113, CellWidthType.Point);

            var row_current = 0;
            TableRow FRow;
            var i = 0;
            Paragraph p1;
            TextRange TR1;
            if (row_current == 0)
            {
                //// Row 0
                //Set the first row as table header
                FRow = table1.Rows[row_current];

                FRow.IsHeader = true;
                //Set the height and color of the first row
                FRow.Height = 30;
                i = 0;

                //table1.ApplyHorizontalMerge(row_current, 0, 1);

                FRow.Cells[i].CellFormat.VerticalAlignment = VerticalAlignment.Middle;
                p1 = FRow.Cells[i].AddParagraph();
                p1.Format.HorizontalAlignment = HorizontalAlignment.Center;
                TR1 = p1.AppendText("STT");

                //TR1.CharacterFormat.FontName = "Arial";

                TR1.CharacterFormat.FontSize = 12;

                TR1.CharacterFormat.Bold = true;

                p1 = FRow.Cells[i].AddParagraph();
                p1.Format.HorizontalAlignment = HorizontalAlignment.Center;
                TR1 = p1.AppendText("No.");

                //TR1.CharacterFormat.FontName = "Arial";

                TR1.CharacterFormat.FontSize = 12;

                TR1.CharacterFormat.Italic = true;
                i++;

                ////

                FRow.Cells[i].CellFormat.VerticalAlignment = VerticalAlignment.Middle;
                p1 = FRow.Cells[i].AddParagraph();
                p1.Format.HorizontalAlignment = HorizontalAlignment.Center;
                TR1 = p1.AppendText("Chỉ tiêu");

                //TR1.CharacterFormat.FontName = "Arial";

                TR1.CharacterFormat.FontSize = 12;

                TR1.CharacterFormat.Bold = true;

                p1 = FRow.Cells[i].AddParagraph();
                p1.Format.HorizontalAlignment = HorizontalAlignment.Center;
                TR1 = p1.AppendText("Tests");

                //TR1.CharacterFormat.FontName = "Arial";

                TR1.CharacterFormat.FontSize = 12;

                TR1.CharacterFormat.Italic = true;
                i++;

                FRow.Cells[i].CellFormat.VerticalAlignment = VerticalAlignment.Middle;
                p1 = FRow.Cells[i].AddParagraph();
                p1.Format.HorizontalAlignment = HorizontalAlignment.Center;
                TR1 = p1.AppendText("Tiêu chuẩn");

                //TR1.CharacterFormat.FontName = "Arial";

                TR1.CharacterFormat.FontSize = 12;

                TR1.CharacterFormat.Bold = true;

                p1 = FRow.Cells[i].AddParagraph();
                p1.Format.HorizontalAlignment = HorizontalAlignment.Center;
                TR1 = p1.AppendText("Specifications");

                //TR1.CharacterFormat.FontName = "Arial";

                TR1.CharacterFormat.FontSize = 12;

                TR1.CharacterFormat.Italic = true;
                i++;

                FRow.Cells[i].CellFormat.VerticalAlignment = VerticalAlignment.Middle;
                p1 = FRow.Cells[i].AddParagraph();
                p1.Format.HorizontalAlignment = HorizontalAlignment.Center;
                TR1 = p1.AppendText("Kết quả");

                //TR1.CharacterFormat.FontName = "Arial";

                TR1.CharacterFormat.FontSize = 12;

                TR1.CharacterFormat.Bold = true;

                p1 = FRow.Cells[i].AddParagraph();
                p1.Format.HorizontalAlignment = HorizontalAlignment.Center;
                TR1 = p1.AppendText("Results");

                //TR1.CharacterFormat.FontName = "Arial";

                TR1.CharacterFormat.FontSize = 12;

                TR1.CharacterFormat.Italic = true;
                i++;

                row_current++;
            }


            if (row_current != 0)
            {
                var stt = 1;
                foreach (var item in parent)
                {
                    //// Row 0
                    //Set the first row as table header
                    FRow = table1.Rows[row_current];

                    //FRow.IsHeader = true;
                    //Set the height and color of the first row
                    //FRow.Height = 30;
                    i = 0;

                    //table1.ApplyHorizontalMerge(row_current, 0, 1);

                    FRow.Cells[i].CellFormat.VerticalAlignment = VerticalAlignment.Middle;
                    p1 = FRow.Cells[i].AddParagraph();
                    p1.Format.HorizontalAlignment = HorizontalAlignment.Center;
                    TR1 = p1.AppendText(stt.ToString());

                    //TR1.CharacterFormat.FontName = "Arial";

                    TR1.CharacterFormat.FontSize = 12;

                    //TR1.CharacterFormat.Bold = true;

                    i++;

                    ////
                    var child = list1.Where(d => d.id_parent == item.id).ToList();
                    var child_trungten = child.Where(d => d.chitieu == item.chitieu).ToList();

                    if (child_trungten.Count() > 0)
                    {
                        //table1.ApplyHorizontalMerge(row_current, 1, 3);
                        table1.ApplyVerticalMerge(0, row_current, row_current + child.Count());
                        table1.ApplyVerticalMerge(1, row_current, row_current + child_trungten.Count());

                        FRow.Cells[i].CellFormat.VerticalAlignment = VerticalAlignment.Middle;
                        p1 = FRow.Cells[i].AddParagraph();
                        p1.Format.HorizontalAlignment = HorizontalAlignment.Justify;
                        TR1 = p1.AppendText(item.chitieu);

                        //TR1.CharacterFormat.FontName = "Arial";

                        TR1.CharacterFormat.FontSize = 12;

                        //TR1.CharacterFormat.Bold = true;

                        if (item.chitieu_en != null && item.chitieu_en != "")
                        {
                            p1 = FRow.Cells[i].AddParagraph();
                            p1.Format.HorizontalAlignment = HorizontalAlignment.Justify;
                            TR1 = p1.AppendText(item.chitieu_en);

                            //TR1.CharacterFormat.FontName = "Arial";

                            TR1.CharacterFormat.FontSize = 12;

                            TR1.CharacterFormat.Italic = true;
                        }

                        i++;
                        ////
                        FRow.Cells[i].CellFormat.VerticalAlignment = VerticalAlignment.Middle;
                        p1 = FRow.Cells[i].AddParagraph();
                        p1.Format.HorizontalAlignment = HorizontalAlignment.Justify;
                        TR1 = p1.AppendText(item.tieuchuan.Replace("\t", " "));

                        //TR1.CharacterFormat.FontName = "Arial";

                        TR1.CharacterFormat.FontSize = 12;

                        //TR1.CharacterFormat.Bold = true;
                        if (item.tieuchuan_en != null && item.tieuchuan_en != "")
                        {
                            p1 = FRow.Cells[i].AddParagraph();
                            p1.Format.HorizontalAlignment = HorizontalAlignment.Justify;
                            TR1 = p1.AppendText(item.tieuchuan_en.Replace("\t", " "));

                            //TR1.CharacterFormat.FontName = "Arial";

                            TR1.CharacterFormat.FontSize = 12;

                            TR1.CharacterFormat.Italic = true;
                        }
                        i++;

                        ///// thực tế
                        if (item.type == "Text")
                        {
                            ////
                            FRow.Cells[i].CellFormat.VerticalAlignment = VerticalAlignment.Middle;
                            p1 = FRow.Cells[i].AddParagraph();
                            p1.Format.HorizontalAlignment = HorizontalAlignment.Center;
                            TR1 = p1.AppendText(item.thucte_text);

                            //TR1.CharacterFormat.FontName = "Arial";

                            TR1.CharacterFormat.FontSize = 12;

                            //TR1.CharacterFormat.Bold = true;
                            if (item.thucte_text_en != null && item.thucte_text_en != "")
                            {
                                p1 = FRow.Cells[i].AddParagraph();
                                p1.Format.HorizontalAlignment = HorizontalAlignment.Center;
                                TR1 = p1.AppendText(item.thucte_text_en);

                                //TR1.CharacterFormat.FontName = "Arial";

                                TR1.CharacterFormat.FontSize = 12;

                                TR1.CharacterFormat.Italic = true;
                            }
                        }
                        else if (item.type == "Number")
                        {
                            ////
                            var text_dat = item.dat == true ? "Đạt" : "Không đạt";
                            FRow.Cells[i].CellFormat.VerticalAlignment = VerticalAlignment.Middle;
                            p1 = FRow.Cells[i].AddParagraph();
                            p1.Format.HorizontalAlignment = HorizontalAlignment.Center;
                            TR1 = p1.AppendText(text_dat);

                            //TR1.CharacterFormat.FontName = "Arial";

                            TR1.CharacterFormat.FontSize = 12;

                            ////
                            var text_item1 = "";
                            if (item.thucte != null)
                            {
                                text_item1 = item.thucte.Value.ToString("#,##0.#####").Replace(",", "*").Replace(".", ",").Replace("*", ".") + " " + item.mota;
                            }
                            if (item.thucte_text != null)
                            {
                                text_item1 += " " + item.thucte_text;
                            }
                            FRow.Cells[i].CellFormat.VerticalAlignment = VerticalAlignment.Middle;
                            p1 = FRow.Cells[i].AddParagraph();
                            p1.Format.HorizontalAlignment = HorizontalAlignment.Center;
                            TR1 = p1.AppendText(text_item1);

                            //TR1.CharacterFormat.FontName = "Arial";

                            TR1.CharacterFormat.FontSize = 12;

                        }
                        else if (item.type == "Boolean")
                        {
                            ////
                            var text_item = item.dat == true ? "Đạt" : "Không đạt";
                            if (item.thucte_text != null)
                            {
                                text_item += " " + item.thucte_text;
                            }
                            FRow.Cells[i].CellFormat.VerticalAlignment = VerticalAlignment.Middle;
                            p1 = FRow.Cells[i].AddParagraph();
                            p1.Format.HorizontalAlignment = HorizontalAlignment.Center;
                            TR1 = p1.AppendText(text_item);

                            //TR1.CharacterFormat.FontName = "Arial";

                            TR1.CharacterFormat.FontSize = 12;

                        }

                        row_current++;

                        foreach (var item1 in child)
                        {
                            var y = 2;
                            FRow = table1.Rows[row_current];
                            FRow.Cells[y].CellFormat.VerticalAlignment = VerticalAlignment.Middle;
                            p1 = FRow.Cells[y].AddParagraph();
                            p1.Format.HorizontalAlignment = HorizontalAlignment.Justify;
                            TR1 = p1.AppendText(item1.tieuchuan.Replace("\t", " "));

                            //TR1.CharacterFormat.FontName = "Arial";

                            TR1.CharacterFormat.FontSize = 12;

                            //TR1.CharacterFormat.Bold = true;

                            if (item1.tieuchuan_en != null && item1.tieuchuan_en != "")
                            {
                                p1 = FRow.Cells[y].AddParagraph();
                                p1.Format.HorizontalAlignment = HorizontalAlignment.Justify;
                                TR1 = p1.AppendText(item1.tieuchuan_en.Replace("\t", " "));

                                //TR1.CharacterFormat.FontName = "Arial";

                                TR1.CharacterFormat.FontSize = 12;

                                TR1.CharacterFormat.Italic = true;
                            }
                            y++;
                            ///
                            if (item1.type == "Text")
                            {

                                ////
                                FRow.Cells[y].CellFormat.VerticalAlignment = VerticalAlignment.Middle;
                                p1 = FRow.Cells[y].AddParagraph();
                                p1.Format.HorizontalAlignment = HorizontalAlignment.Center;
                                TR1 = p1.AppendText(item1.thucte_text);

                                //TR1.CharacterFormat.FontName = "Arial";

                                TR1.CharacterFormat.FontSize = 12;

                                //TR1.CharacterFormat.Bold = true;
                                if (item1.thucte_text_en != null && item1.thucte_text_en != "")
                                {
                                    p1 = FRow.Cells[y].AddParagraph();
                                    p1.Format.HorizontalAlignment = HorizontalAlignment.Center;
                                    TR1 = p1.AppendText(item1.thucte_text_en);

                                    //TR1.CharacterFormat.FontName = "Arial";

                                    TR1.CharacterFormat.FontSize = 12;

                                    TR1.CharacterFormat.Italic = true;
                                }
                            }
                            else if (item1.type == "Number")
                            {
                                ////
                                var text_dat = item1.dat == true ? "Đạt" : "Không đạt";
                                FRow.Cells[y].CellFormat.VerticalAlignment = VerticalAlignment.Middle;
                                p1 = FRow.Cells[y].AddParagraph();
                                p1.Format.HorizontalAlignment = HorizontalAlignment.Center;
                                TR1 = p1.AppendText(text_dat);

                                //TR1.CharacterFormat.FontName = "Arial";

                                TR1.CharacterFormat.FontSize = 12;

                                ////
                                var text_item1 = "";
                                if (item1.thucte != null)
                                {
                                    text_item1 = item1.thucte.Value.ToString("#,##0.#####").Replace(",", "*").Replace(".", ",").Replace("*", ".") + " " + item1.mota;
                                }
                                if (item1.thucte_text != null)
                                {
                                    text_item1 += " " + item1.thucte_text;
                                }
                                FRow.Cells[y].CellFormat.VerticalAlignment = VerticalAlignment.Middle;
                                p1 = FRow.Cells[y].AddParagraph();
                                p1.Format.HorizontalAlignment = HorizontalAlignment.Center;
                                TR1 = p1.AppendText(text_item1);

                                //TR1.CharacterFormat.FontName = "Arial";

                                TR1.CharacterFormat.FontSize = 12;

                            }
                            else if (item1.type == "Boolean")
                            {
                                ////
                                var text_item1 = item1.dat == true ? "Đạt" : "Không đạt";
                                if (item1.thucte_text != null)
                                {
                                    text_item1 += " " + item1.thucte_text;
                                }
                                FRow.Cells[y].CellFormat.VerticalAlignment = VerticalAlignment.Middle;
                                p1 = FRow.Cells[y].AddParagraph();
                                p1.Format.HorizontalAlignment = HorizontalAlignment.Center;
                                TR1 = p1.AppendText(text_item1);

                                //TR1.CharacterFormat.FontName = "Arial";

                                TR1.CharacterFormat.FontSize = 12;

                            }

                            y++;
                            row_current++;
                        }


                    }
                    else if (child.Count() > 0 && (item.tieuchuan != null && item.tieuchuan.Trim() != ""))
                    {
                        //table1.ApplyHorizontalMerge(row_current, 1, 3);
                        //table1.ApplyVerticalMerge(0, row_current, row_current + child.Count());

                        FRow.Cells[i].CellFormat.VerticalAlignment = VerticalAlignment.Middle;
                        p1 = FRow.Cells[i].AddParagraph();
                        p1.Format.HorizontalAlignment = HorizontalAlignment.Justify;
                        TR1 = p1.AppendText(item.chitieu);

                        //TR1.CharacterFormat.FontName = "Arial";

                        TR1.CharacterFormat.FontSize = 12;

                        //TR1.CharacterFormat.Bold = true;

                        if (item.chitieu_en != null && item.chitieu_en != "")
                        {
                            p1 = FRow.Cells[i].AddParagraph();
                            p1.Format.HorizontalAlignment = HorizontalAlignment.Justify;
                            TR1 = p1.AppendText(item.chitieu_en);

                            //TR1.CharacterFormat.FontName = "Arial";

                            TR1.CharacterFormat.FontSize = 12;

                            TR1.CharacterFormat.Italic = true;
                        }

                        i++;
                        ////
                        FRow.Cells[i].CellFormat.VerticalAlignment = VerticalAlignment.Middle;
                        p1 = FRow.Cells[i].AddParagraph();
                        p1.Format.HorizontalAlignment = HorizontalAlignment.Justify;
                        TR1 = p1.AppendText(item.tieuchuan.Replace("\t", " "));

                        //TR1.CharacterFormat.FontName = "Arial";

                        TR1.CharacterFormat.FontSize = 12;

                        //TR1.CharacterFormat.Bold = true;
                        if (item.tieuchuan_en != null && item.tieuchuan_en != "")
                        {
                            p1 = FRow.Cells[i].AddParagraph();
                            p1.Format.HorizontalAlignment = HorizontalAlignment.Justify;
                            TR1 = p1.AppendText(item.tieuchuan_en.Replace("\t", " "));

                            //TR1.CharacterFormat.FontName = "Arial";

                            TR1.CharacterFormat.FontSize = 12;

                            TR1.CharacterFormat.Italic = true;
                        }
                        i++;

                        /////
                        if (item.type == "Text")
                        {
                            ////
                            FRow.Cells[i].CellFormat.VerticalAlignment = VerticalAlignment.Middle;
                            p1 = FRow.Cells[i].AddParagraph();
                            p1.Format.HorizontalAlignment = HorizontalAlignment.Center;
                            TR1 = p1.AppendText(item.thucte_text);

                            //TR1.CharacterFormat.FontName = "Arial";

                            TR1.CharacterFormat.FontSize = 12;

                            //TR1.CharacterFormat.Bold = true;
                            if (item.thucte_text_en != null && item.thucte_text_en != "")
                            {
                                p1 = FRow.Cells[i].AddParagraph();
                                p1.Format.HorizontalAlignment = HorizontalAlignment.Center;
                                TR1 = p1.AppendText(item.thucte_text_en);

                                //TR1.CharacterFormat.FontName = "Arial";

                                TR1.CharacterFormat.FontSize = 12;

                                TR1.CharacterFormat.Italic = true;
                            }
                        }
                        else if (item.type == "Number")
                        {
                            ////
                            var text_dat = item.dat == true ? "Đạt" : "Không đạt";
                            FRow.Cells[i].CellFormat.VerticalAlignment = VerticalAlignment.Middle;
                            p1 = FRow.Cells[i].AddParagraph();
                            p1.Format.HorizontalAlignment = HorizontalAlignment.Center;
                            TR1 = p1.AppendText(text_dat);

                            //TR1.CharacterFormat.FontName = "Arial";

                            TR1.CharacterFormat.FontSize = 12;

                            ////
                            var text_item1 = "";
                            if (item.thucte != null)
                            {
                                text_item1 = item.thucte.Value.ToString("#,##0.#####").Replace(",", "*").Replace(".", ",").Replace("*", ".") + " " + item.mota;
                            }
                            if (item.thucte_text != null)
                            {
                                text_item1 += " " + item.thucte_text;
                            }
                            FRow.Cells[i].CellFormat.VerticalAlignment = VerticalAlignment.Middle;
                            p1 = FRow.Cells[i].AddParagraph();
                            p1.Format.HorizontalAlignment = HorizontalAlignment.Center;
                            TR1 = p1.AppendText(text_item1);

                            //TR1.CharacterFormat.FontName = "Arial";

                            TR1.CharacterFormat.FontSize = 12;

                        }
                        else if (item.type == "Boolean")
                        {
                            ////
                            var text_item = item.dat == true ? "Đạt" : "Không đạt";
                            if (item.thucte_text != null)
                            {
                                text_item += " " + item.thucte_text;
                            }
                            FRow.Cells[i].CellFormat.VerticalAlignment = VerticalAlignment.Middle;
                            p1 = FRow.Cells[i].AddParagraph();
                            p1.Format.HorizontalAlignment = HorizontalAlignment.Center;
                            TR1 = p1.AppendText(text_item);

                            //TR1.CharacterFormat.FontName = "Arial";

                            TR1.CharacterFormat.FontSize = 12;

                        }

                        foreach (var item1 in child)
                        {
                            var text_item = "";
                            if (item1.type == "Text")
                            {
                                text_item = item1.thucte_text;
                            }
                            else if (item1.type == "Number")
                            {
                                text_item = "";
                                if (item1.thucte != null)
                                {
                                    text_item = item1.thucte.Value.ToString("#,##0.#####").Replace(",", "*").Replace(".", ",").Replace("*", ".") + " " + item1.mota;
                                }
                                if (item1.thucte_text != null)
                                {
                                    text_item += " " + item1.thucte_text;
                                }
                            }
                            else if (item.type == "Boolean")
                            {
                                text_item = item1.thucte == 1 ? "Đạt" : "Không đạt";
                                if (item1.thucte_text != null)
                                {
                                    text_item += " " + item1.thucte_text;
                                }
                            }
                            p1 = FRow.Cells[i].AddParagraph();
                            p1.Format.HorizontalAlignment = HorizontalAlignment.Justify;
                            TR1 = p1.AppendText(item1.chitieu + ": " + text_item);

                            //TR1.CharacterFormat.FontName = "Arial";

                            TR1.CharacterFormat.FontSize = 12;

                            //TR1.CharacterFormat.Italic = true;
                        }

                        row_current++;


                    }
                    else if (child.Count() > 0 && (item.tieuchuan == null || item.tieuchuan.Trim() == ""))
                    {
                        table1.ApplyHorizontalMerge(row_current, 1, 3);
                        table1.ApplyVerticalMerge(0, row_current, row_current + child.Count());

                        FRow.Cells[i].CellFormat.VerticalAlignment = VerticalAlignment.Middle;
                        p1 = FRow.Cells[i].AddParagraph();
                        p1.Format.HorizontalAlignment = HorizontalAlignment.Justify;
                        TR1 = p1.AppendText(item.chitieu);

                        //TR1.CharacterFormat.FontName = "Arial";

                        TR1.CharacterFormat.FontSize = 12;

                        //TR1.CharacterFormat.Bold = true;

                        if (item.chitieu_en != null && item.chitieu_en != "")
                        {
                            p1 = FRow.Cells[i].AddParagraph();
                            p1.Format.HorizontalAlignment = HorizontalAlignment.Justify;
                            TR1 = p1.AppendText(item.chitieu_en);

                            //TR1.CharacterFormat.FontName = "Arial";

                            TR1.CharacterFormat.FontSize = 12;

                            TR1.CharacterFormat.Italic = true;
                        }
                        row_current++;
                        foreach (var item1 in child)
                        {
                            var y = i;
                            FRow = table1.Rows[row_current];
                            FRow.Cells[y].CellFormat.VerticalAlignment = VerticalAlignment.Middle;
                            p1 = FRow.Cells[y].AddParagraph();
                            p1.Format.HorizontalAlignment = HorizontalAlignment.Justify;
                            TR1 = p1.AppendText(item1.chitieu);

                            //TR1.CharacterFormat.FontName = "Arial";

                            TR1.CharacterFormat.FontSize = 12;

                            //TR1.CharacterFormat.Bold = true;

                            if (item1.chitieu_en != null && item1.chitieu_en != "")
                            {
                                p1 = FRow.Cells[y].AddParagraph();
                                p1.Format.HorizontalAlignment = HorizontalAlignment.Justify;
                                TR1 = p1.AppendText(item1.chitieu_en);

                                //TR1.CharacterFormat.FontName = "Arial";

                                TR1.CharacterFormat.FontSize = 12;

                                TR1.CharacterFormat.Italic = true;
                            }
                            y++;
                            ////
                            FRow = table1.Rows[row_current];
                            FRow.Cells[y].CellFormat.VerticalAlignment = VerticalAlignment.Middle;
                            p1 = FRow.Cells[y].AddParagraph();
                            p1.Format.HorizontalAlignment = HorizontalAlignment.Justify;
                            TR1 = p1.AppendText(item1.tieuchuan.Replace("\t", " "));

                            //TR1.CharacterFormat.FontName = "Arial";

                            TR1.CharacterFormat.FontSize = 12;

                            //TR1.CharacterFormat.Bold = true;

                            if (item1.tieuchuan_en != null && item1.tieuchuan_en != "")
                            {
                                p1 = FRow.Cells[y].AddParagraph();
                                p1.Format.HorizontalAlignment = HorizontalAlignment.Justify;
                                TR1 = p1.AppendText(item1.tieuchuan_en.Replace("\t", " "));

                                //TR1.CharacterFormat.FontName = "Arial";

                                TR1.CharacterFormat.FontSize = 12;

                                TR1.CharacterFormat.Italic = true;
                            }
                            y++;
                            ///
                            if (item1.type == "Text")
                            {

                                ////
                                FRow.Cells[y].CellFormat.VerticalAlignment = VerticalAlignment.Middle;
                                p1 = FRow.Cells[y].AddParagraph();
                                p1.Format.HorizontalAlignment = HorizontalAlignment.Center;
                                TR1 = p1.AppendText(item1.thucte_text);

                                //TR1.CharacterFormat.FontName = "Arial";

                                TR1.CharacterFormat.FontSize = 12;

                                //TR1.CharacterFormat.Bold = true;
                                if (item1.thucte_text_en != null && item1.thucte_text_en != "")
                                {
                                    p1 = FRow.Cells[y].AddParagraph();
                                    p1.Format.HorizontalAlignment = HorizontalAlignment.Center;
                                    TR1 = p1.AppendText(item1.thucte_text_en);

                                    //TR1.CharacterFormat.FontName = "Arial";

                                    TR1.CharacterFormat.FontSize = 12;

                                    TR1.CharacterFormat.Italic = true;
                                }
                            }
                            else if (item1.type == "Number")
                            {
                                ////
                                var text_dat = item1.dat == true ? "Đạt" : "Không đạt";
                                FRow.Cells[y].CellFormat.VerticalAlignment = VerticalAlignment.Middle;
                                p1 = FRow.Cells[y].AddParagraph();
                                p1.Format.HorizontalAlignment = HorizontalAlignment.Center;
                                TR1 = p1.AppendText(text_dat);

                                //TR1.CharacterFormat.FontName = "Arial";

                                TR1.CharacterFormat.FontSize = 12;

                                ////
                                var text_item1 = "";
                                if (item1.thucte != null)
                                {
                                    text_item1 = item1.thucte.Value.ToString("#,##0.#####").Replace(",", "*").Replace(".", ",").Replace("*", ".") + " " + item1.mota;
                                }
                                if (item1.thucte_text != null)
                                {
                                    text_item1 += " " + item1.thucte_text;
                                }
                                FRow.Cells[y].CellFormat.VerticalAlignment = VerticalAlignment.Middle;
                                p1 = FRow.Cells[y].AddParagraph();
                                p1.Format.HorizontalAlignment = HorizontalAlignment.Center;
                                TR1 = p1.AppendText(text_item1);

                                //TR1.CharacterFormat.FontName = "Arial";

                                TR1.CharacterFormat.FontSize = 12;

                            }
                            else if (item1.type == "Boolean")
                            {
                                ////
                                var text_item1 = item1.dat == true ? "Đạt" : "Không đạt";
                                if (item1.thucte_text != null)
                                {
                                    text_item1 += " " + item1.thucte_text;
                                }
                                FRow.Cells[y].CellFormat.VerticalAlignment = VerticalAlignment.Middle;
                                p1 = FRow.Cells[y].AddParagraph();
                                p1.Format.HorizontalAlignment = HorizontalAlignment.Center;
                                TR1 = p1.AppendText(text_item1);

                                //TR1.CharacterFormat.FontName = "Arial";

                                TR1.CharacterFormat.FontSize = 12;

                            }

                            y++;
                            row_current++;
                        }
                    }
                    else
                    {
                        FRow.Cells[i].CellFormat.VerticalAlignment = VerticalAlignment.Middle;
                        p1 = FRow.Cells[i].AddParagraph();
                        p1.Format.HorizontalAlignment = HorizontalAlignment.Justify;
                        TR1 = p1.AppendText(item.chitieu);

                        //TR1.CharacterFormat.FontName = "Arial";

                        TR1.CharacterFormat.FontSize = 12;

                        //TR1.CharacterFormat.Bold = true;
                        if (item.chitieu_en != null && item.chitieu_en != "")
                        {
                            p1 = FRow.Cells[i].AddParagraph();
                            p1.Format.HorizontalAlignment = HorizontalAlignment.Justify;
                            TR1 = p1.AppendText(item.chitieu_en);

                            //TR1.CharacterFormat.FontName = "Arial";

                            TR1.CharacterFormat.FontSize = 12;

                            TR1.CharacterFormat.Italic = true;
                        }
                        i++;
                        ////
                        FRow.Cells[i].CellFormat.VerticalAlignment = VerticalAlignment.Middle;
                        p1 = FRow.Cells[i].AddParagraph();
                        p1.Format.HorizontalAlignment = HorizontalAlignment.Justify;
                        TR1 = p1.AppendText(item.tieuchuan.Replace("\t", " "));

                        //TR1.CharacterFormat.FontName = "Arial";

                        TR1.CharacterFormat.FontSize = 12;

                        //TR1.CharacterFormat.Bold = true;
                        if (item.tieuchuan_en != null && item.tieuchuan_en != "")
                        {
                            p1 = FRow.Cells[i].AddParagraph();
                            p1.Format.HorizontalAlignment = HorizontalAlignment.Justify;
                            TR1 = p1.AppendText(item.tieuchuan_en.Replace("\t", " "));

                            //TR1.CharacterFormat.FontName = "Arial";

                            TR1.CharacterFormat.FontSize = 12;

                            TR1.CharacterFormat.Italic = true;
                        }
                        i++;
                        if (item.type == "Text")
                        {
                            ////
                            FRow.Cells[i].CellFormat.VerticalAlignment = VerticalAlignment.Middle;
                            p1 = FRow.Cells[i].AddParagraph();
                            p1.Format.HorizontalAlignment = HorizontalAlignment.Center;
                            TR1 = p1.AppendText(item.thucte_text);

                            //TR1.CharacterFormat.FontName = "Arial";

                            TR1.CharacterFormat.FontSize = 12;

                            //TR1.CharacterFormat.Bold = true;
                            if (item.thucte_text_en != null && item.thucte_text_en != "")
                            {
                                p1 = FRow.Cells[i].AddParagraph();
                                p1.Format.HorizontalAlignment = HorizontalAlignment.Center;
                                TR1 = p1.AppendText(item.thucte_text_en);

                                //TR1.CharacterFormat.FontName = "Arial";

                                TR1.CharacterFormat.FontSize = 12;

                                TR1.CharacterFormat.Italic = true;
                            }
                        }
                        else if (item.type == "Number")
                        {
                            ////
                            var text_dat = item.dat == true ? "Đạt" : "Không đạt";

                            FRow.Cells[i].CellFormat.VerticalAlignment = VerticalAlignment.Middle;
                            p1 = FRow.Cells[i].AddParagraph();
                            p1.Format.HorizontalAlignment = HorizontalAlignment.Center;
                            TR1 = p1.AppendText(text_dat);

                            //TR1.CharacterFormat.FontName = "Arial";

                            TR1.CharacterFormat.FontSize = 12;

                            ////
                            ///
                            var text_item1 = "";
                            if (item.thucte != null)
                            {
                                text_item1 = item.thucte.Value.ToString("#,##0.#####").Replace(",", "*").Replace(".", ",").Replace("*", ".") + " " + item.mota;
                            }
                            if (item.thucte_text != null)
                            {
                                text_item1 += " " + item.thucte_text;
                            }
                            FRow.Cells[i].CellFormat.VerticalAlignment = VerticalAlignment.Middle;
                            p1 = FRow.Cells[i].AddParagraph();
                            p1.Format.HorizontalAlignment = HorizontalAlignment.Center;
                            TR1 = p1.AppendText(text_item1);

                            //TR1.CharacterFormat.FontName = "Arial";

                            TR1.CharacterFormat.FontSize = 12;

                        }
                        else if (item.type == "Boolean")
                        {
                            ////
                            var text_item = item.dat == true ? "Đạt" : "Không đạt";
                            if (item.thucte_text != null)
                            {
                                text_item += " " + item.thucte_text;
                            }
                            FRow.Cells[i].CellFormat.VerticalAlignment = VerticalAlignment.Middle;
                            p1 = FRow.Cells[i].AddParagraph();
                            p1.Format.HorizontalAlignment = HorizontalAlignment.Center;
                            TR1 = p1.AppendText(text_item);

                            //TR1.CharacterFormat.FontName = "Arial";

                            TR1.CharacterFormat.FontSize = 12;


                        }

                        i++;
                        row_current++;
                    }


                    stt++;

                }
            }




            string[] fieldName = raw.Keys.ToArray();
            string[] fieldValue = raw.Values.ToArray();

            string[] MergeFieldNames = document.MailMerge.GetMergeFieldNames();
            string[] GroupNames = document.MailMerge.GetMergeGroupNames();

            document.MailMerge.Execute(fieldName, fieldValue);

            foreach (KeyValuePair<string, Spire.Doc.Table> entry in replacements_table)
            {
                TextSelection selection = document.FindString(entry.Key, true, true);
                if (selection == null)
                {
                    var table = entry.Value;
                    //table
                    continue;
                }
                TextRange range = selection.GetAsOneRange();
                Paragraph paragraph = range.OwnerParagraph;
                Body body = paragraph.OwnerTextBody;
                int index = body.ChildObjects.IndexOf(paragraph);
                //var fontsize = paragrap

                body.ChildObjects.Remove(paragraph);
                body.ChildObjects.Insert(index, entry.Value);
            }



            document.SaveToFile("./wwwroot" + documentPath, Spire.Doc.FileFormat.Docx);

            return Json(new { success = true, link = Domain + documentPath, });


        }

        public async Task<JsonResult> danhgiatacdong(string sochange, string ngaydenghi)
        {
            if (sochange == null || ngaydenghi == null)
            {
                return Json(new { success = false });
            }

            var viewPath = _configuration["Source:Path_Private"] + "\\qa\\templates\\danhgiatacdong.docx";
            var documentPath = "/temp/danhgiatacdong_" + DateTime.Now.ToFileTimeUtc() + ".docx";
            string Domain = (HttpContext.Request.IsHttps ? "https://" : "http://") + HttpContext.Request.Host.Value;


            var report = "";


            ///XUẤT PDF

            //Creates Document instance
            Spire.Doc.Document document = new Spire.Doc.Document();
            Section section = document.AddSection();

            document.LoadFromFile(viewPath, Spire.Doc.FileFormat.Docx);



            //document.MailMerge.ExecuteWidthRegion(datatable_details);

            ////Lấy 
            ///

            Dictionary<string, Spire.Doc.Table> replacements_table = new Dictionary<string, Spire.Doc.Table>(StringComparer.OrdinalIgnoreCase) { };

            var sql_CHANGECONTROL_PHANLOAI = $"EXECUTE [CHANGECONTROL_PHANLOAI] '{sochange}' ,'' ,'{ngaydenghi}' ";

            var CHANGECONTROL_PHANLOAI_list = _qlsxContext.CHANGECONTROL_PHANLOAI.FromSqlRaw($"{sql_CHANGECONTROL_PHANLOAI}").ToList();

            var CHANGECONTROL_PHANLOAI = CHANGECONTROL_PHANLOAI_list.FirstOrDefault();

            var sql_CHANGECONTROL_THANHVIEN = $"EXECUTE [CHANGECONTROL_THANHVIEN] '{sochange}' ,'' ,'{ngaydenghi}' ";

            var CHANGECONTROL_THANHVIEN = _qlsxContext.CHANGECONTROL_THANHVIEN.FromSqlRaw($"{sql_CHANGECONTROL_THANHVIEN}").ToList();

            var sql_CHANGECONTROL_CAPA = $"EXECUTE [CHANGECONTROL_CAPA_2] '{sochange}' ,'' ,'{ngaydenghi}' ";

            var list1 = _qlsxContext.CHANGECONTROL_CAPA.FromSqlRaw($"{sql_CHANGECONTROL_CAPA}").ToList();

            var sql_CHANGECONTROL_ANBAN_IN = $"EXECUTE [GMP_DANHMUCTHIETBI_ANBAN_IN] '{sochange}' ,10 ";

            var CHANGECONTROL_ANBAN_IN = _qlsxContext.CHANGECONTROL_ANBAN_IN.FromSqlRaw($"{sql_CHANGECONTROL_ANBAN_IN}").ToList();


            var list_nhom = list1.GroupBy(d => d.manhom_muc).Select(d => new
            {
                manhom_muc = d.Key,
                tennhom_muc = d.FirstOrDefault().tennhom_muc,
                tennhom_en_muc = d.FirstOrDefault().tennhom_en_muc,
                list = d.GroupBy(e => e.id_muc).Select(e => new
                {
                    id_muc = e.Key,
                    ten_muc = e.FirstOrDefault().ten_muc,
                    ten_en_muc = e.FirstOrDefault().ten_en_muc,
                    danhgia = e.FirstOrDefault().danhgia,
                    list = e.ToList()
                }).ToList()
            }).ToList();

            var raw = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
                {
                    {"solenh",sochange },
                    {"anban",CHANGECONTROL_PHANLOAI.anban ?? ""},
                    {"ngayhieuluc",CHANGECONTROL_PHANLOAI.ngayhieuluc != null ? CHANGECONTROL_PHANLOAI.ngayhieuluc.Value.ToString("dd/MM/yyyy") :"" },

                    {"td_thap",CHANGECONTROL_PHANLOAI.phanloai == "Thấp" ? "☒" : "☐" },
                    {"td_trung",CHANGECONTROL_PHANLOAI.phanloai == "Trung bình" ? "☒" : "☐" },
                    {"td_lon",CHANGECONTROL_PHANLOAI.phanloai == "Lớn" ? "☒" : "☐" },

                    {"ahcl_thap",CHANGECONTROL_PHANLOAI.anhhuongchatluong_thap ?? ""},
                    {"ahcl_trung",CHANGECONTROL_PHANLOAI.anhhuongchatluong_trungbinh ?? ""},
                    {"ahcl_lon",CHANGECONTROL_PHANLOAI.anhhuongchatluong_cao ?? ""},

                    {"hoso_thap",CHANGECONTROL_PHANLOAI.anhhuonghoso_thap ?? ""},
                    {"hoso_trung",CHANGECONTROL_PHANLOAI.anhhuonghoso_trungbinh ?? ""},
                    {"hoso_lon",CHANGECONTROL_PHANLOAI.anhhuonghoso_cao ?? ""},

                };

            ///////TABLE
            if ("Table_THANHVIEN" == "Table_THANHVIEN")
            {


                var text = "Table_THANHVIEN";
                Spire.Doc.Table table1 = section.AddTable(true);
                PreferredWidth width = new PreferredWidth(WidthType.Percentage, 100);
                table1.PreferredWidth = width;

                table1.ResetCells(CHANGECONTROL_THANHVIEN.Count() + 1, 3);

                table1.SetColumnWidth(0, 46, CellWidthType.Point);
                table1.SetColumnWidth(1, 209, CellWidthType.Point);
                table1.SetColumnWidth(2, 226, CellWidthType.Point);

                var row_current = 0;
                TableRow FRow;
                var i = 0;
                Paragraph p1;
                TextRange TR1;
                if (row_current == 0)
                {
                    //// Row 0
                    //Set the first row as table header
                    FRow = table1.Rows[row_current];

                    FRow.IsHeader = true;
                    //Set the height and color of the first row
                    FRow.Height = 30;
                    i = 0;

                    //table1.ApplyHorizontalMerge(row_current, 0, 1);

                    FRow.Cells[i].CellFormat.VerticalAlignment = VerticalAlignment.Middle;
                    p1 = FRow.Cells[i].AddParagraph();
                    p1.Format.HorizontalAlignment = HorizontalAlignment.Center;
                    TR1 = p1.AppendText("STT");

                    //TR1.CharacterFormat.FontName = "Arial";

                    TR1.CharacterFormat.FontSize = 12;

                    TR1.CharacterFormat.Bold = true;

                    p1 = FRow.Cells[i].AddParagraph();
                    p1.Format.HorizontalAlignment = HorizontalAlignment.Center;
                    TR1 = p1.AppendText("No.");

                    //TR1.CharacterFormat.FontName = "Arial";

                    TR1.CharacterFormat.FontSize = 12;

                    TR1.CharacterFormat.Italic = true;
                    i++;

                    ////

                    FRow.Cells[i].CellFormat.VerticalAlignment = VerticalAlignment.Middle;
                    p1 = FRow.Cells[i].AddParagraph();
                    p1.Format.HorizontalAlignment = HorizontalAlignment.Center;
                    TR1 = p1.AppendText("Họ tên");

                    //TR1.CharacterFormat.FontName = "Arial";

                    TR1.CharacterFormat.FontSize = 12;

                    TR1.CharacterFormat.Bold = true;

                    p1 = FRow.Cells[i].AddParagraph();
                    p1.Format.HorizontalAlignment = HorizontalAlignment.Center;
                    TR1 = p1.AppendText("Full Name");

                    //TR1.CharacterFormat.FontName = "Arial";

                    TR1.CharacterFormat.FontSize = 12;

                    TR1.CharacterFormat.Italic = true;
                    i++;



                    FRow.Cells[i].CellFormat.VerticalAlignment = VerticalAlignment.Middle;
                    p1 = FRow.Cells[i].AddParagraph();
                    p1.Format.HorizontalAlignment = HorizontalAlignment.Center;
                    TR1 = p1.AppendText("Bộ phận");

                    //TR1.CharacterFormat.FontName = "Arial";

                    TR1.CharacterFormat.FontSize = 12;

                    TR1.CharacterFormat.Bold = true;

                    p1 = FRow.Cells[i].AddParagraph();
                    p1.Format.HorizontalAlignment = HorizontalAlignment.Center;
                    TR1 = p1.AppendText("Department");

                    //TR1.CharacterFormat.FontName = "Arial";

                    TR1.CharacterFormat.FontSize = 12;

                    TR1.CharacterFormat.Italic = true;
                    i++;

                    row_current++;
                }


                if (row_current != 0)
                {
                    var stt = 1;
                    foreach (var item in CHANGECONTROL_THANHVIEN)
                    {
                        //// Row 0
                        //Set the first row as table header
                        FRow = table1.Rows[row_current];

                        //FRow.IsHeader = true;
                        //Set the height and color of the first row
                        //FRow.Height = 30;
                        i = 0;

                        //table1.ApplyHorizontalMerge(row_current, 0, 1);

                        FRow.Cells[i].CellFormat.VerticalAlignment = VerticalAlignment.Middle;
                        p1 = FRow.Cells[i].AddParagraph();
                        p1.Format.HorizontalAlignment = HorizontalAlignment.Center;
                        TR1 = p1.AppendText(stt.ToString());

                        //TR1.CharacterFormat.FontName = "Arial";

                        TR1.CharacterFormat.FontSize = 12;

                        //TR1.CharacterFormat.Bold = true;

                        i++;

                        ////

                        FRow.Cells[i].CellFormat.VerticalAlignment = VerticalAlignment.Middle;
                        p1 = FRow.Cells[i].AddParagraph();
                        p1.Format.HorizontalAlignment = HorizontalAlignment.Justify;
                        TR1 = p1.AppendText(item.hoten);

                        //TR1.CharacterFormat.FontName = "Arial";

                        TR1.CharacterFormat.FontSize = 12;

                        //TR1.CharacterFormat.Bold = true;
                        i++;
                        ////
                        FRow.Cells[i].CellFormat.VerticalAlignment = VerticalAlignment.Middle;
                        p1 = FRow.Cells[i].AddParagraph();
                        p1.Format.HorizontalAlignment = HorizontalAlignment.Justify;
                        TR1 = p1.AppendText(item.tenbp);

                        //TR1.CharacterFormat.FontName = "Arial";

                        TR1.CharacterFormat.FontSize = 12;

                        //TR1.CharacterFormat.Bold = true;
                        i++;

                        row_current++;



                        stt++;

                    }
                }




                replacements_table.Add(text, table1);

            }

            if ("Table_DANHGIA" == "Table_DANHGIA")
            {


                var text = "Table_DANHGIA";
                Spire.Doc.Table table1 = section.AddTable(true);
                PreferredWidth width = new PreferredWidth(WidthType.Percentage, 100);
                table1.PreferredWidth = width;

                table1.ResetCells(list1.Count() + 1 + list_nhom.Count(), 6);

                table1.SetColumnWidth(0, 200, CellWidthType.Point);
                //table1.SetColumnWidth(1, 209, CellWidthType.Point);
                table1.SetColumnWidth(5, 50, CellWidthType.Point);

                var row_current = 0;
                TableRow FRow;
                var i = 0;
                Paragraph p1;
                TextRange TR1;
                if (row_current == 0)
                {
                    //// Row 0
                    //Set the first row as table header
                    FRow = table1.Rows[row_current];

                    FRow.IsHeader = true;
                    //Set the height and color of the first row
                    FRow.Height = 30;
                    i = 0;

                    //table1.ApplyHorizontalMerge(row_current, 0, 1);

                    FRow.Cells[i].CellFormat.VerticalAlignment = VerticalAlignment.Middle;
                    p1 = FRow.Cells[i].AddParagraph();
                    p1.Format.HorizontalAlignment = HorizontalAlignment.Center;
                    TR1 = p1.AppendText("Mục");

                    //TR1.CharacterFormat.FontName = "Arial";

                    TR1.CharacterFormat.FontSize = 12;

                    TR1.CharacterFormat.Bold = true;

                    p1 = FRow.Cells[i].AddParagraph();
                    p1.Format.HorizontalAlignment = HorizontalAlignment.Center;
                    TR1 = p1.AppendText("Items");

                    //TR1.CharacterFormat.FontName = "Arial";

                    TR1.CharacterFormat.FontSize = 12;

                    TR1.CharacterFormat.Italic = true;
                    i++;

                    ////

                    FRow.Cells[i].CellFormat.VerticalAlignment = VerticalAlignment.Middle;
                    p1 = FRow.Cells[i].AddParagraph();
                    p1.Format.HorizontalAlignment = HorizontalAlignment.Center;
                    TR1 = p1.AppendText("Đánh giá");

                    //TR1.CharacterFormat.FontName = "Arial";

                    TR1.CharacterFormat.FontSize = 12;

                    TR1.CharacterFormat.Bold = true;

                    p1 = FRow.Cells[i].AddParagraph();
                    p1.Format.HorizontalAlignment = HorizontalAlignment.Center;
                    TR1 = p1.AppendText("Evaluation");

                    //TR1.CharacterFormat.FontName = "Arial";

                    TR1.CharacterFormat.FontSize = 12;

                    TR1.CharacterFormat.Italic = true;
                    i++;



                    FRow.Cells[i].CellFormat.VerticalAlignment = VerticalAlignment.Middle;
                    p1 = FRow.Cells[i].AddParagraph();
                    p1.Format.HorizontalAlignment = HorizontalAlignment.Center;
                    TR1 = p1.AppendText("Số hành động");

                    //TR1.CharacterFormat.FontName = "Arial";

                    TR1.CharacterFormat.FontSize = 12;

                    TR1.CharacterFormat.Bold = true;

                    p1 = FRow.Cells[i].AddParagraph();
                    p1.Format.HorizontalAlignment = HorizontalAlignment.Center;
                    TR1 = p1.AppendText("Action No.");

                    //TR1.CharacterFormat.FontName = "Arial";

                    TR1.CharacterFormat.FontSize = 12;

                    TR1.CharacterFormat.Italic = true;
                    i++;


                    FRow.Cells[i].CellFormat.VerticalAlignment = VerticalAlignment.Middle;
                    p1 = FRow.Cells[i].AddParagraph();
                    p1.Format.HorizontalAlignment = HorizontalAlignment.Center;
                    TR1 = p1.AppendText("Hành động");

                    //TR1.CharacterFormat.FontName = "Arial";

                    TR1.CharacterFormat.FontSize = 12;

                    TR1.CharacterFormat.Bold = true;

                    p1 = FRow.Cells[i].AddParagraph();
                    p1.Format.HorizontalAlignment = HorizontalAlignment.Center;
                    TR1 = p1.AppendText("Activity");

                    //TR1.CharacterFormat.FontName = "Arial";

                    TR1.CharacterFormat.FontSize = 12;

                    TR1.CharacterFormat.Italic = true;
                    i++;

                    FRow.Cells[i].CellFormat.VerticalAlignment = VerticalAlignment.Middle;
                    p1 = FRow.Cells[i].AddParagraph();
                    p1.Format.HorizontalAlignment = HorizontalAlignment.Center;
                    TR1 = p1.AppendText("Trách nhiệm");

                    //TR1.CharacterFormat.FontName = "Arial";

                    TR1.CharacterFormat.FontSize = 12;

                    TR1.CharacterFormat.Bold = true;

                    p1 = FRow.Cells[i].AddParagraph();
                    p1.Format.HorizontalAlignment = HorizontalAlignment.Center;
                    TR1 = p1.AppendText("Responsibility");

                    //TR1.CharacterFormat.FontName = "Arial";

                    TR1.CharacterFormat.FontSize = 12;

                    TR1.CharacterFormat.Italic = true;
                    i++;

                    FRow.Cells[i].CellFormat.VerticalAlignment = VerticalAlignment.Middle;
                    p1 = FRow.Cells[i].AddParagraph();
                    p1.Format.HorizontalAlignment = HorizontalAlignment.Center;
                    TR1 = p1.AppendText("Thời hạn dự kiến");

                    //TR1.CharacterFormat.FontName = "Arial";

                    TR1.CharacterFormat.FontSize = 12;

                    TR1.CharacterFormat.Bold = true;

                    p1 = FRow.Cells[i].AddParagraph();
                    p1.Format.HorizontalAlignment = HorizontalAlignment.Center;
                    TR1 = p1.AppendText("Expected deadline");

                    //TR1.CharacterFormat.FontName = "Arial";

                    TR1.CharacterFormat.FontSize = 12;

                    TR1.CharacterFormat.Italic = true;
                    i++;
                    row_current++;
                }


                if (row_current != 0)
                {
                    var stt = 1;
                    foreach (var item in list_nhom)
                    {
                        //// Row 0
                        //Set the first row as table header
                        FRow = table1.Rows[row_current];

                        //FRow.IsHeader = true;
                        //Set the height and color of the first row
                        //FRow.Height = 30;
                        i = 0;

                        table1.ApplyHorizontalMerge(row_current, 0, 5);

                        FRow.Cells[i].CellFormat.VerticalAlignment = VerticalAlignment.Middle;
                        p1 = FRow.Cells[i].AddParagraph();
                        p1.Format.HorizontalAlignment = HorizontalAlignment.Left;
                        TR1 = p1.AppendText(item.tennhom_muc);

                        //TR1.CharacterFormat.FontName = "Arial";

                        TR1.CharacterFormat.FontSize = 12;

                        TR1.CharacterFormat.Bold = true;

                        TR1 = p1.AppendText(" / ");

                        //TR1.CharacterFormat.FontName = "Arial";

                        TR1.CharacterFormat.FontSize = 12;

                        //TR1.CharacterFormat.Bold = true;

                        var TR2 = p1.AppendText(item.tennhom_en_muc);

                        //TR1.CharacterFormat.FontName = "Arial";

                        TR2.CharacterFormat.FontSize = 12;

                        TR2.CharacterFormat.Italic = true;
                        //TR1.CharacterFormat.Bold = true;
                        //i++;

                        row_current++;

                        foreach (var item1 in item.list)
                        {
                            FRow = table1.Rows[row_current];

                            //FRow.IsHeader = true;
                            //Set the height and color of the first row
                            //FRow.Height = 30;
                            i = 0;

                            //table1.ApplyHorizontalMerge(row_current, 0, 1);
                            table1.ApplyVerticalMerge(0, row_current, row_current + item1.list.Count() - 1);

                            FRow.Cells[i].CellFormat.VerticalAlignment = VerticalAlignment.Middle;
                            p1 = FRow.Cells[i].AddParagraph();
                            p1.Format.HorizontalAlignment = HorizontalAlignment.Left;
                            TR1 = p1.AppendText(item1.ten_muc);

                            //TR1.CharacterFormat.FontName = "Arial";

                            TR1.CharacterFormat.FontSize = 12;

                            //TR1.CharacterFormat.Bold = true;


                            //TR1.CharacterFormat.Bold = true;
                            p1 = FRow.Cells[i].AddParagraph();
                            p1.Format.HorizontalAlignment = HorizontalAlignment.Left;
                            TR1 = p1.AppendText(item1.ten_en_muc);

                            //TR1.CharacterFormat.FontName = "Arial";

                            TR1.CharacterFormat.FontSize = 12;

                            TR1.CharacterFormat.Italic = true;
                            //TR1.CharacterFormat.Bold = true;
                            i++;


                            table1.ApplyVerticalMerge(1, row_current, row_current + item1.list.Count() - 1);
                            FRow.Cells[i].CellFormat.VerticalAlignment = VerticalAlignment.Middle;
                            p1 = FRow.Cells[i].AddParagraph();
                            p1.Format.HorizontalAlignment = HorizontalAlignment.Left;
                            var danhgia_yes = item1.danhgia != null && item1.danhgia != "" ? "☒" : "☐";
                            TR1 = p1.AppendText(danhgia_yes + " Có / ");

                            //TR1.CharacterFormat.FontName = "Arial";

                            TR1.CharacterFormat.FontSize = 12;

                            //TR1.CharacterFormat.Bold = true;
                            TR1 = p1.AppendText("Yes");

                            //TR1.CharacterFormat.FontName = "Arial";

                            TR1.CharacterFormat.FontSize = 12;
                            TR1.CharacterFormat.Italic = true;

                            //TR1.CharacterFormat.Bold = true;
                            p1 = FRow.Cells[i].AddParagraph();
                            p1.Format.HorizontalAlignment = HorizontalAlignment.Left;
                            var danhgia_no = item1.danhgia != null && item1.danhgia != "" ? "☐" : "☒";
                            TR1 = p1.AppendText(danhgia_no + " Không / ");

                            //TR1.CharacterFormat.FontName = "Arial";

                            TR1.CharacterFormat.FontSize = 12;
                            TR1 = p1.AppendText("No");

                            //TR1.CharacterFormat.FontName = "Arial";

                            TR1.CharacterFormat.FontSize = 12;
                            TR1.CharacterFormat.Italic = true;
                            i++;
                            foreach (var item2 in item1.list)
                            {
                                i = 2;
                                FRow = table1.Rows[row_current];
                                //table1.ApplyHorizontalMerge(row_current, 0, 1);
                                p1 = FRow.Cells[i].AddParagraph();
                                p1.Format.HorizontalAlignment = HorizontalAlignment.Center;
                                TR1 = p1.AppendText(item2.sohanhdong);

                                //TR1.CharacterFormat.FontName = "Arial";

                                TR1.CharacterFormat.FontSize = 12;
                                i++;

                                p1 = FRow.Cells[i].AddParagraph();
                                p1.Format.HorizontalAlignment = HorizontalAlignment.Left;
                                TR1 = p1.AppendText(item2.tenhanhdong);

                                //TR1.CharacterFormat.FontName = "Arial";

                                TR1.CharacterFormat.FontSize = 12;
                                i++;

                                p1 = FRow.Cells[i].AddParagraph();
                                p1.Format.HorizontalAlignment = HorizontalAlignment.Left;
                                TR1 = p1.AppendText(item2.trachnhiem);

                                //TR1.CharacterFormat.FontName = "Arial";

                                TR1.CharacterFormat.FontSize = 12;
                                i++;

                                p1 = FRow.Cells[i].AddParagraph();
                                p1.Format.HorizontalAlignment = HorizontalAlignment.Center;
                                TR1 = p1.AppendText(item2.dukien.Value.ToString("dd/MM/yyyy"));

                                //TR1.CharacterFormat.FontName = "Arial";

                                TR1.CharacterFormat.FontSize = 12;
                                i++;
                                row_current++;
                            }

                        }


                        stt++;

                    }
                }




                replacements_table.Add(text, table1);

            }
            if ("Table_ANBAN" == "Table_ANBAN")
            {


                var text = "Table_ANBAN";
                Spire.Doc.Table table1 = section.AddTable(true);
                PreferredWidth width = new PreferredWidth(WidthType.Percentage, 100);
                table1.PreferredWidth = width;

                table1.ResetCells(CHANGECONTROL_ANBAN_IN.Count() + 1, 3);

                table1.SetColumnWidth(0, 46, CellWidthType.Point);
                table1.SetColumnWidth(1, 100, CellWidthType.Point);
                table1.SetColumnWidth(2, 226, CellWidthType.Point);

                var row_current = 0;
                TableRow FRow;
                var i = 0;
                Paragraph p1;
                TextRange TR1;
                if (row_current == 0)
                {
                    //// Row 0
                    //Set the first row as table header
                    FRow = table1.Rows[row_current];

                    FRow.IsHeader = true;
                    //Set the height and color of the first row
                    FRow.Height = 30;
                    i = 0;

                    //table1.ApplyHorizontalMerge(row_current, 0, 1);

                    FRow.Cells[i].CellFormat.VerticalAlignment = VerticalAlignment.Middle;
                    p1 = FRow.Cells[i].AddParagraph();
                    p1.Format.HorizontalAlignment = HorizontalAlignment.Center;
                    TR1 = p1.AppendText("Ấn bản");

                    //TR1.CharacterFormat.FontName = "Arial";

                    TR1.CharacterFormat.FontSize = 12;

                    TR1.CharacterFormat.Bold = true;

                    p1 = FRow.Cells[i].AddParagraph();
                    p1.Format.HorizontalAlignment = HorizontalAlignment.Center;
                    TR1 = p1.AppendText("Version");

                    //TR1.CharacterFormat.FontName = "Arial";

                    TR1.CharacterFormat.FontSize = 12;

                    TR1.CharacterFormat.Italic = true;
                    i++;

                    ////

                    FRow.Cells[i].CellFormat.VerticalAlignment = VerticalAlignment.Middle;
                    p1 = FRow.Cells[i].AddParagraph();
                    p1.Format.HorizontalAlignment = HorizontalAlignment.Center;
                    TR1 = p1.AppendText("Ngày hiệu lực");

                    //TR1.CharacterFormat.FontName = "Arial";

                    TR1.CharacterFormat.FontSize = 12;

                    TR1.CharacterFormat.Bold = true;

                    p1 = FRow.Cells[i].AddParagraph();
                    p1.Format.HorizontalAlignment = HorizontalAlignment.Center;
                    TR1 = p1.AppendText("Effective date");

                    //TR1.CharacterFormat.FontName = "Arial";

                    TR1.CharacterFormat.FontSize = 12;

                    TR1.CharacterFormat.Italic = true;
                    i++;



                    FRow.Cells[i].CellFormat.VerticalAlignment = VerticalAlignment.Middle;
                    p1 = FRow.Cells[i].AddParagraph();
                    p1.Format.HorizontalAlignment = HorizontalAlignment.Center;
                    TR1 = p1.AppendText("Mô tả thay đổi");

                    //TR1.CharacterFormat.FontName = "Arial";

                    TR1.CharacterFormat.FontSize = 12;

                    TR1.CharacterFormat.Bold = true;

                    p1 = FRow.Cells[i].AddParagraph();
                    p1.Format.HorizontalAlignment = HorizontalAlignment.Center;
                    TR1 = p1.AppendText("Description of changes");

                    //TR1.CharacterFormat.FontName = "Arial";

                    TR1.CharacterFormat.FontSize = 12;

                    TR1.CharacterFormat.Italic = true;
                    i++;

                    row_current++;
                }


                if (row_current != 0)
                {
                    var stt = 1;
                    foreach (var item in CHANGECONTROL_ANBAN_IN)
                    {
                        //// Row 0
                        //Set the first row as table header
                        FRow = table1.Rows[row_current];

                        //FRow.IsHeader = true;
                        //Set the height and color of the first row
                        //FRow.Height = 30;
                        i = 0;

                        //table1.ApplyHorizontalMerge(row_current, 0, 1);

                        FRow.Cells[i].CellFormat.VerticalAlignment = VerticalAlignment.Middle;
                        p1 = FRow.Cells[i].AddParagraph();
                        p1.Format.HorizontalAlignment = HorizontalAlignment.Center;
                        TR1 = p1.AppendText(item.anban);

                        //TR1.CharacterFormat.FontName = "Arial";

                        TR1.CharacterFormat.FontSize = 12;

                        //TR1.CharacterFormat.Bold = true;

                        i++;

                        ////

                        FRow.Cells[i].CellFormat.VerticalAlignment = VerticalAlignment.Middle;
                        p1 = FRow.Cells[i].AddParagraph();
                        p1.Format.HorizontalAlignment = HorizontalAlignment.Justify;
                        TR1 = p1.AppendText(item.ngayhieuluc.ToString("dd/MM/yyyy"));

                        //TR1.CharacterFormat.FontName = "Arial";

                        TR1.CharacterFormat.FontSize = 12;

                        //TR1.CharacterFormat.Bold = true;
                        i++;
                        ////
                        FRow.Cells[i].CellFormat.VerticalAlignment = VerticalAlignment.Middle;
                        p1 = FRow.Cells[i].AddParagraph();
                        p1.Format.HorizontalAlignment = HorizontalAlignment.Justify;
                        TR1 = p1.AppendText(item.mota);

                        //TR1.CharacterFormat.FontName = "Arial";

                        TR1.CharacterFormat.FontSize = 12;

                        //TR1.CharacterFormat.Bold = true;
                        TR1 = p1.AppendText(item.mota_en);

                        //TR1.CharacterFormat.FontName = "Arial";

                        TR1.CharacterFormat.FontSize = 12;

                        TR1.CharacterFormat.Italic = true;
                        i++;

                        row_current++;



                        stt++;

                    }
                }




                replacements_table.Add(text, table1);

            }

            string[] fieldName = raw.Keys.ToArray();
            string[] fieldValue = raw.Values.ToArray();

            string[] MergeFieldNames = document.MailMerge.GetMergeFieldNames();
            string[] GroupNames = document.MailMerge.GetMergeGroupNames();

            document.MailMerge.Execute(fieldName, fieldValue);

            foreach (KeyValuePair<string, Spire.Doc.Table> entry in replacements_table)
            {
                TextSelection selection = document.FindString(entry.Key, true, true);
                if (selection == null)
                {
                    var table = entry.Value;
                    //table
                    continue;
                }
                TextRange range = selection.GetAsOneRange();
                Paragraph paragraph = range.OwnerParagraph;
                Body body = paragraph.OwnerTextBody;
                int index = body.ChildObjects.IndexOf(paragraph);
                //var fontsize = paragrap

                body.ChildObjects.Remove(paragraph);
                body.ChildObjects.Insert(index, entry.Value);
            }



            document.SaveToFile("./wwwroot" + documentPath, Spire.Doc.FileFormat.Docx);

            return Json(new { success = true, link = Domain + documentPath, });


        }
        public async Task<JsonResult> baocaosuco(string so, string ngaylap)
        {
            if (so == null || ngaylap == null)
            {
                return Json(new { success = false });
            }

            var viewPath = _configuration["Source:Path_Private"] + "\\qa\\templates\\Suco.docx";
            var documentPath = "/temp/danhgiatacdong_" + DateTime.Now.ToFileTimeUtc() + ".docx";
            string Domain = (HttpContext.Request.IsHttps ? "https://" : "http://") + HttpContext.Request.Host.Value;


            var report = "";


            ///XUẤT PDF

            //Creates Document instance
            Spire.Doc.Document document = new Spire.Doc.Document();
            Section section = document.AddSection();

            document.LoadFromFile(viewPath, Spire.Doc.FileFormat.Docx);



            //document.MailMerge.ExecuteWidthRegion(datatable_details);

            ////Lấy 
            ///

            Dictionary<string, Spire.Doc.Table> replacements_table = new Dictionary<string, Spire.Doc.Table>(StringComparer.OrdinalIgnoreCase) { };

            var sql_SUCO = $"EXECUTE [SUCO_HIENTHI_IN] '{so}','{ngaylap}' ";

            var SUCO_list = _qlsxContext.SUCO.FromSqlRaw($"{sql_SUCO}").ToList();

            var SUCO = SUCO_list.FirstOrDefault();

            var sql_DANHGIA = $"EXECUTE [SUCO_HIENTHI_IN_CHITIET] '{so}' ,'{ngaylap}' ";

            var SUCO_DANHGIA = _qlsxContext.SUCO_DANHGIA.FromSqlRaw($"{sql_DANHGIA}").ToList();

            var sql_HANHDONG = $"EXECUTE [SUCO_HIENTHI_KEHOACH_IN] '{so}' ,'{ngaylap}' ";

            var SUCO_HANHDONG = _qlsxContext.SUCO_HANHDONG.FromSqlRaw($"{sql_HANHDONG}").ToList();


            var tacdongchitiet_list = SUCO_DANHGIA.Select(d => d.tacdongchitiet).ToList();
            var dinhkem_list = SUCO_DANHGIA.Select(d => d.dinhkem).ToList();

            var raw = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
                {
                    {"so",so },
                    {"tenkhuvuc",SUCO.tenkhuvuc ?? ""},
                    {"tieude",SUCO.tieude ?? ""},
                    {"is_suco",SUCO.suco == true ?  "☒" : "☐"},
                    {"is_sailech",SUCO.suco != true ?  "☒" : "☐"},

                    {"suco_baolau",SUCO.suco_baolau ?? "" },
                    {"suco_khinao",SUCO.suco_khinao ?? "" },
                    {"suco_phathien",SUCO.suco_phathien ?? "" },
                    {"suco_odau",SUCO.suco_odau ?? "" },
                    {"suco_anhhuong",SUCO.suco_anhhuong ?? "" },
                    {"suco_anhhuong_sanpham",SUCO.suco_anhhuong_sanpham ?? "" },
                    {"sanpham",SUCO.sanpham ?? "" },
                    {"giaitrinhtre",SUCO.giaitrinhtre ?? "" },
                    {"suco_tacdong",string.Join("\n",tacdongchitiet_list) },
                    {"suco_dinhkem",string.Join("\n",dinhkem_list) },

                    {"is_02",SUCO.mabp_02 != "" && SUCO.mabp_02 != null ? "☒" : "☐" },
                    {"is_05",SUCO.mabp_05 != "" && SUCO.mabp_05 != null ? "☒" : "☐" },
                    {"is_06",SUCO.mabp_06 != "" && SUCO.mabp_06 != null ? "☒" : "☐" },
                    {"is_07",SUCO.mabp_07 != "" && SUCO.mabp_07 != null ? "☒" : "☐" },
                    {"is_09",SUCO.mabp_09 != "" && SUCO.mabp_09 != null ? "☒" : "☐" },
                    {"is_10",SUCO.mabp_10 != "" && SUCO.mabp_10 != null ? "☒" : "☐" },
                    {"is_11",SUCO.mabp_11 != "" && SUCO.mabp_11 != null ? "☒" : "☐" },
                    {"is_13",SUCO.mabp_13 != "" && SUCO.mabp_13 != null ? "☒" : "☐" },
                    {"is_16",SUCO.mabp_16 != "" && SUCO.mabp_16 != null ? "☒" : "☐" },

                };

            ///////TABLE
            if ("TABLE_HANHDONG" == "TABLE_HANHDONG")
            {

                System.Data.DataTable datatable_details = new System.Data.DataTable("hanhdong");
                datatable_details.Columns.Add("stt");
                datatable_details.Columns.Add("noidung");
                datatable_details.Columns.Add("trachnhiem");
                datatable_details.Columns.Add("dukien");
                foreach (var item in SUCO_HANHDONG)
                {
                    var values = new object[4];
                    values[0] = item.stt;
                    values[1] = item.noidung;
                    values[2] = item.trachnhiem;
                    values[3] = item.dukien != null ? item.dukien.Value.ToString("dd/MM/yyyy") : "";
                    datatable_details.Rows.Add(values);
                }

                document.MailMerge.ExecuteWidthRegion(datatable_details);

            }

            if ("TABLE_DANHGIA" == "TABLE_DANHGIA")
            {
                System.Data.DataTable datatable_details = new System.Data.DataTable("danhgia");
                datatable_details.Columns.Add("tenkhuvuc_1");
                datatable_details.Columns.Add("hanhdong");
                datatable_details.Columns.Add("dieutra");
                datatable_details.Columns.Add("tacdong_yes");
                datatable_details.Columns.Add("tacdong_no");
                datatable_details.Columns.Add("giaithich");
                foreach (var item in SUCO_DANHGIA)
                {
                    var values = new object[6];
                    values[0] = item.tenkhuvuc_1;
                    values[1] = item.hanhdong;
                    values[2] = item.dieutra;
                    values[3] = item.tacdong == true ? "☒" : "☐";
                    values[4] = item.tacdong != true ? "☒" : "☐";
                    values[5] = item.giaithich;
                    datatable_details.Rows.Add(values);
                }

                document.MailMerge.ExecuteWidthRegion(datatable_details);

            }

            string[] fieldName = raw.Keys.ToArray();
            string[] fieldValue = raw.Values.ToArray();

            string[] MergeFieldNames = document.MailMerge.GetMergeFieldNames();
            string[] GroupNames = document.MailMerge.GetMergeGroupNames();

            document.MailMerge.Execute(fieldName, fieldValue);



            document.SaveToFile("./wwwroot" + documentPath, Spire.Doc.FileFormat.Docx);

            return Json(new { success = true, link = Domain + documentPath, });


        }

        public async Task<JsonResult> dieutrasuco(int id)
        {
            if (id == null)
            {
                return Json(new { success = false });
            }
            ////Lấy 
            ///

            Dictionary<string, Spire.Doc.Table> replacements_table = new Dictionary<string, Spire.Doc.Table>(StringComparer.OrdinalIgnoreCase) { };

            var sql_SUCO = $"EXECUTE [SUCO_IA_HIENTHI] {id}";

            var SUCO_list = _qlsxContext.DIEUTRASUCO.FromSqlRaw($"{sql_SUCO}").ToList();

            var SUCO = SUCO_list.FirstOrDefault();

            var sql_partA = $"EXECUTE [SUCO_IA_CAPA_KHONG_HIENTHI] {id} ";

            var SUCO_PARTA = _qlsxContext.DIEUTRASUCO_PARTA.FromSqlRaw($"{sql_partA}").ToList();

            var sql_partB = $"EXECUTE [SUCO_IA_CAPA_NANGCAP_HIENTHI] {id} ";

            var SUCO_PARTB = _qlsxContext.DIEUTRASUCO_PARTB.FromSqlRaw($"{sql_partB}").ToList();

            var sql_CAPA = $"EXECUTE [SUCO_IA_CAPA_HIENTHI_IN] {id} ";

            var SUCO_CAPA = _qlsxContext.DIEUTRASUCO_CAPA.FromSqlRaw($"{sql_CAPA}").ToList();


            var viewPath = _configuration["Source:Path_Private"] + "\\qa\\templates\\DieutraSuco_A.docx";
            var documentPath = "/temp/dieutrasuco_" + DateTime.Now.ToFileTimeUtc() + ".docx";
            string Domain = (HttpContext.Request.IsHttps ? "https://" : "http://") + HttpContext.Request.Host.Value;
            if (SUCO.capa != true && SUCO.suco != true)
            {
                viewPath = _configuration["Source:Path_Private"] + "\\qa\\templates\\DieutraSailech_A.docx";
                documentPath = "/temp/dieutrasailech_" + DateTime.Now.ToFileTimeUtc() + ".docx";
            }
            else if (SUCO.capa == true && SUCO.suco != true)
            {
                viewPath = _configuration["Source:Path_Private"] + "\\qa\\templates\\DieutraSailech_B.docx";
                documentPath = "/temp/dieutrasailech_" + DateTime.Now.ToFileTimeUtc() + ".docx";
            }
            else if (SUCO.capa == true && SUCO.suco == true)
            {
                viewPath = _configuration["Source:Path_Private"] + "\\qa\\templates\\DieutraSuco_B.docx";
                documentPath = "/temp/dieutrasuco_" + DateTime.Now.ToFileTimeUtc() + ".docx";
            }

            var report = "";


            ///XUẤT PDF

            //Creates Document instance
            Spire.Doc.Document document = new Spire.Doc.Document();
            Section section = document.AddSection();

            document.LoadFromFile(viewPath, Spire.Doc.FileFormat.Docx);



            //document.MailMerge.ExecuteWidthRegion(datatable_details);


            //var tacdongchitiet_list = SUCO_DANHGIA.Select(d => d.tacdongchitiet).ToList();
            //var dinhkem_list = SUCO_DANHGIA.Select(d => d.dinhkem).ToList();

            var tongdiem = SUCO.tansuat_diem * SUCO.mucdo_diem;
            var raw = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
                {
                    {"so",SUCO.so },
                    {"bienphap",SUCO.bienphap ?? ""},
                    {"ll_yes",SUCO.laplai == true ?  "☒" : "☐"},
                    {"ll_no",SUCO.laplai != true ?  "☒" : "☐"},

                    {"sothamkhao",SUCO.sothamkhao ?? "" },
                    {"mucdo_mota",SUCO.mucdo_mota ?? "" },
                    {"tansuat_mota",SUCO.tansuat_mota ?? "" },
                    {"mucdo_diem",(SUCO.mucdo_diem ?? 1).ToString()},
                    {"tansuat_diem",(SUCO.tansuat_diem ?? 1).ToString()},
                    {"tong_diem",tongdiem.ToString() },
                    {"capa_no",SUCO.capa != true ?  "☒" : "☐" },
                    {"capa_yes",SUCO.capa == true ?  "☒" : "☐" },
                    {"is_nho",tongdiem>=4  && tongdiem<=6 ? "☒" : "☐" },
                    {"is_lon",tongdiem>=7  && tongdiem<=15 ? "☒" : "☐" },
                    {"is_nt",tongdiem>=16  && tongdiem<=50 ? "☒" : "☐" },

                    {"pheduyet_danhgia",SUCO.pheduyet_danhgia ?? "" },
                    {"pheduyet_quyetdinh",SUCO.pheduyet_quyetdinh ?? "" },

                    {"baocaocaptren_no",SUCO.baocaocaptren != true ?  "☒" : "☐" },
                    {"baocaocaptren_yes",SUCO.baocaocaptren == true ?  "☒" : "☐" },

                    {"baocaonsx_no",SUCO.baocaonsx != true ?  "☒" : "☐" },
                    {"baocaonsx_yes",SUCO.baocaonsx == true ?  "☒" : "☐" },
                };

            ///////TABLE
            if ("partA" == "partA")
            {
                raw.Add("nguyennhancotloi", SUCO_PARTA.FirstOrDefault()?.nguyennhancotloi);
                System.Data.DataTable datatable_details = new System.Data.DataTable("partA");
                datatable_details.Columns.Add("stt");
                datatable_details.Columns.Add("hanhdong");
                datatable_details.Columns.Add("tenbp");
                datatable_details.Columns.Add("ngayhoanthanh_dukien");
                var stt = 1;
                foreach (var item in SUCO_PARTA)
                {
                    var values = new object[4];
                    values[0] = stt++;
                    values[1] = item.hanhdong;
                    values[2] = item.tenbp;
                    values[3] = item.ngayhoanthanh_dukien != null ? item.ngayhoanthanh_dukien.Value.ToString("dd/MM/yyyy") : "";
                    datatable_details.Rows.Add(values);
                }

                document.MailMerge.ExecuteWidthRegion(datatable_details);

            }

            if ("partB" == "partB")
            {
                raw.Add("nguyennhancotloi_B", SUCO_PARTB.FirstOrDefault()?.nguyennhancotloi);
                System.Data.DataTable datatable_details = new System.Data.DataTable("partB");
                datatable_details.Columns.Add("stt");
                datatable_details.Columns.Add("hoten");
                datatable_details.Columns.Add("tenbp");
                var stt = 1;
                foreach (var item in SUCO_PARTB)
                {
                    var values = new object[3];
                    values[0] = stt++;
                    values[1] = item.hoten;
                    values[2] = item.tenbp;
                    datatable_details.Rows.Add(values);
                }

                document.MailMerge.ExecuteWidthRegion(datatable_details);

            }
            if ("CAPA" == "CAPA")
            {
                System.Data.DataTable datatable_details = new System.Data.DataTable("capa");
                datatable_details.Columns.Add("socapa");
                datatable_details.Columns.Add("noidung");
                datatable_details.Columns.Add("tenbp");
                datatable_details.Columns.Add("ngayhoanthanh_dukien");
                var stt = 1;
                foreach (var item in SUCO_CAPA)
                {
                    var values = new object[4];
                    values[0] = item.socapa;
                    values[1] = item.noidung;
                    values[2] = item.tenbp;
                    values[3] = item.ngayhoanthanh_dukien != null ? item.ngayhoanthanh_dukien.Value.ToString("dd/MM/yyyy") : "";
                    datatable_details.Rows.Add(values);
                }

                document.MailMerge.ExecuteWidthRegion(datatable_details);

            }
            string[] fieldName = raw.Keys.ToArray();
            string[] fieldValue = raw.Values.ToArray();

            string[] MergeFieldNames = document.MailMerge.GetMergeFieldNames();
            string[] GroupNames = document.MailMerge.GetMergeGroupNames();

            document.MailMerge.Execute(fieldName, fieldValue);



            document.SaveToFile("./wwwroot" + documentPath, Spire.Doc.FileFormat.Docx);

            return Json(new { success = true, link = Domain + documentPath, });


        }
    }


}
