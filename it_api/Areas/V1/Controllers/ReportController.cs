

using iText.Commons.Bouncycastle.Asn1.X509;
using Microsoft.AspNetCore.Identity;
using Microsoft.AspNetCore.Mvc;
using Microsoft.CodeAnalysis;
using Microsoft.EntityFrameworkCore;
using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Doc.Fields;
using System.Data;
using Vue.Data;
using Vue.Models;

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
                    {"vi2",first.vi2.Value.ToString("#,##0").Replace(",",".") },
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
        public async Task<JsonResult> phieuCOA(string solenh)
        {
            if (solenh == null)
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

            var sqla = $"EXECUTE [QC_COA_IN_2] '{solenh}' ";

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
                    {"sophieu",solenh },
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
            table1.SetColumnWidth(0, 30, CellWidthType.Point);

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

                    if (child.Count() > 0 && (item.tieuchuan != null && item.tieuchuan.Trim() != ""))
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
    }


}
