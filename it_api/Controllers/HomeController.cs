
using Vue.Models;
using Microsoft.AspNetCore.Mvc;
using Vue.Data;
using System.Net.Mail;
using Vue.Services;
using System.Net.Mime;
using Microsoft.EntityFrameworkCore.Infrastructure;
using System.Diagnostics;

namespace Vue.Controllers
{

	public class HomeController : Controller
	{
		protected readonly ItContext _context;
		private readonly ViewRender _view;


		public HomeController(ItContext context, ViewRender view)
		{
			_context = context;
			_view = view;
            var listener = _context.GetService<DiagnosticSource>();
            (listener as DiagnosticListener).SubscribeWithAdapter(new CommandInterceptor());
        }

		public JsonResult Index()
		{
			return Json(new { test = 1, message = DateTime.Now });

		}

		private SuccesMail SendMail(string to, string subject, string body, List<string> attachments)
		{
			try
			{

				string[] list_to = to.Split(",");

				MailMessage message = new MailMessage();
				message.From = new MailAddress("pymepharco.mail@gmail.com", "Pymepharco System");
				//message.From = new MailAddress("daolytran@pymepharco.com", "Pymepharco System");
				foreach (string str in list_to)
				{
					message.To.Add(new MailAddress(str));
				}
				message.Subject = subject;
				message.Body = body;
				message.BodyEncoding = System.Text.Encoding.UTF8;
				message.SubjectEncoding = System.Text.Encoding.UTF8;
				message.IsBodyHtml = true;
				foreach (var attach in attachments)
				{
					// Create  the file attachment for this email message.
					Attachment data = new Attachment("." + attach, MediaTypeNames.Application.Pdf);
					// Add time stamp information for the file.
					ContentDisposition disposition = data.ContentDisposition;
					disposition.CreationDate = System.IO.File.GetCreationTime(attach);
					disposition.ModificationDate = System.IO.File.GetLastWriteTime(attach);
					disposition.ReadDate = System.IO.File.GetLastAccessTime(attach);
					// Add the file attachment to this email message.

					message.Attachments.Add(data);
				}

				SmtpClient client = new SmtpClient("smtp.gmail.com", 587);
				//SmtpClient client = new SmtpClient("mail.pymepharco.com", 993);
				client.EnableSsl = true;
				client.UseDefaultCredentials = false;
				client.Credentials = new System.Net.NetworkCredential("pymepharco.mail@gmail.com", "xenrezrhmvueqmvw");
				//client.Credentials = new System.Net.NetworkCredential("daolytran@pymepharco.com", "Asd12345");
				client.Send(message);
			}
			catch (Exception ex)
			{
				return new SuccesMail { ex = ex, success = 0 };
			}
			return new SuccesMail { success = 1 };
		}


	}
	class SuccesMail
	{
		public int success { get; set; }
		public Exception ex { get; set; }
	}
}
