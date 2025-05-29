using System.Collections;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Text.Json.Serialization;
namespace Vue.Models
{
    public class LoginResponse
    {
        [Key]
        public bool authed { get; set; }

        public string? error { get; set; }
        public string? parameter { get; set; }

        public string? session { get; set; }
        public UserInfo? user { get; set; }
        public string? token { get; set; }
    }
    public class UserInfo
    {
        [Key]
        public string id { get; set; }
        public string email { get; set; }

        public string FullName { get; set; }

        public string image_url { get; set; }
        public string image_sign { get; set; }
        public string key_private { get; set; }
        public string? report_for { get; set; }
        public bool is_sign
        {
            get
            {
                if (image_sign == "/private/images/tick.png")
                {
                    return true;
                }
                return false;
            }
        }
        public bool is_truongbophan { get; set; }
        public DateTime last_updated { get; set; }
        public IList<string>? roles { get; set; }
    }
}