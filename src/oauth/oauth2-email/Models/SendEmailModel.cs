using Microsoft.Graph;

namespace oauth2_email.Models
{
    public class SendEmailModel
    {
        public string? SignedInAs { get; set; }

        public bool IsSignedIn { get; set; }

        public string? To { get; set; }

        public string? Subject { get; set; }

        public string? Body { get; set; }

        public bool IsSent { get; set; }

        public List<Message> Emails { get; set; }
    }
}
