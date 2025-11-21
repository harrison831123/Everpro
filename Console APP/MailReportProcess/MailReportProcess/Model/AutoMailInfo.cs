using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MailReportProcess.Model
{
    public class AutoMailInfo
    {
        /// <summary> 
        /// 程式名稱                              
        /// <summary> 
        public string FuncId { get; set; }

        /// <summary> 
        /// 信件主旨                              
        /// <summary> 
        public string MailTitle { get; set; }

        /// <summary> 
        ///                               
        /// <summary> 
        public string SenderAddress { get; set; }

        /// <summary> 
        ///                               
        /// <summary> 
        public string SenderDisplayname { get; set; }

        /// <summary> 
        ///                               
        /// <summary> 
        public string SmtpAddress { get; set; }

        /// <summary> 
        ///                               
        /// <summary> 
        public int SmtpPort { get; set; }

        /// <summary> 
        ///                               
        /// <summary> 
        public string SmtpCredentials { get; set; }

        /// <summary> 
        ///                               
        /// <summary> 
        public string CredentialsId { get; set; }

        /// <summary> 
        ///                               
        /// <summary> 
        public string CredentialsPwd { get; set; }

        /// <summary> 
        ///                               
        /// <summary> 
        public string Aid { get; set; }

        /// <summary> 
        ///                               
        /// <summary> 
        public string Adt { get; set; }

        /// <summary> 
        ///                               
        /// <summary> 
        public string Cid { get; set; }

        /// <summary> 
        ///                               
        /// <summary> 
        public string cdt { get; set; }

        /// <summary> 
        /// 收件人MailAddress                              
        /// <summary> 
        public string v { get; set; }

        /// <summary> 
        /// 收件人MailAddress                              
        /// <summary> 
        public string MaillAddrCc{ get; set; }
        /// <summary> 
        /// 收件人MailAddress                              
        /// <summary> 
        public string MaillAddrBcc { get; set; }
    }
}
