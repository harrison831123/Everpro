using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PolicyNoteShift.Model
{
    public class AutoMailInfo
    {
        /// <summary> 
        /// 程式名稱                              
        /// <summary> 
        public string func_id { get; set; }

        /// <summary> 
        /// 信件主旨                              
        /// <summary> 
        public string mail_title { get; set; }

        /// <summary> 
        ///                               
        /// <summary> 
        public string sender_address { get; set; }

        /// <summary> 
        ///                               
        /// <summary> 
        public string sender_displayname { get; set; }

        /// <summary> 
        ///                               
        /// <summary> 
        public string smtp_address { get; set; }

        /// <summary> 
        ///                               
        /// <summary> 
        public string smtp_credentials { get; set; }

        /// <summary> 
        ///                               
        /// <summary> 
        public string credentials_id { get; set; }

        /// <summary> 
        ///                               
        /// <summary> 
        public string credentials_pwd { get; set; }

        /// <summary> 
        ///                               
        /// <summary> 
        public string aid { get; set; }

        /// <summary> 
        ///                               
        /// <summary> 
        public string adt { get; set; }

        /// <summary> 
        ///                               
        /// <summary> 
        public string cid { get; set; }

        /// <summary> 
        ///                               
        /// <summary> 
        public string cdt { get; set; }

        /// <summary> 
        /// 收件人MailAddress                              
        /// <summary> 
        public string mail_addr_TO { get; set; }

        /// <summary> 
        /// 收件人MailAddress                              
        /// <summary> 
        public string mail_addr_CC { get; set; }
        /// <summary> 
        /// 收件人MailAddress                              
        /// <summary> 
        public string mail_addr_BCC { get; set; }
    }
}
