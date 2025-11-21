using System.Collections.Generic;
using System.Xml.Serialization;

namespace MailReportProcess.Model
{
    [XmlRoot(ElementName = "RptDetail")]
    public class RptDetail
    {

        [XmlAttribute(AttributeName = "rpt_seq")]
        public int RptSeq { get; set; }

        [XmlAttribute(AttributeName = "rpt_name")]
        public string RptName { get; set; }

        [XmlAttribute(AttributeName = "output_flag")]
        public string OutputFlag { get; set; }

    }

    [XmlRoot(ElementName = "RptInfo")]
	public class RptInfo
	{
        [XmlElement(ElementName = "RptDetail")]
        public List<RptDetail> RptDetail { get; set; }

        [XmlAttribute(AttributeName = "vaild_flag")]
		public string VaildFlag { get; set; }

        [XmlAttribute(AttributeName = "execute_type")]
        public string ExecuteType { get; set; }

        //執行日
        [XmlAttribute(AttributeName = "execute_day")]
        public string ExecuteDay { get; set; }

        [XmlAttribute(AttributeName = "execute_time")]
        public string ExecuteTime { get; set; }

        [XmlAttribute(AttributeName = "func_id")]
        public string FuncId { get; set; }

        [XmlAttribute(AttributeName = "title")]
        public string Title { get; set; }

        [XmlAttribute(AttributeName = "content1")]
        public string Content1 { get; set; }

        [XmlAttribute(AttributeName = "file_type")]
        public string FileType { get; set; }

        [XmlAttribute(AttributeName = "file_name")]
        public string FileName { get; set; }

        [XmlAttribute(AttributeName = "zip_name")]
        public string ZipName { get; set; }

        [XmlAttribute(AttributeName = "zip_flag")]
        public string ZipFlag { get; set; }

        [XmlAttribute(AttributeName = "zip_pw")]
        public string ZipPw { get; set; }

        [XmlAttribute(AttributeName = "db_conn")]
        public string DbConn { get; set; }

        [XmlAttribute(AttributeName = "sqlstring")]
        public string Sqlstring { get; set; }

        //sql中用來判斷的欄位，報表不產生
        [XmlAttribute(AttributeName = "rpt_key")]
        public string RptKey { get; set; }
    }

	[XmlRoot(ElementName = "MailRpt")]
	public class MailRpt
	{

		[XmlElement(ElementName = "RptInfo")]
        public List<RptInfo> RptInfo { get; set; }
	}
}
