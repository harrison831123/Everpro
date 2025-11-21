using NLog;
using PolicyNoteShift.Model;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using iTextSharp.text.pdf;
using System.Net.Mail;
using System.Net;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System.Threading;

namespace PolicyNoteShift
{
	public class Process
	{
		private static Logger logger = NLog.LogManager.GetCurrentClassLogger();
		private static DBHelper oDB = new DBHelper();
		private static FileHelper oFh = new FileHelper();
		private static string ResultDir = AppDomain.CurrentDomain.BaseDirectory + "\\Result_Data\\";
		private static string ExtractDir = AppDomain.CurrentDomain.BaseDirectory + "\\Extract_Data\\";
		private static string FileDir = "";
		private static string DateTimeNow = DateTime.Now.ToString("yyyyMMdd");
		private static string BacthSeq = "";
		private static string ResultFileName = string.Empty;

		public Process()
		{
			List<FileTransInfo> olilst = oDB.setModelData();
			for (int i = 0; i < olilst.Count; i++)
			{
				BacthSeq = DateTime.Now.ToString("yyyyMMddHHmmssfff");

				//202411 by Fion 20241017002-調整各家保公照會檔入檔方式-不檢核ExecTime只要保公有資料就處理
				//if (olilst[i].ExecTime != DateTime.Now.ToString("HH") && oDB.ExecEnv != "T")
				//{
				//	continue;
				//}

				logger.Info("CompanyCode：{0}；BatchSeq：{1}；", olilst[i].CompanyCode, BacthSeq);
				CompanyFileProcess(olilst[i]);
				//202411 by Fion 20241017002-調整各家保公照會檔入檔方式-增加等待時間減少網路磁碟機連線錯誤
				logger.Info("Wait Next Company!!");
				Thread.Sleep(60000);
			}
		}

		public void CompanyFileProcess(FileTransInfo filetransinfo)
		{
			string ls_sourcedisk = "";
			string ls_backupdisk = "";
			Boolean bSource = false;
			Boolean bTarget = false;

			try
			{
				List<PbdNoteData> list_ut = new List<PbdNoteData>();				

				// Map Network Drive And Read File To DataTable
				bSource = oFh.MapNetworkDrive(filetransinfo.SourcePath, filetransinfo.SourceAccount, filetransinfo.SourcePassword
														, ref ls_sourcedisk);

				bTarget = oFh.MapNetworkDrive(filetransinfo.BackupPath, filetransinfo.BackupAccount, filetransinfo.BackupPassword
														, ref ls_backupdisk);

				if (bSource && bTarget)
				{
					string execfilename = filetransinfo.ZipFileName + ".zip";

					string[] lsa_filename = System.IO.Directory.GetFiles(ls_sourcedisk, execfilename);

					if (lsa_filename.Length > 0)
					{
						foreach (string ls_readfile in lsa_filename)
						{
							logger.Info("==檔案：" + ls_readfile + "==Start==");
							DataTable ldt_import = oFh.ToDataTable(list_ut);

							string zipPW = "";
							if (filetransinfo.ZipPw.IndexOf("fyyymmdd") > -1)
							{
								zipPW = filetransinfo.ZipPw.Replace("fyyymmdd",
										ls_readfile.Substring(ls_readfile.IndexOf(filetransinfo.ZipFileName.Replace("*", "")) + filetransinfo.ZipFileName.Replace("*", "").Length, 8));
							}
							else if (filetransinfo.ZipPw.IndexOf("yyyymmdd") > -1)
							{
								zipPW = filetransinfo.ZipPw.Replace("yyyymmdd", DateTime.Now.ToString("yyyyMMdd"));
							}
							else
							{
								zipPW = filetransinfo.ZipPw;
							}

							logger.Info("檔案：" + ls_readfile + " 解壓縮");
							var outDir = Path.Combine(ExtractDir, filetransinfo.CompanyCode, zipPW);
							//解壓縮
							oFh.ExtractFiles(Path.Combine(ls_sourcedisk, ls_readfile), outDir, zipPW);

							string[] extract_filename = System.IO.Directory.GetFiles(outDir, filetransinfo.FileName + "." + filetransinfo.FileNameExtension);
							if (extract_filename.Length == 1)
							{
								ldt_import = oFh.ReadFileToDataTable(extract_filename[0], filetransinfo.FileEncoding, filetransinfo.CompanyCode, ldt_import, ls_readfile);
							}
							oDB.TruncateTable(filetransinfo.UserDefineTableName);

							logger.Info("判斷保公來檔txt有沒有寫副檔名pdf，並將PDF檔副檔名統一改成小寫");
							//判斷保公來檔txt有沒有寫副檔名pdf，並將PDF檔副檔名統一改成小寫
							foreach (DataRow row in ldt_import.Rows)
							{
								string pdffileName = row["note_pdf_name"].ToString();
								string[] pdfname = pdffileName.Split('.');
								if (pdfname.Length > 1)
								{
									row["note_pdf_name"] = pdfname[0] + "." + pdfname[1].ToLower();
								}
								else
								{
									row["note_pdf_name"] = row["note_pdf_name"].ToString() + ".pdf";
								}
							}

							logger.Info("檔案轉入DB");
							//檔案轉入DB
							oDB.dtBulkCopy2DT(ldt_import);

							logger.Info("檢核PDF檔是否存在");
							string notePdfNames = "";
							foreach (DataRow row in ldt_import.Rows)
							{
								//檢核PDF檔是否存在
								string pdffileName = row["note_pdf_name"].ToString();
								bool fileExists = File.Exists(Path.Combine(outDir, pdffileName));
								if (!fileExists)
								{
									//item.result_flag = "F";
									//item.result_desc = "該筆資料不存在|";
									notePdfNames = notePdfNames + "," + pdffileName;
								}
							}
							logger.Info("SP檢核照會資料");
							//檢核新契約照會資料
							var list_pndh = oDB.checkPbdNoteData(notePdfNames);
							if (list_pndh.Count != 0)
							{
								logger.Info("檔案搬移&PDF加密");
								foreach (PbdNoteDataHistory item in list_pndh)
								{
									if (item.result_flag != "F")
									{
										var targetDir = string.Empty;
										var targetDir4Agent = string.Empty;
										//string PolicySerial = oDB.GetPolicySerial(item.agent_license_no);
										string comppolicycode = item.company_code + "_" + item.policy_no;
										string noticeCYdate = oDB.WDateToCDate(item.notice_date);

										if (item.note_type == "NB")
										{
											targetDir = Path.Combine(oDB.GetConfigValueByName("FPANoteFileDir"), noticeCYdate, item.policy_serial);
											targetDir4Agent = Path.Combine(oDB.GetConfigValueByName("FPANoteFileDir4Agent"), noticeCYdate, item.policy_serial);
										}
										else
										{
											targetDir = Path.Combine(oDB.GetConfigValueByName("FPANoteFileDir"), noticeCYdate, comppolicycode);
											targetDir4Agent = Path.Combine(oDB.GetConfigValueByName("FPANoteFileDir4Agent"), noticeCYdate, comppolicycode);
										}

										//pdf檔處理
										if (!Directory.Exists(targetDir))
										{
											Directory.CreateDirectory(targetDir);
										}
										if (!Directory.Exists(targetDir4Agent))
										{
											Directory.CreateDirectory(targetDir4Agent);
										}

										//檔案搬移
										//202409 by Fion 全球照會要保書序號為空值，比對保單號碼
										string sourcefile = Path.Combine(outDir, item.note_pdf_name);
										string AgentId = oDB.GetAgentId(item.policy_no, item.content_seq, 1, item.note_type);
										string Agent2Id = oDB.GetAgentId(item.policy_no, item.content_seq, 2, item.note_type);
										AgentId = !String.IsNullOrEmpty(AgentId) ? AgentId.Substring(0, 10) : "";
										string NoteFilePDFName = string.Empty;
										string NoteFile02PDFName = string.Empty;

										if (!String.IsNullOrEmpty(Agent2Id))
										{
											Agent2Id = Agent2Id.Substring(0, 10);
											NoteFilePDFName = item.note_pdf_name;
											NoteFile02PDFName = "02_" + item.note_pdf_name;
											System.IO.File.Copy(sourcefile, Path.Combine(targetDir, item.note_pdf_name), true);

											if (File.Exists(sourcefile))
											{
												try
												{
													//PDF檔案加密
													using (PdfReader reader = new PdfReader(sourcefile))
													{
														using (var os = new FileStream(Path.Combine(targetDir4Agent, NoteFilePDFName), FileMode.Create))
														{
															PdfEncryptor.Encrypt(reader, os, true, AgentId, AgentId, PdfWriter.ALLOW_DEGRADED_PRINTING);
														}
													}
													//PDF檔案加密
													using (PdfReader reader = new PdfReader(sourcefile))
													{
														using (var os = new FileStream(Path.Combine(targetDir4Agent, NoteFile02PDFName), FileMode.Create))
														{
															PdfEncryptor.Encrypt(reader, os, true, Agent2Id, Agent2Id, PdfWriter.ALLOW_DEGRADED_PRINTING);
														}
													}
												}
												catch (Exception ex)
												{
													logger.Error(ex.Message);
													item.result_flag = "F";
													item.result_desc = "PDF檔加密失敗，檔案異常|";
												}
											}
										}
										else
										{
											NoteFilePDFName = item.note_pdf_name;
											System.IO.File.Copy(sourcefile, Path.Combine(targetDir, NoteFilePDFName), true);

											if (File.Exists(sourcefile))
											{
												try
												{	
													//PDF檔案加密
													using (PdfReader reader = new PdfReader(sourcefile))
													{
														using (var os = new FileStream(Path.Combine(targetDir4Agent, NoteFilePDFName), FileMode.Create))
														{
															PdfEncryptor.Encrypt(reader, os, true, AgentId, AgentId, PdfWriter.ALLOW_DEGRADED_PRINTING);
														}
													}
												}
												catch(Exception ex)
												{
													logger.Error(ex.Message);
													item.result_flag = "F";
													item.result_desc = "PDF檔加密失敗，檔案異常|";
												}
											}
										}
									}
									//string[] pdfname = item.note_pdf_name.Split('.');
									//if (pdfname.Length > 1)
									//{
									//	item.note_pdf_name = pdfname[0] + "." + pdfname[1].ToLower();
									//}
									//else
									//{
									//	item.note_pdf_name = item.note_pdf_name + ".pdf";
									//}
								}
								logger.Info("資料匯入PbdNoteDataHistory");
								oDB.InsertPbdNoteDataHistory(list_pndh);
							}
							logger.Info("刪除解壓縮後的pdf & txt");
							//刪除解壓縮後的pdf & txt
							System.IO.Directory.Delete(outDir, true);

							string uploaddir = Path.GetDirectoryName(ls_backupdisk + Path.Combine(filetransinfo.BackupFolder));
							if (!Directory.Exists(uploaddir))
								Directory.CreateDirectory(uploaddir);

							//Move原始檔至指定備份位置
							if (File.Exists(ls_backupdisk + Path.Combine(filetransinfo.BackupFolder, ls_readfile.Replace(ls_sourcedisk, ""))))
								File.Delete(ls_backupdisk + Path.Combine(filetransinfo.BackupFolder, ls_readfile.Replace(ls_sourcedisk, "")));
							logger.Info("Move原始檔:" + ls_readfile + "至指定備份位置:" + ls_backupdisk + Path.Combine(filetransinfo.BackupFolder, ls_readfile.Replace(ls_sourcedisk, "")));
							//Move原始檔至指定備份位置
							System.IO.File.Move(ls_readfile, ls_backupdisk + Path.Combine(filetransinfo.BackupFolder, ls_readfile.Replace(ls_sourcedisk, "")));

							MailProcess(filetransinfo, list_pndh);

							logger.Info("新增批次Log檔Layout");
							oDB.InsertBatchJobExecLog(filetransinfo.FuncId, ls_readfile);

							logger.Info("==檔案：" + ls_readfile + "==END==");
						}

					}
				}
			}
			catch (Exception ex)
			{
				logger.Error(ex.StackTrace);
				logger.Error(ex.Message);
				SendLogMail(filetransinfo.FuncDesc);
			}
			finally
			{
				if (bTarget)
					oFh.UnmapNetworkDrive(ls_backupdisk);
				if (bSource)
					oFh.UnmapNetworkDrive(ls_sourcedisk);
			}
		}

		public void MailProcess(FileTransInfo filetransinfo, List<PbdNoteDataHistory> list_pndh)
		{		
			int SuccessCnt = 0, FailedCnt = 0;
			var list_er = new List<ExcelReport>();

			foreach (var item in list_pndh)
			{
				var ER = new ExcelReport
				{
					note_type = item.note_type,
					po_serial = item.po_serial,
					policy_no = item.policy_no,
					notice_date = item.notice_date,
					replay_date = item.replay_date,
					content_seq = item.content_seq,
					agent_license_no = item.agent_license_no,
					note_pdf_name = item.note_pdf_name,
					zipfile_name = item.zipfile_name,
					company_code = item.company_code,
					result_flag = item.result_flag,
					result_desc = item.result_desc
				};
				list_er.Add(ER);
				if (item.result_flag == "Y")
				{
					SuccessCnt += 1;
				}
				else
				{
					FailedCnt += 1;
				}
			}
			
			FileDir = ResultDir + filetransinfo.CompanyCode + "\\";
			ResultFileName = DateTime.Now.ToString("yyyyMMdd") + "{0}照會資料.xlsx";
			ResultFileName = String.Format(ResultFileName, "-" + filetransinfo.CompanyCode);
			if (!Directory.Exists(FileDir))
			{
				try
				{
					Directory.CreateDirectory(FileDir);
					logger.Info(string.Format("成功建立目錄: {0}。", FileDir));
				}
				catch (Exception ex)
				{
					logger.Error(string.Format("建立目錄: {0} 失敗: {1}", FileDir, ex.Message));
				}
			}
			else
			{
				string[] dirfile = Directory.GetFiles(FileDir);
				foreach (var item in dirfile)
				{
					if (File.GetCreationTime(item).ToString("yyyyMMdd") != DateTime.Now.ToString("yyyyMMdd"))
					{
						File.Delete(item);
						logger.Info(string.Format("刪除舊資料: {0}。", item));
					}
				}
			}


			var dataset_pndh = oFh.ToDataSet(list_er);
			logger.Info("生成Excel附件");
			RenderDataToExcel(dataset_pndh,  FileDir);

			//Mail通知
			logger.Info("Mail通知");
			SendMail(filetransinfo.FuncId, list_er, SuccessCnt, FailedCnt);
		}

		public void SendMail(string funcid, List<ExcelReport> list, int SuccessCnt, int FailedCnt)
		{
			SmtpClient client = new SmtpClient();
			try
			{
				AutoMailInfo mailinfo = oDB.getMailInfo(funcid);

				String ls_mailbody = DateTime.Now.ToString("yyyy/MM/dd") + "處理結果";
				ls_mailbody += string.Format("<table border=\"1\">" +
												"<tr bgcolor=\"#58D3F7\"><td>本日更新共：{0} 筆</td></tr>" +
												"<tr bgcolor=\"#F5A9A9\"><td>比對成功共：{1} 筆</td></tr>" +
												"<tr bgcolor=\"#FFFF00\"><td>比對不成功共：{2} 筆</td></tr>" +
												"</table>"
								, list.Count, SuccessCnt, FailedCnt);

				using (MailMessage lo_mm = new MailMessage())
				{
					//lo_mm.From = new MailAddress(mailinfo.sender_address, mailinfo.sender_displayname, System.Text.Encoding.UTF8);
					lo_mm.From = new MailAddress(mailinfo.sender_address, mailinfo.sender_displayname);
					if (!string.IsNullOrEmpty(mailinfo.mail_addr_TO))
					{
						string[] mailto = mailinfo.mail_addr_TO.Split(new char[] { (';') }, StringSplitOptions.RemoveEmptyEntries);
						for (int i = 0; i < mailto.Length; i++)
						{
							lo_mm.To.Add(new MailAddress(mailto[i].ToString().Split(new char[] { ('|') }, StringSplitOptions.RemoveEmptyEntries)[0]
										, mailto[i].ToString().Split(new char[] { ('|') }, StringSplitOptions.RemoveEmptyEntries)[1]));
						}
					}
					else
					{
						lo_mm.To.Add(new MailAddress("harrison831123@mail.everprobks.com.tw"));
					}

					if (!string.IsNullOrEmpty(mailinfo.mail_addr_CC))
					{
						string[] mailcc = mailinfo.mail_addr_CC.Split(new char[] { (';') }, StringSplitOptions.RemoveEmptyEntries);
						for (int i = 0; i < mailcc.Length; i++)
						{
							lo_mm.CC.Add(new MailAddress(mailcc[i].ToString().Split(new char[] { ('|') }, StringSplitOptions.RemoveEmptyEntries)[0]
										, mailcc[i].ToString().Split(new char[] { ('|') }, StringSplitOptions.RemoveEmptyEntries)[1]));
						}
					}

					if (!string.IsNullOrEmpty(mailinfo.mail_addr_BCC))
					{
						string[] mailbcc = mailinfo.mail_addr_BCC.Split(new char[] { (';') }, StringSplitOptions.RemoveEmptyEntries);
						for (int i = 0; i < mailbcc.Length; i++)
						{
							lo_mm.Bcc.Add(new MailAddress(mailbcc[i].ToString().Split(new char[] { ('|') }, StringSplitOptions.RemoveEmptyEntries)[0]
										, mailbcc[i].ToString().Split(new char[] { ('|') }, StringSplitOptions.RemoveEmptyEntries)[1]));
						}
					}

					lo_mm.Subject = mailinfo.mail_title;
					if (oDB.ExecEnv == "T")
					{
						lo_mm.Subject = "[TEST]" + mailinfo.mail_title;
					}
					lo_mm.SubjectEncoding = System.Text.Encoding.UTF8;
					lo_mm.Body = ls_mailbody;
					lo_mm.BodyEncoding = System.Text.Encoding.UTF8;
					lo_mm.IsBodyHtml = true;
					lo_mm.Priority = System.Net.Mail.MailPriority.High;
					lo_mm.Attachments.Add(new Attachment(FileDir + ResultFileName));
					client.Host = mailinfo.smtp_address; //設定smtp Server
					client.Port = 587; //設定Port
					client.UseDefaultCredentials = true;
					client.EnableSsl = true; //gmail預設開啟驗證

					// Check Sender Mail ID and PWD
					if (mailinfo.smtp_credentials == "Y")
					{
						client.Credentials = new NetworkCredential(mailinfo.credentials_id, mailinfo.credentials_pwd);
					}
					client.Send(lo_mm);
				}
			}
			catch (Exception ex)
			{
				logger.Error(ex.Message);
			}
			finally
			{
				client.Dispose();
			}
		}

		public void SendLogMail(string FuncDesc)
		{
			SmtpClient client = new SmtpClient();

			try
			{
				AutoMailInfo mailinfo = oDB.getMailInfo("pbd_policynote_update_Sysinfo");
				var logfile = (LogManager.Configuration.FindTargetByName("File") as NLog.Targets.FileTarget).FileName
					.Render(new LogEventInfo() { TimeStamp = DateTime.Now, LoggerName = "loggerName" });

				using (MailMessage lo_mm = new MailMessage())
				{
					lo_mm.From = new MailAddress(mailinfo.sender_address, mailinfo.sender_displayname);
					if (!string.IsNullOrEmpty(mailinfo.mail_addr_TO))
					{
						string[] mailto = mailinfo.mail_addr_TO.Split(new char[] { (';') }, StringSplitOptions.RemoveEmptyEntries);
						for (int i = 0; i < mailto.Length; i++)
						{
							lo_mm.To.Add(new MailAddress(mailto[i].ToString().Split(new char[] { ('|') }, StringSplitOptions.RemoveEmptyEntries)[0]
										, mailto[i].ToString().Split(new char[] { ('|') }, StringSplitOptions.RemoveEmptyEntries)[1]));
						}
					}
					else
					{
						lo_mm.To.Add(new MailAddress("harrison831123@mail.everprobks.com.tw"));
					}

					if (!string.IsNullOrEmpty(mailinfo.mail_addr_CC))
					{
						string[] mailcc = mailinfo.mail_addr_CC.Split(new char[] { (';') }, StringSplitOptions.RemoveEmptyEntries);
						for (int i = 0; i < mailcc.Length; i++)
						{
							lo_mm.CC.Add(new MailAddress(mailcc[i].ToString().Split(new char[] { ('|') }, StringSplitOptions.RemoveEmptyEntries)[0]
										, mailcc[i].ToString().Split(new char[] { ('|') }, StringSplitOptions.RemoveEmptyEntries)[1]));
						}
					}

					if (!string.IsNullOrEmpty(mailinfo.mail_addr_BCC))
					{
						string[] mailbcc = mailinfo.mail_addr_BCC.Split(new char[] { (';') }, StringSplitOptions.RemoveEmptyEntries);
						for (int i = 0; i < mailbcc.Length; i++)
						{
							lo_mm.Bcc.Add(new MailAddress(mailbcc[i].ToString().Split(new char[] { ('|') }, StringSplitOptions.RemoveEmptyEntries)[0]
										, mailbcc[i].ToString().Split(new char[] { ('|') }, StringSplitOptions.RemoveEmptyEntries)[1]));
						}
					}

					lo_mm.Subject = FuncDesc + "_" + mailinfo.mail_title + "";
					if (oDB.ExecEnv == "T")
					{
						lo_mm.Subject = "[TEST]" + FuncDesc + "_" + mailinfo.mail_title;
					}
					lo_mm.SubjectEncoding = System.Text.Encoding.UTF8;
					lo_mm.BodyEncoding = System.Text.Encoding.UTF8;
					lo_mm.IsBodyHtml = true;
					lo_mm.Priority = System.Net.Mail.MailPriority.High;
					lo_mm.Attachments.Add(new Attachment(logfile));
					client.Host = mailinfo.smtp_address; //設定smtp Server
					client.Port = 587; //設定Port
					client.UseDefaultCredentials = true;
					client.EnableSsl = true; //gmail預設開啟驗證

					// Check Sender Mail ID and PWD
					if (mailinfo.smtp_credentials == "Y")
					{
						client.Credentials = new NetworkCredential(mailinfo.credentials_id, mailinfo.credentials_pwd);
					}
					client.Send(lo_mm);
				}
			}
			catch (Exception ex)
			{
				logger.Error(ex.Message);
			}
			finally
			{
				client.Dispose();
			}
		}

		public void RenderDataToExcel(DataSet _ds, string filedir)
		{
			MemoryStream _msexcel = new MemoryStream();
			IWorkbook _wb = new XSSFWorkbook();
			ISheet _sheet;
			IRow _row;
			Int32 _rownumber;
			ICell _cell;
			ICellStyle _borderstyle = (ICellStyle)_wb.CreateCellStyle();
			ICellStyle _green = (ICellStyle)_wb.CreateCellStyle();
			ICellStyle _red = (ICellStyle)_wb.CreateCellStyle();
			ICellStyle _headstyle = (ICellStyle)_wb.CreateCellStyle();
			IFont _font = _wb.CreateFont();
			string[] colname = { "照會類別", "要保書受理序號", "保單號碼", "照會日期", "照會回覆期限", "資料內容序號", "業務員登錄字號"
					, "附檔PDF檔名", "原始壓縮檔名","保險公司代碼","轉檔資料檢核處理結果","轉檔資料異常描述" };

			try
			{
				// 資料的格式：有框線和置中
				_headstyle.BorderBottom = BorderStyle.Thin;
				_headstyle.BorderLeft = BorderStyle.Thin;
				_headstyle.BorderRight = BorderStyle.Thin;
				_headstyle.BorderTop = BorderStyle.Thin;
				_headstyle.Alignment = HorizontalAlignment.Center;
				_headstyle.FillForegroundColor = 47;
				_headstyle.FillPattern = FillPattern.SolidForeground;

				// 資料的格式：有框線和置中
				_borderstyle.BorderBottom = BorderStyle.Thin;
				_borderstyle.BorderLeft = BorderStyle.Thin;
				_borderstyle.BorderRight = BorderStyle.Thin;
				_borderstyle.BorderTop = BorderStyle.Thin;
				_borderstyle.Alignment = HorizontalAlignment.Center;

				// 綠色字體靠左
				_green.Alignment = HorizontalAlignment.Left;
				_font.IsBold = true;
				_font.Color = NPOI.HSSF.Util.HSSFColor.Green.Index;
				_green.SetFont(_font);

				// 紅色字體靠右
				_font = _wb.CreateFont();
				_red.Alignment = HorizontalAlignment.Left;
				_font.IsBold = true;
				_font.Color = NPOI.HSSF.Util.HSSFColor.Red.Index;
				_red.SetFont(_font);
				_rownumber = 0;
				_sheet = _wb.CreateSheet("照會傳輸資訊檔回饋");

				_row = _sheet.CreateRow(_rownumber);
				_cell = _row.CreateCell(0);
				_cell.SetCellValue("更新結果:");
				_cell.CellStyle = _green;
				_rownumber++;

				_row = _sheet.CreateRow(_rownumber);
				int j = 0;
				foreach (DataColumn _col in _ds.Tables[0].Columns)
				{
					_cell = _row.CreateCell(_col.Ordinal);
					_cell.SetCellValue(colname[j]);
					_cell.CellStyle = _headstyle;
					j++;
				}
				_rownumber++;

				for (int i = 0; i < _ds.Tables[0].Rows.Count; i++)
				{
					_row = _sheet.CreateRow(_rownumber);
					int icolumn = 0;
					foreach (DataColumn _col in _ds.Tables[0].Columns)
					{
						ICell _cells = _row.CreateCell(_col.Ordinal);
						_cells.CellStyle = _borderstyle;
						if (_col.DataType.FullName == "System.String")
						{
							_cells.SetCellType(NPOI.SS.UserModel.CellType.String);
							_cells.SetCellValue(_ds.Tables[0].Rows[i][_col].ToString());
						}
						else if (_col.DataType.FullName == "System.Int32")
						{
							_cells.SetCellType(NPOI.SS.UserModel.CellType.Numeric);
							_cells.SetCellValue(Convert.ToInt32(_ds.Tables[0].Rows[i][_col]));
						}
						else
						{
							_cells.SetCellType(NPOI.SS.UserModel.CellType.String);
							_cells.SetCellValue(_ds.Tables[0].Rows[i][_col].ToString());
						}
						_sheet.AutoSizeColumn(icolumn);
						icolumn++;
					}
					_rownumber++;
				}
				_rownumber++;
			}
			catch (Exception ex)
			{
				logger.Error(ex.Message);
				throw;
			}

			FileStream sw = File.Create(filedir + ResultFileName);
			_wb.Write(sw);
			sw.Dispose();
			sw.Close();
		}
	}
}
