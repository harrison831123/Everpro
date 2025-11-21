using Newtonsoft.Json.Linq;
using System;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
/// <summary>
/// 共用模組
/// 需求單
/// 
/// </summary>
namespace SACTAPI.Utilities
{
    public class CryHelper
    {

        public static string AESkey = "4ob6oHsr@r5&W!0sX9uq3x8uYUg5EsP$"; 
        public static string AESiv = "GcH73Jpyyy6RvMK0"; 

        /// <summary>
        /// 加密處理
        /// </summary>
        /// <param name="text">明文</param>
        /// <returns>密文(base64)</returns>
        public static string EncryptAES(string text) =>
            EncryptAesBase64(text, Encoding.UTF8.GetBytes(AESkey), Encoding.UTF8.GetBytes(AESiv), true);

        /// <summary>
        /// 解密處理
        /// </summary>
        /// <param name="text">密文(base64)</param>
        /// <returns>明文</returns>
        public static string DecryptAES(string text) =>
            DecryptAesBase64(text, Encoding.UTF8.GetBytes(AESkey), Encoding.UTF8.GetBytes(AESiv), true);


        /// <summary>
        /// 檔案的md5 checksum值
        /// </summary>
        /// <param name="fileName">檔案路徑</param>
        /// <returns></returns>
        public static string CalculateMD5(string fileName)
        {
            using (var md5 = MD5.Create())
            {
                using (var stream = File.OpenRead(fileName))
                {
                    var hash = md5.ComputeHash(stream);
                    return BitConverter.ToString(hash).Replace("-", "").ToLowerInvariant();
                }
            }
        }

        /// <summary>
        /// base64字串轉成pdf
        /// </summary>
        /// <param name="base64BinaryStr">base64字串</param>
        /// <param name="pdfFilePath">轉成pdf的路徑</param>
        /// <returns></returns>
        public static bool Base64StringToPdf(string base64BinaryStr, string pdfFilePath)
        {
            bool blret = false;

            byte[] byteArray = Convert.FromBase64String(base64BinaryStr);

            using (FileStream fs = File.Create(pdfFilePath))
            {
                fs.Write(byteArray, 0, byteArray.Length);
            }
           
            blret = true;
            return blret;
        }

        /// <summary>
        ///  pdf轉成base64字串
        /// </summary>
        /// <param name="pdfFilePath"> pdf路徑</param>
        /// <returns></returns>
        public static string PdfToBase64String(string pdfFilePath)
        {
            byte[] pdfBytes = File.ReadAllBytes(pdfFilePath);
            string pdfBase64 = Convert.ToBase64String(pdfBytes);

            return pdfBase64;
        }

        /// <summary>
        /// AES加密成base64字串
        /// </summary>
        /// <param name="text">明文</param>
        /// <param name="key">KEY</param>
        /// <param name="iv">IV</param>
        /// <param name="normalize">是否需要正規化資料</param>
        /// <returns>base64字串的密文</returns>
        public static string EncryptAesBase64(string text, byte[] key, byte[] iv, bool normalize = false)
        {
            var sourceBytes = Encoding.UTF8.GetBytes(text);
            using (var aes = Aes.Create())
            {
                aes.Mode = CipherMode.CBC;
                aes.Padding = PaddingMode.PKCS7;
                aes.Key = key;
                aes.IV = iv;
                using (var transform = aes.CreateEncryptor())
                {
                    string base64 = Convert.ToBase64String(transform.TransformFinalBlock(sourceBytes, 0, sourceBytes.Length));
                    if (normalize)
                        base64 = base64.Replace(" ", "+").Replace("_", "/").Replace("-", "+");
                    return base64;
                }
            }
        }

        /// <summary>
        /// AES解密base64字串為明文字串
        /// </summary>
        /// <param name="base64">base64密文</param>
        /// <param name="key">KEY</param>
        /// <param name="iv">IV</param>
        /// <param name="normalize">是否需要正規化資料</param>
        /// <returns>明文字串</returns>
        public static string DecryptAesBase64(string base64, byte[] key, byte[] iv, bool normalize = false)
        {
            if (normalize)
                base64 = base64.Replace(" ", "+").Replace("_", "/").Replace("-", "+");

            byte[] encryptedBytes = Convert.FromBase64String(base64);
            using (var aes = Aes.Create())
            {
                aes.Mode = CipherMode.CBC;
                aes.Padding = PaddingMode.PKCS7;
                aes.Key = key;
                aes.IV = iv;
                using (var transform = aes.CreateDecryptor())
                {
                    return Encoding.UTF8.GetString(transform.TransformFinalBlock(encryptedBytes, 0, encryptedBytes.Length));
                }
            }
        }
    }
}