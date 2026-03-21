using System;
using System.Collections.Generic;
using System.IO;
using System.Security.Cryptography;
using System.Text;
using System.Web.Script.Serialization;
using StockManager.Config;

namespace StockManager
{
    public partial class MainWindow
    {
        private Tuple<string, string> LoadCapitalCredential()
        {
            try
            {
                if (!File.Exists(_capitalCredentialFile))
                {
                    return null;
                }

                var protectedBytes = File.ReadAllBytes(_capitalCredentialFile);
                var bytes = DecryptCapitalCredentialBytes(protectedBytes);
                var json = Encoding.UTF8.GetString(bytes);
                var serializer = new JavaScriptSerializer();
                var data = serializer.Deserialize<Dictionary<string, string>>(json);
                if (data == null)
                {
                    return null;
                }

                string id;
                string pwd;
                if (!data.TryGetValue("login_id", out id) || !data.TryGetValue("password", out pwd))
                {
                    return null;
                }

                if (string.IsNullOrWhiteSpace(id) || string.IsNullOrWhiteSpace(pwd))
                {
                    return null;
                }

                return Tuple.Create(id, pwd);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[群益登入記憶讀取失敗] {ex.Message}");
                return null;
            }
        }

        private void SaveCapitalCredential(string loginId, string password)
        {
            try
            {
                if (!Directory.Exists(AppConfig.UserConfigDir))
                {
                    Directory.CreateDirectory(AppConfig.UserConfigDir);
                }

                var serializer = new JavaScriptSerializer();
                var payload = serializer.Serialize(new Dictionary<string, string>
                {
                    { "login_id", loginId ?? string.Empty },
                    { "password", password ?? string.Empty }
                });
                var bytes = Encoding.UTF8.GetBytes(payload);
                var protectedBytes = EncryptCapitalCredentialBytes(bytes);
                File.WriteAllBytes(_capitalCredentialFile, protectedBytes);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[群益登入記憶儲存失敗] {ex.Message}");
            }
        }

        private byte[] EncryptCapitalCredentialBytes(byte[] plain)
        {
            var entropy = (Environment.UserName + "|" + Environment.MachineName + "|StockManager");
            var salt = Encoding.UTF8.GetBytes("StockManager-Capital-Credential-Salt");
            using (var derive = new Rfc2898DeriveBytes(entropy, salt, 1000))
            using (var aes = new AesManaged())
            {
                aes.Key = derive.GetBytes(32);
                aes.IV = derive.GetBytes(16);
                aes.Mode = CipherMode.CBC;
                aes.Padding = PaddingMode.PKCS7;

                using (var ms = new MemoryStream())
                using (var cs = new CryptoStream(ms, aes.CreateEncryptor(), CryptoStreamMode.Write))
                {
                    cs.Write(plain, 0, plain.Length);
                    cs.FlushFinalBlock();
                    return ms.ToArray();
                }
            }
        }

        private byte[] DecryptCapitalCredentialBytes(byte[] cipher)
        {
            var entropy = (Environment.UserName + "|" + Environment.MachineName + "|StockManager");
            var salt = Encoding.UTF8.GetBytes("StockManager-Capital-Credential-Salt");
            using (var derive = new Rfc2898DeriveBytes(entropy, salt, 1000))
            using (var aes = new AesManaged())
            {
                aes.Key = derive.GetBytes(32);
                aes.IV = derive.GetBytes(16);
                aes.Mode = CipherMode.CBC;
                aes.Padding = PaddingMode.PKCS7;

                using (var input = new MemoryStream(cipher))
                using (var cs = new CryptoStream(input, aes.CreateDecryptor(), CryptoStreamMode.Read))
                using (var output = new MemoryStream())
                {
                    cs.CopyTo(output);
                    return output.ToArray();
                }
            }
        }

        private void ClearCapitalCredential()
        {
            try
            {
                if (File.Exists(_capitalCredentialFile))
                {
                    File.Delete(_capitalCredentialFile);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[群益登入記憶清除失敗] {ex.Message}");
            }
        }
    }
}
