// Decompiled with JetBrains decompiler
// Type: MOS_EXCEL_LEARN.Home
// Assembly: MOS_EXCEL_LEARN, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null
// MVID: A91AAF83-2707-4347-A301-00149AC4CDCE
// Assembly location: C:\Users\Admin\Downloads\WE_19\HOC_EXCEL_19\Chay File Nay.exe

using System;
using System.IO;
using System.Security.Cryptography;
using System.Text;

namespace MOS_EXCEL_LEARN
{
  internal class Home
  {
    public static byte[] Encrypt(byte[] inputArray, string key)
    {
      TripleDESCryptoServiceProvider cryptoServiceProvider = new TripleDESCryptoServiceProvider();
      cryptoServiceProvider.Key = Encoding.UTF8.GetBytes(key);
      cryptoServiceProvider.Mode = CipherMode.ECB;
      cryptoServiceProvider.Padding = PaddingMode.PKCS7;
      byte[] numArray = cryptoServiceProvider.CreateEncryptor().TransformFinalBlock(inputArray, 0, inputArray.Length);
      cryptoServiceProvider.Clear();
      return numArray;
    }

    public static byte[] Decrypt(byte[] inputArray, string key)
    {
      TripleDESCryptoServiceProvider cryptoServiceProvider = new TripleDESCryptoServiceProvider();
      cryptoServiceProvider.Key = Encoding.UTF8.GetBytes(key);
      cryptoServiceProvider.Mode = CipherMode.ECB;
      cryptoServiceProvider.Padding = PaddingMode.PKCS7;
      byte[] numArray = cryptoServiceProvider.CreateDecryptor().TransformFinalBlock(inputArray, 0, inputArray.Length);
      cryptoServiceProvider.Clear();
      return numArray;
    }

    public static void EncryptFile(string fileNameSource, string fileNameDest)
    {
      BinaryReader binaryReader = new BinaryReader((Stream) File.Open(fileNameSource, FileMode.Open));
      int length = (int) binaryReader.BaseStream.Length;
      byte[] inputArray = binaryReader.ReadBytes(length);
      binaryReader.Close();
      byte[] buffer = Home.Encrypt(inputArray, "4180835854685610rivbgfthd6712096111949160@&^#%$1599429789645140");
      BinaryWriter binaryWriter = new BinaryWriter((Stream) File.Open(fileNameDest, FileMode.Create));
      binaryWriter.Write(buffer);
      binaryWriter.Close();
    }

    public static void DecryptFile(string fileNameSource, string fileNameDest)
    {
      BinaryReader binaryReader = new BinaryReader((Stream) File.Open(fileNameSource, FileMode.Open));
      int length = (int) binaryReader.BaseStream.Length;
      byte[] inputArray = binaryReader.ReadBytes(length);
      binaryReader.Close();
      byte[] buffer = Home.Decrypt(inputArray, "123456789012345678901234");
      BinaryWriter binaryWriter = new BinaryWriter((Stream) File.Open(fileNameDest, FileMode.Create));
      binaryWriter.Write(buffer);
      binaryWriter.Close();
    }

    public static void EncryptFileT(string fileNameSource, string fileNameDest)
    {
      TextReader textReader = (TextReader) new StreamReader(fileNameSource);
      string end = textReader.ReadToEnd();
      textReader.Close();
      string str = Home.EncryptString("a12ws5246b4e4133bbce2ea2315a2021", end);
      TextWriter textWriter = (TextWriter) new StreamWriter(fileNameDest);
      textWriter.Write(str);
      textWriter.Close();
    }

    public static void DecryptFileT(string fileNameSource, string fileNameDest)
    {
      TextReader textReader = (TextReader) new StreamReader(fileNameSource);
      string end = textReader.ReadToEnd();
      textReader.Close();
      string str = Home.DecryptString("a12ws5246b4e4133bbce2ea2315a2021", end);
      TextWriter textWriter = (TextWriter) new StreamWriter(fileNameDest);
      textWriter.Write(str);
      textWriter.Close();
    }

    public static string EncryptString(string key, string plainText)
    {
      byte[] numArray = new byte[16];
      byte[] array;
      using (Aes aes = Aes.Create())
      {
        aes.Key = Encoding.UTF8.GetBytes(key);
        aes.IV = numArray;
        ICryptoTransform encryptor = aes.CreateEncryptor(aes.Key, aes.IV);
        using (MemoryStream memoryStream = new MemoryStream())
        {
          using (CryptoStream cryptoStream = new CryptoStream((Stream) memoryStream, encryptor, CryptoStreamMode.Write))
          {
            using (StreamWriter streamWriter = new StreamWriter((Stream) cryptoStream))
              streamWriter.Write(plainText);
            array = memoryStream.ToArray();
          }
        }
      }
      return Convert.ToBase64String(array);
    }

    public static string DecryptString(string key, string cipherText)
    {
      byte[] numArray = new byte[16];
      byte[] buffer = Convert.FromBase64String(cipherText);
      using (Aes aes = Aes.Create())
      {
        aes.Key = Encoding.UTF8.GetBytes(key);
        aes.IV = numArray;
        ICryptoTransform decryptor = aes.CreateDecryptor(aes.Key, aes.IV);
        using (MemoryStream memoryStream = new MemoryStream(buffer))
        {
          using (CryptoStream cryptoStream = new CryptoStream((Stream) memoryStream, decryptor, CryptoStreamMode.Read))
          {
            using (StreamReader streamReader = new StreamReader((Stream) cryptoStream))
              return streamReader.ReadToEnd();
          }
        }
      }
    }
  }
}
