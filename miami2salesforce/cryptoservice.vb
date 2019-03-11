Imports System.Security.Cryptography
Imports System.Text
Imports System.IO

Public Class CryptoService

    Private Shared _IV As Byte() = System.Text.Encoding.UTF8.GetBytes("GB47sEP1") '8 bytes
    Private Shared _Key As Byte() = System.Text.Encoding.UTF8.GetBytes("tfS7TesN69YPOdqxcVEjM7Qh") ' 24 bytes
    Private Shared tDESalg As TripleDESCryptoServiceProvider = New TripleDESCryptoServiceProvider()

    'Encrypt
    Public Shared Function Encrypt(dataToEncrypt As String) As String

        Dim cryptoTransform As ICryptoTransform = tDESalg.CreateEncryptor(_Key, _IV)

        Dim data As Byte() = Encoding.UTF8.GetBytes(dataToEncrypt)
        Dim memoryStream As MemoryStream = New MemoryStream()
        Dim CryptoStream As CryptoStream = New CryptoStream(memoryStream, cryptoTransform, CryptoStreamMode.Write)
        CryptoStream.Write(data, 0, data.Length)
        CryptoStream.FlushFinalBlock()
        CryptoStream.Close()

        Return Convert.ToBase64String(memoryStream.ToArray())
    End Function

    'Decrypt
    Public Shared Function Decrypt(dataToDecrypt As String) As String

        Dim cryptoTransform As ICryptoTransform = tDESalg.CreateDecryptor(_Key, _IV)
        Dim data As Byte() = Convert.FromBase64String(dataToDecrypt)
        Dim memoryStream As MemoryStream = New MemoryStream()
        Dim CryptoStream As CryptoStream = New CryptoStream(MemoryStream, cryptoTransform, CryptoStreamMode.Write)
        CryptoStream.Write(data, 0, data.Length)
        CryptoStream.FlushFinalBlock()
        CryptoStream.Close()

        Return Encoding.UTF8.GetString(memoryStream.ToArray())
    End Function


End Class
