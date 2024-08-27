Imports System.Security.Cryptography
Imports System.Text
Module ModHashCode

    Public Function fbuGetHash(vriInput As String) As String
        Using vnHasher As MD5 = MD5.Create()    ' create hash object

            ' Convert to byte array and get hash
            Dim vnDBytes As Byte() =
                 vnHasher.ComputeHash(Encoding.UTF8.GetBytes(vriInput))

            ' sb to create string from bytes
            Dim vnStrBuilder As New StringBuilder()

            ' convert byte data to hex string
            For vn As Integer = 0 To vnDBytes.Length - 1
                vnStrBuilder.Append(vnDBytes(vn).ToString("X2"))
            Next vn

            Return vnStrBuilder.ToString()
        End Using

    End Function
End Module
