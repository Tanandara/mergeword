Imports Spire.Doc
Imports System.IO

Module Module1
    'เอามาจาก https://www.e-iceblue.com/Tutorials/Spire.Doc/Spire.Doc-Program-Guide/NET-Merge-Word-Merge-Multiple-Word-Documents-into-One-in-C-and-VB.NET.html
    Sub Main()
        Dim path As String = Directory.GetCurrentDirectory() ' path ปัจจุบัน
        Dim d As DirectoryInfo = New DirectoryInfo(path) ' ไฟล์ใน path



        Dim Doc1 As New Document() ' สร้างไฟล์แรกสำหรับรวม
        Doc1.LoadFromFile(path + "\" + d.GetFiles("*.docx")(0).Name, FileFormat.Docx2013)

        For Each file In d.GetFiles("*.docx")
            If (file.Name = d.GetFiles("*.docx")(0).Name) Then ' ไม่เอาไฟล์แรก
                Continue For
            End If

            Dim Doc2 As New Document() ' ไฟล์ที่เหลือ
            Doc2.LoadFromFile(path + "\" + file.Name, FileFormat.Docx2013)

            For Each section As Section In Doc2.Sections
                Doc1.Sections.Add(section.Clone()) ' เพิ่มลงไปในไฟล์แรก
            Next section

            Doc2.Close()

            Console.WriteLine(file.Name)

        Next

        Doc1.SaveToFile("Merge.docx", FileFormat.Docx2013) ' เซฟไฟล์ที่รวมแล้วแยกออกมา
        Doc1.Close()

        System.Diagnostics.Process.Start("Merge.docx") ' เปิดไฟล์

        Console.ReadKey()


    End Sub

End Module
