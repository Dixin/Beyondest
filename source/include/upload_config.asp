<SCRIPT RUNAT=SERVER LANGUAGE=VBSCRIPT> 

Dim upfile_classes_Stream

Class upload_classes
Dim Form
Dim File

Private Sub Class_Initialize
    Dim iStart
    Dim iFileNameStart
    Dim iFileNameEnd
    Dim iEnd
    Dim vbEnter
    Dim iFormStart
    Dim iFormEnd
    Dim theFile
    Dim strDiv
    Dim mFormName
    Dim mFormValue
    Dim mFileName
    Dim mFileSize
    Dim mFilePath
    Dim iDivLen
    Dim mStr

    If Request.TotalBytes < 1 Then Exit Sub
        Set Form                   = CreateObject("Scripting.Dictionary")
        Set File                   = CreateObject("Scripting.Dictionary")
        Set upfile_classes_Stream  = CreateObject("Adodb.Stream")
        upfile_classes_Stream.mode = 3
        upfile_classes_Stream.type = 1
        upfile_classes_Stream.open
        upfile_classes_Stream.Write Request.BinaryRead(Request.TotalBytes)
        vbEnter               = Chr(13) & Chr(10)
        iDivLen               = inString(1,vbEnter) + 1
        strDiv                = subString(1,iDivLen)
        iFormStart            = iDivLen

        iFormEnd              = inString(iformStart,strDiv) - 1
        While iFormStart < iFormEnd
        iStart                = inString(iFormStart,"name=""")
        iEnd                  = inString(iStart + 6,"""")
        mFormName             = subString(iStart + 6,iEnd - iStart - 6)
        iFileNameStart        = inString(iEnd + 1,"filename=""")

        If iFileNameStart > 0 And iFileNameStart < iFormEnd Then
            iFileNameEnd      = inString(iFileNameStart + 10,"""")
            mFileName         = subString(iFileNameStart + 10,iFileNameEnd - iFileNameStart - 10)
            iStart            = inString(iFileNameEnd + 1,vbEnter & vbEnter)
            iEnd              = inString(iStart + 4,vbEnter & strDiv)

            If iEnd > iStart Then
                mFileSize     = iEnd - iStart - 4
            Else
                mFileSize     = 0
            End If

            Set theFile       = new FileInfo
            theFile.FileName  = getFileName(mFileName)
            theFile.FilePath  = getFilePath(mFileName)
            theFile.FileSize  = mFileSize
            theFile.FileStart = iStart + 4
            theFile.FormName  = FormName
            file.add mFormName,theFile
        Else
            iStart         = inString(iEnd + 1,vbEnter & vbEnter)
            iEnd           = inString(iStart + 4,vbEnter & strDiv)

            If iEnd > iStart Then
                mFormValue = subString(iStart + 4,iEnd - iStart - 4)
            Else
                mFormValue = ""
            End If

            form.Add mFormName,mFormValue
        End If

        iFormStart = iformEnd + iDivLen
        iFormEnd   = inString(iformStart,strDiv) - 1
        Wend
    End Sub

    Private Function subString(theStart,theLen)
        Dim i
        Dim c
        Dim stemp
        upfile_classes_Stream.Position = theStart - 1
        stemp                          = ""

        For i = 1 To theLen
            If upfile_classes_Stream.EOS Then Exit For
            c = ascB(upfile_classes_Stream.Read(1))

            If c > 127 Then
                If upfile_classes_Stream.EOS Then Exit For
                stemp = stemp & Chr(AscW(ChrB(AscB(upfile_classes_Stream.Read(1))) & ChrB(c)))
                i     = i + 1
            Else
                stemp = stemp & Chr(c)
            End If

        Next

        subString = stemp
    End Function

    Private Function inString(theStart,varStr)
        Dim i
        Dim j
        Dim bt
        Dim theLen
        Dim str
        InString = 0
        Str      = toByte(varStr)
        theLen   = LenB(Str)

        For i = theStart To upfile_classes_Stream.Size - theLen
            If i > upfile_classes_Stream.size Then Exit Function
            upfile_classes_Stream.Position = i - 1

            If AscB(upfile_classes_Stream.Read(1)) = AscB(midB(Str,1)) Then
                InString                   = i

                For j = 2 To theLen

                    If upfile_classes_Stream.EOS Then
                        inString = 0
                        Exit For
                    End If

                    If AscB(upfile_classes_Stream.Read(1)) <> AscB(MidB(Str,j,1)) Then
                        InString = 0
                        Exit For
                    End If

                Next

                If InString <> 0 Then Exit Function
            End If

        Next

    End Function

    Private Sub Class_Terminate
        form.RemoveAll
        file.RemoveAll
        Set form                  = Nothing
        Set file                  = Nothing
        upfile_classes_Stream.Close
        Set upfile_classes_Stream = Nothing
    End Sub

    Private Function GetFilePath(FullPath)

        If FullPath <> "" Then
            GetFilePath = Left(FullPath,InStrRev(FullPath, "\"))
        Else
            GetFilePath = ""
        End If

    End Function

    Private Function GetFileName(FullPath)

        If FullPath <> "" Then
            GetFileName = Mid(FullPath,InStrRev(FullPath, "\") + 1)
        Else
            GetFileName = ""
        End If

    End Function

    Private Function toByte(Str)
        Dim i
        Dim iCode
        Dim c
        Dim iLow
        Dim iHigh
        toByte    = ""

        For i = 1 To Len(Str)
            c     = Mid(Str,i,1)
            iCode = Asc(c)
            If iCode < 0 Then iCode = iCode + 65535

            If iCode > 255 Then
                iLow   = Left(Hex(Asc(c)),2)
                iHigh  = Right(Hex(Asc(c)),2)
                toByte = toByte & chrB("&H" & iLow) & chrB("&H" & iHigh)
            Else
                toByte = toByte & chrB(AscB(c))
            End If

        Next

    End Function

    End Class

    Class FileInfo
    Dim FormName
    Dim FileName
    Dim FilePath
    Dim FileSize
    Dim FileStart

    Private Sub Class_Initialize
        FileName  = ""
        FilePath  = ""
        FileSize  = 0
        FileStart = 0
        FormName  = ""
    End Sub

    Public Function SaveAs(FullPath)
        Dim dr
        Dim ErrorChar
        Dim i
        SaveAs = 1
        If Trim(fullpath) = "" Or FileSize = 0 Or FileStart = 0 Or FileName = "" Then Exit Function
        If FileStart = 0 Or Right(fullpath,1) = "/" Then Exit Function
        Set dr                         = CreateObject("Adodb.Stream")
        dr.Mode                        = 3
        dr.Type                        = 1
        dr.Open
        upfile_classes_Stream.position = FileStart - 1
        upfile_classes_Stream.copyto dr,FileSize
        dr.SaveToFile FullPath,2
        dr.Close
        Set dr = Nothing
        SaveAs = 0
    End Function

    End Class 
 </SCRIPT>