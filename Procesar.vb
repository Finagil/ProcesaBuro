Imports System.IO
Module Procesar
    Dim RutaBuro As String = "\\server-nas\BuroCredito\OCT2019\"
    Dim RutaBuroBackup As String = "\\server-nas\BuroCredito\OCT2019\Backup\"
    Dim Mes As String = "20191001"
    Dim TA As New FinagilDSTableAdapters.BC_BuroDatosTableAdapter
    Dim TF As New FinagilDSTableAdapters.FIRA_CALTableAdapter

    Sub Main()

        Dim D As New System.IO.DirectoryInfo(RutaBuro)
        Dim DB As New System.IO.DirectoryInfo(RutaBuroBackup)
        If Not DB.Exists Then
            System.IO.Directory.CreateDirectory(RutaBuroBackup)
        End If
        Dim F As FileInfo() = D.GetFiles("*.csv")
        'TA.DeleteALL()
        'TF.DeleteAll()
        For i As Integer = 0 To F.Length - 1
            Console.WriteLine(F(i).CreationTime & " " & F(i).Name)
            LeeArchivo(F(i).FullName)
            File.Move(F(i).FullName, RutaBuroBackup & F(i).Name)
        Next
        'PARTE 2++++++++++++++++++++++++++++++++
            F = D.GetFiles("*.txt")
        For i As Integer = 0 To F.Length - 1
            Console.WriteLine(F(i).CreationTime & " " & F(i).Name)
            LeeArchivo2(F(i).FullName)
            File.Move(F(i).FullName, RutaBuroBackup & F(i).Name)
        Next
        'PARTE 2++++++++++++++++++++++++++++++++
    End Sub


    Sub LeeArchivo(ByVal F As String)
        Dim X As Integer = 0
        Dim RFC As String = ""
        Dim RFCant As String = ""
        Dim cad As String
        Dim Vars(26) As String
        Dim Datos() As String
        Dim Linea As String
        Dim f1 As New System.IO.StreamReader(F, Text.Encoding.GetEncoding(1252))
        Linea = f1.ReadLine
        While Not f1.EndOfStream
            Linea = f1.ReadLine
            Datos = Linea.Split("|")
            If Trim(Datos(55)) <> "" Then ' nombre de la caracteristica
                Console.WriteLine(Datos(4) & " " & Datos(55))
                cad = Datos(56)
                If Mid(cad, 1, 1) = "'" Then
                    cad = Mid(cad, 2, cad.Length)
                End If
                If RFCant <> Datos(4) And RFCant <> "" Then
                    TA.DeleteRFC(RFC)
                    TA.Insert(RFC, Mes, Vars(0), Vars(1), Vars(2), Vars(3), Vars(4), Vars(5), Vars(6), Vars(7), Vars(8), Vars(9),
                                                    Vars(10), Vars(11), Vars(12), Vars(13), Vars(14), Vars(15), Vars(16), Vars(17), Vars(18), Vars(19),
                                                    Vars(20), Vars(21), Vars(22), Vars(23), Vars(24), Vars(25), Vars(26))
                    X = 0
                End If
                Vars(X) = Trim(cad)

                X += 1
                RFC = Datos(4)
                If Mid(Datos(1), 1, 1) = "'" Then Datos(1) = Mid(Datos(1), 2, Datos(1).Length)
                If Mid(Datos(2), 1, 1) = "'" Then Datos(2) = Mid(Datos(2), 2, Datos(2).Length)
                If Mid(Datos(3), 1, 1) = "'" Then Datos(3) = Mid(Datos(3), 2, Datos(3).Length)
                If Mid(Datos(54), 1, 1) = "'" Then Datos(54) = Mid(Datos(54), 2, Datos(54).Length)
                If Mid(Datos(56), 1, 1) = "'" Then Datos(56) = Mid(Datos(56), 2, Datos(56).Length)
                If RFCant <> RFC Then TF.DeleteRFC(RFC)
                TF.Insert(Datos(0), Datos(1), Datos(2), Datos(3), Datos(4), Datos(5), Datos(6), Datos(7), Datos(8), Datos(54), Datos(55), Datos(56), Datos(57))
                RFCant = RFC
            End If
        End While
        TA.DeleteRFC(RFC)
        TA.Insert(RFC, Mes, Vars(0), Vars(1), Vars(2), Vars(3), Vars(4), Vars(5), Vars(6), Vars(7), Vars(8), Vars(9), _
                                        Vars(10), Vars(11), Vars(12), Vars(13), Vars(14), Vars(15), Vars(16), Vars(17), Vars(18), Vars(19), _
                                        Vars(20), Vars(21), Vars(22), Vars(23), Vars(24), Vars(25), Vars(26))


        f1.Close()
        f1.Dispose()


    End Sub

    Sub LeeArchivo2(ByVal F As String)
        Dim X As Integer = 0
        Dim RFC As String = ""
        Dim Datos() As String
        Dim Linea As String
        Dim f1 As New System.IO.StreamReader(F, Text.Encoding.GetEncoding(1252))
        Linea = f1.ReadLine
        While Not f1.EndOfStream
            Linea = f1.ReadLine
            Datos = Linea.Split("|")
            If Datos.Length < 27 Then
                Continue While
            End If

            If InStr(Datos(2), "-") Then
                RFC = ""
                For X = 1 To Datos(2).Length
                    If Mid(Datos(2), X, 1) <> "-" Then
                        RFC += Mid(Datos(2), X, 1)
                    End If
                Next
            Else
                RFC = Datos(2)
            End If
            TA.Insert(RFC, Mes, Datos(0 + 5), Datos(1 + 5), Datos(2 + 5), Datos(3 + 5), Datos(4 + 5), Datos(5 + 5), Datos(6 + 5), Datos(7 + 5), Datos(8 + 5), Datos(9 + 5), _
                                        Datos(10 + 5), Datos(11 + 5), Datos(12 + 5), Datos(13 + 5), Datos(14 + 5), Datos(15 + 5), Datos(16 + 5), Datos(17 + 5), Datos(18 + 5), Datos(19 + 5), _
                                        Datos(20 + 5), Datos(21 + 5), Datos(22 + 5), Datos(23 + 5), Datos(24 + 5), Datos(25 + 5), Datos(26 + 5))
            Console.WriteLine(RFC)
        End While
        TA.CorrigeCampos()
        f1.Close()
        f1.Dispose()


    End Sub


End Module
