
'Mewnforio lyfrgell ar gyfer delio â ffeiliau
Imports System.IO
'************************************************************************************************
Public Class Form1

    'Cofnod manylion cwsmer
    Public Structure Customer
        Public RhACwsmer As String 'Maes allweddol
        Public Teitl As String 'Teitl y cwsmer
        Public Enw As String 'Enw cyntaf y cwsmer
        Public Cyfenw As String 'Cyfenw'r cwsmer
        Public Cyfeiriad1 As String 'Llinell 1 y cyfeiriad
        Public Cyfeiriad2 As String 'Llinell 2 y cyfeiriad
        Public Cyfeiriad3 As String 'Llinell 3 y cyfeiriad
        Public CodPost As String 'Cod Post y cyfeiriad
        Public DyddiadGeni As String 'Dyddiad geni'r cwsmer
        Public Rhyw As String 'Rhyw'r cwsmer
    End Structure
    '******************************************************************************************
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'Oes ffeil o'r enw Customer.txt yn bodoli yn y ffolder cyfredol
        If Dir$("Customer.txt") = "" Then 'Nag oes
            Dim sw As New StreamWriter(Application.StartupPath & "\Customer.txt", True)  'creu ffeil os nad yw'n bodoli
            sw.Close() 'Cau'r sianel
            MsgBox("A new database has been created", vbExclamation, "Warning!")
        End If
    End Sub

    Private Sub btnArbed_Click(sender As Object, e As EventArgs) Handles btnArbed.Click
        'Datganiadau
        Dim Cwsmer As New Customer 'Cofnod newydd

        'Darllen y data o'r ffurflen
        Darllen(Cwsmer)
        'Dilysu'r data
        If Dilysu(Cwsmer) Then 'Data yn dilys
            'Arbed y data
            'Agor ffrwd newydd i ffeil
            Dim sw As New StreamWriter(Dir$("Customer.txt"), True) '- hyn dim yn gweithio
            'Dim sw As New StreamWriter(CurDir("Customer.txt"), True)

            'Addasu hyd y meysydd i'w maint osodedig
            With Cwsmer
                .RhACwsmer = LSet(.RhACwsmer, 4)
                .Teitl = LSet(.Teitl, 4)
                .Enw = LSet(.Enw, 15)
                .Cyfenw = LSet(.Cyfenw, 25)
                .Cyfeiriad1 = LSet(.Cyfeiriad1, 30)
                .Cyfeiriad2 = LSet(.Cyfeiriad2, 30)
                .Cyfeiriad3 = LSet(.Cyfeiriad3, 30)
                .CodPost = LSet(.CodPost, 8)
                .DyddiadGeni = LSet(.DyddiadGeni, 10)
                .Rhyw = LSet(.Rhyw, 1)
                'Atodi i'r ffeil
                sw.WriteLine(.RhACwsmer & .Teitl & .Enw & .Cyfenw & .Cyfeiriad1 & .Cyfeiriad2 & .Cyfeiriad3 & .CodPost & .DyddiadGeni & .Rhyw)
                sw.Close()
            End With
            MsgBox("Record stored")
        Else 'Data yn annilys
            MsgBox("Record NOT stored")
        End If

    End Sub

    Private Sub Darllen(ByRef Cwsmer As Customer)
        'Darllen data o'r ffurflen a storio yn y maes priodol
        'Dileu unrhyw bylchau ar y blaen ac wrth gefn y data
        With Cwsmer
            .RhACwsmer = Trim(txtRhACleient.Text)
            .Teitl = Trim(txtTeitl.Text)
            .Enw = Trim(txtEnw.Text)
            .Cyfenw = (Trim(txtCyfenw.Text))
            .Cyfeiriad1 = Trim(txtCyfeiriad1.Text)
            .Cyfeiriad2 = Trim(txtCyfeiriad2.Text)
            .Cyfeiriad3 = Trim(txtCyfeiriad3.Text)
            .CodPost = Trim(txtCodPost.Text)
            .DyddiadGeni = Trim(txtDyddiadGeni.Text)
            .Rhyw = Trim(txtRhyw.Text)
        End With
    End Sub

    Private Function Dilysu(ByVal Cwsmer As Customer) As Boolean
        'Profi dilysrwydd y data
        Dim Dilys As Boolean = True 'Gosod data i fod yn nilys i ddechrau

        'Dilysu presenoldeb data maes
        If PresenoldebDilys(Cwsmer) Then 'Data yn bresenol yn y meysydd yma
            'Dilysu hyd y meysydd
            If HydDilys(Cwsmer) Then 'hyd y data yn ddilys yn y meysydd yma
                'Profion meysydd unigol
                'Profi Rhif Cwsmer
                If RhACustomerDilys(Cwsmer.RhACwsmer) Then 'Maes allweddol yn ddilys
                    If TeitlDilys(Cwsmer.Teitl) Then 'Teitl yn ddilys
                        If CodPostDilys(Cwsmer.CodPost) Then 'Cod Post dilys
                            If DyddiadGeniDilys(Cwsmer.DyddiadGeni) Then 'Dyddiad geni dilys
                                If RhywDilys(Cwsmer.Rhyw) Then 'Rhyw dilys
                                    'Pob eitem yn ddilys
                                    Return Dilys
                                Else 'Rhyw annilys
                                    Dilys = False
                                End If 'Rhyw
                            Else 'Dyddiad geni annilys
                                Dilys = False
                            End If 'Dyddiad geni
                        Else 'Cod Post annilys
                            Dilys = False
                        End If 'Cod Post
                    Else 'Teitl annilys
                        Dilys = False
                    End If 'Teitl
                Else 'RhACwsmer annilys
                    Dilys = False
                End If 'RhACwsmer
            Else 'Hyd annilys
                Dilys = False
            End If 'Hyd
        Else 'Presenoldeb annilys
            Dilys = False
        End If 'Presenoldeb

        Return Dilys
    End Function
    Private Function PresenoldebDilys(ByVal Cwsmer As Customer) As Boolean
        'Oes data yn bresenol yn y meysydd dilysu presenoldeb ac hyd

        Dim Dilys As Boolean = True 'gwerth cychwynol yn wir - newid i false os oes data annilys

        With Cwsmer 'Gwirio presenoldeb pob maes sy'n berthnasol
            If .Enw = "" Then
                Dilys = False
                MsgBox("Name is missing")
            End If
            If .Cyfenw = "" Then
                Dilys = False
                MsgBox("Surname is missing")
            End If
            If .Cyfeiriad1 = "" Then
                Dilys = False
                MsgBox("Address Line 1 is missing")
            End If
            If .Cyfeiriad2 = "" Then
                Dilys = False
                MsgBox("Address Line 2 is missing")
            End If
            If .Cyfeiriad3 = "" Then
                Dilys = False
                MsgBox("Address Line 3 is missing")
            End If
        End With
        Return Dilys
        'Diwedd PresenoldebDilys
        '****************************************************************************************
    End Function

    Private Function HydDilys(ByVal Cwsmer As Customer) As Boolean
        'Dilysu hyd y meysydd dilysu presenoldeb ac hyd
        Dim Dilys As Boolean = True 'gwerth cychwynol yn wir - newid i false os oes data annilys

        With Cwsmer 'Gwirio hyd y meysydd perthnasol
            If .RhACwsmer.Length > 4 Then '
                Dilys = False
                MsgBox("Customer number is too long")
            End If
            If .Teitl.Length > 4 Then
                Dilys = False
                MsgBox("Title is too long")
            End If
            If .Enw.Length > 15 Then
                Dilys = False
                MsgBox("Name is too long")
            End If
            If .Cyfenw.Length > 25 Then
                Dilys = False
                MsgBox("Surname is too long")
            End If
            If .Cyfeiriad1.Length > 30 Then
                Dilys = False
                MsgBox("Address Line 1 is too long")
            End If
            If .Cyfeiriad2.Length > 30 Then
                Dilys = False
                MsgBox("Address Line 2 is too long")
            End If
            If .Cyfeiriad3.Length > 30 Then
                Dilys = False
                MsgBox("Address Line 3 is too long")
            End If
        End With

        Return Dilys

        'Diwedd HydDilys
        '********************************************************************************************
    End Function

    Private Function RhACustomerDilys(ByVal RhACwsmer As String) As Boolean
        'Profi Rhif Cwsmer
        Dim Dilys As Boolean = True 'gwerth cychwynol yn wir - newid i false os oes data annilys

        If RhACwsmer = "" Then 'Gwag
            Dilys = False
            MsgBox("Customer number is missing")
        Else ' data'n bresenol
            'Profi bod Rhif Allweddol y cwsmer yn cyfanrif
            If IsNumeric(RhACwsmer) Then
                If InStr(RhACwsmer, ".") = 0 Then 'Nid yw'n cynnwys pwynt degol
                    'Dilysu bod yn fwy na 0.
                    If CInt(RhACwsmer) <= 0 Then ' Negyddol - felly annilys
                        Dilys = False
                        MsgBox("Customer Number is negative")
                    Else 'Dilysu unigrywedd
                        If Dyblygu(RhACwsmer) Then ' nid yw'n unigryw
                            Dilys = False
                            MsgBox("Customer Number already exists")
                        End If 'unigrywedd
                    End If 'negyddol
                Else ' Cynnwys pwynt degol
                    Dilys = False
                    MsgBox("Customer Number is not an integer")
                End If
            Else
                Dilys = False
                MsgBox("Customer Number is not a number")
            End If ' cyfanrif
        End If 'gwag

        Return Dilys
        'Diwedd Rhif Cwsmer
        '**************************************************************************
    End Function

    Private Function TeitlDilys(ByVal Teitl As String) As Boolean
        'Dilysu Teitl

        Dim Dilys As Boolean = True 'gwerth cychwynol yn wir - newid i false os oes data annilys
        If Teitl = "" Then
            Dilys = False
            MsgBox("Title is missing")
        Else
            'Gwerthoedd penodol
            If (Teitl <> "Mr") And (Teitl <> "Miss") And (Teitl <> "Mrs") And (Teitl <> "Ms") And (Teitl <> "Dr") Then 'Teitl annilys
                Dilys = False
                MsgBox("Invalid Title")
            End If 'gwerthoedd
        End If 'gwag
        Return Dilys
        'Diwedd Teitl
        '*******************************************************************************
    End Function

    Private Function CodPostDilys(ByVal CodPost As String) As Boolean
        'Dilysu Cod Post

        Dim Dilys As Boolean = True 'gwerth cychwynol yn wir - newid i false os oes data annilys

        If CodPost = "" Then
            Dilys = False
            MsgBox("Post Code is missing")
        Else
            'Hyd
            If (CodPost.Length < 7) Or (CodPost.Length > 8) Then
                Dilys = False
                MsgBox("Post Code is wrong length")
            Else
                'Profi fformat - cywir AA90 0AA
                Dim CodPostDrosDro As String = UCase(CodPost) 'arbed y cod post mewn newidyn PRIF LLYTHRENNAU dros dro

                If (CodPostDrosDro(0) < "A") Or (CodPostDrosDro(0) > "Z") Then ' Nid yw'n Llythyren
                    Dilys = False
                End If
                If (CodPostDrosDro(1) < "A") Or (CodPostDrosDro(1) > "Z") Then ' Nid yw'n Llythyren
                    Dilys = False
                End If
                If (CodPostDrosDro(2) < "1") Or (CodPostDrosDro(2) > "9") Then ' nid yw'n digid dilys
                    Dilys = False
                End If
                If CodPostDrosDro.Length = 7 Then 'Cod post 7 nod
                    If (CodPostDrosDro(3) <> " ") Then ' nid yw'n bwlch
                        Dilys = False
                    End If
                    If (CodPostDrosDro(4) < "1") Or (CodPostDrosDro(4) > "9") Then ' nid yw'n digid dilys
                        Dilys = False
                    End If
                    If (CodPostDrosDro(5) < "A") Or (CodPostDrosDro(5) > "Z") Then ' Nid yw'n Llythyren
                        Dilys = False
                    End If
                    If (CodPostDrosDro(6) < "A") Or (CodPostDrosDro(6) > "Z") Then ' Nid yw'n Llythyren
                        Dilys = False
                    End If
                Else ' Cod post 8 Nod
                    If (CodPostDrosDro(3) < "0") Or (CodPostDrosDro(3) > "9") Then ' nid yw'n digid dilys
                        Dilys = False
                    End If
                    If (CodPostDrosDro(4) <> " ") Then
                        Dilys = False
                    End If
                    If (CodPostDrosDro(5) < "1") Or (CodPostDrosDro(5) > "9") Then ' nid yw'n digid dilys
                        Dilys = False
                    End If
                    If (CodPostDrosDro(6) < "A") Or (CodPostDrosDro(6) > "Z") Then ' Nid yw'n Llythyren
                        Dilys = False
                    End If
                    If (CodPostDrosDro(7) < "A") Or (CodPostDrosDro(7) > "Z") Then ' Nid yw'n Llythyren
                        Dilys = False
                    End If
                End If
                If Not (Dilys) Then ' cod post yn y fformat anghywir
                    MsgBox("Cod Post incorrect format")
                End If
            End If
        End If

        Return Dilys
        'Diwedd cod post
        '*****************************************************************************
    End Function

    Private Function DyddiadGeniDilys(ByVal DyddiadGeni As String) As Boolean
        'Dilysu Dyddiad Geni
        Dim Dilys As Boolean = True 'gwerth cychwynol yn wir - newid i false os oes data annilys

        If DyddiadGeni = "" Then 'Gwag
            Dilys = False
            MsgBox("DOB is missing")
        Else
            'Profi hyd = 8
            If (DyddiadGeni.Length <> 10) Then 'Hyd anghywir
                Dilys = False
                MsgBox("DOB is wrong length")
            Else 'Hyd cywir
                'Gwirio fformat - dd/mm/bbbb
                Dim DyddiadDrosDro As String = DyddiadGeni 'newidyn dros dro
                Dim DyddiadDilys As Boolean = True 'Er mwyn cael unneges ar gyfer y dyddiad ar ddiwedd y prawf
                Dim i As Integer 'Rheoli'r ddolen
                'Profi fformat ac amrediad
                For i = 0 To 9
                    Select Case i
                        Case 0
                            If (DyddiadDrosDro(i) < "0") Or (DyddiadDrosDro(i) > "3") Then 'annilys
                                DyddiadDilys = False
                            End If
                        Case 1, 3, 6, 7, 8, 9
                            If (DyddiadDrosDro(i) < "0") Or (DyddiadDrosDro(i) > "9") Then 'annilys
                                DyddiadDilys = False
                            End If
                        Case 4
                            If (DyddiadDrosDro(i) < "0") Or (DyddiadDrosDro(i) > "2") Then 'annilys
                                DyddiadDilys = False
                            End If
                        Case 6
                            If (DyddiadDrosDro(i) < "1") Or (DyddiadDrosDro(i) > "2") Then 'annilys
                                DyddiadDilys = False
                            End If
                        Case 2, 5
                            If DyddiadDrosDro(i) <> "/" Then 'annilys
                                DyddiadDilys = False
                            End If
                    End Select
                Next
                If DyddiadDilys Then 'dal yn dilys
                    'profi gwerthoedd
                    'Trosi dydd mis a blwyddyn i gyfanrifau
                    Dim dydd As Integer
                    dydd = CInt(DyddiadDrosDro.Substring(0, 2))

                    Dim mis As Integer
                    mis = CInt(DyddiadDrosDro.Substring(3, 2))

                    Dim blwyddyn As Integer
                    blwyddyn = CInt(DyddiadDrosDro.Substring(6, 4))

                    'Dilysu bod y dyddiad yn bodoli
                    'Dilysu mis
                    If (mis < 1) Or (mis > 12) Then 'annilys
                        DyddiadDilys = False
                    Else ' Mis yn dilys
                        Select Case mis
                            Case 1, 3, 5, 7, 8, 10, 12
                                '1 i 31
                                If dydd > 31 Then
                                    DyddiadDilys = False
                                End If
                            Case 4, 6, 9, 11
                                '1 i 30
                                If dydd > 30 Then
                                    DyddiadDilys = False
                                End If
                            Case 2
                                'Chwefror - blwyddyn naid
                                If (blwyddyn Mod 4) = 0 Then 'Blwyddyn naid
                                    If dydd > 29 Then
                                        DyddiadDilys = False
                                    End If
                                Else 'Nid yw'n blwyddyn naid
                                    If dydd > 28 Then
                                        DyddiadDilys = False
                                    End If
                                End If
                        End Select
                    End If
                End If
                'Os yw'r dydiad yn annilys
                If Not (DyddiadDilys) Then
                    Dilys = False
                    MsgBox("Invalid Date")
                End If
            End If 'hyd
        End If 'gwag

        Return Dilys
        'Diwedd dilysu dyddiad geni
        '********************************************************************************************
    End Function
    Private Function RhywDilys(ByVal Rhyw As String) As Boolean
        'Dilysu rhyw

        Dim Dilys As Boolean = True 'gwerth cychwynol yn wir - newid i false os oes data annilys

        If Rhyw = "" Then 'Gwag
            Dilys = False
            MsgBox("Gender is missing")
        Else
            If Rhyw.Length > 1 Then
                Dilys = False
                MsgBox("Gender is too long")
            Else 'Gwirio am werth dilys
                Dim RhywDrosDro As String = UCase(Rhyw)
                If (RhywDrosDro <> "F") And (RhywDrosDro <> "M") Then 'Annilys
                    Dilys = False
                    MsgBox("Gender is invalid - F or M required")
                End If
            End If
        End If

        Return Dilys
        'Diwedd dilysu rhyw
        '********************************************************************************************
    End Function
    Private Function Dyblygu(ByVal RhifCwsmer As String) As Boolean
        Dim WediDyblygu As Boolean = False
        Dim i As Integer = 0

        'Agor a darllrn y ffeil i arae
        Dim CwsmerTabl() As String = File.ReadAllLines(Application.StartupPath & "\Customer.txt")

        'Tra bod cofnod i wirio, ac heb ddarganfod y rhif 
        Do While i <= UBound(CwsmerTabl) And Not (WediDyblygu)
            If RhifCwsmer = Trim(Mid(CwsmerTabl(i), 1, 4)) Then 'Mae allweddol yn bodoli'n barod
                WediDyblygu = True
            End If
            i = i + 1 ' nesaf
        Loop
        Return WediDyblygu
    End Function
    Private Sub btnCau_Click(sender As Object, e As EventArgs) Handles btnCau.Click
        If MsgBox("Select Yes to confirm program close?", MsgBoxStyle.YesNo, "Close Program") = MsgBoxResult.Yes Then
            Application.Exit()
        End If
    End Sub
    Private Sub btnChwilio_Click(sender As Object, e As EventArgs) Handles btnChwilio.Click
        'chwilio am gofnod mewn ffeil yn ol maes allweddol
        Dim TargedChwilio As String = txtRhACleient.Text
        Dim i As Integer 'rheoli dolen
        Dim Cwsmer As New Customer 'Cofnod
        Dim Darganfod As Boolean = False 'Dynodi bod wedi darganfod
        Dim Lleoliad As Integer 'Lleoliad y cofnod

        If (TargedChwilio <> "") And (TargedChwilio.Length <= 4) And CInt(TargedChwilio) Then ' Targed yn dilys
            'Agor ffeil i ddarllen
            'Dim sr As New StreamReader(Dir$("Customer.txt"))
            Dim CwsmerTabl() As String = File.ReadAllLines(Dir$("Customer.txt"))

            Do While i <= UBound(CwsmerTabl) And Not (Darganfod)
                If TargedChwilio = Trim(Mid(CwsmerTabl(i), 1, 4)) Then 'Mae allweddol yn bodoli'n barod
                    Darganfod = True
                    Lleoliad = i
                End If
                i = i + 1 ' nesaf
            Loop
            If Darganfod Then 'Wedi darganfod
                MsgBox(Lleoliad)
                Allbynnu(CwsmerTabl(Lleoliad))
            Else 'Heb ddarganfod
                MsgBox("Heb ddarganfod")
            End If

        Else ' data i chwilio
            MsgBox("Customer number invalid")
        End If
    End Sub

    Private Sub Allbynnu(ByVal Cwsmer As String)
        'Ysgrifennu'r cofnod i'r blychau testun
        With Cwsmer
            txtRhACleient.Text = Trim(Mid(Cwsmer, 1, 4))
            txtTeitl.Text = Trim(Mid(Cwsmer, 5, 4))
            txtEnw.Text = Trim(Mid(Cwsmer, 9, 15))
            txtCyfenw.Text = Trim(Mid(Cwsmer, 24, 25))
            txtCyfeiriad1.Text = Trim(Mid(Cwsmer, 49, 30))
            txtCyfeiriad2.Text = Trim(Mid(Cwsmer, 79, 30))
            txtCyfeiriad3.Text = Trim(Mid(Cwsmer, 109, 30))
            txtCodPost.Text = Trim(Mid(Cwsmer, 139, 8))
            txtDyddiadGeni.Text = Trim(Mid(Cwsmer, 147, 10))
            txtRhyw.Text = Trim(Mid(Cwsmer, 157, 1))
        End With
    End Sub
End Class
