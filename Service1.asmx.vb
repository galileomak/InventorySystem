Imports System.Web.Services
Imports System.Web.Services.Protocols
Imports System.ComponentModel

<System.Web.Services.WebService(Namespace:="http://tempuri.org/")> _
<System.Web.Services.WebServiceBinding(ConformsTo:=WsiProfiles.BasicProfile1_1)> _
<ToolboxItem(False)> _
Public Class Service1
    Inherits System.Web.Services.WebService

    Public Class Result
        Public Brand As String
        Public Code As String
        Public ProductName_ As String
        Public ProductType As String
        Public UserCode As String
        Public AgeAndSex As String
        Public SportsType As String


        Public ColorTable As String()()
        Public SizeTable As String()()()
        Public SitePositionTable As String()()()
        Public InventoryTable As String()()()

        Public Message As String = ""
        Public GoMode As Boolean
    End Class


    <WebMethod()> _
    Public Sub ProccessNumber(test As String)
        Try
            OpenConnection()
            Dim value As String = "ni-315122"
            Dim v() As String = value.Split("-")
            Dim r As New Result
            Dim a As PUCode
            If v.Length >= 2 Then
                a = New PUCode(v(0), SF.CutZero(v(1)))
                Dim s As New InfoSys.ISSelectArray
                Dim w As New InfoSys.ISConcatAnd
                Dim dt As DataTable
                dt = SC.Allc.Query(SC.SSO.ProductColorSize, a.GenBrandAndCodeWhere(SC.SSO.ProductColorSize))
                If dt.Rows.Count > 0 Then
                    r.Brand = dt.Rows(0)(SC.SSO.ProductColorSize.Brand.Alias)
                    r.Code = dt.Rows(0)(SC.SSO.ProductColorSize.Code.Alias)
                    r.ProductName_ = dt.Rows(0)(SC.SSO.ProductColorSize.Name.Alias)
                    r.ProductType = dt.Rows(0)(SC.SSO.ProductColorSize.ProductType.Alias)
                    r.UserCode = dt.Rows(0)(SC.SSO.ProductColorSize.UserCode.Alias)
                    r.AgeAndSex = dt.Rows(0)(SC.SSO.ProductColorSize.ProductAgeAndSex.Alias)
                    r.SportsType = dt.Rows(0)(SC.SSO.ProductColorSize.SportsType.Alias)

                    r.GoMode = False
                    ReadNewNumber(r)
                Else
                    r.Message = "找不到貨品"
                End If
            End If
            If v.Length = 1 Or v.Length = 3 Then
                Dim t As String
                If v.Length = 1 Then t = v(0)
                If v.Length = 3 Then t = v(2)

                If t.Length >= 8 Then
                    Dim e As String = ""
                    Dim PU As PUCode
                    Dim otherpu As PUCode = Nothing
                    PU = New PUCode(t, e, otherpu, False, SC.SSO.SystemSetting.MySite)

                    If e = "" Then
                        r.Brand = PU.Brand
                        r.Code = PU.Code
                        r.ProductName_ = PU.ProductName
                        r.ProductType = PU.ProductType
                        r.UserCode = PU.UserCode
                        r.AgeAndSex = PU.ProductAgeAndSex
                        r.SportsType = PU.SportsType

                        r.GoMode = True
                        ReadNewNumber(r)
                    Else
                        r.Message = "找不到貨品"
                    End If
                Else

                End If

            End If


            Dim re As String
            're = Newtonsoft.Json.JsonConvert.SerializeObject(r, Newtonsoft.Json.Formatting.None)


            HttpContext.Current.Response.ContentType = "application/json;charset=utf-8"
            re = Newtonsoft.Json.JsonConvert.SerializeObject(r, Newtonsoft.Json.Formatting.None)
            HttpContext.Current.Response.Write(re)

            'Return re
        Catch ex As Exception
            HttpContext.Current.Response.ContentType = "application/json;charset=utf-8"
            HttpContext.Current.Response.Write(ex.ToString)
        End Try
      

    End Sub

    Public Sub OpenConnection()
        If SC.DirectConnection Is Nothing Then
            SC.DirectConnection = New Odbc.OdbcConnection

            Console.WriteLine(Server.MapPath("~/"))

            SC.DirectConnection.ConnectionString = MakeConnectionString(Server.MapPath("~/bin/data.mdb"))
            SC.DirectConnection.Open()
            SC.Allc = New ISConnectionWithLog(SC.DirectConnection, "Allc")
            Synchronization.Module1.SetSiteTable(GetType(SiteEx))
            SC.SSO = New SS(False, True, "1")
        End If
    End Sub

    Public Shared DatabasePassword As String = "qap03srd01"
    Public Shared Function MakeConnectionString(ByVal pathfile As String) As String
        If Is64BitProcess() Then
            Return "Driver={Microsoft Access Driver (*.mdb, *.accdb)};Dbq=" + pathfile + ";Uid=Admin;Pwd=" + DatabasePassword + ";"
        Else
            Return "Driver={Microsoft Access Driver (*.mdb)};Dbq=" + pathfile + ";Uid=Admin;Pwd=" + DatabasePassword + ";"
        End If
    End Function
    Public Shared Function Is64BitProcess() As Boolean
        Return IntPtr.Size = 8
    End Function
    Public Function ReadNewNumber(r As Result) As String
        Dim ColorTable As DataTable = Nothing

        Dim a() As DataTable = Nothing
        Dim b() As DataTable = Nothing
        Dim c() As DataTable = Nothing
        Dim br() As DataTable = Nothing
        Dim cr() As DataTable = Nothing
        Dim sitetempdt As DataTable = Nothing
        Try
            SF.QueryInventoryTableByColor(r.Brand, r.Code, r.ProductType, a, b, c, ColorTable, QueryInventoryType.Qty, sitetempdt, br, cr)
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try

        r.ColorTable = ToArray(ColorTable)
        Dim i As Integer
        ReDim r.SizeTable(a.Length - 1)
        For i = 0 To a.Length - 1
            r.SizeTable(i) = ToArray(a(i))
        Next
        ReDim r.SitePositionTable(b.Length - 1)
        For i = 0 To b.Length - 1
            r.SitePositionTable(i) = ToArray(b(i))
        Next
        ReDim r.InventoryTable(c.Length - 1)
        For i = 0 To c.Length - 1
            r.InventoryTable(i) = ToArray(c(i))
        Next

    End Function

    Public Function ToArray(dt As DataTable) As String()()
        Dim i As Integer
        Dim r()() As String
        ReDim r(dt.Rows.Count)
        ReDim r(0)(dt.Columns.Count - 1)
        For i = 0 To dt.Columns.Count - 1
            r(0)(i) = dt.Columns(i).ColumnName
        Next
        Dim j As Integer
        For i = 0 To dt.Rows.Count - 1
            ReDim r(i + 1)(dt.Columns.Count - 1)
            For j = 0 To dt.Columns.Count - 1
                r(i + 1)(j) = dt.Rows(i)(j)
            Next
        Next
        Return r
    End Function
End Class