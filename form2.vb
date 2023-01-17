Imports System.Diagnostics.Eventing.Reader
Imports excel = Microsoft.Office.Interop.Excel
Public Class Form1
    Dim cxrili As New DataTable
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
    End Sub
    Sub kodeqsi()
        Dim k, i As Int16
        Dim sab As Integer
        Dim amor, narch As Single
        If Vnomeri.Text = Nothing Or Vsab.Text = Nothing Or Vcveta.Text = Nothing Or Vweli.Text = Nothing Or Vdas.Text = Nothing Then
            MsgBox("შეავსე ცარიელი ველი!", MsgBoxStyle.Critical)
            Exit Sub
        End If
        sab = Vsab.Text
        Call cxrilinarch()
        k = 1
        For i = Vweli.Text To Today.Year
            If k = 1 Then
                amor = sab * CInt(Vcveta.Text) / 100
                narch = sab - amor
                cxrili.Rows.Add(New String() {Vnomeri.Text, Vdas.Text, i, sab, Vcveta.Text & "%", amor, narch})
            Else
                sab = narch
                amor = sab * Vcveta.Text / 100
                narch = sab - amor
                cxrili.Rows.Add(New String() {Vnomeri.Text, Vdas.Text, i, sab, Vcveta.Text & "%", amor, narch})
            End If
            k = k + 1

        Next
        DGsia.DataSource = cxrili
        Dim ex As New excel.Application
        'ab25 = ex.WorksheetFunction.Fv()
        Call shedegi()


    End Sub
    Sub shedegi()
        Dim k As Integer = 1
        Dim ex = New excel.Application '
        Dim gverdi As excel.Worksheet '
        ex.Workbooks.Add() '
        gverdi = ex.Worksheets.Add '
        gverdi.Name = "ოქმი" '
        'ex.Visible = True
        With cxrili
            For t = 0 To cxrili.Columns.Count - 1
                gverdi.Cells(1, t + 1) = cxrili.Columns(t).Caption 'cxrili.Columns(t).Caption

            Next
            k = 2
            For i = 0 To cxrili.Rows.Count - 1
                For t = 0 To .Columns.Count - 1
                    gverdi.Cells(k, t + 1) = cxrili(i)(t)

                Next
                k = k + 1

            Next
            gverdi.Range(gverdi.Cells(1, 1), gverdi.Cells(.Rows.Count + 1, .Columns.Count)).ColumnWidth = 10
            gverdi.Range(gverdi.Cells(1, 1), gverdi.Cells(.Rows.Count + 1, .Columns.Count)).Font.Bold = True
            gverdi.Range(gverdi.Cells(1, 1), gverdi.Cells(.Rows.Count + 1, .Columns.Count)).WrapText = True
            gverdi.Range(gverdi.Cells(1, 1), gverdi.Cells(.Rows.Count + 1, .Columns.Count)).BorderAround2()
            gverdi.Range(gverdi.Cells(1, 1), gverdi.Cells(.Rows.Count + 1, .Columns.Count)).Borders.LineStyle = 1
            gverdi.Range(gverdi.Cells(1, 1), gverdi.Cells(.Rows.Count + 1, .Columns.Count)).Borders.Weight = 1
        End With

        ex.Visible = True

    End Sub
    Sub cxrilinarch()
        cxrili.Clear() : cxrili.Columns.Clear()
        With cxrili
            .Columns.Add("საინვენტარო ნომერი")
            .Columns.Add("დასახელება")
            .Columns.Add("წელი")
            .Columns.Add("საბალანსო ღირებულება")
            .Columns.Add("ცვეთის კოეფიციენტი")
            .Columns.Add("წლიური ცვეთა ")
            .Columns.Add("ნარჩენი ღირებულება")

        End With

    End Sub
    Private Sub Label1_Click(sender As Object, e As EventArgs) Handles Lsab.Click, Lcveta.Click, Ldas.Click, Lnomeri.Click, Lweli.Click, Lnarch.Click

    End Sub

    Private Sub amor_Click(sender As Object, e As EventArgs) Handles amor.Click
        Call kodeqsi()
    End Sub
End Class
