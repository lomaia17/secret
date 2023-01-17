Imports Excel = Microsoft.Office.Interop.Excel
Public Class Form1

    Dim Cxrili As New DataTable
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub Ramor_Click(sender As Object, e As EventArgs) Handles Ramor.Click
        Call Atanabari()
    End Sub

    Sub Atanabari()
        Dim k, vada As Int16
        Dim sab, lik, weli As Integer
        Dim amor, nar As Single

        Call satauri()

        vada = Vvada.Text : lik = Vlik.Text : sab = Vsab.Text : weli = Vweli.Text

        amor = (sab - lik) / vada

        For i = weli To Today.Year
            nar = sab - amor * k
            Cxrili.Rows.Add(New String() {Vnomeri.Text, Vdas.Text, i, sab, lik, vada, amor, nar})
            k = k + 1
        Next
        DGsia.DataSource = Cxrili
    End Sub

    Sub satauri()
        Cxrili.Clear() : Cxrili.Columns.Clear()
        With Cxrili
            .Columns.Add("საინვენტარო ნომერი")
            .Columns.Add("დასახელება")
            .Columns.Add("წელი")
            .Columns.Add("საბალანსო")
            .Columns.Add("სალიკვიდაციო")
            .Columns.Add("ვადა")
            .Columns.Add("ამორტიზაცია")
            .Columns.Add("ნარჩენი")
        End With
    End Sub

    Sub Sedegi()
        Dim k, p As Integer
        Dim Ex = New Excel.Application
        Dim gverdi As Excel.Worksheet
        Ex.Workbooks.Add()
        gverdi = Ex.Worksheets.Add
        gverdi.Name = "ოქმი"
        k = 3 : p = 1
        With Cxrili
            gverdi.Range("A1:G1").Merge()
            gverdi.Range("A1:G1").Value = "თანაბარი ამორტიზაცია"
            gverdi.Range("A1:G1").Style.Font.Bold = True
            For t = 0 To Cxrili.Columns.Count - 1
                gverdi.Cells(2, t + 1) = Cxrili.Columns(t).Caption
            Next
            For i = 0 To Cxrili.Rows.Count - 1
                For t = 0 To .Columns.Count - 1
                    gverdi.Cells(k, t + 1) = Cxrili(i)(t)
                Next
                k = k + 1
            Next
        End With
        Ex.Visible = True
    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DGsia.CellContentClick

    End Sub

    Private Sub Dabechdva_Click(sender As Object, e As EventArgs) Handles Dabechdva.Click
        Call Sedegi()
    End Sub
End Class
