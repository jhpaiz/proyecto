Imports System.Math
Public Class Form1
    Dim total As Integer
    Dim muestra(0) As Integer
    Dim frecuencia(10) As Integer
    Dim frecacum(10) As Integer
    Dim marca(10) As Double
    Dim frecpormarca(11) As Double
    Dim marcamedia(11) As Double
    Dim i As Integer = 0
    Dim media, lri, lri1, mediana, moda, dv As Double
    Dim fx, f, fai, fai1, fas As Integer

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If Val(TextBox1.Text) Then
            total = TextBox1.Text
            TextBox1.ReadOnly = True
            ReDim muestra(total)
            Label3.Visible = True
            TextBox2.Visible = True
            ListBox1.Visible = True
            Button1.Enabled = False
        Else : MsgBox("Dato incorrecto", MsgBoxStyle.OkOnly)
            TextBox1.Clear()
            TextBox1.Focus()
        End If

    End Sub

    Private Sub TextBox2_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextBox2.KeyDown
        Select Case e.KeyData
            Case Keys.Enter
                If Val(TextBox2.Text) Then
                    If TextBox2.Text >= 1 And TextBox2.Text <= 100 Then
                        muestra(i) = TextBox2.Text
                        ListBox1.Items.Add(TextBox2.Text)
                        i = i + 1
                    Else : MsgBox("Fuera de rango", MsgBoxStyle.OkOnly)
                    End If
                Else : MsgBox("Dato incorrecto", MsgBoxStyle.OkOnly)
                End If
                If i = total Then
                    MsgBox("Ingreso completado, se ordenaran los datos", MsgBoxStyle.OkOnly)
                    Array.Sort(muestra)
                    ListBox1.Items.Clear()
                    For j As Integer = 1 To total Step 1
                        ListBox1.Items.Add(muestra(j))
                    Next
                    MsgBox("Datos ordenados", MsgBoxStyle.OkOnly)
                    Label3.Visible = False
                    TextBox2.Visible = False
                    Label4.Visible = True
                    Label5.Visible = True
                    Label6.Text = muestra(total)
                    Label7.Text = muestra(1)
                    Label6.Visible = True
                    Label7.Visible = True
                    Button3.Visible = True
                End If
                TextBox2.Clear()
                TextBox2.Focus()
        End Select
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Label8.Visible = True
        Label9.Visible = True
        Label10.Visible = True
        Label11.Visible = True
        Label16.Visible = True
        Label17.Visible = True
        Label18.Visible = True
        Label19.Visible = True
        Label20.Visible = True
        ListBox2.Visible = True
        ListBox3.Visible = True
        ListBox4.Visible = True
        ListBox5.Visible = True
        For x As Integer = 1 To total Step 1
            Select Case muestra(x)
                Case 1 To 10
                    frecuencia(1) = frecuencia(1) + 1
                Case 11 To 20
                    frecuencia(2) = frecuencia(2) + 1
                Case 21 To 30
                    frecuencia(3) = frecuencia(3) + 1
                Case 31 To 40
                    frecuencia(4) = frecuencia(4) + 1
                Case 41 To 50
                    frecuencia(5) = frecuencia(5) + 1
                Case 51 To 60
                    frecuencia(6) = frecuencia(6) + 1
                Case 61 To 70
                    frecuencia(7) = frecuencia(7) + 1
                Case 71 To 80
                    frecuencia(8) = frecuencia(8) + 1
                Case 81 To 90
                    frecuencia(9) = frecuencia(9) + 1
                Case 91 To 100
                    frecuencia(10) = frecuencia(10) + 1
            End Select
        Next
        For y As Integer = 1 To 10 Step 1
            frecacum(y) = frecuencia(y) + frecacum(y - 1)
        Next
        i = 0
        For z As Integer = 1 To 100 Step 10
            i = i + 1
            marca(i) = (z + (z + 9)) / 2
            ListBox5.Items.Add(z & " - " & z + 9)
        Next
        ListBox1.Items.Clear()
        ListBox2.Items.Clear()
        For k As Integer = 1 To 10 Step 1
            ListBox1.Items.Add(frecuencia(k))
            ListBox2.Items.Add(frecacum(k))
            ListBox3.Items.Add(marca(k))
            frecpormarca(k) = frecuencia(k) * marca(k)
            frecpormarca(11) = frecpormarca(11) + frecpormarca(k)
        Next
        media = frecpormarca(11) / total
        For k As Integer = 1 To 10 Step 1
            marcamedia(k) = frecuencia(k) * ((marca(k) - media) ^ 2)
            marcamedia(11) = marcamedia(11) + marcamedia(k)
            ListBox4.Items.Add(marcamedia(k))
        Next
        Select Case media
            Case 1 To 10
                lri = 0.5
                fai = frecacum(0)
                fx = frecuencia(1)
            Case 11 To 20
                lri = 10.5
                fai = frecacum(1)
                fx = frecuencia(2)

                lri1 = 0.5
                f = frecuencia(1)
                fai1 = frecuencia(0)
                fas = frecuencia(2)
            Case 21 To 30
                lri = 20.5
                fai = frecuencia(2)
                fx = frecuencia(3)

                lri1 = 10.5
                f = frecuencia(2)
                fai1 = frecuencia(1)
                fas = frecuencia(3)
            Case 31 To 40
                lri = 30.5
                fai = frecacum(3)
                fx = frecuencia(4)

                lri1 = 20.5
                f = frecuencia(3)
                fai1 = frecuencia(2)
                fas = frecuencia(4)
            Case 41 To 50
                lri = 40.5
                fai = frecacum(4)
                fx = frecuencia(5)

                lri1 = 30.5
                f = frecuencia(4)
                fai1 = frecuencia(3)
                fas = frecuencia(5)
            Case 51 To 60
                lri = 50.5
                fai = frecacum(5)
                fx = frecuencia(6)

                lri1 = 40.5
                f = frecuencia(5)
                fai1 = frecuencia(4)
                fas = frecuencia(6)
            Case 61 To 70
                lri = 60.5
                fai = frecacum(6)
                fx = frecuencia(7)

                lri1 = 50.5
                f = frecuencia(6)
                fai1 = frecuencia(5)
                fas = frecuencia(7)
            Case 71 To 80
                lri = 70.5
                fai = frecacum(7)
                fx = frecuencia(8)

                lri1 = 60.5
                f = frecuencia(7)
                fai1 = frecuencia(6)
                fas = frecuencia(8)

            Case 81 To 90
                lri = 80.5
                fai = frecacum(8)
                fx = frecuencia(9)

                lri1 = 70.5
                f = frecuencia(8)
                fai1 = frecuencia(7)
                fas = frecuencia(9)
            Case 91 To 100
                lri = 90.5
                fai = frecacum(9)
                fx = frecuencia(10)

                lri1 = 80.5
                f = frecuencia(9)
                fai1 = frecuencia(8)
                fas = frecuencia(10)
        End Select
        mediana = lri + (((total / 2) - fai) / fx) * 10
        moda = lri1 + (f - fai1) / ((f - fai1) + (f - fas)) * 10
        dv = Sqrt(marcamedia(11) / total)
        Label12.Text = media
        Label13.Text = mediana
        Label14.Text = moda
        Label15.Text = dv
        Button3.Enabled = False
    End Sub
End Class
