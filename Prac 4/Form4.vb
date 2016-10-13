Public Class Form4
    Dim euros As Double
    Dim dolares As Double
    Dim yenes As Double
    Dim libras As Double
    Dim dragmes As Double
    Dim operacion As Double
    Dim primeraVezDol As Boolean = False
    Dim primeraVezYen As Boolean = False
    Dim primeraVezLib As Boolean = False
    Dim primeraVezDra As Boolean = False
    Dim primeraComa As Boolean = True
    Dim coma As Boolean = False
    Dim contDecimal As Integer = 0


    Private Sub ButtonConvertir_Click(sender As Object, e As EventArgs) Handles ButtonConvertir.Click

        Try
            Select Case ComboBox1.Text
                Case "Dolares"
                    If primeraVezDol = False Then

                        dolares = InputBox("¿A cuanto esta el dolar?")
                        primeraVezDol = True
                        operacion = CDbl(TextBoxNum.Text) * dolares
                        TextBoxNum.Text = Format(operacion, "0.00")

                    Else
                        operacion = CDbl(TextBoxNum.Text) * dolares
                        TextBoxNum.Text = Format(operacion, "0.00")

                    End If

                Case "Yenes"
                    If primeraVezYen = False Then
                        yenes = InputBox("¿A cuanto esta el yen?")
                        primeraVezDol = True
                        operacion = CDbl(TextBoxNum.Text) * yenes
                        TextBoxNum.Text = Format(operacion, "0.00")

                    Else
                        operacion = CDbl(TextBoxNum.Text) * yenes
                        TextBoxNum.Text = Format(operacion, "0.00")

                    End If

                Case "Libras"
                    If primeraVezDol = False Then
                        libras = InputBox("¿A cuanto esta la libra?")
                        primeraVezDol = True
                        operacion = CDbl(TextBoxNum.Text) * libras
                        TextBoxNum.Text = Format(operacion, "0.00")

                    Else
                        operacion = CDbl(TextBoxNum.Text) * libras
                        TextBoxNum.Text = Format(operacion, "0.00")

                    End If

                Case "Dragmes"
                    If primeraVezDra = False Then
                        dragmes = InputBox("¿A cuanto esta el dragma?")
                        primeraVezDol = True
                        operacion = CDbl(TextBoxNum.Text) * dragmes
                        TextBoxNum.Text = Format(operacion, "0.00")

                    Else
                        operacion = CDbl(TextBoxNum.Text) * dragmes
                        TextBoxNum.Text = Format(operacion, "0.00")

                    End If
            End Select
        Catch
            MsgBox("Introduce un valor que se pueda convertir.")
        End Try

    End Sub

    Private Sub ButtonC_Click(sender As Object, e As EventArgs) Handles ButtonC.Click
        TextBoxNum.Text = ""
        coma = False
        contDecimal = 0
    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        If coma = True And contDecimal <= 1 Then
            contDecimal = contDecimal + 1
            TextBoxNum.Text = TextBoxNum.Text + CStr(1)
        End If
        If coma = False And contDecimal <= 2 Then
            TextBoxNum.Text = TextBoxNum.Text + CStr(1)
        End If

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If coma = True And contDecimal <= 1 Then
            contDecimal = contDecimal + 1
            TextBoxNum.Text = TextBoxNum.Text + CStr(2)
        End If
        If coma = False And contDecimal <= 2 Then
            TextBoxNum.Text = TextBoxNum.Text + CStr(2)
        End If
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        If coma = True And contDecimal <= 1 Then
            contDecimal = contDecimal + 1
            TextBoxNum.Text = TextBoxNum.Text + CStr(3)
        End If
        If coma = False And contDecimal <= 2 Then
            TextBoxNum.Text = TextBoxNum.Text + CStr(3)
        End If
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        If coma = True And contDecimal <= 1 Then
            contDecimal = contDecimal + 1
            TextBoxNum.Text = TextBoxNum.Text + CStr(4)
        End If
        If coma = False And contDecimal <= 2 Then
            TextBoxNum.Text = TextBoxNum.Text + CStr(4)
        End If
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        If coma = True And contDecimal <= 1 Then
            contDecimal = contDecimal + 1
            TextBoxNum.Text = TextBoxNum.Text + CStr(5)
        End If
        If coma = False And contDecimal <= 2 Then
            TextBoxNum.Text = TextBoxNum.Text + CStr(5)
        End If
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        If coma = True And contDecimal <= 1 Then
            contDecimal = contDecimal + 1
            TextBoxNum.Text = TextBoxNum.Text + CStr(6)
        End If
        If coma = False And contDecimal <= 2 Then
            TextBoxNum.Text = TextBoxNum.Text + CStr(6)
        End If
    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        If coma = True And contDecimal <= 1 Then
            contDecimal = contDecimal + 1
            TextBoxNum.Text = TextBoxNum.Text + CStr(7)
        End If
        If coma = False And contDecimal <= 2 Then
            TextBoxNum.Text = TextBoxNum.Text + CStr(7)
        End If
    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        If coma = True And contDecimal <= 1 Then
            contDecimal = contDecimal + 1
            TextBoxNum.Text = TextBoxNum.Text + CStr(8)
        End If
        If coma = False And contDecimal <= 2 Then
            TextBoxNum.Text = TextBoxNum.Text + CStr(8)
        End If
    End Sub

    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click
        If coma = True And contDecimal <= 1 Then
            contDecimal = contDecimal + 1
            TextBoxNum.Text = TextBoxNum.Text + CStr(9)
        End If
        If coma = False And contDecimal <= 2 Then
            TextBoxNum.Text = TextBoxNum.Text + CStr(9)
        End If
    End Sub

    Private Sub Button0_Click(sender As Object, e As EventArgs) Handles Button0.Click
        If coma = True And contDecimal <= 1 Then
            contDecimal = contDecimal + 1
            TextBoxNum.Text = TextBoxNum.Text + CStr(0)
        End If

    End Sub

    Private Sub ButtonFlecha_Click(sender As Object, e As EventArgs) Handles ButtonFlecha.Click
        Try
            If TextBoxNum.Text.Last = "," Then
                coma = False
                contDecimal = 0
            End If

            If coma = True Then
                contDecimal = contDecimal - 1
            End If

            TextBoxNum.Text = Mid(TextBoxNum.Text, 1, Len(TextBoxNum.Text) - 1)
        Catch

        End Try
    End Sub

    Private Sub ButtonComa_Click(sender As Object, e As EventArgs) Handles ButtonComa.Click


        If (coma = False And TextBoxNum.Text <> "") Then
            TextBoxNum.Text = TextBoxNum.Text + ","

            coma = True
        End If

    End Sub

    Private Sub ButtonCambiar_Click(sender As Object, e As EventArgs) Handles ButtonCambiar.Click
        Try
            Select Case ComboBox1.Text
                Case "Dolares"
                    dolares = InputBox("¿A cuanto esta el dolar?")
                    operacion = CDbl(TextBoxNum.Text) * dolares
                    TextBoxNum.Text = Format(operacion, "0.00")

                Case "Yenes"
                    yenes = InputBox("¿A cuanto esta el yen?")
                    operacion = CDbl(TextBoxNum.Text) * yenes
                    TextBoxNum.Text = Format(operacion, "0.00")

                Case "Libras"
                    libras = InputBox("¿A cuanto esta la libra?")
                    operacion = CDbl(TextBoxNum.Text) * libras
                    TextBoxNum.Text = Format(operacion, "0.00")

                Case "Dragmes"
                    dragmes = InputBox("¿A cuanto esta el dragma?")
                    operacion = CDbl(TextBoxNum.Text) * dragmes
                    TextBoxNum.Text = Format(operacion, "0.00")

            End Select
        Catch
            MsgBox("Introduce un valor que se pueda convertir.")
        End Try
    End Sub
End Class
