Public Class Form1
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If ComboBox1.SelectedIndex = -1 Or ComboBox2.SelectedIndex = -1 Then
            MsgBox("Seleccione todos los campos")
            Exit Sub
        End If


        If contador < 8 Then
            numeroventa(contador) = contador + 1
            tamaño(contador) = ComboBox1.SelectedItem
            material(contador) = ComboBox2.SelectedItem


            Select Case tamaño(contador)

                Case "Sofa"
                    Select Case material(contador)
                        Case "Cuero"
                            precioManoObra(contador) = Psofa
                            precioCosto(contador) = Math.Round(Val(8 * Pcuero + precioManoObra(contador)), 2)
                        Case "Cuerina"
                            precioManoObra(contador) = Psofa
                            precioCosto(contador) = Math.Round(Val(8 * Pcuerina + precioManoObra(contador)), 2)
                    End Select

                Case "Individual"
                    Select Case material(contador)
                        Case "Cuero"
                            precioManoObra(contador) = Pindividual
                            precioCosto(contador) = Math.Round(Val(3.5 * Pcuero + precioManoObra(contador)), 2)
                        Case "Cuerina"
                            precioManoObra(contador) = Pindividual
                            precioCosto(contador) = Math.Round(Val(3.5 * Pcuerina + precioManoObra(contador)), 2)
                    End Select

                Case "Doble"
                    Select Case material(contador)
                        Case "Cuero"
                            precioManoObra(contador) = Pdoble
                            precioCosto(contador) = Math.Round(Val(6 * Pcuero + precioManoObra(contador)), 2)

                        Case "Cuerina"
                            precioManoObra(contador) = Pdoble
                            precioCosto(contador) = Math.Round(Val(6 * Pcuerina + precioManoObra(contador)), 2)


                    End Select


            End Select


            precioVenta(contador) = Math.Round(Val(precioCosto(contador) * 1.65), 2)

            ListBox1.Items.Add(numeroventa(contador))
            ListBox2.Items.Add(tamaño(contador))
            ListBox3.Items.Add(material(contador))
            ListBox4.Items.Add(precioManoObra(contador))
            ListBox5.Items.Add(precioCosto(contador))
            ListBox6.Items.Add(precioVenta(contador))

            contador = contador + 1



        Else
            MsgBox("Se ha llegado al limite")
        End If
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        salir()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        limpiar()
    End Sub
End Class
