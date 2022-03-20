Module Module1
    Public Const Psofa = 205.99
    Public Const Pindividual = 150.99
    Public Const Pdoble = 200.99

    Public Const Pcuero = 75.0
    Public Const Pcuerina = 45.99

    Public numeroventa(8)
    Public tamaño(8)
    Public material(8)
    Public precio(8)
    Public precioCosto(8)
    Public precioManoObra(8)
    Public precioVenta(8)

    Public contador = 0
    Sub limpiar()
        Form1.ListBox1.Items.Clear()
        Form1.ListBox2.Items.Clear()
        Form1.ListBox3.Items.Clear()
        Form1.ListBox4.Items.Clear()
        Form1.ListBox5.Items.Clear()
        Form1.ListBox6.Items.Clear()
    End Sub

    Sub salir()
        If (MsgBox("Esta seguro que desea salir ", vbQuestion + vbYesNo, "MENSAJE DE SALIDA") = vbYes) Then
            End
        Else
            limpiar()
        End If
    End Sub



End Module
