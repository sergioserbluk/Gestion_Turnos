Imports MySqlConnector
Public Class frmPacientes
    Dim conexiondb As New Conexion
    Private Sub CargarPacientes()
        Dim sql As String = "select * from Pacientes"
        Using conexion = conexiondb.ObtenerConexion()
            conexion.Open()
            Dim adaptador As New MySqlDataAdapter(sql, conexion)
            Dim tabla As New DataTable()
            adaptador.Fill(tabla)
            dgvPacientes.DataSource = tabla
        End Using
    End Sub
    Private Sub LimpiarCuadros()
        txtDni.Clear()
        txtApellido.Clear()
        txtNombre.Clear()
        txtTelefono.Clear()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If txtApellido.Text = "" Or txtNombre.Text = "" Or txtDni.Text = "" Or txtTelefono.Text = "" Then
            MsgBox("Debe completar todos los campos")
            txtDni.Focus()
            Exit Sub
        End If
        Dim sql As String = "insert into Pacientes (dni,nombre,apellido,telefono) values (@dni, @nombre,@apellido,@telefono)"
        Using conexion = conexiondb.ObtenerConexion()
            conexion.Open()
            Dim comando As New MySqlCommand(sql, conexion)
            comando.Parameters.AddWithValue("@dni", txtDni.Text)
            comando.Parameters.AddWithValue("@nombre", txtNombre.Text)
            comando.Parameters.AddWithValue("@apellido", txtApellido.Text)
            comando.Parameters.AddWithValue("@telefono", txtTelefono.Text)
            comando.ExecuteNonQuery()
            LimpiarCuadros()
            MsgBox("El paciente se guardo correctamente!")

        End Using
        CargarPacientes()
    End Sub

    Private Sub frmPacientes_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        CargarPacientes()
    End Sub

    Private Sub dgvPacientes_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles dgvPacientes.CellContentClick
        If e.RowIndex >= 0 Then
            Dim fila As DataGridViewRow = dgvPacientes.Rows(e.RowIndex)
            txtDni.Text = fila.Cells("dni").Value.ToString()
            txtNombre.Text = fila.Cells("nombre").Value.ToString()
            txtApellido.Text = fila.Cells("apellido").Value.ToString()
            txtTelefono.Text = fila.Cells("telefono").Value.ToString()
        End If
        Button1.Enabled = False
        Button2.Enabled = True
        Button3.Enabled = True

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim sql As String = "update pacientes set nombre=@nombre, apellido=@apellido, telefono=@telefono where dni=@dni"
        Using Conexion = conexiondb.ObtenerConexion()
            Conexion.Open()
            Dim comando As New MySqlCommand(sql, Conexion)
            comando.Parameters.AddWithValue("@dni", txtDni.Text)
            comando.Parameters.AddWithValue("@nombre", txtNombre.Text)
            comando.Parameters.AddWithValue("@apellido", txtApellido.Text)
            comando.Parameters.AddWithValue("@telefono", txtTelefono.Text)
            comando.ExecuteNonQuery()
            LimpiarCuadros()
            MsgBox("Los datos se catualizaron correctamente!")
        End Using
        CargarPacientes()
        Button1.Enabled = True
        Button2.Enabled = False
        Button3.Enabled = False
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Dim resp As Integer = MsgBox("¿Desea eliminar el paciente?" + txtApellido.Text + ", " + txtNombre.Text + " ?", MsgBoxStyle.YesNo, "Eliminar paciente")
        If resp = 6 Then
            Dim sql As String = "delete from pacientes where dni=@dni"
            Using Conexion = conexiondb.ObtenerConexion()
                Conexion.Open()
                Dim comando As New MySqlCommand(sql, Conexion)
                comando.Parameters.AddWithValue("@dni", txtDni.Text)
                comando.ExecuteNonQuery()
                LimpiarCuadros()
                MsgBox("El paciente se elimino correctamente!")
            End Using
        End If
        CargarPacientes()
        Button1.Enabled = True
        Button2.Enabled = False
        Button3.Enabled = False
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        frmprincipal.Show()
        Me.Close()
    End Sub
End Class