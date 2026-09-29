Imports MysqlConnector

Public Class Conexion
    Private cadena As String = "Server=localhost; Database=Clinica; Uid=root; Pwd=;"
    Public Function ObtenerConexion() As MySqlConnection

        Return New MySqlConnection(cadena)
    End Function
End Class
