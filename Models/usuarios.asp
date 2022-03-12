<%

	Class clsUsuariosModel
		Private conn
		Public usuarios

		Private Sub Class_Initialize()
			Set usuarios = Entities.Management.Usuarios
			usuarios.id_usuario.primaryKey = true

			Set conn = New Database
			conn.connect(strConnection)
		End Sub

		Public Sub Create()
			conn.insert(usuarios)
		End Sub

		Public Function Read()
			Set sql = New SQLHelper
			sql.setFrom = usuarios
			Set Read = conn.read(sql.getSelect())
		End Function

		Public Sub Update()
			conn.update(usuarios)
		End Sub

		Public Sub Delete()
			conn.delete(usuarios)
		End Sub

		Private Sub Class_Terminate()
			Set usuarios = Nothing
			Set conn = Nothing
		End Sub
	End Class

%>