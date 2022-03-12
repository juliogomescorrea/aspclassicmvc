<%

	Class clsPerfisModel
		Private conn
		Public perfis

		Private Sub Class_Initialize()
			Set perfis = Entities.Management.Perfis
			perfis.id_perfil.primaryKey = true

			Set conn = New Database
			conn.connect(strConnection)
		End Sub

		Public Sub Create()
			conn.insert(perfis)
		End Sub

		Public Function Read()
			Set sql = New SQLHelper
			sql.setFrom = perfis
			Set Read = conn.read(sql.getSelect())
		End Function

		Public Sub Update()
			conn.update(perfis)
		End Sub

		Public Sub Delete()
			conn.delete(perfis)
		End Sub

		Private Sub Class_Terminate()
			Set perfis = Nothing
			Set conn = Nothing
		End Sub
	End Class

%>