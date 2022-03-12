<%

	Class clsMenusModel
		Private conn
		Public menus

		Private Sub Class_Initialize()
			Set menus = Entities.Management.Menus
			menus.id_menu.primaryKey = true

			Set conn = New Database
			conn.connect(strConnection)
		End Sub

		Public Sub Create()
			conn.insert(menus)
		End Sub

		Public Function Read()
			Set sql = New SQLHelper
			sql.setFrom = menus
			Set Read = conn.read(sql.getSelect())
		End Function

		Public Sub Update()
			conn.update(menus)
		End Sub

		Public Sub Delete()
			conn.delete(menus)
		End Sub

		Private Sub Class_Terminate()
			Set menus = Nothing
			Set conn = Nothing
		End Sub
	End Class

%>