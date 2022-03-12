<%

	Class clsReferencesModel
		Private conn
		Public References

		Private Sub Class_Initialize()
			Set References = Entities.Management.Refs
			References.id_reference.primaryKey = true

			Set conn = New Database
			conn.connect(strConnection)
		End Sub

		Public Sub Create()
			conn.insert(References)
		End Sub

		Public Function Read()
			Set sql = New SQLHelper
			sql.setFrom = References
			Set Read = conn.read(sql.getSelect())
		End Function

		Public Sub Update()
			conn.update(References)
		End Sub

		Public Sub Delete()
			conn.delete(References)
		End Sub

		Private Sub Class_Terminate()
			Set References = Nothing
			Set conn = Nothing
		End Sub
	End Class

%>