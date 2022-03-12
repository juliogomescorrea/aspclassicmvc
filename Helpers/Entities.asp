<%

	Class clsEntities 

		Public Function Integrations()
			Set Integrations = Utils.LoadEntities("Integrations")
		End Function

		Public Function Management()
			Set Management = Utils.LoadEntities("Management")
		End Function

		Public Function Planning()
			Set Planning = Utils.LoadEntities("Planning")
		End Function

		Public Function Processes()
			Set Processes = Utils.LoadEntities("Processes")
		End Function

	End Class

	Set Entities = New clsEntities

%>