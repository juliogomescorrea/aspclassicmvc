<%

	Class clsReferencesController
		Private obj
		Private prvID

		Private Sub Class_Initialize()
			Set obj = New clsReferencesModel
			Set menus = New clsMenusController
			Call menus.Read()
		End Sub

		Public Property Let id(byVal tmpID)
			prvID = tmpID
		End Property

		Public Sub Create()
			obj.References.reference.value 				= Utils.ASP_POST("reference", "")
			obj.References.status.value 				= Utils.ASP_POST("status", "")
			Call obj.Create()
			
			Call Me.Read()
		End Sub

		Public Sub Read()
			If (prvID <> "") Then
				obj.References.id_reference.find = prvID
				Set list = obj.Read()
			End If

			obj.References.id_reference.find = ""
			Set list = obj.Read()
			Call HTML.DataBind(list, "references", "list", "")
		End Sub
		
		Public Sub Update()
			obj.References.reference.value 			= Utils.ASP_POST("reference", "")
			obj.References.status.value 			= Utils.ASP_POST("status", "")
			obj.References.id_reference.find 		= prvID
			Call obj.Update()
		End Sub

		Public Sub Delete()
			obj.References.id_reference.find = prvID
			Call obj.Delete()

			Call Me.Read()
		End Sub

		Private Sub Class_Terminate()
			Set obj = Nothing
		End Sub

	End Class

%>