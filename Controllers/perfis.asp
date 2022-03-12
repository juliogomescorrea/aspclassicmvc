<%

	Class clsPerfisController
		Private obj
		Private prvID

		Private Sub Class_Initialize()
			Set obj = New clsPerfisModel
			Set menus = New clsMenusController
			Call menus.Read()
		End Sub

		Public Property Let id(byVal tmpID)
			prvID = tmpID
		End Property

		Public Sub Create()
			obj.perfis.perfil.value = "PROD1"
			obj.perfis.situacao.value = "S"
			Call obj.Create()

			Call Me.Read()
		End Sub

		Public Sub Read()
			If (prvID <> "") Then
				obj.perfis.id_perfil.find = prvID
				Set List = obj.Read()
				Call HTML.DataBind(List, "perfis", "form", "")
			End If

			obj.perfis.id_perfil.find = ""
			Set List = obj.Read()
			Call HTML.DataBind(List, "perfis", "list", "")
		End Sub
		
		Public Sub Update()
			obj.perfis.perfil.value = "PRODUTO 1"
			obj.perfis.situacao.value = "S"
			obj.perfis.id_perfil.find = prvID
			Call obj.Update()

			Call Me.Read()
		End Sub

		Public Sub Delete()
			obj.perfis.id_perfil.find = prvID
			Call obj.Delete()

			Call Me.Read()
		End Sub

		Private Sub Class_Terminate()
			Set obj = Nothing
		End Sub

	End Class

%>