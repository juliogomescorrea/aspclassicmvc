<%

	Class clsMenusController
		Private obj
		Private prvID

		Private Sub Class_Initialize()
			Set obj = New clsMenusModel
		End Sub

		Public Property Let id(byVal tmpID)
			prvID = tmpID
		End Property

		'Public Sub Create()
			'obj.menus.nome.value = Request.Form("nome")
			'obj.menus.email.value = Request.Form("email")
			'obj.menus.senha.value = "123"
			'obj.menus.situacao.value = "A"
			'obj.menus.id_usuario_insercao.value = 1
			'obj.menus.id_perfil.value = 1
			'Call obj.Create()
			
			'Call Me.Read()
		'End Sub

		Public Sub Read()
			'If (prvID <> "") Then
				'obj.menus.id_menu.find = prvID
				'Set list = obj.Read()
				'Call HTML.DataBind(list, "menus", "form", "")
			'End If

			'obj.menus.id_menu.find = ""
			Set list = obj.Read()
			Call HTML.DataBind(list, "menus", "list", "menu-principal")
		End Sub
		
		'Public Sub Update()
			'obj.menus.nome.value = Request.Form("nome")
			'obj.menus.email.value = Request.Form("email")
			'obj.menus.senha.value = "123"
			'obj.menus.situacao.value = "A"
			'obj.menus.id_menu.find = prvID
			'Call obj.Update()
		'End Sub

		'Public Sub Delete()
			'obj.menus.id_menu.find = prvID
			'Call obj.Delete()

			'Call Me.Read()
		'End Sub

		Private Sub Class_Terminate()
			Set obj = Nothing
		End Sub

	End Class

%>