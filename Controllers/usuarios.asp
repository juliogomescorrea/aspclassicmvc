<%

	Class clsUsuariosController
		Private obj
		Private prvID

		Private Sub Class_Initialize()
			Set obj = New clsUsuariosModel
		End Sub

		Public Property Let id(byVal tmpID)
			prvID = tmpID
		End Property

		Public Sub Create()
			obj.usuarios.nome.value 				= Utils.ASP_POST("nome", "")
			obj.usuarios.email.value 				= Utils.ASP_POST("email", "")
			obj.usuarios.senha.value 				= Utils.ASP_POST("senha", "")
			obj.usuarios.situacao.value 			= Utils.ASP_POST("situacao", "")
			obj.usuarios.id_perfil.value 			= Utils.ASP_POST("id_perfil", "")
			obj.usuarios.id_usuario_insercao.value 	= 1
			Call obj.Create()
			
			Call Me.Read()
		End Sub

		Public Sub Read()
			If (prvID <> "") Then
				obj.usuarios.id_usuario.find = prvID
				Set list = obj.Read()
			End If

			Call HTML.DataBind(list, "details:usuarios", "form", "")

			obj.usuarios.id_usuario.find = ""
			Set list = obj.Read()
			Call HTML.DataBind(list, "usuarios", "list", "")
		End Sub
		
		Public Sub Update()
			obj.usuarios.nome.value 			= Utils.ASP_POST("nome", "")
			obj.usuarios.email.value 			= Utils.ASP_POST("email", "")
			obj.usuarios.senha.value 			= Utils.ASP_POST("senha", "")
			obj.usuarios.situacao.value 		= Utils.ASP_POST("situacao", "")
			obj.usuarios.id_usuario.find 		= prvID
			Call obj.Update()
		End Sub

		Public Sub Delete()
			obj.usuarios.id_usuario.find = prvID
			Call obj.Delete()

			Call Me.Read()
		End Sub

		Private Sub Class_Terminate()
			Set obj = Nothing
		End Sub

	End Class

%>