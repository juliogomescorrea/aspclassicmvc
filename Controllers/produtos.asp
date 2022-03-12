<%

	Class clsProdutosController
		Private obj
		Private prvID

		Private Sub Class_Initialize()
			Set obj = New clsProdutosModel
			'Set menus = New clsMenusController
			'Call menus.Read()
		End Sub

		Public Property Let id(byVal tmpID)
			prvID = tmpID
		End Property

		Public Sub Create()
			obj.Produtos.name.value = Utils.ASP_POST("nome")
			Call obj.Create()
			
			Call Me.Read()
		End Sub

		Public Sub Read()
			If (prvID <> "") Then
				obj.Produtos.id_product.find = prvID
				Set list = obj.Read()
				Call HTML.DataBind(list, "Produtos", "form", "")
			End If

			obj.Produtos.id_product.find = ""
			Set list = obj.Read()
			Call HTML.DataBind(list, "Produtos", "list", "")
		End Sub
		
		Public Sub Update()
			obj.Produtos.name.value = Request.Form("nome")
			obj.Produtos.id_product.find = prvID
			Call obj.Update()
		End Sub

		Public Sub Delete()
			obj.Produtos.id_product.find = prvID
			Call obj.Delete()

			Call Me.Read()
		End Sub

		Private Sub Class_Terminate()
			Set obj = Nothing
		End Sub

	End Class

%>