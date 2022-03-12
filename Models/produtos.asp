<%

	Class clsProdutosModel
		Private conn
		Public Produtos

		Private Sub Class_Initialize()
			Set Produtos = Entities.Management.Products
			Produtos.id_product.primaryKey = true

			Set conn = New Database
			conn.connect(strConnection)
		End Sub

		Public Sub Create()
			conn.insert(Produtos)
		End Sub

		Public Function Read()
			Set sql = New SQLHelper
			sql.setFrom = Produtos
			Set Read = conn.read(sql.getSelect())
		End Function

		Public Sub Update()
			conn.update(Produtos)
		End Sub

		Public Sub Delete()
			conn.delete(Produtos)
		End Sub

		Private Sub Class_Terminate()
			Set Produtos = Nothing
			Set conn = Nothing
		End Sub
	End Class

%>