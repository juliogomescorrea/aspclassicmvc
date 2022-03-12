<%

	Class clsIndexController

		Public Sub Read()
			HTML.Visible "banner", True

			Set menus = New clsMenusController
			Call menus.Read()
		End Sub

	End Class

%>