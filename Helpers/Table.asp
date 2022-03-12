<%
	
	Class clsTable
		Private prvTable
		Private prvColumns
		Private prvIndex
		
		Public Property Let table(ByVal strTable)
			prvTable = strTable
		End Property

		Public Property Get table()
			table = prvTable
		End Property

		Public Property Let columns(ByVal arrColumns)
			prvColumns = arrColumns
		End Property

		Public Property Get columns()
			columns = prvColumns
		End Property

		Public Property Let index(ByVal intIndex)
			prvIndex = intIndex
		End Property

		Public Property Get index()
			index = prvIndex
		End Property
	End Class
	
%>