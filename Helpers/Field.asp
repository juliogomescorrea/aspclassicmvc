<%
	
	Class clsField
		Dim prvValue()
		Dim prvType()
		Dim prvFind()
		Dim prvName()
		Dim prvTable()
		Dim prvPrimaryKey()
		Dim prvForeingKey()
		Dim prvObject()
		Private prvIndex

		Public Property Let Index(ByVal tmpIndex)
			prvIndex = tmpIndex
		End Property

		Public Property Get Index()
			Index = prvIndex
		End Property

		Public Property Let value(ByVal strValue)
			Redim Preserve prvValue(Me.Index)
			prvValue(Me.Index) = strValue
		End Property

		Public Default Property Get value()
			On Error Resume Next
			value = prvValue(Me.Index)
			On Error Goto 0
		End Property

		Public Property Let Object(ByVal objObject)
			Redim Preserve prvObject(Me.Index)
			Set prvObject(Me.Index) = objObject
		End Property

		Public Property Get Object()
			On Error Resume Next
			Object = prvObject(Me.Index)
			On Error Goto 0
		End Property

		Public Property Let Sqltype(ByVal strType)
			Redim Preserve prvType(Me.Index)
			prvType(Me.Index) = strType
		End Property

		Public Property Get Sqltype()
			Sqltype = prvType(Me.Index)
		End Property

		Public Property Let name(ByVal strName)
			Redim Preserve prvName(Me.Index)
			prvName(Me.Index) = strName
		End Property

		Public Property Get name()
			name = prvName(Me.Index)
		End Property

		Public Property Let table(ByVal strTable)
			Redim Preserve prvTable(Me.Index)
			prvTable(Me.Index) = strTable
		End Property

		Public Property Get table()
			table = prvTable(Me.Index)
		End Property

		Public Property Let Find(ByVal strFind)
			Redim Preserve prvFind(Me.Index)
			prvFind(Me.Index) = strFind
		End Property

		Public Property Get Find()
			On Error Resume Next
			Find = prvFind(Me.Index)
			On Error Goto 0
		End Property

		Public Property Let primaryKey(ByVal trueFalse)
			Redim Preserve prvPrimaryKey(Me.Index)
			prvPrimaryKey(Me.Index) = trueFalse
		End Property

		Public Property Get primaryKey()
			On Error Resume Next
			primaryKey = prvPrimaryKey(Me.Index)
			On Error Goto 0
		End Property

		Public Property Let foreingKey(ByVal objTable)
			Redim Preserve prvForeingKey(Me.Index)
			Set prvForeingKey(Me.Index) = objTable
		End Property

		'Public Property Let foreingKey(ByVal table)
		'	Redim Preserve prvForeingKey(Me.Index)
		'	prvForeingKey(Me.Index) = table & "." & Me.name & " = " & Me.table & "." & Me.name
		'End Property

		Public Property Get foreingKey()
			On Error Resume Next
			Set foreingKey = prvForeingKey(Me.Index)
			On Error Goto 0
		End Property

		Private Sub Class_Terminate()
			Redim prvValue(0)
			Redim prvType(0)
			Redim prvFind(0)
			prvIndex = ""
		End Sub
	End Class
	
%>