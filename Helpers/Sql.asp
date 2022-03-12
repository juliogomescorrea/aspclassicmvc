<%
	
	Class SQLHelper
		Private prvFrom
		Dim arrJoin()
		Private prvIndexJoin
		Dim arrWhere()
		Private prvIndexWhere
		
		Public Property Let setFrom(ByVal objTable)
			Set prvFrom = objTable
		End Property

		Public Property Get getFrom()
			Set getFrom = prvFrom
		End Property
		
		Public Property Let setJoin(ByVal objTable)
			If (prvIndexJoin&"" = "") Then
				prvIndexJoin = 0
			End If
			Redim Preserve arrJoin(prvIndexJoin)
			Set arrJoin(prvIndexJoin) = objTable
			prvIndexJoin = prvIndexJoin + 1
		End Property
		
		Public Property Get getJoin()
			If (prvIndexJoin&"" <> "") Then
				For i = 0 To UBound(arrJoin)
					tableJoin = arrJoin(i).Schema.table()
					strJoins = strJoins & " INNER JOIN " & tableJoin & " ON " & getJoinON(tableJoin)
				Next
				
				getJoin = strJoins
			End If
		End Property

		Private Function getJoinON(ByVal tableJoin)
			Set object = Me.getFrom
			For i=0 To UBound(object.Schema.columns())
				obj = object.Schema.columns()(i)
				Execute("strForeingKey = object." & obj & ".foreingKey")

				If (InStr(strForeingKey, tableJoin) > 0) Then
					getJoinON = strForeingKey
					Exit For
				End If
			Next
		End Function

		Public Function setWhere(ByVal strName, ByVal strOperator, ByVal strValid)
			If (prvIndexWhere&"" = "") Then
				prvIndexWhere = 0
			End If
			Redim Preserve arrWhere(prvIndexWhere)
			arrWhere(prvIndexWhere) = strName & " " & strOperator & " " & strValid
			prvIndexWhere = prvIndexWhere + 1
		End Function
		
		Public Function getWhere()
			Dim arrParamsFind()
			Dim f : f = 0

			Set object = Me.getFrom
			For i=0 To UBound(object.Schema.columns())
				obj = object.Schema.columns()(i)
				Execute("strFind = object." & obj & ".find")
				If (strFind <> "") Then
					Redim Preserve arrParamsFind(f)
					arrParamsFind(f) = object.schema.columns()(i) & " = " & strFind
					f = f + 1
				End If
			Next

			If (f > 0) Then
				strWhere = Join(arrParamsFind, " AND ")
				
				getWhere = " WHERE " & strWhere
			End If
		End Function
		
		Public Function getSelect()
			getSelect = "SELECT " & Me.getFrom.Schema.table() & "." & Join(Me.getFrom.Schema.columns(), "," & Me.getFrom.Schema.table() & ".") & " FROM " & Me.getFrom.Schema.table() & Me.getJoin() & Me.getWhere()
		End Function
	End Class

%>