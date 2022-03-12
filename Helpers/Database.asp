<%

	Class Database
		Private objConn
		Dim connectionDisabled
		Dim sql
		
		Private Sub Class_Initialize()
			Set objConn = Server.CreateObject("ADODB.Connection")
		End Sub
		
		Public Sub connect(ByVal strConnection)
			On Error Resume Next
			objConn.Open strConnection
			If (Err.Number&"" <> "0") Then
				connectionDisabled = True

				If (EnabledDebug) Then
					HTML.Debug "Err.Number: " & Err.Number & "<br>Err.Description: " & Err.Description
				End If
			End If
			On Error Goto 0
		End Sub
		
		Public Function rows(ByVal strSQL)
			If (Not connectionDisabled) Then
				Dim rs
				Set rs = objConn.Execute(strSQL)
				If (Not rs.Eof) Then
					rows = rs.getRows()
				Else
					rows = Array()
				End If
				Me.Close(rs)
			End If
		End Function
		
		Public Function read(ByVal strSQL)
			If (Not connectionDisabled) Then
				Dim rs

				If (EnabledDebug) Then
					HTML.debug strSQL
				End If
				Set rs = objConn.Execute(strSQL)
				Set read = rs
			End If
		End Function

		Public Function update(ByVal object)
			If (Not connectionDisabled) Then
				Dim arrParamsValues()
				Dim arrParamsFind()
				
				f = -1
				x = -1
				For i = 0 To UBound(object.schema.columns)
					obj = object.schema.columns()(i)
					Execute("isPrimaryKey = object." & obj & ".primaryKey")
					Execute("strValue = object." & obj & ".value")
					If (Not isPrimaryKey) Then
						x = x + 1
						Redim Preserve arrParamsValues(x)
						arrParamsValues(x) = object.schema.columns()(i) & " = '" & strValue & "'"
					End If

					Execute("strFind = object." & obj & ".find")
					If (strFind <> "") Then
						f = f + 1
						Redim Preserve arrParamsFind(f)
						arrParamsFind(f) = object.schema.columns()(i) & " = " & strFind
					End If
				Next
				
				If (f > -1) Then
					strFind = Utils.iif(UBound(arrParamsFind) >= 0, " WHERE " & Join(arrParamsFind, " AND "), "")
				End If

				sql = "UPDATE " & object.schema.table & " SET " & Join(arrParamsValues, ",") & strFind
				If (EnabledDebug) Then
					HTML.debug sql
				End If
				Me.exec(sql)
			End If						
		End Function

		Public Function insert(ByVal object)
			If (Not connectionDisabled) Then
				Dim arrParamsFields()
				Dim arrParamsValues()
				Dim arrParamsFind()
				
				f = 0
				x = 0
				For i = 0 To UBound(object.schema.columns)
					obj = object.schema.columns()(i)
					Execute("isPrimaryKey = object." & obj & ".primaryKey")
					Execute("strValue = object." & obj & ".value")
					If (Not isPrimaryKey) Then
						Redim Preserve arrParamsFields(x)
						arrParamsFields(x) = obj

						Redim Preserve arrParamsValues(x)
						arrParamsValues(x) = "'" & strValue & "'"
						x = x + 1
					End If

					Execute("strFind = object." & obj & ".find")
					If (strFind <> "") Then
						Redim Preserve arrParamsFind(f)
						arrParamsFind(f) = object.schema.columns()(i) & " = " & strFind
						f = f + 1
					End If
				Next

				sql = "INSERT INTO " & object.schema.table & " (" & Join(arrParamsFields, ",") & ") VALUES (" & Join(arrParamsValues, ",") & ")"
				If (EnabledDebug) Then
					HTML.debug sql
				End If
				Me.exec(sql)
			End If						
		End Function

		Public Function delete(ByVal object)
			If (Not connectionDisabled) Then
				Dim arrParamsFind()
				
				f = -1
				For i = 0 To UBound(object.schema.columns)
					obj = object.schema.columns()(i)
					Execute("strFind = object." & obj & ".find")
					If (strFind <> "") Then
						f = f + 1
						Redim Preserve arrParamsFind(f)
						arrParamsFind(f) = object.schema.columns()(i) & " = " & strFind
					End If
				Next
				
				If (f > -1) Then
					strFind = Utils.iif(UBound(arrParamsFind) >= 0, " WHERE " & Join(arrParamsFind, " AND "), "")
				End If

				sql = "DELETE FROM " & object.schema.table & strFind
				If (EnabledDebug) Then
					HTML.debug sql
				End If
				Me.exec(sql)
			End If
		End Function

		Public Sub exec(ByVal strSQL)
			If (Not connectionDisabled) Then
				On Error Resume Next
				objConn.Execute(strSQL)
				If (Err.Number <> "0") Then
					errMessage = Err.Description
				Else
					errMessage = "Success"
				End If
				On Error Goto 0
			Else
				errMessage = "Connection Error"
			End If

			If (InStr(strSQL, "SIS_logs") = 0) Then
				Call Utils.log(Replace(strSQL, "'", "''"), errMessage)
			End If
		End Sub

		Public Sub Close(ByVal object)
			If (Not connectionDisabled) Then
				If (IsObject(object)) Then
					object.Close()
				End If
			End If
			Set object = Nothing
		End Sub
		
		Private Sub Class_Terminate()
			If (Not connectionDisabled) Then
				Me.Close(objConn)
			End If
		End Sub
		
	End Class

	Set Conn = New Database
	Conn.connect(strConnection)
	
%>