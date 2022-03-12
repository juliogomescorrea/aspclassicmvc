<%

	Class clsUtils

		Public Function ASP_GET(ByVal param, ByVal valueDefault)
			Dim temp : temp = Trim(Replace(Request.QueryString(param), "'", "''"))
			ASP_GET = Me.iif(temp&"" <> "", temp, valueDefault) 
		End Function

		Public Function ASP_POST(ByVal param, ByVal valueDefault)
			Dim temp : temp = Trim(Replace(Request.Form(param), "'", "''"))
			ASP_POST = Me.iif(temp&"" <> "", temp, valueDefault) 
		End Function

		Public Function ASP_REQUEST(ByVal param, ByVal valueDefault)
			Dim temp : temp = Trim(Replace(Request(param), "'", "''"))
			ASP_REQUEST = Me.iif(temp&"" <> "", temp, valueDefault) 
		End Function

		Public Function iif(ByVal condicao, ByVal retornoTrue, ByVal retornoFalse)
			If (condicao) Then
				iif = retornoTrue
			Else
				iif = retornoFalse
			End If
		End Function

		Public Sub log(ByVal sql, ByVal result)
			sql = "INSERT INTO SIS_logs (sql, resultado) VALUES ('" & sql & "', '" & result & "');"
			conn.exec(sql)
		End Sub

		Public Function RunShellScript(ByVal strCmd)
			Set LogMessages = Entities.LogMessages

			On Error Resume Next
			Set wScript = Server.CreateObject("WScript.Shell")
			result = wScript.Run(strCmd, 0, True)
			If (result = 0) Then
				LogMessages.Number.value = 0
				LogMessages.Description.value = "Success: Comando Shell Script processado com sucesso!"
			Else
				LogMessages.Number.value = 1
				LogMessages.Description.value = "Error: Erro nao identificado ao processar comando Shell Script! strCmd: " & strCmd
			End If

			If (Err.Number <> 0) Then
				LogMessages.Number.value = 1
				LogMessages.Description.value = "Error: " & Server.HTMLEncode(Err.Description)
			End If
			On Error Goto 0

			Set RunShellScript = LogMessages
		End Function

		Public Function LoadEntities(ByVal Entitie)
			pathEntities = Server.MapPath("/Entities")
			Set objFso = Server.CreateObject("Scripting.FileSystemObject")
			Set objFile = objFso.OpenTextFile(pathEntities & "\" & Entitie & ".ent")
			Set LoadEntities = serializeEntitieColection(formatObjects(objFile.ReadAll))
			Set objFile = Nothing
			Set objFso = Nothing
		End Function

		Public Function formatObjects(ByVal script)
			formatObjects = Replace(Replace(Replace(script, vbCrLf, ""), vbTab, ""), " ", "")
		End Function

	End Class

	Set Utils = New clsUtils

%>