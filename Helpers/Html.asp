<%
	
	Class HTMLHelper
		
		Private prvEnabledDebug
		Private objHTMLDOM
		Private tmpHTMLMain

		Public Property Let EnabledDebug(byVal tmpEnabledDebug)
			prvEnabledDebug = tmpEnabledDebug
		End Property

		Private Sub Class_Initialize()
			Set objHTMLDOM = CreateObject("HTMLFILE")

			Dim objFSO
			Set objFSO = Server.CreateObject("Scripting.FileSystemObject")

			Dim htmlFile

			arquivo = Server.MapPath("master.htm")

			Set htmlFile= objFSO.OpenTextFile(arquivo)
			objHTMLDOM.Write htmlFile.ReadAll

			Set htmlFile = Nothing
			Set objFSO = Nothing
		End Sub
		
		Public Sub debug(ByVal str)
			If (prvEnabledDebug) Then
				tmpHTMLMain = tmpHTMLMain & "<div class=""warning""><span>DEBUG:<br>" & str & "</span></div>" & vbCrLf
			End If
		End Sub

		Public Sub debug_txt(ByVal str)
			Response.Write "<textarea style=""width:100%;height:400px;"">" & str & "</textarea><br/>"
		End Sub

		Public Sub debug_die(ByVal str)
			Response.Write str & "<br>" & vbCrLf
			Response.End
		End Sub

		Public Sub rw(ByVal str)
			tmpHTMLMain = tmpHTMLMain & str
		End Sub

		Public Sub Visible(ByVal id, ByVal bool)
			objHTMLDOM.getElementById(id).Style.Display = Utils.iif(bool, "", "none")
		End Sub

		Public Sub HTML(ByVal id, ByVal txtHTML)
			If (id&"" = "main") Then
				tmpHTMLMain = tmpHTMLMain & txtHTML
			Else
				objHTMLDOM.getElementById(id).innerHTML = txtHTML
			End If
		End Sub

		Public Sub Warning(ByVal id, ByVal txtHTML)
			If (id&"" = "main") Then
				tmpHTMLMain = tmpHTMLMain & "<div class=""warning"">" & txtHTML & "</div>"
			Else
				objHTMLDOM.getElementById(id).innerHTML = "<div class=""warning"">" & txtHTML & "</div>"
			End If
		End Sub

		Public Sub Bug(ByVal id, ByVal txtHTML)
			If (id&"" = "main") Then
				tmpHTMLMain = tmpHTMLMain & "<div class=""bug"">" & txtHTML & "</div>"
			Else
				objHTMLDOM.getElementById(id).innerHTML = "<div class=""bug"">" & txtHTML & "</div>"
			End If
		End Sub

		Public Sub Success(ByVal id, ByVal txtHTML)
			If (id&"" = "main") Then
				tmpHTMLMain = tmpHTMLMain & "<div class=""bug"">" & txtHTML & "</div>"
			Else
				objHTMLDOM.getElementById(id).innerHTML = "<div class=""bug"">" & txtHTML & "</div>"
			End If
		End Sub

		Public Sub Info(ByVal id, ByVal txtHTML)
			If (id&"" = "main") Then
				tmpHTMLMain = tmpHTMLMain & "<div class=""info"">" & txtHTML & "</div>"
			Else
				objHTMLDOM.getElementById(id).innerHTML = "<div class=""info"">" & txtHTML & "</div>"
			End If
		End Sub

		Public Sub DataBind(ByVal List, ByVal controller, ByVal action, ByVal content)
			Dim doc
			Set doc = CreateObject("HTMLFILE")

			Dim objFSO
			Set objFSO = Server.CreateObject("Scripting.FileSystemObject")

			arquivo = Server.MapPath("Views/" & Replace(Replace(controller, "foreach:", ""), "details:", "") & "/" & action & ".htm")

			Dim htmlFile
			Set htmlFile= objFSO.OpenTextFile(arquivo)
			doc.write htmlFile.ReadAll

			Set element 		= doc.getElementById(controller)
			htmlDoc 			= doc.Body.innerHTML
			valid_controller 	= controller
			templatePai 		= element.innerHTML
			template 			= element.innerHTML

			If (IsObject(list)) Then
				If (Not list.Eof) Then
					Do While Not list.Eof
						tempLine = template
						For Each campo In list.Fields
							tempLine = Replace(tempLine, "{" & Replace(Replace(controller, "foreach:", ""), "details:", "") & "." & campo.name & "}", campo)
						Next
						tempFinal = tempFinal & tempLine
						tempLine = ""
					list.MoveNext
					Loop
					list.MoveFirst
				End If
				htmlMain = Replace(Replace(Replace(htmlDoc, Chr(13), ""), Chr(10), ""), Replace(Replace(templatePai, Chr(13), ""), Chr(10), ""), Replace(Replace(tempFinal, Chr(13), ""), Chr(10), ""))
			Else
				htmlMain = Replace(Replace(Replace(htmlDoc, Chr(13), ""), Chr(10), ""), Replace(Replace(templatePai, Chr(13), ""), Chr(10), ""), Replace(Replace(templatePai, Chr(13), ""), Chr(10), ""))
			End If

			If (content&"" = "") Then
				tmpHTMLMain = tmpHTMLMain & htmlMain
			Else
				objHTMLDOM.getElementById(content).innerHTML = htmlMain
			End If

			Set htmlFile = Nothing
			Set objFSO = Nothing
		End Sub

		Public Function Render()
			controller 	= Utils.ASP_GET("controller", "index")
			action 		= Utils.ASP_GET("action", "read")
			id 			= Utils.ASP_GET("id", "")

			If (ForceErrors) Then
				Execute("Set obj = New cls" & controller & "Controller")

				ErrNumber 		= Err.Number
				ErrDescription 	= Err.Description

				If (id&"" <> "") Then
					Execute("obj.id = " & id)
				End If

				Execute("obj." & action)

				If (Err.Number&"" <> "0") Then
					Call Me.Warning("main", "Ops! Página não encontrada ou temporariamente indisponivel!<br><br><span>A página que você procura, foi removida ou está encontrando dificuldades no momento. Tente novamente mais tarde ou entre em contato para reportar o problema.<br><br>Desculpe o transtorno!</span>")
					If (EnabledDebug) Then
						Me.debug("Err.Number: " & ErrNumber & "<br>Err.Description: " & ErrDescription)
					End If
				End If

				If (Conn.connectionDisabled) Then
					Call Me.Warning("main", "Ops! Conexão perdida!<br><br><span>A página que você procura, está encontrando dificuldades no momento. Tente novamente mais tarde ou entre em contato para reportar o problema.<br><br>Desculpe o transtorno!</span>")
				End If

				Set obj 	= Nothing
			Else
				On Error Resume Next
				Execute("Set obj = New cls" & controller & "Controller")

				ErrNumber 		= Err.Number
				ErrDescription 	= Err.Description

				If (id&"" <> "") Then
					Execute("obj.id = " & id)
				End If

				Execute("obj." & action)

				If (Err.Number&"" <> "0") Then
					Call Me.Warning("main", "Ops! Página não encontrada ou temporariamente indisponivel!<br><br><span>A página que você procura, foi removida ou está encontrando dificuldades no momento. Tente novamente mais tarde ou entre em contato para reportar o problema.<br><br>Desculpe o transtorno!</span>")
					If (EnabledDebug) Then
						Me.debug("Err.Number: " & ErrNumber & "<br>Err.Description: " & ErrDescription)
					End If
				End If

				If (Conn.connectionDisabled) Then
					Call Me.Warning("main", "Ops! Conexão perdida!<br><br><span>A página que você procura, está encontrando dificuldades no momento. Tente novamente mais tarde ou entre em contato para reportar o problema.<br><br>Desculpe o transtorno!</span>")
				End If
				On Error Goto 0

				Set obj 	= Nothing
			End If
		End Function
		
		Private Sub Class_Terminate()
			If (tmpHTMLMain&"" <> "") Then
				objHTMLDOM.getElementById("main").innerHTML = tmpHTMLMain
			End If

			Response.Write "<!DOCTYPE html PUBLIC ""-//W3C//DTD XHTML 1.0 Transitional//EN"" ""http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd"">"
			Response.Write "<html xmlns=""http://www.w3.org/1999/xhtml"">"
			Response.Write "<head>"
			Response.write objHTMLDOM.documentElement.childNodes.item(0).innerHTML
			Response.Write "</head>"
			Response.Write "<body>"
			Response.write objHTMLDOM.Body.innerHTML
			Response.Write "</body>"
			Response.Write "</html>"
			Set objHTMLDOM = Nothing
		End Sub
		
	End Class
	
	Set HTML = New HTMLHelper
	HTML.EnabledDebug = EnabledDebug

%>