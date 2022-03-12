<%
	
	If (LocalSystem) Then

		Dim objFSO
		Set objFSO 			= Server.CreateObject("Scripting.FileSystemObject")

		'CONTROLLERS
		Dim fileControllers
		Dim folderControllers
		Dim tmpHTMLFileControllers
		Dim objFolderControllers

		fileControllers 	= Server.MapPath("/Helpers/controllers.asp")
		folderControllers 	= Server.MapPath("/Controllers/")

		Set objFolderControllers = objFSO.GetFolder(folderControllers)
		For Each Files In objFolderControllers.Files
			tmpHTMLFileControllers  = tmpHTMLFileControllers & "<!--#include file=""../Controllers/" & Files.Name & """-->" & vbCrLf
		Next

		Dim objFileControllers
		Set objFileControllers = objFSO.OpenTextFile(fileControllers,2,True)
		objFileControllers.Write tmpHTMLFileControllers

		Set objFileControllers = Nothing
		Set objFolderControllers = Nothing

		'MODELS
		Dim fileModels
		Dim folderModels 
		Dim tmpHTMLFileModels
		Dim objFolderModels

		fileModels 			= Server.MapPath("/Helpers/models.asp")
		folderModels		= Server.MapPath("/Models/")

		Set objFolderModels = objFSO.GetFolder(folderModels)
		For Each Files In objFolderModels.Files
			tmpHTMLFileModels  = tmpHTMLFileModels & "<!--#include file=""../Models/" & Files.Name & """-->" & vbCrLf
		Next

		Dim objFileModels
		Set objFileModels = objFSO.OpenTextFile(fileModels,2,True)
		objFileModels.Write tmpHTMLFileModels

		Set objFileModels = Nothing
		Set objFolderModels = Nothing

		'ENTITIES
		Dim fileEntities
		Dim folderEntities 
		Dim tmpFileEntities
		Dim objFolderEntities

		fileEntities 		= Server.MapPath("/Helpers/Entities.asp")
		folderEntities		= Server.MapPath("/Entities/")

		scriptFile = 	chr(60) & "%" & vbCrLf & vbCrLf &_
						vbTab & "Class clsEntities " & vbCrLf & vbCrLf &_
						"{templates}" &_
						vbTab & "End Class" & vbCrLf & vbCrLf &_
						vbTab & "Set Entities = New clsEntities" & vbCrLf & vbCrLf &_
						"%" & chr(62)

		Set objFolderEntities = objFSO.GetFolder(folderEntities)
		For Each Files In objFolderEntities.Files
			tmpFileEntities  = tmpFileEntities & 	vbTab & vbTab & "Public Function " & Replace(Files.Name, ".ent", "") & "()" & vbCrLf &_
													vbTab & vbTab & vbTab & "Set " & Replace(Files.Name, ".ent", "") & " = Utils.LoadEntities(""" & Replace(Files.Name, ".ent", "") & """)" & vbCrLf  &_
													vbTab & vbTab & "End Function" & vbCrLf & vbCrLf
		Next

		Dim objFileEntities
		Set objFileEntities = objFSO.OpenTextFile(fileEntities,2,True)
		objFileEntities.Write Replace(scriptFile, "{templates}", tmpFileEntities)

		Set objFileEntities = Nothing
		Set objFolderEntities = Nothing

		Set objFSO = Nothing

	End If

%>