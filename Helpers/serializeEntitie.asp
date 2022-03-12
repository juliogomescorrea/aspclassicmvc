<%

	Function serializeEntitie(ByVal strObject)
		table = Split(strObject, ":")(0)
		fields = Split(Split(strObject, ":")(1), ",")
		info = ""
		infoIni = ""
		
		info = " Private objTable" & vbCrLf
		info = info & " Private objField" & vbCrLf

		For x = 0 To UBound(fields)
			info = info & "	Private prv" & fields(x) & " " & vbCrLf

			info = info & "	Public Function " & fields(x) & "() " & vbCrLf
			info = info & "		objField.Index = " & x & " " & vbCrLf
			info = info & "		Set " & fields(x) & " = objField " & vbCrLf
			info = info & "	End Function " & vbCrLf

			infoIni = infoIni & "		objField.Index = " & x & " " & vbCrLf
			infoIni = infoIni & "		objField.Name = """ & fields(x) & """ " & vbCrLf
			infoIni = infoIni & "		objField.Find = """" " & vbCrLf
			infoIni = infoIni & "		objField.Table = """ & table & """ " & vbCrLf
		Next

		info = info & " Private Sub Class_Initialize()" & vbCrLf
		info = info & "		Set objTable = New clsTable " & vbCrLf
		info = info & "		Set objField = New clsField " & vbCrLf
		info = info & "		objTable.table = """ & table & """ " & vbCrLf
		info = info & "		objTable.columns = Split(""" & Join(fields, ",") & """, "","") " & vbCrLf
		info = info & infoIni
		info = info & " End Sub" & vbCrLf

		info = info & " Public Function Schema()" & vbCrLf
		info = info & "		Set Schema = objTable " & vbCrLf
		info = info & " End Function" & vbCrLf

		info = info & " Private Sub Class_Terminate()" & vbCrLf
		info = info & "		Set objTable = Nothing " & vbCrLf
		info = info & "		Set objField = Nothing " & vbCrLf
		info = info & " End Sub" & vbCrLf
		
		code = "Class cls" & table & vbCrLf
		code = code & info & vbCrLf
		code = code & "End Class" & vbCrLf

		Execute(code)
		Execute("Set objTemp = New cls" & table)

		Set serializeEntitie = objTemp
		
		Set objTemp = Nothing
	End Function

	Function serializeEntitieColection(ByVal strObject)
		NameSpace = Split(strObject, ":")(0)
		Objects = Split(Split(strObject, ":")(1), "},")
		info = ""
		
		For x = 0 To UBound(Objects)

			nameObject 		= Split(Objects(x), "{")(0)
			fieldsObject 	= Replace(Split(Objects(x), "{")(1), "}", "")

			info = info & "	Public Function " & nameObject & "() " & vbCrLf
			info = info & "		Set " & nameObject & " = serializeEntitie(""" & nameObject & ":" & fieldsObject & """) " & vbCrLf
			info = info & "	End Function " & vbCrLf

		Next
		
		code = "Class cls" & NameSpace & vbCrLf
		code = code & info & vbCrLf
		code = code & "End Class" & vbCrLf

		Execute(code)
		Execute("Set objTemp = New cls" & NameSpace)

		Set serializeEntitieColection = objTemp
		
		Set objTemp = Nothing
	End Function

	Function serializeEntitieObjects(ByVal strObject)
		table = Split(strObject, ":")(0)
		fields = Split(Split(strObject, ":")(1), ",")
		info = ""
		
		For x = 0 To UBound(fields)

			arrNameMethod = Split(fields(x), ".")
			nameMethod = arrNameMethod(Ubound(arrNameMethod))

			info = info & "	Public Function " & nameMethod & "() " & vbCrLf
			info = info & "		Set " & nameMethod & " = " & fields(x) & " " & vbCrLf
			info = info & "	End Function " & vbCrLf

		Next
		
		code = "Class cls" & table & vbCrLf
		code = code & info & vbCrLf
		code = code & "End Class" & vbCrLf

		Execute(code)
		Execute("Set objTemp = New cls" & table)

		Set serializeEntitieObjects = objTemp
		
		Set objTemp = Nothing
	End Function

%>