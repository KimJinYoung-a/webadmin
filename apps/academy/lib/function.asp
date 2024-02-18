<%
Function SortDictionary(objDict,intSort)
	Dim strDict()
	Dim objKey
	Dim strKey,strItem
	Dim X,Y,Z
	Dim dictKey
	Dim dictItem
	dictKey=1
	dictItem=2
	Z = objDict.Count
	If Z > 1 Then
		ReDim strDict(Z,2)
		X = 0
		For Each objKey In objDict
			strDict(X,dictKey) = CStr(objKey)
			strDict(X,dictItem) = CStr(objDict(objKey))
			X = X + 1
		Next
		For X = 0 to (Z - 2)
			For Y = X to (Z - 1)
				If StrComp(strDict(X,intSort),strDict(Y,intSort),vbTextCompare) > 0 Then
					strKey  = strDict(X,dictKey)
					strItem = strDict(X,dictItem)
					strDict(X,dictKey)  = strDict(Y,dictKey)
					strDict(X,dictItem) = strDict(Y,dictItem)
					strDict(Y,dictKey)  = strKey
					strDict(Y,dictItem) = strItem
				End If
			Next
		Next
		objDict.RemoveAll
		For X = 0 to (Z - 1)
			objDict.Add strDict(X,dictKey), strDict(X,dictItem)
		Next
	End If
End Function
%>