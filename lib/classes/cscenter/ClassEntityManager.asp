<%

Function ParseDecimalToSQL(pValue)
	Dim value
	value = pValue
	value = Replace(value, ".", "")
	value = Replace(value, "," , ".")
	ParseDecimalToSQL = value
End Function

Function ToSqlString(pValue)
	Dim value
	value = pValue
	''if Not InStr(value, "''") then
	''	value = Replace(value, "'", "''")
	''end if
	value = "'" & value & "'"
	ToSqlString = value
End Function

Function ToSqlBit(pValue)
	If pValue = true Or LCase(pValue) = "true" Then
		ToSqlBit = 1
	Else
		ToSqlBit = 0
	End If
End Function


Class ClassEntityManager

	Public ObjConn, ObjRs2

	Private StrSQL, ObjRs, Identity, ParentClass
	Public Table, Fields, Values, Types

	Sub Class_Initialize()
		ResetDictionary()
	End Sub

	Public Default Function Init(pTable, pIdentity, pParentClass)
		Table = pTable
		Identity = pIdentity
		Set ParentClass = pParentClass
		Set Init = me
	End Function

	Public Sub ResetDictionary()
		Set Fields = Server.CreateObject("Scripting.Dictionary")
		Set Values = Server.CreateObject("Scripting.Dictionary")
		Set Types  = Server.CreateObject("Scripting.Dictionary")
	end sub

	Public Sub Register(pKey, pField, pType)
		If Not Fields.Exists(pKey) Then
			Call Fields.Add(pKey, pField)
			Call Types.Add(pKey, pType)
		End If
	End Sub

	Private Sub UpdateDictionary()
		Dim Key
		For Each Key In Fields.Keys
			Call SetValue(Key, Eval("ParentClass.FOneItem."& Fields( Key )))
		Next
	End Sub

	Private Sub SetValue(pField, pValue)
		If pValue = "" Then
			Exit Sub
		End If

		If Values.Exists(pField) Then
			Values.Item(pField) = pValue
		Else
			Call Values.Add(pField, pValue)
		End If
	End Sub

	Public Sub Save()
		UpdateDictionary()
		If Values.Exists( Identity ) Then
			If Values.Item( Identity ) <> "" Then
				StrSQL = UpdateQuery()
			Else
				StrSQL = InsertQuery() & InsertValuesQuery()
			End If
		Else
			StrSQL = InsertQuery() & InsertValuesQuery()
		End If

		If StrSQL = "" Then Exit Sub
		Set objRs = ObjConn.Execute( StrSQL )
		If Fields.Exists( Identity ) Then
			Execute("ParentClass.FOneItem." & Fields( Identity ) & " = " & objRs("ID"))
		End If
	End Sub

	Public Sub Delete()
		UpdateDictionary()
		StrSQL = DeleteQuery()
		If StrSQL = "" Then Exit Sub
		ObjConn.Execute( StrSQL )
	End Sub

	Public Sub LoadOne()
		UpdateDictionary()
		StrSQL = SelectQuery()
		If StrSQL = "" Then Exit Sub
		Set objRs = ObjConn.Execute( StrSQL )
		If Not objRs.Eof Then
			Dim Key, Val
			For Each Key In Fields.Keys
				Val = objRs( Key )
				if IsNull(Val) then
					Execute("ParentClass.FOneItem." & Fields.Item( Key ) & " = """ & objRs( Key ) & """")
				else
					Val = Replace(objRs( Key ), vbCrLf, "vbCrLf")
					Execute("ParentClass.FOneItem." & Fields.Item( Key ) & " = """ & Val & """")
					Execute("ParentClass.FOneItem." & Fields.Item( Key ) & " = Replace(ParentClass.FOneItem." & Fields.Item( Key ) & ", ""vbCrLf"", vbCrLf)")
				end if
			Next
		End If
	End Sub

	Public Sub LoadList(countQuery, selectQuery)
		dim i
		UpdateDictionary()

		'// ====================================================================
		StrSQL = countQuery
		If StrSQL = "" Then Exit Sub
		ObjRs2.CursorLocation = adUseClient
		ObjRs2.Open StrSQL, ObjConn, adOpenForwardOnly, adLockReadOnly
			Execute("ParentClass.FTotalCount = " & ObjRs2("cnt"))
		ObjRs2.close

		'// ====================================================================
		StrSQL = selectQuery
		If StrSQL = "" Then Exit Sub
		Execute("ObjRs2.PageSize = ParentClass.FPageSize")
		ObjRs2.CursorLocation = adUseClient
		ObjRs2.Open StrSQL, ObjConn, adOpenForwardOnly, adLockReadOnly

		Execute("ParentClass.FtotalPage =  CLng(ParentClass.FTotalCount\ParentClass.FPageSize)")
		Execute("if (ParentClass.FTotalCount\ParentClass.FPageSize)<>(ParentClass.FTotalCount/ParentClass.FPageSize) then FtotalPage = FtotalPage +1")
		Execute("ParentClass.FResultCount = " & ObjRs2.RecordCount & "-(ParentClass.FPageSize*(ParentClass.FCurrPage-1))")
		Execute("ParentClass.SetFItemListSize()")

		if not ObjRs2.EOF then
			Execute("ObjRs2.absolutepage = ParentClass.FCurrPage")

			i = 0
			do until ObjRs2.eof
				Execute("set ParentClass.FItemList(" & i & ") = new CEmergencyQuestionMasterItem")

				Dim Key, Val
				For Each Key In Fields.Keys
					if (Types.Item( Key ) = "string") then
						Val = objRs2( Key )
						if IsNull(Val) then
							Execute("ParentClass.FItemList(" & i & ")." & Fields.Item( Key ) & " = db2html(""" & Val & """)")
						else
							Val = Replace(objRs2( Key ), vbCrLf, "vbCrLf")
							Execute("ParentClass.FItemList(" & i & ")." & Fields.Item( Key ) & " = db2html(""" & Val & """)")
							Execute("ParentClass.FItemList(" & i & ")." & Fields.Item( Key ) & " = Replace(ParentClass.FItemList(" & i & ")." & Fields.Item( Key ) & ", ""vbCrLf"", vbCrLf)")
						end if
					else
						Execute("ParentClass.FItemList(" & i & ")." & Fields.Item( Key ) & " = """ & ObjRs2( Key ) & """")
					end if
				Next

				i=i+1
				ObjRs2.moveNext
			loop
		end if
	    ObjRs2.Close
	End Sub

	Private Function InsertQuery()
		Dim Key, tempArray : tempArray = Array()
		For Each Key In Values.Keys
			If Key <> Identity Then
				Redim Preserve tempArray( UBound(tempArray) + 1 )
				tempArray( Ubound(tempArray) ) = Key
			End If
		Next
		InsertQuery = " SET NOCOUNT ON; INSERT INTO " & Table & " (" & Join(tempArray, ", ") & " ) "
	End Function

	Private Function InsertValuesQuery()
		Dim Key, tempArray : tempArray = Array()
		For Each Key In Values.Keys
			If Key <> Identity Then
				Redim Preserve tempArray( UBound(tempArray) + 1 )
				tempArray( Ubound(tempArray) ) = FormatType( Values.Item( Key ), Types.Item( Key ) )
			End If
		Next
		InsertValuesQuery = " VALUES ( " & Join(tempArray, ", ") & " ) SELECT @@IDENTITY AS ID "
	End Function

	Private Function UpdateQuery()
		Dim Key, IdentityQuery, tempArray : tempArray = Array()
		For Each Key in Values.Keys
			If Key = Identity Then
				IdentityQuery = Identity & " = " & FormatType( Values.Item( Key ), Types.Item( Key ) )
			Else
				Redim Preserve tempArray( UBound(tempArray) + 1 )
				tempArray( Ubound(tempArray) ) = Key & " = " & FormatType( Values.Item( Key ), Types.Item( Key ) )
			End If
		Next
		UpdateQuery = " SET NOCOUNT ON; UPDATE " & Table & _
					  " SET " & Join(tempArray, ", ") &_
					  " WHERE " & IdentityQuery &_
					  " SELECT " & Values.Item( Identity ) & " AS ID "
	End Function

	Private Function DeleteQuery()
		If Not Values.Exists( Identity ) Then Exit Function
		DeleteQuery = "DELETE FROM " & Table & " WHERE " & Identity & " = " & FormatType( Values.Item( Identity ), Types.Item( Identity ) )
	End Function

	Private Function SelectQuery()
		If Not Values.Exists( Identity ) Then Exit Function
		SelectQuery = _
					" SELECT " & Join(Fields.Keys, ", ") & " FROM " & table &_
					" WHERE " & Identity & " = " & FormatType( Values.Item( Identity ), Types.Item( Identity ) )
	End Function

	Private Function FormatType(pValue, pType)
		If UCase(pValue) = "NULL" Then
			FormatType = pValue
			Exit Function
		End If

		Select Case LCase(pType)
			Case "decimal"
				FormatType = ParseDecimalToSQL(pValue)
			Case "string"
				FormatType = ToSqlString(pValue)
			Case "bool"
				FormatType = ToSqlBit(pValue)
			Case Else
				FormatType = pValue
		End Select
	End Function
End Class

%>
