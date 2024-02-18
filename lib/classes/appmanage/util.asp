<%
'////////////////////////////////////////////////////////////////////////////////////////////////////////////
'//
'// Base Utils
'//

Class FactoryClass
	Public Function Create(sType)
		Select Case sType
		Case "Args"
			Set Create = New Args
		End Select
	End Function
End Class
Dim Factory : Set Factory = New FactoryClass



'////////////////////////////////////////////////////////////////////////////////////////////////////////////
'//
'// Language Utils
'//

'
'	oArgs.SetArg("Key1", "Val1").arg("Key2", "Val2").done
'	or
'	oArgs.SetArgs(Array("Key1", "Val1", "Key2", "Val2"))
'
Class Args
	Private m_oArgs

	Public Sub Class_Initialize
		Set m_oArgs = Server.CreateObject("Scripting.Dictionary")
	End Sub

	' oArgs는 Array, Args 객체 모두 가능하다
	Public Function SetArgs(oArgs)
		
		If IsArray(oArgs) Then
			' Array 복사
			Dim num : num = UBound(oArgs)
			Dim i
			For i = 0 To num Step 2
				m_oArgs.Add oArgs(i), oArgs(i + 1)
			Next

		ElseIf IsObject(oArgs) Then
			' Args 복사
			Dim aKeys : aKeys = oArgs.Keys
			Dim sKey
			For Each sKey In aKeys
				If ( m_oArgs.Exists(sKey) ) Then
					m_oArgs.Item(sKey) = oArgs.Item(sKey)
				Else
					m_oArgs.Add sKey, oArgs.Item(sKey)
				End If
			Next
		End If

		Set SetArgs = me
	End Function

	Public Function SetArg(sKey, val)
		m_oArgs.Add sKey, val
		Set SetArg = me
	End Function

	Public Sub done
	End Sub

	Public Function Item(sKey)
		If ( m_oArgs.Exists(sKey) = False ) Then
			Err.Raise 1, "키 '" & sKey & "' 는 존재하지 않습니다", "Args.item('" & sKey & "')"
		End If

		If IsObject(m_oArgs.Item(sKey)) Then
			Set Item = m_oArgs.Item(sKey)
		Else
			Item = m_oArgs.Item(sKey)
		End If
	End Function

	Public Function Keys()
		Keys = m_oArgs.Keys
	End function

	Public Function HasKey(sKey)
		If ( m_oArgs.Exists(sKey) = True ) Then
			HasKey = True
			Exit Function
		End If
		HasKey = False
	End Function

	Public Function ToString
		Dim aTemp : aTemp = Array()
		ReDim aTemp(m_oArgs.Count)
		
		Dim i : i = 0
		Dim sKey, sVal
		Dim aKeys : aKeys = m_oArgs.Keys
		For Each sKey In aKeys
			If IsObject(m_oArgs.Item(sKey)) Then
				sVal = "(Object)"
			Else
				sVal = m_oArgs.Item(sKey)
			End If

			aTemp(i) = sKey & ":'" & sVal & "'"

			i = i + 1
		Next
		
		ToString = Join(aTemp, ", ")
	End Function
End Class


Function IF_(bCondition, trueCase, falseCase)
	If ( bCondition ) Then
		IF_ = trueCase
	Else
		IF_ = falseCase
	End If
End Function


'////////////////////////////////////////////////////////////////////////////////////////////////////////////
'//
'// DB Utils
'//

Function GetRsFromArgs(oArgs)
	If ( oArgs.HasKey("oRS") = False ) Then
		' oRS가 없으면 전역으로 사용되는 rs를 사용
		Set GetRsFromArgs = rsget
	Else
		Set GetRsFromArgs = oArgs.item("oRS")
	End If
End Function

Function ForRS(oRS, sFunc)
	Dim nIdx : nIdx = 0
	Do until oRS.EOF
		GetRef(sFunc)(nIdx),(oRS)
		nIdx = nIdx + 1
		oRS.MoveNext
	Loop
End Function

Function GetTableRows(oRS, sTableName)
	Dim sSQL : sSQL = "SELECT COUNT(*) AS rows FROM " & sTableName

	oRS.open sSQL, dbget, 0

	GetTableRows = 0
	Do until oRS.EOF
		GetTableRows = oRS("rows")
		oRS.MoveNext
	Loop
End Function



'////////////////////////////////////////////////////////////////////////////////////////////////////////////
'//
'// Page
'//

Class Pagination
	Private m_oArgs

	Public Sub Class_Initialize
		Set m_oArgs = Factory.Create("Args")
	End Sub

	Public Function SetOptions(oArgs)
		m_oArgs.SetArgs(Array( _
			"base_url", "", _
			"total_rows", 100, _
			"per_page", 20, _
			"cur_page", 1, _
			"num_links", 10, _

			"full_tag_open", "<p>", _
			"full_tag_close", "</p>", _

			"first_link", "First", _
			"first_tag_open", "<div>", _
			"first_tag_close", "</div>", _

			"last_link", "Last", _
			"last_tag_open", "<div>", _
			"last_tag_close", "</div>", _

			"next_link", "&gt;", _
			"next_tag_open", "<div>", _
			"next_tag_close", "</div>", _

			"prev_link", "&lt;", _
			"prev_tag_open", "<div>", _
			"prev_tag_close", "</div>", _

			"cur_tag_open", "<b>", _
			"cur_tag_close", "</b>", _

			"num_tag_open", "<div>", _
			"num_tag_close", "</div>" _
		))
		m_oArgs.SetArgs(oArgs)
	End Function

	'<div class="pagination">
	'  <ul>
	'    <li><a href="#">Prev</a></li>
	'    <li>[<a href="#">1</a>]</li>
	'    <li>[<a href="#">2</a>]</li>
	'    <li>[<a href="#">3</a>]</li>
	'    <li>[<a href="#">4</a>]</li>
	'    <li><a href="#">Next</a></li>
	'  </ul>
	'</div>
	Public Function ToString
		Dim sBaseURL, nTotalRows, nPerPage, nCurPage, nNumLinks
		sBaseURL = m_oArgs.Item("base_url")
		nTotalRows = m_oArgs.Item("total_rows")
		nPerPage = m_oArgs.Item("per_page")
		nCurPage = IF_(m_oArgs.Item("cur_page") > 1, m_oArgs.Item("cur_page"), 1)
		nNumLinks = m_oArgs.Item("num_links")

		Dim nFirstNum : nFirstNum = 1
		Dim nLastNum : nLastNum = Round((nTotalRows / nPerPage) + 0.5, 0)

		Dim nFirstDigit : nFirstDigit = Int((nCurPage - 1) / nNumLinks) * nNumLinks + 1
		Dim nLastDigit : nLastDigit = IF_((nFirstDigit + nNumLinks - 1) > nLastNum, nLastNum, (nFirstDigit + nNumLinks - 1))

		Dim nPrevNum : nPrevNum = IF_((nFirstDigit - nNumLinks) < 1, 1, (nFirstDigit - nNumLinks))
		Dim nNextNum : nNextNum = IF_((nFirstDigit + nNumLinks) >= nLastNum, nLastNum, (nFirstDigit + nNumLinks))

		Dim sOut, i
		sOut = m_oArgs.Item("full_tag_open")
		
		sOut = sOut & m_oArgs.Item("first_tag_open")
		sOut = sOut & "<a href='" & sBaseURL & nFirstNum & "'>"
		sOut = sOut & m_oArgs.Item("first_link")
		sOut = sOut & "</a>"
		sOut = sOut & m_oArgs.Item("first_tag_close")

		sOut = sOut & m_oArgs.Item("prev_tag_open")
		sOut = sOut & "<a href='" & sBaseURL & nPrevNum & "'>"
		sOut = sOut & m_oArgs.Item("prev_link")
		sOut = sOut & "</a>"
		sOut = sOut & m_oArgs.Item("prev_tag_close")

		For i = nFirstDigit To nLastDigit
			If ( i = nCurPage ) Then
				sOut = sOut & m_oArgs.Item("cur_tag_open")
				sOut = sOut & "<a href='" & sBaseURL & i & "'>"
				sOut = sOut & i
				sOut = sOut & "</a>"
				sOut = sOut & m_oArgs.Item("cur_tag_close")
			Else
				sOut = sOut & m_oArgs.Item("num_tag_open")
				sOut = sOut & "<a href='" & sBaseURL & i & "'>"
				sOut = sOut & i
				sOut = sOut & "</a>"
				sOut = sOut & m_oArgs.Item("num_tag_close")
			End If
		Next

		sOut = sOut & m_oArgs.Item("next_tag_open")
		sOut = sOut & "<a href='" & sBaseURL & nNextNum & "'>"
		sOut = sOut & m_oArgs.Item("next_link")
		sOut = sOut & "</a>"
		sOut = sOut & m_oArgs.Item("next_tag_close")

		sOut = sOut & m_oArgs.Item("last_tag_open")
		sOut = sOut & "<a href='" & sBaseURL & nLastNum & "'>"
		sOut = sOut & m_oArgs.Item("last_link")
		sOut = sOut & "</a>"
		sOut = sOut & m_oArgs.Item("last_tag_close")

		sOut = sOut & m_oArgs.Item("full_tag_close")

		ToString = sOut
	End Function
End Class

%>