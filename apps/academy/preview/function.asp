<%
'// 2016 카테고리 선택 상자 (sDisp:전시카테고리, sType:확장여부, sCallback:콜백함수명)
Sub fnPrntDispCateNaviV16(sDisp,sType,sCallback)
	Dim sName, sDepth, sResult, sTmp
	Dim strSql

	'// 카테고리 명 접수
	If sDisp = "" Then
		sName = "카테고리"
	Else
		sName = getDisplayCateNameDB(sDisp)
	End If

	'// 카테고리 조회 범위 설정
	if sDisp="" then
		sDepth = 1
	else
		sDepth = cInt(len(sDisp)/3)
	end if

	'// 표시 형태 (F: 1뎁스 고정, E: 하위분류 확장, S:검색엔진)
	if sType="" then sType="F"
	if sType="E" and sDisp<>"" then sDepth = sDepth +1
	if sType="S" and sDisp<>"" then
		sDepth = sDepth +1
		if sDepth>3 then sDepth=3
	End if

	'// 결과 출력
	sResult = "<button type=""button"" class=""btnSort"" id=""btnDispCate"">" & sName & "</button>" & vbCrLf &_
		"	<div class=""sortList" & chkIIF(sDepth>1," depth2","")& """>" & vbCrLf &_
		"	<ul id=""lyrDispCateList"">" & vbCrLf

		Select Case sType
			Case "F","E"
				'/// DB에서 전시카테고리 접수

				'1Depth는 전체 항목 추가
				if sDepth=1 then
					sResult = sResult & "<li " & chkIIF(sDisp="","class=""current""","") & "><a href=""#"" onclick=""" & sCallback & "('');return false;"">전체</a></li>" & vbCrLf
				end if

				'최종뎁스 확인
				If sDepth > 1 Then
					strSql = " select count(catecode) as cnt from [db_academy].[dbo].tbl_display_cate_Academy "
					strSql = strSql & " where depth = '" & sDepth & "' and useyn = 'Y' "
					strSql = strSql & " and Left(catecode,"&(sDepth-1)*3&") = '" & Left(sDisp,(sDepth-1)*3) & "' "
					rsget.Open strSql,dbget,1
					if rsget("cnt")=0 then
						sDepth = sDepth -1
					end if
					rsget.Close
				end if
		
				'전시카테고리 접수
				strSql = " select catecode, catename from [db_academy].[dbo].tbl_display_cate_Academy "
				strSql = strSql & " where depth = '" & sDepth & "' and useyn = 'Y' "
				If sDepth > 1 Then
					strSql = strSql & " and Left(catecode,"&(sDepth-1)*3&") = '" & Left(sDisp,(sDepth-1)*3) & "' "
				End If
				strSql = strSql & " order by sortno Asc"
				rsget.CursorLocation = adUseClient
				rsget.Open strSql,dbget,adOpenForwardOnly,adLockReadOnly
				if  not rsget.EOF  then
					do until rsget.EOF
						if Left(Cstr(sDisp),3*sDepth) = Cstr(rsget("catecode")) then
							sTmp = "class=""current"""
						end if
						sResult = sResult & "<li "&sTmp&"><a href=""#"" onclick=""" & sCallback & "(" &rsget("catecode") &");return false;"">"& db2html(rsget("catename")) &"</a></li>"
						sTmp = ""
					rsget.MoveNext
					loop
				end if
				rsget.close
			Case "S"
				'/// Ajax 사용 (호출 페이지에서 처리: 여기선 내용없음)
		End Select
		sResult = sResult & "	</ul>" & vbCrLf &_	
		"</div>"
	Response.Write sResult
End Sub


Sub fnPrntLecCateNaviV16(depth,code_large,code_mid,sType,sCallback)
	Dim sName, sDepth, sResult, sTmp
	Dim strSql

	'// 카테고리 명 접수
	If code_large = "" Then
		sName = "카테고리"
	Else
		sName = getLecCateNameDB(depth,code_large,code_mid)
	End If

	'// 카테고리 조회 범위 설정
	if sDepth="" then
		sDepth = 1
	end if
	
	if code_large<>"" then
		sDepth = 2
	end if

	'// 표시 형태 (F: 1뎁스 고정, E: 하위분류 확장, S:검색엔진)
	if sType="" then sType="F"

	'// 결과 출력
	sResult = "<button type=""button"" class=""btnSort"" id=""btnDispCate"">" & sName & "</button>" & vbCrLf &_
		"	<div class=""sortList" & chkIIF(sDepth>1," depth2","")& """>" & vbCrLf &_
		"	<ul id=""lyrDispCateList"">" & vbCrLf

		Select Case sType
			Case "F","E"

				'1Depth는 전체 항목 추가
				if sDepth=1 then
					sResult = sResult & "<li " & chkIIF(code_large="","class=""current""","") & "><a href=""#"" onclick=""" & sCallback & "('','');return false;"">전체</a></li>" & vbCrLf
				end if

				'전시카테고리 접수
				if sDepth=1 then
					strSql = " select code_large as catecode, code_nm from [db_academy].[dbo].[tbl_lec_Cate_large] "
					strSql = strSql & " where display_yn = 'Y' and code_large > 70 "
				else
					strSql = " select code_mid as catecode, code_nm from [db_academy].[dbo].[tbl_lec_Cate_mid] "
					strSql = strSql & " where display_yn = 'Y' and code_large = '" & code_large & "' "
				end if
				strSql = strSql & " order by orderNo Asc"
				rsget.CursorLocation = adUseClient
				rsget.Open strSql,dbget,adOpenForwardOnly,adLockReadOnly
				if  not rsget.EOF  then
					do until rsget.EOF
						if Cstr(CHKIIF(sDepth=1,code_large,code_mid)) = Cstr(rsget("catecode")) then
							sTmp = "class=""current"""
						end if
						
						if sDepth = 2 then
							sResult = sResult & "<li "&sTmp&"><a href=""#"" onclick=""" & sCallback & "('" & code_large & "','" &rsget("catecode") &"');return false;"">"& db2html(rsget("code_nm")) &"</a></li>"
						else
							sResult = sResult & "<li "&sTmp&"><a href=""#"" onclick=""" & sCallback & "('" &rsget("catecode") &"','');return false;"">"& db2html(rsget("code_nm")) &"</a></li>"
						end if
						sTmp = ""
					rsget.MoveNext
					loop
				end if
				rsget.close
			Case "S"
				'/// Ajax 사용 (호출 페이지에서 처리: 여기선 내용없음)
		End Select
		sResult = sResult & "	</ul>" & vbCrLf &_	
		"</div>"
	Response.Write sResult
End Sub


Sub fnPrntMagaCateNaviV16(magacode,sType,sCallback)
	Dim sName, sDepth, sResult, sTmp
	Dim strSql

	'// 카테고리 명 접수
	If magacode = "" Then
		sName = "카테고리"
	Else
		sName = getMagaCateNameDB(magacode)
	End If

	'// 카테고리 조회 범위 설정
	if sDepth="" then
		sDepth = 1
	end if

	'// 표시 형태 (F: 1뎁스 고정, E: 하위분류 확장, S:검색엔진)
	if sType="" then sType="F"

	'// 결과 출력
	sResult = "<button type=""button"" class=""btnSort"" id=""btnDispCate"">" & sName & "</button>" & vbCrLf &_
		"	<div class=""sortList" & chkIIF(sDepth>1," depth2","")& """>" & vbCrLf &_
		"	<ul id=""lyrDispCateList"">" & vbCrLf

		Select Case sType
			Case "F","E"

				'1Depth는 전체 항목 추가
				if sDepth=1 then
					sResult = sResult & "<li " & chkIIF(sDisp="","class=""current""","") & "><a href=""#"" onclick=""" & sCallback & "('');return false;"">전체</a></li>" & vbCrLf
				end if

				'카테고리 접수
				strSql = " select idx as catecode, catename from [db_academy].[dbo].[tbl_academy_magazine_catecode] "
				strSql = strSql & " where isusing = 'Y' and idx = '" & magacode & "' "
				strSql = strSql & " order by orderNo Asc"
				rsget.CursorLocation = adUseClient
				rsget.Open strSql,dbget,adOpenForwardOnly,adLockReadOnly
				if  not rsget.EOF  then
					do until rsget.EOF
						if Cstr(magacode) = Cstr(rsget("catecode")) then
							sTmp = "class=""current"""
						end if
						sResult = sResult & "<li "&sTmp&"><a href=""#"" onclick=""" & sCallback & "(" &rsget("catecode") &");return false;"">"& db2html(rsget("catename")) &"</a></li>"
						sTmp = ""
					rsget.MoveNext
					loop
				end if
				rsget.close
			Case "S"
				'/// Ajax 사용 (호출 페이지에서 처리: 여기선 내용없음)
		End Select
		sResult = sResult & "	</ul>" & vbCrLf &_	
		"</div>"
	Response.Write sResult
End Sub


'// 2016 정렬선택 상자 (sType:정렬방법, sUse:사용처 구분, sCallback:콜백함수명)
Sub fnPrntSortNaviV16(sType,sUse,sCallback)
	Dim sName, sResult, lp
	if sType="" then sType="be"

	if sUse="dft" then
		sUse = "abcdef"
	end if

	'// 현재 정렬명
	Select Case sType
		Case "ne": sName = "신규순"
		Case "be": sName = "인기순"
		Case "ws": sName = "위시등록순"
		Case "hs": sName = "할인율순"
		Case "hp": sName = "높은가격순"
		Case "lp": sName = "낮은가격순"
		Case "br": sName = "리뷰등록순"
		Case "pj": sName = "인기포장순"
		Case "rg": sName = "등록순"
		Case "nm": sName = "이름순"
		Case "mi": sName = "마감임박순"
	End Select

	sResult = "<button type=""button"" class=""btnSort"">" & sName & "</button>" & vbCrLf
	sResult = sResult& "	<div class=""sortList"">" & vbCrLf
	sResult = sResult& "	<ul>" & vbCrLf

	for lp=1 to len(sUse)
		Select Case mid(sUse,lp,1)
			Case "a":  sResult = sResult& "		<li " & chkIIF(sType="ne","class=""current""","") & "><a href=""#"" onclick=""" & sCallback & "('ne');return false;"">신규순</a></li>" & vbCrLf
			Case "b":  sResult = sResult& "		<li " & chkIIF(sType="be","class=""current""","") & "><a href=""#"" onclick=""" & sCallback & "('be');return false;"">인기순</a></li>" & vbCrLf
			Case "c":  sResult = sResult& "		<li " & chkIIF(sType="ws","class=""current""","") & "><a href=""#"" onclick=""" & sCallback & "('ws');return false;"">위시등록순</a></li>" & vbCrLf
			Case "d":  sResult = sResult& "		<li " & chkIIF(sType="hs","class=""current""","") & "><a href=""#"" onclick=""" & sCallback & "('hs');return false;"">할인율순</a></li>" & vbCrLf
			Case "e":  sResult = sResult& "		<li " & chkIIF(sType="hp","class=""current""","") & "><a href=""#"" onclick=""" & sCallback & "('hp');return false;"">높은가격순</a></li>" & vbCrLf
			Case "f":  sResult = sResult& "		<li " & chkIIF(sType="lp","class=""current""","") & "><a href=""#"" onclick=""" & sCallback & "('lp');return false;"">낮은가격순</a></li>" & vbCrLf
			Case "g":  sResult = sResult& "		<li " & chkIIF(sType="br","class=""current""","") & "><a href=""#"" onclick=""" & sCallback & "('br');return false;"">리뷰등록순</a></li>" & vbCrLf
			Case "h":  sResult = sResult& "		<li " & chkIIF(sType="pj","class=""current""","") & "><a href=""#"" onclick=""" & sCallback & "('pj');return false;"">인기포장순</a></li>" & vbCrLf
			Case "i":  sResult = sResult& "		<li " & chkIIF(sType="rg","class=""current""","") & "><a href=""#"" onclick=""" & sCallback & "('rg');return false;"">등록순</a></li>" & vbCrLf
			Case "j":  sResult = sResult& "		<li " & chkIIF(sType="nm","class=""current""","") & "><a href=""#"" onclick=""" & sCallback & "('nm');return false;"">이름순</a></li>" & vbCrLf
			Case "k":  sResult = sResult& "		<li " & chkIIF(sType="mi","class=""current""","") & "><a href=""#"" onclick=""" & sCallback & "('mi');return false;"">마감임박순</a></li>" & vbCrLf
		End Select
	next

	sResult = sResult& "	</ul>" & vbCrLf
	sResult = sResult& "</div>"
	Response.Write sResult
End Sub

%>