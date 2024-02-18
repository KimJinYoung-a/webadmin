<%
'=========================================================
' 2011 New 페이징 함수 
' 2011.03.21 강준구 생성
' 2012.03.26 허진원 DIV레이아웃으로 변경
' 2016.06.07 김진영 핑거스 레이아웃으로 변경
' ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' sbDisplayPaging_New(현재 페이지번호, 총 레코드 갯수, 한페이지에 보이는 상품 갯수(select top 수), js 페이지이동 함수명)
' ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' 페이지 이동 js 함수명은 strJsFuncName 으로 임의로 정하고 페이지 번호만 담아서 넘김. 각 페이지에 페이징 전용 form을 만들거나 서칭폼을 같이 쓰거나 하여 post 또는 get으로 넘김.
'=========================================================

Function fnDisplayPaging_New(strCurrentPage, intTotalRecord, intRecordPerPage, strJsFuncName)
	'변수 선언
	Dim intCurrentPage, strCurrentPath, vPageBody
	Dim intStartBlock, intEndBlock, intTotalPage
	Dim strParamName, intLoop

	'현재 페이지 설정
	intCurrentPage = strCurrentPage		'현재 페이지 값

	'총 페이지 수 설정
	intTotalPage =   int((intTotalRecord-1)/intRecordPerPage) +1
	if (intTotalPage<1) then intTotalPage=1

	vPageBody = ""
	strJsFuncName = trim(strJsFuncName)

	If intCurrentPage = 1 Then
		vPageBody = vPageBody & "<a href='javascript:'  class='btnPrev'><span>이전 페이지</span></a>"
	Else
		vPageBody = vPageBody & "<a href='javascript:" & strJsFuncName & "(" & intCurrentPage - 1 & ")'  class='btnPrev'><span>이전 페이지</span></a>"
	End If
	vPageBody = vPageBody & "<span><input type='number' class='pageNum' value=""" & intCurrentPage & """ min=""1"" max=""" & intTotalPage & """ onkeypress=""if(event.keyCode==13){fnDirPg" & strJsFuncName & "(this.value); return false;}"" /> / "&intTotalPage&"</span>"

	If intTotalPage >= intCurrentPage + 1 Then
		vPageBody = vPageBody & "<a href='javascript:" & strJsFuncName & "(" & intCurrentPage + 1 & ")'  class='btnNext'><span>다음 페이지</span></a>"
	Else
		vPageBody = vPageBody & "<a href='javascript:'  class='btnNext'><span>다음 페이지</span></a>"
	End If
	vPageBody = vPageBody & "<script>" & vbCrLf
	vPageBody = vPageBody & "function fnDirPg" & strJsFuncName & "(pg) { " & vbCrLf
	vPageBody = vPageBody & "	if(pg.match(/^\d+$/ig) == null || pg > "&intTotalPage&" || pg == 0 ){alert('정확한 페이지 숫자를 입력해주세요.');} " & vbCrLf
	vPageBody = vPageBody & "	if(pg>0 && pg<=" & intTotalPage & ") " & strJsFuncName & "(pg);" & vbCrLf
	vPageBody = vPageBody & "}" & vbCrLf
	vPageBody = vPageBody & "</script>" & vbCrLf
	fnDisplayPaging_New = vPageBody
End Function

Function fnDisplayPaging_NewMobile(strCurrentPage, intTotalRecord, intRecordPerPage, strJsFuncName)
	'변수 선언
	Dim intCurrentPage, strCurrentPath, vPageBody
	Dim intStartBlock, intEndBlock, intTotalPage
	Dim strParamName, intLoop

	'현재 페이지 설정
	intCurrentPage = strCurrentPage		'현재 페이지 값

	'총 페이지 수 설정
	intTotalPage =   int((intTotalRecord-1)/intRecordPerPage) +1
	if (intTotalPage<1) then intTotalPage=1

	vPageBody = ""
	strJsFuncName = trim(strJsFuncName)

	vPageBody = vPageBody & "<div class='pagination'> "

	If intCurrentPage = 1 Then
		vPageBody = vPageBody & "<a href='javascript:'  class='btnPrev'><span>이전 페이지</span></a>"
	Else
		vPageBody = vPageBody & "<a href='javascript:" & strJsFuncName & "(" & intCurrentPage - 1 & ")'  class='btnPrev'><span>이전 페이지</span></a>"
	End If
	vPageBody = vPageBody & "<span><input type='number' class='pageNum' value=""" & intCurrentPage & """ min=""1"" max=""" & intTotalPage & """ onkeypress=""if(event.keyCode==13){fnDirPg" & strJsFuncName & "(this.value); return false;}"" /> / "&intTotalPage&"</span>"

	If intTotalPage >= intCurrentPage + 1 Then
		vPageBody = vPageBody & "<a href='javascript:" & strJsFuncName & "(" & intCurrentPage + 1 & ")'  class='btnNext'><span>다음 페이지</span></a>"
	Else
		vPageBody = vPageBody & "<a href='javascript:'  class='btnNext'><span>다음 페이지</span></a>"
	End If
	vPageBody = vPageBody & "</div>"
	fnDisplayPaging_NewMobile = vPageBody
	
End Function

%>