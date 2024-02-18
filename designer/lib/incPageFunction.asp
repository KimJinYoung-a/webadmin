<%
'입력: strCurrentPage-현재 페이지 변수명, intCurrentPage-현재페이지, intTotalRecord-총 검색건수
'			, intRecordPerPage-한페이지당 보여지는 레코드 수, intBlockPerPage-한 블럭사이즈
Sub sbDisplayPaging(ByVal strCurrentPage, ByVal intCurrentPage, ByVal intTotalRecord, ByVal intRecordPerPage, ByVal intBlockPerPage,ByVal menupos)

	'변수 선언
	Dim strCurrentPath
	Dim intStartBlock, intEndBlock, intTotalPage
	Dim strParamName, intLoop
   
	'현재 페이지 명
	strCurrentPath = Request.ServerVariables("Script_Name")
		
	'해당페이지에 표시되는 시작페이지와 마지막페이지 설정
	intStartBlock = Int((intCurrentPage - 1) / intBlockPerPage) * intBlockPerPage + 1
	intEndBlock = Int((intCurrentPage - 1) / intBlockPerPage) * intBlockPerPage + intBlockPerPage
	
	'총 페이지 수 설정
	intTotalPage =  -(int(-(intTotalRecord/intRecordPerPage)))
	
	'폼 설정 & hidden 파라미터 설정
	Response.Write	"<form name='frmPaging' method='get' action ='" & strCurrentPath & "'>" &_
							"<input type='hidden' name='" & strCurrentPage & "'>" 
		
	'파라미터 값들(예: 검색어)을 hidden 파라미터로 저장한다
	strParamName = ""
	For Each strParamName In Request.Form	
		If strParamName <> strCurrentPage Then
			
			'hidden 파라미터 값도 파라미터 검열
			Response.Write "<input type='hidden' name='" & strParamName & "' value='" & requestCheckVar(Request.Form(strParamName),50) & "'>"
		End If
	Next
	strParamName = ""
	
	For Each strParamName In Request.Querystring
		If strParamName <> strCurrentPage Then			
			'hidden 파라미터 값도 파라미터 검열
			Response.Write "<input type='hidden' name='" & strParamName & "' value='" & requestCheckVar(Request.QueryString(strParamName),50) & "'>"		
		END IF	
	Next
		
	Response.Write "<table border='0' cellpadding='0' cellspacing='0' class='a'><tr align='center'><td>"

	'이전 페이지 이미지 설정
	If intStartBlock > 1 Then
		Response.Write "<a href='javascript:document.frmPaging." & strCurrentPage & ".value=" & intStartBlock - intBlockPerPage & ";document.frmPaging.submit();' onfocus='this.blur();'>[pre]</a>" 
							   
	Else
		Response.Write "[pre]"
	End If

	Response.Write "</td><td>&nbsp;"
	
	'페이징 출력
	If intTotalPage > 1 Then
		For intLoop = intStartBlock To intEndBlock
			If intLoop > intTotalPage Then Exit For
			
			If Int(intLoop) <> Int(intStartBlock) Then Response.Write "|"
			
			If Int(intLoop) = Int(intCurrentPage) Then		'현재 페이지
				Response.Write "&nbsp;<span class='text01'><strong>" & intLoop & "</strong></span>&nbsp;"
			Else															'그 외 페이지
				Response.Write "&nbsp;<a href='javascript:document.frmPaging." & strCurrentPage & ".value=" & intLoop & ";document.frmPaging.submit();'><font class='text01'>" & intLoop & "</font></a>&nbsp;"
			End If
		
		Next
	Else		'한 페이지만 존재 할때
		Response.Write "&nbsp;<span class='text01'><strong>1</strong></span>&nbsp;"
	End If

	Response.Write "&nbsp;</td><td>"

	'다음 페이지 이미지 설정
	If Int(intEndBlock) < Int(intTotalPage) Then
		Response.Write "<a href='javascript:document.frmPaging." & strCurrentPage & ".value=" & intEndBlock+1 & ";document.frmPaging.submit();'onfocus='this.blur();'>[next]</a>"  
	Else
		Response.Write "[next]"
	End If
	
	Response.Write "</td></tr></table></form>"

End Sub 
%> 