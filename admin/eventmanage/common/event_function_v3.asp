<%
'####################################################
' Page : /lib/event_function.asp
' Description :  이벤트 함수
' History : 2007.02.07 정윤정 생성
'####################################################

'-----------------------------------------------------------------------
' 1. sbGetDesignerid :웹디자인팀 부서번호(12)로 디자이너이름 리스트가져오기
' 2007.02.07 정윤정 생성
'2011.01.18 정윤정 수정 : 부서정보 테이블변경(tbl_partner -> tbl_user_tenbyten)
'-----------------------------------------------------------------------
 Sub sbGetDesignerid(ByVal selName, ByVal sIDValue, ByVal sScript)
  Dim strSql, arrList, intLoop
   strSql = " SELECT userid,  username"
   strSql = strSql & "  from db_partner.dbo.tbl_user_tenbyten with (noLock) "
   strSql = strSql & " WHERE part_sn ='12' and isUsing=1 and (statediv ='Y' or (statediv ='N' and datediff(dd,retireday,getdate())<=0)) and userid <> ''  order by posit_sn, empno"

   rsget.Open strSql,dbget
   IF not rsget.eof THEN
   	arrList = rsget.getRows()
   End IF
   rsget.close
%>
	<select name="<%=selName%>" <%=sScript%> class="Select">
	<option value="">선택</option>
<%
   If isArray(arrList) THEN
   		For intLoop = 0 To UBound(arrList,2)
%>
	<option value="<%=arrList(0,intLoop)%>" <%if Cstr(arrList(0,intLoop)) = Cstr(sIDValue) then %>selected<%end if%>><%=arrList(1,intLoop)%></option>
<%
   		Next
   End IF
%>
	</select>
<%
 End Sub


'-----------------------------------------------------------------------
' 2.fnSetEventCommonCode : 이벤트 공통코드 어플변수에 세팅
' 2007.02.07 정윤정 생성
'-----------------------------------------------------------------------
 Function fnSetEventCommonCode
 On Error Resume Next
	Dim strSql, arrList, intLoop
	Dim intI, intJ, arrCode(), strtype
	strSql = " SELECT code_type, code_value, code_desc FROM [db_event].[dbo].[tbl_event_commoncode] WHERE code_using ='Y' Order by code_type, code_sort "
	rsget.Open strSql,dbget
	IF not rsget.EOF THEN
		arrList = rsget.getRows()
	END IF
	rsget.close

	intJ = 0
	For intI = 0 To UBound(arrList, 2)
		If intI > 0 AND intI <> Ubound(arrList, 2) Then
			Application.Lock
			Application(Trim(strtype)) = arrCode
			Application.UnLock
		End If

		If strtype <> Trim(arrList(0, intI)) And Not IsEmpty(strtype) Then intJ = 0	' 인덱스 초기화
		ReDim Preserve arrCode(Ubound(arrList)-1, intJ)	' 배열 확장
		strtype = Trim(arrList(0, intI))  	' 구분 설정
		arrCode(0, intJ) = Trim(arrList(1, intI)) ' 코드 대입
		arrCode(1, intJ) = Trim(arrList(2, intI)) ' 코드명 대입

		intJ = intJ + 1 ' 인덱스 증가

		If intI = Ubound(arrList, 2) Then
			Application.Lock
			Application(Trim(strtype)) = arrCode
			Application.UnLock
		End If

	Next
End Function

'--------------------------------------------------------------------------------
' 3.sbGetOptEventCodeValue(변수명, 선택값, '선택'구문 사용유무, 스크립트)
' : 이벤트 공통코드 어플변수 select 구문화
' 2007.02.07 정윤정 생성
'--------------------------------------------------------------------------------
 Sub sbGetOptEventCodeValue(ByVal sType, ByVal selValue, ByVal sViewOpt, ByVal sScript)
   Dim arrList, intLoop, iValue
 	arrList = Application(sType)
 	iValue  = selValue
 	IF  isNull(selValue) THEN 		selValue = ""
 	IF selValue = ""  THEN	iValue = 0

%>
	<select name="<%=sType%>" id="<%=sType%>" <%=sScript%> class="select">
	<%IF sViewOpt THEN%>
	<option value="">선택</option>
    <%END IF%>
<%
IF isArray(arrList) then
 	For intLoop =0 To UBound(arrList,2)
%>
	<option value="<%=arrList(0,intLoop)%>" <%If CStr(selValue) = CStr(arrList(0,intLoop)) THEN%>selected<%END IF%>>
	<%
		If arrList(1,intLoop) = "100px" Then
			Response.Write "130px"
		ElseIf arrList(1,intLoop) = "150px" Then
			Response.Write "180px"
		ElseIf arrList(1,intLoop) = "200px" Then
			Response.Write "400px"
		ElseIf arrList(1,intLoop) = "155px" Then
			Response.Write "270px"
		ElseIf arrList(1,intLoop) = "160px" Then
			Response.Write "320px"
		Else
			Response.Write arrList(1,intLoop)
		End If
	%>
	</option>
<%
	Next
end if	
	 %>
	</select>
<%
 End Sub
'--------------------------------------------------------------------------------
' 3-1.sbGetOptEventCodeValue(변수명, 선택값, '선택'구문 사용유무, 스크립트)
' : 이벤트 공통코드 어플변수 select 구문화, 변수명 따로 추가
' 2007.02.07 정윤정 생성
'--------------------------------------------------------------------------------
 Sub sbGetVarOptEventCodeValue(ByVal sVarName, ByVal sType, ByVal selValue, ByVal sViewOpt, ByVal sScript)
   Dim arrList, intLoop, iValue
 	arrList = Application(sType)
 	iValue  = selValue
 	IF  isNull(selValue) THEN 		selValue = ""
 	IF selValue = ""  THEN	iValue = 0

%>
	<select name="<%=sVarName%>" <%=sScript%>>
	<%IF sViewOpt THEN%>
	<option value="">선택</option>
    <%END IF%>
<%
 	For intLoop =0 To UBound(arrList,2)
%>
	<option value="<%=arrList(0,intLoop)%>" <%If CStr(selValue) = CStr(arrList(0,intLoop)) THEN%>selected<%END IF%>>
	<%
		If arrList(1,intLoop) = "100px" Then
			Response.Write "130px"
		ElseIf arrList(1,intLoop) = "150px" Then
			Response.Write "180px"
		ElseIf arrList(1,intLoop) = "200px" Then
			Response.Write "400px"
		ElseIf arrList(1,intLoop) = "155px" Then
			Response.Write "270px"
		ElseIf arrList(1,intLoop) = "160px" Then
			Response.Write "320px"
		Else
			Response.Write arrList(1,intLoop)
		End If
	%>
	</option>
<%
	Next %>
	</select>
<%
 End Sub
'-----------------------------------------------------------------------
' 4.fnGetEventCodeDesc : 이벤트 공통코드 값에 따른 이름 가져오기
' 2007.02.07 정윤정 생성
'-----------------------------------------------------------------------
 Function fnGetEventCodeDesc(ByVal sType, ByVal selValue)

	Dim arrList, intLoop
	arrList = Application(sType)

	' 애플리케이션 변수 존재 체크
	If IsEmpty(arrList) OR not isArray(arrList) Then Exit Function

	' 코드 설정
	selValue = Trim(selValue)

	' 코드명 찾는 루틴
	For intLoop = 0 To Ubound(arrList, 2)
		' 코드명 비교 후 맞으면 반환
			If CStr(selValue) = CStr(arrList(0, intLoop)) Then fnGetEventCodeDesc = arrList(1, intLoop) : Exit For
	Next
End Function

'-----------------------------------------------------------------------
' 5.sbGetOptCategoryLarge : 카테고리 대분류명가져오기
' 2007.02.07 정윤정 생성
'-----------------------------------------------------------------------
 Sub sbGetOptCategoryLarge(byval selectBoxName,byval selectedId, ByVal strEtc)
   dim tmp_str,query1
   %><select name="<%=selectBoxName%>" <%=strEtc%>>
     <option value="" <% if selectedId="" then response.write " selected"%>>선택</option><%
   query1 = " select code_large, code_nm from [db_item].[dbo].tbl_Cate_large"
   query1 = query1 + " where display_yn = 'Y'"
   query1 = query1 + " order by code_large Asc"

   rsget.Open query1,dbget
   if  not rsget.EOF  then
       rsget.Movefirst

       do until rsget.EOF
           if Cstr(selectedId) = Cstr(rsget("code_large")) then
               tmp_str = " selected"
           end if
           response.write("<option value='"&rsget("code_large")&"' "&tmp_str&">"& db2html(rsget("code_nm")) &"</option>")
           tmp_str = ""
           rsget.MoveNext
       loop
   end if
   rsget.close
   response.write("</select>")

end Sub

'-----------------------------------------------------------------------
' 5.sbGetOnlyOptCategoryLarge : 카테고리 대분류명가져오기
' 2007.02.07 정윤정 생성
'-----------------------------------------------------------------------
 Sub sbGetOnlyOptCategoryLarge(byval selectedId)
   dim tmp_str,query1
   %>
     <option value="" <% if selectedId="" then response.write " selected"%>>선택</option><%
   query1 = " select code_large, code_nm from [db_item].[dbo].tbl_Cate_large"
   query1 = query1 + " where display_yn = 'Y'"
   query1 = query1 + " order by code_large Asc"

   rsget.Open query1,dbget
   if  not rsget.EOF  then
       rsget.Movefirst

       do until rsget.EOF
           if Cstr(selectedId) = Cstr(rsget("code_large")) then
               tmp_str = " selected"
           end if
           response.write("<option value='"&rsget("code_large")&"' "&tmp_str&">"& db2html(rsget("code_nm")) &"</option>")
           tmp_str = ""
           rsget.MoveNext
       loop
   end if
   rsget.close

end Sub
'-----------------------------------------------------------------------
' 6.sbGetStaticEvent : 정기이벤트 종류 가져오기
' 2007.02.07 정윤정 생성
'-----------------------------------------------------------------------
	Sub sbGetStaticEvent(ByVal selValue)
%>
	<select name="selStatic">
	<option value="">선택</option>
	<option value="포토상품후기" <%IF Trim(selValue) = "포토상품후기" THEN%>selected<%END IF%>>포토상품후기</option>
	<option value="한줄낙서" <%IF Trim(selValue) = "한줄낙서" THEN%>selected<%END IF%>>한줄낙서</option>
	<option value="보물경매" <%IF Trim(selValue) = "보물경매" THEN%>selected<%END IF%>>보물경매</option>
	<option value="100%SHOP" <%IF Trim(selValue) = "100%SHOP" THEN%>selected<%END IF%>>100% SHOP</option>
	<option value="러브하우스" <%IF Trim(selValue) = "러브하우스" THEN%>selected<%END IF%>>러브하우스</option>
	</select>
<%
	End Sub

'-----------------------------------------------------------------------
' 7.GetImageFolerName : 폴더명 가져오기
' 2007.02.07 정윤정 생성
'-----------------------------------------------------------------------
	function GetImageFolerName(byval itemid)
		GetImageFolerName = "0" + CStr(Clng(itemid\10000))
	end function

'-----------------------------------------------------------------------
' 8.fnSetDispUrl : 이벤트 종류에 따라 링크 변경
' 2007.02.07 정윤정 생성
'-----------------------------------------------------------------------
	Function fnSetDispUrl(ByVal sKind, ByVal eCode)
		SELECT CASE sKind
			CASE 7
				fnSetDispUrl = "/admin/sitemaster/weekly_codi_detail.asp?mode=add&eC="&eCode
			CASE ELSE
				fnSetDispUrl = "event_modify.asp?eC="&eCode&"&menupos="&menupos
		END SELECT
	End Function

'-----------------------------------------------------------------------
' 9.sbAlertMsg : 알림문구 후 페이지 이동 처리
' 2007.02.07 정윤정 생성
'-----------------------------------------------------------------------
	Sub sbAlertMsg(byVal strMsg, ByVal strUrl, ByVal strTarget)
		Dim strLink
		IF strUrl = "close" THEN	'팝업 창 닫을경우
			strLink = strTarget & ".close();"
		ELSEIF strUrl ="back" THEN	'이전 페이지로 이동
			strLink = "history.back(-1);"
		ELSE
			strLink = strTarget & ".location.href='" & strUrl & "';"
		END IF
%>
	<script type="text/javascript">
		alert("<%=strMsg%>");
		<%=strLink%>;
	</script>
<%		response.End
	End Sub

'--------------------------------------------------------------------------------
' 10.sbGetOptGiftCodeValue(변수명, 선택값,  그룹선택구문 사용유무, 스크립트,이벤트코드)
' : 사은품공통코드 어플변수 select 구문화
' 예외사항: 그룹선택일 경우 선택적 view
' 2007.05.09 정윤정 생성
'--------------------------------------------------------------------------------
 Sub sbGetOptGiftCodeValue(ByVal sType, ByVal selValue, ByVal sViewOpt, ByVal sScript,ByVal eCode)
   Dim arrList, intLoop
 	arrList = fnSetCommonCodeArr(sType, True)
 	IF  isNull(selValue) THEN selValue = ""
%>
    <select name="<%=sType%>" class="select" <%=sScript%>>
	<% if selValue="1" then %>
	        <option value="1" selected >전체증정</option>
	<% end if %>

<%
 	For intLoop =0 To UBound(arrList,2)

 	IF ((NOT sViewOpt and  Cstr(arrList(0,intLoop)) <> "4") OR sViewOpt) AND (not (Cstr(eCode) <> ""  AND  Cstr(arrList(0,intLoop)) = "5") ) AND ( not (Cstr(eCode) ="" AND ( Cstr(arrList(0,intLoop)) = "2" OR Cstr(arrList(0,intLoop)) = "6") )) THEN
%>
	<option value="<%=arrList(0,intLoop)%>" <%If CStr(selValue) = CStr(arrList(0,intLoop)) THEN%>selected<%END IF%>><%=arrList(1,intLoop)%></option>
<%   END IF
	Next %>
	<% if selValue="9" then %>
	        <option value="9" selected >다이어리구매시(+텐배송)</option>
	 <% end if %>
	</select>
<%
 End Sub

'--------------------------------------------------------------------------------
' 10.sbOptPartner
'	:특정 제휴몰
' 2008.03.24 정윤정 생성
'--------------------------------------------------------------------------------
Sub sbOptPartner(ByVal selPartner)
	Dim strSql, arrList,intLoop, selvalue
	strSql = "SELECT id, company_name from [db_partner].[dbo].tbl_partner where userdiv=999 and isusing='Y'"
	 rsget.Open strSql,dbget
	 IF not rsget.Eof THEN
	 	arrList = rsget.getRows()
	 END IF
	 rsget.close

	 IF isArray(arrList) THEN
	 	For intLoop =0 To UBound(arrList,2)
%>
	<option value="<%=arrList(0,intLoop)%>" <%IF Cstr(selPartner) = arrList(0,intLoop) THEN%>selected<%END IF%>><%=arrList(1,intLoop)%> (<%=arrList(0,intLoop)%>)</option>
<%	 	Next
	 END IF
End Sub
'--------------------------------------------------------------------------------
' 11.sbGetOptStatusCodeValue(변수명, 선택값, '선택'구문 사용유무, 스크립트)
' : 이벤트 상태값 공통코드 어플변수 select 구문화 - 현재값의 상태보다 이전값으로 이동 못하도록
' 2007.02.07 정윤정 생성
'--------------------------------------------------------------------------------
 Sub sbGetOptStatusCodeValue(ByVal sType, ByVal selValue, ByVal sViewOpt, ByVal sScript)
   Dim arrList, intLoop, iValue
 	arrList= fnSetCommonCodeArr(sType, True)
 	iValue  = selValue
 	IF  isNull(selValue) THEN 		selValue = ""
 	IF selValue = ""  THEN	iValue = 0

%>
	<select name="<%=sType%>" <%=sScript%>>
	<%IF sViewOpt THEN%>
	<option value="">선택</option>
    <%END IF%>
<%
 	For intLoop =0 To UBound(arrList,2)
 	 IF Cint(arrList(0,intLoop)) >= Cint(iValue) OR sViewOpt THEN
%>
	<option value="<%=arrList(0,intLoop)%>" <%If CStr(selValue) = CStr(arrList(0,intLoop)) THEN%>selected<%END IF%>><%=replace(arrList(1,intLoop),"오픈예정","오픈")%></option>
<%   END IF
	Next %>
	</select>
<%
 End Sub

'--------------------------------------------------------------------------------
' 11-1.sbGetOptStatusCodeAuth(변수명, 선택값, 사용구분, 스크립트)
' 2013.04.25 허진원 생성
'--------------------------------------------------------------------------------
 Sub sbGetOptStatusCodeAuth(sType, selValue, sViewOpt, sScript)
   Dim arrList, intLoop, iValue
 	arrList= fnSetCommonCodeArr(sType, True)
 	iValue  = selValue
 	IF isNull(selValue) THEN	selValue = ""
 	IF selValue="" THEN			iValue = 0

	'// 이벤트 관리자 여부(인증구분; MD파트/마케팅 선임 및 관리자)
	Dim uMng : uMng=false
	if (session("ssAdminLsn")<=3 and (session("ssAdminPsn")=11 or session("ssAdminPsn")=14)) or (session("ssAdminLsn")=1) then
		uMng = true
	end if

	Response.Write "<select name='" & sType & "' " & sScript & ">" & vbCrLf

	For intLoop =0 To UBound(arrList,2)
		Select Case sViewOpt
			Case "N"
				'# 신규등록 시 (등록대기, 승인요청만 출력)
				If Cint(arrList(0,intLoop))=0 or Cint(arrList(0,intLoop))=2 then
					Response.Write "<option value='" & arrList(0,intLoop) & "' " & chkIIF(CStr(selValue)=CStr(arrList(0,intLoop)),"selected","") & ">" & replace(arrList(1,intLoop),"오픈예정","오픈")  & "</option>" & vbCrLf
				end if
			Case Else
				if uMng then
					'//이벤트 관리자라면 전체 구분 출력
					IF Cint(arrList(0,intLoop)) >= Cint(iValue) or (Cint(iValue)<=2 and Cint(arrList(0,intLoop))=1) THEN
			 	 		Response.Write "<option value='" & arrList(0,intLoop) & "' " & chkIIF(CStr(selValue)=CStr(arrList(0,intLoop)),"selected","") & ">" & replace(arrList(1,intLoop),"오픈예정","오픈")  & "</option>" & vbCrLf
					END IF
				else
					IF Cint(iValue)>2 THEN
						IF Cint(arrList(0,intLoop)) >= Cint(iValue) THEN
			 	 			Response.Write "<option value='" & arrList(0,intLoop) & "' " & chkIIF(CStr(selValue)=CStr(arrList(0,intLoop)),"selected","") & ">" & replace(arrList(1,intLoop),"오픈예정","오픈")  & "</option>" & vbCrLf
			 	 		End if
					ElseIf (Cint(arrList(0,intLoop))<>1 and Cint(arrList(0,intLoop))<=2 and Cint(arrList(0,intLoop)) >= Cint(iValue)) or (Cint(iValue)=1 and Cint(arrList(0,intLoop))=1) then
						Response.Write "<option value='" & arrList(0,intLoop) & "' " & chkIIF(CStr(selValue)=CStr(arrList(0,intLoop)),"selected","") & ">" & replace(arrList(1,intLoop),"오픈예정","오픈")  & "</option>" & vbCrLf
					END IF
				end if
		End Select
	Next

 	 Response.Write "</select>" & vbCrLf

 End Sub

'--------------------------------------------------------------------------------
' 11-2.sbGetOptStatusCodeSort(변수명, 선택값, '선택'구문 사용유무, 스크립트)
' : 이벤트 상태값 공통코드 어플변수 select 구문화 - 현재값의 상태보다 이전값으로 이동 못하도록
' 2015.07.09 정윤정 생성
'--------------------------------------------------------------------------------
 Sub sbGetOptStatusCodeSort(ByVal sType, ByVal selValue, ByVal sViewOpt, ByVal sScript)
   Dim arrList, intLoop, iValue
   Dim strSql
  
  	iValue  = selValue
 	IF  isNull(selValue) THEN 		selValue = ""
 	IF selValue = ""  THEN	iValue = 0
 	strSql = ""    
	strSql = " SELECT code_value, code_desc,code_dispYN "&vbcrlf
	strSql = strSql & " FROM [db_event].[dbo].[tbl_event_commoncode] "&vbcrlf
	strSql = strSql & " WHERE code_type='"&sType&"' and code_dispYN ='Y' and code_using ='Y' "&vbcrlf
	strSql = strSql & "    and code_sort >= ( select code_Sort from db_event.dbo.tbl_Event_commoncode where code_type='"&sType&"' and  code_dispYN ='Y' and code_value = '"&iValue&"' and code_using ='Y')"&vbcrlf
	strSql = strSql &" Order by code_dispYN desc, code_type, code_sort "
     rsget.Open strSql,dbget
	IF not rsget.EOF THEN
	    arrList = rsget.getRows()
	 END IF
	rsget.close   
%>
	<select name="<%=sType%>" <%=sScript%> class="select">
	<%IF sViewOpt THEN%>
	<option value="">선택</option>
    <%END IF%>
<%
 	For intLoop =0 To UBound(arrList,2) 
%>
	<option value="<%=arrList(0,intLoop)%>" <%If CStr(selValue) = CStr(arrList(0,intLoop)) THEN%>selected<%END IF%>><%=replace(arrList(1,intLoop),"오픈예정","오픈")%></option>
<%    
	Next %>
	</select>
<%
 End Sub
'-----------------------------------------------------------------------
' 12.fnGetCommCodeArrDesc : 특정종류의 공통코드값의 배열에서 특정값의 코드명 가져오기
'-----------------------------------------------------------------------
	Function fnGetCommCodeArrDesc(ByVal arrCode, ByVal iCodeValue)
		Dim intLoop
		IF iCodeValue = "" or isNull(iCodeValue) THEN iCodeValue = -1
		For intLoop =0 To UBound(arrCode,2)
	 
			IF Cint(iCodeValue) = arrCode(0,intLoop) THEN
				fnGetCommCodeArrDesc = arrCode(1,intLoop)
				Exit For
			END IF
		Next
	End Function

'-----------------------------------------------------------------------
' 13.fnSetCommonCodeArr : 이벤트 공통코드 가져오기
' 2007.02.07 정윤정 생성
'-----------------------------------------------------------------------
 Function fnSetCommonCodeArr(ByVal code_type, ByVal blnUse)
	Dim strSql, arrList, intLoop, strAdd
	Dim intI, intJ, arrCode(), strtype
	strAdd = ""
	IF blnUse THEN
		strAdd= " and code_using ='Y' "
	END IF
	strSql = " SELECT code_value, code_desc,code_dispYN FROM [db_event].[dbo].[tbl_event_commoncode] WHERE code_type='"&code_type&"' and code_dispYN ='Y' "&strAdd&_
			" Order by code_dispYN desc, code_type, code_sort "
	rsget.Open strSql,dbget
	IF not rsget.EOF THEN
		fnSetCommonCodeArr = rsget.getRows()
	END IF
	rsget.close
End Function

'--------------------------------------------------------------------------------
' 14.sbGetOptCommonCodeArr(변수명, 선택값, '선택'구문 사용유무, 스크립트)
' : 특정종류의 공통코드값의 배열에서 select 처리
' 2008.04.15 정윤정 생성
'--------------------------------------------------------------------------------
 Sub sbGetOptCommonCodeArr(ByVal sType, ByVal selValue, ByVal sViewOpt, ByVal blnUse, ByVal sScript)
   Dim arrCode, intLoop, iValue
   	arrCode= fnSetCommonCodeArr(sType, blnUse)
 	iValue  = selValue
 	IF  isNull(selValue) THEN 	selValue = ""
 	IF selValue = ""  THEN	iValue = 0
%>
	<select name="<%=sType%>" class="select" <%=sScript%>>
	<%
		IF sViewOpt THEN
			Response.Write "<option value="""">선택</option>" &vbCrLf
		END IF
		IF sType="eventkind" THEN
			'if (session("ssAdminPsn")="11" or session("ssAdminPsn")="21")   then 'MD부서라면 (쇼핑찬스,전체,상품,브랜드,다이어리,테스터,신규디자이너) 
				Response.Write "<option value=""1,12,13,23,27,28,29""  " & CHKIIF(CStr(selValue) = "1,12,13,23,27,28,29", "selected", "") & ">#관심항목</option>" &vbCrLf 
			'end if
			'Response.Write "<option value=""1,12,13,16,17,23,24,28"" " & CHKIIF(CStr(selValue) = "1,12,13,16,17,23,24,28", "selected", "") & " >#일반 이벤트</option>" &vbCrLf
			'Response.Write "<option value=""19,25,26"" " & CHKIIF(CStr(selValue) = "19,25,26", "selected", "") & " >#모바일 or 앱 전용</option>" &vbCrLf
		END IF
	%>
<% 	IF isArray(arrCode) THEN
 	For intLoop =0 To UBound(arrCode,2)
%>
	<option value="<%=arrCode(0,intLoop)%>" <%If CStr(selValue) = CStr(arrCode(0,intLoop)) THEN%>selected<%END IF%> <%if arrCode(2,intLoop) ="N" then%>style="color:gray;"<%end if%>> <%=arrCode(1,intLoop)%></option>
<%
	Next
	End IF
	%>
	</select>
<%
 End Sub

 '--------------------------------------------------------------------------------
' 15.fnSetStatusDesc
' : 상태값에 따른 상태명
' 2008.04.15 정윤정 생성
'--------------------------------------------------------------------------------
	Function fnSetStatusDesc(ByVal FState, ByVal FSDate, ByVal FEDate, ByVal FStateDesc)
		IF FState = "7" AND datediff("d",FSDate,date()) >= 0 and datediff("d",FEDate,date()) <=0 THEN
			fnSetStatusDesc = "오픈"
		ELSEIF FState ="7" AND datediff("d",FEDate,date()) > 0 THEN
			fnSetStatusDesc = "종료"
		ELSE
			fnSetStatusDesc = FStateDesc
		END IF
	End Function

 '-----------------------------------------------------------------------
' 16. sbGetMDid :담당MD 리스트가져오기 (팀장 미만,월급계약 이상)
' 2010.01.25 허진원 생성 / '2011.01.18 정윤정 수정 : 부서정보 테이블변경(tbl_partner -> tbl_user_tenbyten)
'-----------------------------------------------------------------------
 Sub sbGetMDid(ByVal selName, ByVal sIDValue, ByVal sScript)
  Dim strSql, arrList, intLoop
   strSql = " SELECT userid, username"
   strSql = strSql & " FROM db_partner.dbo.tbl_user_tenbyten "
   strSql = strSql & " WHERE part_sn ='11' and  posit_sn>='4' and  posit_sn<='12' and   isUsing=1  and (statediv ='Y' or (statediv ='N' and datediff(dd,retireday,getdate())<=0)) and userid <> '' order by posit_sn, empno"

   rsget.Open strSql,dbget
   IF not rsget.eof THEN
   	arrList = rsget.getRows()
   End IF
   rsget.close
%>
	<select name="<%=selName%>" <%=sScript%>>
	<option value="">선택</option>
<%
   If isArray(arrList) THEN
   		For intLoop = 0 To UBound(arrList,2)
%>
	<option value="<%=arrList(0,intLoop)%>" <%if Cstr(arrList(0,intLoop)) = Cstr(sIDValue) then %>selected<%end if%>><%=arrList(1,intLoop)%></option>
<%
   		Next
   End IF
%>
	</select>
<%
 End Sub


  '-----------------------------------------------------------------------
' 17. sbGetMKTid :담당MKT,MD 리스트가져오기 (팀장 미만,직원 이상)
' 2011.06.20 강준구 생성
'-----------------------------------------------------------------------
 Sub sbGetMKTid(ByVal selName, ByVal sIDValue, ByVal sScript)
  Dim strSql, arrList, intLoop
   strSql = " SELECT userid, username"
   strSql = strSql & " FROM db_partner.dbo.tbl_user_tenbyten "
   strSql = strSql & " WHERE part_sn in('11','14') and  posit_sn>='4' and  posit_sn<='8' and   isUsing=1  and (statediv ='Y' or (statediv ='N' and datediff(dd,retireday,getdate())<=0)) and userid <> '' order by part_sn, posit_sn, empno"

   rsget.Open strSql,dbget
   IF not rsget.eof THEN
   	arrList = rsget.getRows()
   End IF
   rsget.close
%>
	<select name="<%=selName%>" <%=sScript%>>
	<option value="">선택</option>
<%
   If isArray(arrList) THEN
   		For intLoop = 0 To UBound(arrList,2)
%>
	<option value="<%=arrList(0,intLoop)%>" <%if Cstr(arrList(0,intLoop)) = Cstr(sIDValue) then %>selected<%end if%>><%=arrList(1,intLoop)%></option>
<%
   		Next
   End IF
%>
	</select>
<%
End Sub

	Sub sbGetwork(ByVal selName, ByVal sIDValue, ByVal sScript)
		Dim strSql, arrList, intLoop
		strSql = " SELECT userid, username "
		strSql = strSql & " FROM db_partner.dbo.tbl_user_tenbyten "
		strSql = strSql & " WHERE userid = '"&sIDValue&"' and userid <> '' "
		rsget.Open strSql,dbget
		IF not rsget.eof THEN
			arrList = rsget.getRows()
		End IF
		rsget.close

		IF isArray(arrList) THEN
%>
			<input type="text" class="text" name="doc_workername" value="<%=arrList(1,0)%>" size="10" readonly>
			<input type="button" class="button" value="지정" onClick="workerlist()">
			<input type="button" class="button" value="X" onClick="workerDel()">
<% 			If selName = "selMKTId" Then %>
				<input type="hidden" name="selMKTId" value="<%=arrList(0,0)%>">
<%			Else %>
				<input type="hidden" name="selMId" value="<%=arrList(0,0)%>">
<%
			End If
		Else
%>
			<input type="text" class="text" name="doc_workername" value="" size="10" readonly>
			<input type="button" class="button" value="지정" onClick="workerlist()">
			<input type="button" class="button" value="X" onClick="workerDel()">

<%			If selName = "selMKTId" Then %>
				<input type="hidden" name="selMKTId" value="">
<%			Else %>
				<input type="hidden" name="selMId" value="">
<%			End If
		End IF
	End Sub

'--------------------------------------------------------------------------------
' 14.GetEvnetKindName(변수명, 선택값)
' 2019.01.22 정태훈 생성
'--------------------------------------------------------------------------------
Sub GetEvnetKindName(ByVal sType, ByVal selValue)
	Dim arrCode, intLoop
	arrCode= fnSetCommonCodeArr(sType, True)
	IF isArray(arrCode) THEN
		For intLoop=0 To UBound(arrCode,2)
			if selValue = arrCode(0,intLoop) then
				if (sType="jobkind" or sType="placekind") and arrCode(0,intLoop)>1 then
					response.write arrCode(1,intLoop)
				end if
			end if
		Next
	End IF
End Sub

%>
