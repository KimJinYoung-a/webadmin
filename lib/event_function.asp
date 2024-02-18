<%
'####################################################
' Page : /lib/event_function.asp
' Description :  이벤트 함수 
' History : 2007.02.07 정윤정 생성
'####################################################
 
'-----------------------------------------------------------------------  
' 1. sbGetDesignerid :웹디자인팀 부서번호(12)로 디자이너이름 리스트가져오기
' 2007.02.07 정윤정 생성
'-----------------------------------------------------------------------
Sub sbGetDesignerid(ByVal selName, ByVal sIDValue, ByVal sScript)
	Dim strSql, arrList, intLoop
	strSql = " SELECT userid, username from db_partner.[dbo].tbl_user_tenbyten WHERE  part_sn ='12' and isUsing=1" & vbcrlf

	' 퇴사예정자 처리	' 2018.10.16 한용민
	strSql = strSql & "	and (statediv ='Y' or (statediv ='N' and datediff(dd,retireday,getdate())<=0))" & vbcrlf
	strSql = strSql & "	and userid <> ''" & vbcrlf  
	strSql = strSql & "	order by posit_sn, empno" & vbcrlf

	'response.write strSql & "<br>"
   rsget.Open strSql,dbget,1
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
' 2.fnSetEventCommonCode : 이벤트 공통코드 어플변수에 세팅
' 2007.02.07 정윤정 생성
'----------------------------------------------------------------------- 
 Function fnSetEventCommonCode
	Dim strSql, arrList, intLoop
	Dim intI, intJ, arrCode(), strtype
	strSql = " SELECT code_type, code_value, code_desc FROM [db_event].[dbo].[tbl_event_commoncode] WHERE code_using ='Y' "
	rsget.Open strSql,dbget,1
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
   Dim arrList, intLoop
 	arrList = Application(sType)
 	IF  isNull(selValue) THEN selValue = ""
%>
	<select class="select" name="<%=sType%>" <%=sScript%>>
	<%IF sViewOpt THEN%>
	<option value="">선택</option>
    <%END IF%>
<% 	
 	For intLoop =0 To UBound(arrList,2)
%>
	<option value="<%=arrList(0,intLoop)%>" <%If CStr(selValue) = CStr(arrList(0,intLoop)) THEN%>selected<%END IF%>><%=arrList(1,intLoop)%></option>
<%  Next %>
	<select>
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
' 2008.04.15 허진원 수정; 新카테고리로 변경
'----------------------------------------------------------------------- 
 Sub sbGetOptCategoryLarge(byval selectBoxName,byval selectedId, ByVal strEtc)
   dim tmp_str,query1
   %><select class="select" name="<%=selectBoxName%>" <%=strEtc%>>   	
     <option value="" <% if selectedId="" then response.write " selected"%>>선택</option><%
   query1 = " select code_large, code_nm from [db_item].[dbo].tbl_Cate_large "
   query1 = query1 + " where display_yn = 'Y'"
   query1 = query1 + " order by code_large Asc"

   rsget.Open query1,dbget,1

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
' 6.sbGetStaticEvent : 정기이벤트 종류 가져오기
' 2007.02.07 정윤정 생성
'----------------------------------------------------------------------- 
	Sub sbGetStaticEvent(ByVal selValue)
%>
	<select class="select" name="selStatic">
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
	<script language="javascript">
	<!--
		alert("<%=strMsg%>");
		<%=strLink%>;
	//-->
	</script>
<%		dbget.close()	:	response.End
	End Sub

'--------------------------------------------------------------------------------   
' 10.sbGetOptGiftCodeValue(변수명, 선택값,  그룹선택구문 사용유무, 스크립트)
' : 사은품공통코드 어플변수 select 구문화 
' 예외사항: 그룹선택일 경우 선택적 view 
' 2007.05.09 정윤정 생성
'--------------------------------------------------------------------------------  
 Sub sbGetOptGiftCodeValue(ByVal sType, ByVal selValue, ByVal sViewOpt, ByVal sScript)
   Dim arrList, intLoop
 	arrList = Application(sType)
 	IF  isNull(selValue) THEN selValue = ""
%>
	<select class="select" name="<%=sType%>" <%=sScript%>>	
	<% if selValue="1" then %>
	        <option value="1" selected >전체증정</option>
	<% end if %>
	
<% 	
 	For intLoop =0 To UBound(arrList,2)  	
 	 IF ((NOT sViewOpt and  Cstr(arrList(0,intLoop)) <> "4") OR sViewOpt )THEN 	 	
%>
	<option value="<%=arrList(0,intLoop)%>" <%If CStr(selValue) = CStr(arrList(0,intLoop)) THEN%>selected<%END IF%>><%=arrList(1,intLoop)%></option>
<%   END IF	
	Next %>
	<select>
<%
 End Sub
%>