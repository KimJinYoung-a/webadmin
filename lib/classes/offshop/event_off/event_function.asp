<%
'####################################################
' Description :  오프라인 이벤트 함수모음
' History : 2010.03.09 한용민 생성
'####################################################

'//오프라인매장 구분 매장공통 포함
Sub drawSelectBoxOffShop_off(selectBoxName,selectedId)
	dim tmp_str,query1
	%>
	<select class="select" name="<%=selectBoxName%>">
	<option value='' <%if selectedId="" then response.write " selected"%>>선택하세요</option>
	<option value='all' <%if selectedId="all" then response.write " selected"%>>전체매장공통</option>
	<%
		query1 = " select userid,shopname from [db_shop].[dbo].tbl_shop_user  "
		query1 = query1 & " where isusing='Y' "
		query1 = query1 & " and userid<>'streetshop000'"
		query1 = query1 & " and userid<>'streetshop800'"
		query1 = query1 & " and userid<>'streetshop870'"
	
		rsget.Open query1,dbget,1
		
		if  not rsget.EOF  then
		rsget.Movefirst
		
		do until rsget.EOF
		if Lcase(selectedId) = Lcase(rsget("userid")) then
		tmp_str = " selected"
		end if
		response.write("<option value='"&rsget("userid")&"' "&tmp_str&">"&rsget("userid")&"/"&rsget("shopname")&"</option>")
		tmp_str = ""
		rsget.MoveNext
		loop
		end if
		rsget.close
	response.write("</select>")
end sub

'//오프라인매장 구분 매장공통 미포함
Sub drawSelectBoxoneOffShop_off(selectBoxName,selectedId)
	dim tmp_str,query1
	%>
	<select class="select" name="<%=selectBoxName%>">
	<option value='' <%if selectedId="" then response.write " selected"%>>선택하세요</option>	
	<%
		query1 = " select userid,shopname from [db_shop].[dbo].tbl_shop_user  "
		query1 = query1 & " where isusing='Y' "
		query1 = query1 & " and userid<>'streetshop000'"
		query1 = query1 & " and userid<>'streetshop800'"
		query1 = query1 & " and userid<>'streetshop870'"
	
		rsget.Open query1,dbget,1
		
		if  not rsget.EOF  then
		rsget.Movefirst
		
		do until rsget.EOF
		if Lcase(selectedId) = Lcase(rsget("userid")) then
		tmp_str = " selected"
		end if
		response.write("<option value='"&rsget("userid")&"' "&tmp_str&">"&rsget("userid")&"/"&rsget("shopname")&"</option>")
		tmp_str = ""
		rsget.MoveNext
		loop
		end if
		rsget.close
	response.write("</select>")
end sub

'//오프라인매장 구분 매장공통 미포함 , 해외매장 미포함
Sub drawSelectBoxinternalOffShop_off(selectBoxName,selectedId)
	dim tmp_str,query1
	%>
	<select class="select" name="<%=selectBoxName%>">
	<option value='' <%if selectedId="" then response.write " selected"%>>선택하세요</option>	
	<%
		query1 = " select userid,shopname"
		query1 = query1 & " from [db_shop].[dbo].tbl_shop_user"		
		query1 = query1 & " where isusing='Y' "
		query1 = query1 & " and userid<>'streetshop000'"
		query1 = query1 & " and userid<>'streetshop800'"
		query1 = query1 & " and userid<>'streetshop870'"
		query1 = query1 & " and shopdiv <> '7'"
		
		rsget.Open query1,dbget,1
		
		if  not rsget.EOF  then
		rsget.Movefirst
		
		do until rsget.EOF
		if Lcase(selectedId) = Lcase(rsget("userid")) then
		tmp_str = " selected"
		end if
		response.write("<option value='"&rsget("userid")&"' "&tmp_str&">"&rsget("userid")&"/"&rsget("shopname")&"</option>")
		tmp_str = ""
		rsget.MoveNext
		loop
		end if
		rsget.close
	response.write("</select>")
end sub

'//이벤트 공통코드 어플변수에 세팅
Function fnSetEventCommonCode_off
 On Error Resume Next
	Dim strSql, arrList, intLoop
	Dim intI, intJ, arrCode(), strtype
	strSql = " SELECT code_type, code_value, code_desc " + vbcrlf
	strSql = strSql & " FROM [db_shop].[dbo].[tbl_event_off_commoncode] WHERE code_using ='Y' Order by code_type, code_sort "			
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

'//특정종류의 공통코드값의 배열에서 특정값의 코드명 가져오기
Function fnGetCommCodeArrDesc_off(ByVal arrCode, ByVal iCodeValue)
	Dim intLoop		
	IF iCodeValue = "" or isNull(iCodeValue) THEN iCodeValue = -1
	For intLoop =0 To UBound(arrCode,2)		
		IF Cint(iCodeValue) = arrCode(0,intLoop) THEN				
			fnGetCommCodeArrDesc_off = arrCode(1,intLoop)
			Exit For
		END IF	
	Next	
End Function

'//sbGetOptStatusCodeValue_off(변수명, 선택값, '선택'구문 사용유무, 스크립트)
'//이벤트 상태값 공통코드 어플변수 select 구문화 - 현재값의 상태보다 이전값으로 이동 못하도록
 Sub sbGetOptStatusCodeValue_off(ByVal sType, ByVal selValue, ByVal sViewOpt, ByVal sScript)
   Dim arrList, intLoop, iValue
 	arrList= fnSetCommonCodeArr_off(sType, True) 
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
	<select>
<%
 End Sub

'//sbGetOptEventCodeValue_off(변수명, 선택값, '선택'구문 사용유무, 스크립트)
'// : 이벤트 공통코드 어플변수 select 구문화 
 Sub sbGetOptEventCodeValue_off(ByVal sType, ByVal selValue, ByVal sViewOpt, ByVal sScript)
   Dim arrList, intLoop, iValue
 	arrList = Application(sType)
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
%>
	<option value="<%=arrList(0,intLoop)%>" <%If CStr(selValue) = CStr(arrList(0,intLoop)) THEN%>selected<%END IF%>><%=arrList(1,intLoop)%></option>
<%   
	Next %>
	</select>
<%
 End Sub

''// 이벤트 공통코드 가져오기
 Function fnSetCommonCodeArr_off(ByVal code_type, ByVal blnUse) 
	Dim strSql, arrList, intLoop, strAdd
	Dim intI, intJ, arrCode(), strtype
	strAdd = ""
	IF blnUse THEN
		strAdd= " and code_using ='Y' "
	END IF	
	strSql = " SELECT code_value, code_desc " + vbcrlf
	strSql = strSql & " FROM [db_shop].[dbo].[tbl_event_off_commoncode] " + vbcrlf
	strSql = strSql & " WHERE code_type='"&code_type&"' "&strAdd&"" + vbcrlf
	strSql = strSql & " Order by code_type, code_sort "						
	rsget.Open strSql,dbget
	IF not rsget.EOF THEN
		fnSetCommonCodeArr_off = rsget.getRows()
	END IF		
	rsget.close	
End Function	

''// 이벤트 상품이 등록이 되어 있나 체크
 Function geteventcheckitem(ByVal eCode) 
	Dim strSql
	
	geteventcheckitem = false

	strSql = "select top 1 evt_code" + vbcrlf
	strSql = strSql &"  from db_shop.dbo.tbl_eventitem_off" + vbcrlf
	strSql = strSql &"  where evt_code = "&eCode&"" + vbcrlf
	
	'response.write strSql				
	rsget.Open strSql,dbget
	
	IF not rsget.EOF THEN
		geteventcheckitem = true
	END IF		
	rsget.close		
End Function

'//sbGetOptCommonCodeArr_off(변수명, 선택값, '선택'구문 사용유무, 스크립트)
'//특정종류의 공통코드값의 배열에서 select 처리
 Sub sbGetOptCommonCodeArr_off(ByVal sType, ByVal selValue, ByVal sViewOpt, ByVal blnUse, ByVal sScript)
   Dim arrCode, intLoop, iValue 	
   	arrCode= fnSetCommonCodeArr_off(sType, blnUse) 
 	iValue  = selValue
 	IF  isNull(selValue) THEN 	selValue = ""
 	IF selValue = ""  THEN	iValue = 0 	
%>
	<select name="<%=sType%>" <%=sScript%>>
	<%IF sViewOpt THEN%>
	<option value="">선택</option>
    <%END IF%>
<% 	IF isArray(arrCode) THEN
 	For intLoop =0 To UBound(arrCode,2) 	 
%>
	<option value="<%=arrCode(0,intLoop)%>" <%If CStr(selValue) = CStr(arrCode(0,intLoop)) THEN%>selected<%END IF%>><%=arrCode(1,intLoop)%></option>
<%   
	Next 
	End IF
	%>
	<select>
<%
 End Sub

'//담당MD 리스트가져오기 (팀장 미만,직원 이상)
 Sub sbGetMDid_off(ByVal selName, ByVal sIDValue, ByVal sScript)
  Dim strSql, arrList, intLoop
   strSql = " SELECT p.id, isNull(p.company_name,u.username) as company_name from db_partner.[dbo].tbl_partner as p " + vbcrlf
   strSql = strSql & " inner join db_partner.[dbo].tbl_user_tenbyten as u on p.id = u.userid "
   strSql = strSql & " WHERE  p.part_sn ='13' and p.posit_sn>='4' and p.posit_sn<='8' and p.isUsing ='Y' order by p.level_sn"
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
 
Sub sbOptCodeType(ByVal selCodeType)
%>	
	<option value="evt_kind" <%IF Cstr(selCodeType)="evt_kind" THEN%>selected<%END IF%>>이벤트종류</option>	
	<option value="evt_state" <%IF Cstr(selCodeType)="evt_state" THEN%>selected<%END IF%>>상태</option>	
<%			
End Sub
  
'//sbGetOptGiftCodeValue(변수명, 선택값, 스크립트,이벤트코드)
'//사은품공통코드 어플변수 select 구문화 
 Sub sbGetOptGiftCodeValue_off(ByVal sType, ByVal selValue, ByVal sScript,ByVal eCode)
   Dim arrList, intLoop
 	arrList = fnSetCommonCodeArr_off(sType, True) 
 
%>
<select name="<%=sType%>" <%=sScript%>>	
	<% For intLoop =0 To UBound(arrList,2) %>
	<option value="<%=arrList(0,intLoop)%>" <%If CStr(selValue) = CStr(arrList(0,intLoop)) THEN%>selected<%END IF%>><%=arrList(1,intLoop)%></option>
	<% Next %>
<select>
<% 
End Sub

'//매대구분
function getracknum(boxname,stats)
dim i
%>
<select name="<%=boxname%>">	
	<option value='' <% if stats="" then response.write " selected" %>>선택</option>
	<% for i = 1 to 20 %>
	<option value='<%=i%>' <% if stats=i then response.write " selected" %>><%=i%></option>
	<% next %>
<%
end function

'마진형태에 따른 매입가 생성
Function fnSetSaleSupplyPrice(ByVal MarginType, ByVal MarginValue, ByVal orgPrice, ByVal orgSupplyPrice, ByVal salePrice,comm_cd)
	Dim orgMRate
	if orgPrice <>0 then '원 마진율
		orgMRate = 100-fix(orgSupplyPrice/orgPrice*10000)/100
	end if
		
	SELECT CASE MarginType
		Case 1	'동일마진					
			fnSetSaleSupplyPrice = salePrice-fix(salePrice * (orgMRate/100))
		Case 2	'업체부담
			fnSetSaleSupplyPrice = salePrice-(orgPrice-orgSupplyPrice)
		Case 3	'반반부담
			fnSetSaleSupplyPrice = orgSupplyPrice- fix((orgPrice-salePrice)/2)
		Case 4	'10x10부담
			fnSetSaleSupplyPrice = orgSupplyPrice
		Case 5	'직접설정
			fnSetSaleSupplyPrice = salePrice - fix(salePrice*(MarginValue/100))
		Case 6	'업체위탁반반부담/나머지텐바이텐부담
			if comm_cd = "B012" then
				fnSetSaleSupplyPrice = orgSupplyPrice- fix((orgPrice-salePrice)/2)
			else
				fnSetSaleSupplyPrice = orgSupplyPrice
			end if		
		Case 7	'업체위탁,출고위탁,텐바이텐위탁 반반부담/나머지텐바이텐부담
			if comm_cd = "B011" or comm_cd = "B012" or comm_cd = "B013" then
				fnSetSaleSupplyPrice = orgSupplyPrice- fix((orgPrice-salePrice)/2)
			else
				fnSetSaleSupplyPrice = orgSupplyPrice
			end if		
	END SELECT	
End Function

'//날짜 기준 '/이벤트기간,매출기간
function draweventmaechul_datefg(selectBoxName,selectedId,changefg)
%>
<select name="<%=selectBoxName%>" <%=changefg%>>
	<option value='' <%if selectedId="" then response.write " selected"%>>선택</option>	
	<option value='event' <%if selectedId="event" then response.write " selected"%>>이벤트기간</option>
	<option value='jumun' <%if selectedId="jumun" then response.write " selected"%>>매출기간</option>
</select>
<%
end function

'/할인 처리(전체)
function offitemsaleSet_all()
dim sql
		
	sql = "exec db_shop.dbo.sp_Ten_item_SetPrice_off "
	
	'response.write sql &"<Br>"
	dbget.execute sql
end function
%>