<%

function DDotFormat(byval str,byval n)
	DDotFormat = str
	if IsNULL(str) then Exit function

	if Len(str)> n then
		DDotFormat = Left(str,n) + "..."
	end if
end function

function GetImageSubFolderByItemid(byval iitemid)
	GetImageSubFolderByItemid = "0" + CStr(Clng(Clng(iitemid) \ 10000))
end function


Sub DrawYMBox(byval yyyy1,mm1)
	dim buf,i

	buf = "<select name='yyyy1'>"
    for i=2001 to year(dateadd("m",1,now()))
		if (CStr(i)=CStr(yyyy1)) then
			buf = buf + "<option value='" + CStr(i) +"' selected>" + CStr(i) + "</option>"
		else
    	buf = buf + "<option value=" + CStr(i) + " >" + CStr(i) + "</option>"
        end if
	next
    buf = buf + "</select>"

    buf = buf + "<select name='mm1' >"

    for i=1 to 12
		if (Format00(2,i)=Format00(2,mm1)) then
			buf = buf + "<option value='" + Format00(2,i) +"' selected>" + Format00(2,i) + "</option>"
		else
    	    buf = buf + "<option value='" + Format00(2,i) +"' >" + Format00(2,i) + "</option>"
		end if
	next

    buf = buf + "</select>"

    response.write buf
end Sub


Sub DrawOneDateBox(byval yyyy1,mm1,dd1)
	dim buf,i

	buf = "<select name='yyyy1'>"
    buf = buf + "<option value='" + CStr(yyyy1) +"' selected>" + CStr(yyyy1) + "</option>"
    for i=2001 to year(dateadd("m",1,now()))
    	buf = buf + "<option value=" + CStr(i) + " >" + CStr(i) + "</option>"
	next
    buf = buf + "</select>"

    buf = buf + "<select name='mm1' >"
    buf = buf + "<option value='" + CStr(mm1) + "' selected>" + CStr(mm1) + "</option>"

    for i=1 to 12
    	buf = buf + "<option value='" + Format00(2,i) +"' >" + Format00(2,i) + "</option>"
	next

    buf = buf + "</select>"

    buf = buf + "<select name='dd1'>"
    buf = buf + "<option value='" + CStr(dd1) +"' selected>" + CStr(dd1) + "</option>"
    for i=1 to 31
        buf = buf + "<option value='" + Format00(2,i) + "' >" + Format00(2,i) + "</option>"
    next
    buf = buf + "</select>"

    response.write buf
end Sub

Sub DrawOneDateBox2(byval yyyy2,mm2,dd2)
	dim buf,i

	buf = "<select name='yyyy2'>"
    buf = buf + "<option value='" + CStr(yyyy2) +"' selected>" + CStr(yyyy2) + "</option>"
    for i=2001 to year(dateadd("m",1,now()))
    	buf = buf + "<option value=" + CStr(i) + " >" + CStr(i) + "</option>"
	next
    buf = buf + "</select>"

    buf = buf + "<select name='mm2' >"
    buf = buf + "<option value='" + CStr(mm2) + "' selected>" + CStr(mm2) + "</option>"

    for i=1 to 12
    	buf = buf + "<option value='" + Format00(2,i) +"' >" + Format00(2,i) + "</option>"
	next

    buf = buf + "</select>"

    buf = buf + "<select name='dd2'>"
    buf = buf + "<option value='" + CStr(dd2) +"' selected>" + CStr(dd2) + "</option>"
    for i=1 to 31
        buf = buf + "<option value='" + Format00(2,i) + "' >" + Format00(2,i) + "</option>"
    next
    buf = buf + "</select>"

    response.write buf
end Sub

Sub drawSelectBoxSellYN(selectBoxName,selectedId)
   dim tmp_str,query1
   %>
   <select class="select" name="<%=selectBoxName%>">
   <option value="">전체</option>
   <option value="Y" <% if selectedId="Y" then response.write "selected" %> >판매</option>
   <option value="S" <% if selectedId="S" then response.write "selected" %> >일시품절</option>
   <option value="N" <% if selectedId="N" then response.write "selected" %> >품절</option>
   <option value="YS" <% if selectedId="YS" then response.write "selected" %> >판매+일시품절</option>
   </select>
   <%
End Sub

Sub drawSelectBoxUsingYN(selectBoxName,selectedId)
   dim tmp_str,query1
   %>
   <select class="select" name="<%=selectBoxName%>">
   <option value="">전체</option>
   <option value="Y" <% if selectedId="Y" then response.write "selected" %> >사용함</option>
   <option value="N" <% if selectedId="N" then response.write "selected" %> >사용안함</option>
   </select>
   <%
End Sub

Sub drawSelectBoxDanjongYN(selectBoxName,selectedId)
   dim tmp_str,query1
   %>
   <select class="select" name="<%=selectBoxName%>">
   <option value="">전체</option>
   <option value="N" <% if selectedId="N" then response.write "selected" %> >생산중</option>
   <option value="S" <% if selectedId="S" then response.write "selected" %> >재고부족</option>
   <option value="Y" <% if selectedId="Y" then response.write "selected" %> >단종</option>
   <option value="M" <% if selectedId="M" then response.write "selected" %> >MD품절</option>
   <option value="YM" <% if selectedId="YM" then response.write "selected" %> >단종+MD품절</option>
   <option value="SN" <% if selectedId="SN" then response.write "selected" %> >단종아님</option>
   </select>
   <%
End Sub

Sub drawSelectBoxLimitYN(selectBoxName,selectedId)
   dim tmp_str,query1
   %>
   <select class="select" name="<%=selectBoxName%>">
   <option value="">전체</option>
   <option value="N" <% if selectedId="N" then response.write "selected" %> >비한정</option>
   <option value="Y" <% if selectedId="Y" then response.write "selected" %> >한정</option>
   <option value="Y0" <% if selectedId="Y0" then response.write "selected" %> >한정(0)</option>
   </select>
   <%
End Sub

Sub drawSelectBoxMWU(selectBoxName,selectedId)
   dim tmp_str,query1
   %>
   <select class="select" name="<%=selectBoxName%>">
   <option value="">전체</option>
   <option value="MW" <% if selectedId="MW" then response.write "selected" %> >매입+특정</option>
   <option value="W" <% if selectedId="W" then response.write "selected" %> >특정</option>
   <option value="M" <% if selectedId="M" then response.write "selected" %> >매입</option>
   <option value="U" <% if selectedId="U" then response.write "selected" %> >업체</option>
   </select>
   <%
End Sub

Sub drawSelectBoxSailYN(selectBoxName,selectedId)
   dim tmp_str,query1
   %>
   <select class="select" name="<%=selectBoxName%>">
   <option value="">전체</option>
   <option value="Y" <% if selectedId="Y" then response.write "selected" %> >할인</option>
   <option value="N" <% if selectedId="N" then response.write "selected" %> >할인안함</option>
   </select>
   <%
End Sub

Sub drawSelectBoxCouponYN(selectBoxName,selectedId)
   dim tmp_str,query1
   %>
   <select class="select" name="<%=selectBoxName%>">
   <option value="">전체</option>
   <option value="Y" <% if selectedId="Y" then response.write "selected" %> >쿠폰할인</option>
   <option value="N" <% if selectedId="N" then response.write "selected" %> >쿠폰없음</option>
   </select>
   <%
End Sub

Sub drawSelectBoxVatYN(selectBoxName,selectedId)
   dim tmp_str,query1
   %>
   <select class="select" name="<%=selectBoxName%>">
   <option value="">전체</option>
   <option value="Y" <% if selectedId="Y" then response.write "selected" %> >과세</option>
   <option value="N" <% if selectedId="N" then response.write "selected" %> >면세</option>
   </select>
   <%
End Sub

Sub drawSelectBoxIsOverSeaYN(selectBoxName,selectedId)
   %>
   <select class="select" name="<%=selectBoxName%>">
   <option value="">전체</option>
   <option value="Y" <% if selectedId="Y" then response.write "selected" %> >사용</option>
   <option value="N" <% if selectedId="N" then response.write "selected" %> >안함</option>
   </select>
   <%
End Sub

Sub drawSelectBoxIsWeightYN(selectBoxName,selectedId)
   %>
   <select class="select" name="<%=selectBoxName%>">
   <option value="">전체</option>
   <option value="Y" <% if selectedId="Y" then response.write "selected" %> >사용</option>
   <option value="N" <% if selectedId="N" then response.write "selected" %> >안함</option>
   </select>
   <%
End Sub

Sub drawBeadalDiv(selectBoxName,selectedId)
   %>
   <select class="select" name="<%=selectBoxName%>" >
   	 <option value='' <%if selectedId="" then response.write " selected"%>>선택</option>
     <option value='1' <%if selectedId="1" then response.write " selected"%>>텐바이텐배송</option>
	 <option value='2' <%if selectedId="2" OR  selectedId="5" then response.write " selected"%>>업체무료배송</option>
     <option value='4' <%if selectedId="4" then response.write " selected"%>>텐바이텐무료배송</option>
     <!--<option value='5' <%if selectedId="5" then response.write " selected"%>>업체무료배송</option>-->
     <option value='7' <%if selectedId="7" then response.write " selected"%>>업체착불배송</option>
     <option value='9' <%if selectedId="9" then response.write " selected"%>>업체개별배송</option>
   </select>
   <%
end Sub

'// 구분에 따른 문자열 색상 지정
function fnColor(str, div)
	Select Case div
		Case "yn"
			if str<>"Y" or isNull(str) then
				fnColor = "<Font color=#F08050>" & str & "</font>"
			else
				fnColor = "<Font color=#5080F0>" & str & "</font>"
			end if
		Case "mw"
			Select Case str
				Case "M"
					fnColor = "<Font color=#F08050>매입</font>"
				Case "W"
					fnColor = "<Font color=#808080>특정</font>"
				Case "U"
					fnColor = "<Font color=#5080F0>업체</font>"
			end Select
		Case "tx"
			if str="Y" then
				fnColor = "<Font color=#808080>과세</font>"
			else
				fnColor = "<Font color=#F08050>면세</font>"
			end if
		Case "dj"
			if str="Y" then
				fnColor = "<Font color=#33CC33>단종</font>"
			elseif str="S" then
				fnColor = "<Font color=#3333CC>재고부족</font>"
			elseif str="M" then
				fnColor = "<Font color=#CC3333>MD품절</font>"
			end if
		Case "delivery"
			IF str THEN
				fnColor = "<Font color=#F08050>업체</font>"
			ELSE
				fnColor = "<Font color=#5080F0>10x10</font>"
			end IF
		Case "sellyn"
			IF str="N" THEN
				fnColor = "<Font color=#F08050>품절</font>"
			elseif str="S" then
			    fnColor = "<Font color=#3333CC>일시품절</font>"
			end IF
		Case "cancelyn"
			IF str="N" THEN
				fnColor = "<Font color=#000000>정상</font>"
			elseif str="D" then
			    fnColor = "<Font color=#FF0000>삭제</font>"
			elseif str="Y" then
			    fnColor = "<Font color=#FF0000>취소</font>"
			elseif str="A" then
			    fnColor = "<Font color=#FF0000>추가</font>"
			end IF
	end Select
end Function

Sub drawSelectBoxLecturer(selectBoxName,selectedId)
   dim tmp_str,query1
   %><select name="<%=selectBoxName%>">
     <option value='' <%if selectedId="" then response.write " selected"%>>선택</option><%
   query1 = " select userid, socname, socname_kor from [db_user].[dbo].tbl_user_c a where a.userdiv='14' "
   'query1 = query1 + "and a.isusing='Y'"
   rsget.Open query1,dbget,1

   if  not rsget.EOF  then
       rsget.Movefirst

       do until rsget.EOF
           if Lcase(selectedId) = Lcase(rsget("userid")) then
               tmp_str = " selected"
           end if
           response.write("<option value='" & rsget("userid") & "' " & tmp_str & ">" & rsget("userid") & " [" + db2html(rsget("socname_kor")) + "]</option>")
           tmp_str = ""
           rsget.MoveNext
       loop
   end if
   rsget.close
   response.write("</select>")
End Sub

function getWeekdayStr(yyyymmdd)
	dim wd
	if IsNULL(yyyymmdd) then Exit function
	wd = weekday(yyyymmdd)

	select case wd
		case 1
			getWeekdayStr = "<font color=red>일</font>"
		case 2
			getWeekdayStr = "월"
		case 3
			getWeekdayStr = "화"
		case 4
			getWeekdayStr = "수"
		case 5
			getWeekdayStr = "목"
		case 6
			getWeekdayStr = "금"
		case 7
			getWeekdayStr = "<font color=blue>토</font>"
		case else
			getWeekdayStr = yyyymmdd
	end select

end function


Sub DrawDateBox(byval yyyy1,yyyy2,mm1,mm2,dd1,dd2)
	dim buf,i

    dim today_year,today_month,monstart,MonFirstDay,lastdaytemp,result,MonLastDay

today_year = request("Year")   '이번 년
	if today_year = "" then today_year = year(date) end if
today_month = request("Month")    '이번 달
	if today_month = "" then today_month = month(date) end if
monstart=DateSerial(today_year, today_month, 1)
MonFirstDay = day(monstart)

		for lastdaytemp = 28 to 31
			result = DateSerial(today_year, today_month, lastdaytemp)
			if int(today_month) = month(result) then
               MonLastDay = lastdaytemp  '이번 달의 마지막 날..
			end if
		next


	buf = "<select name='yyyy1'>"
    for i=2001 to year(dateadd("m",1,now()))
		if (CStr(i)=CStr(yyyy1)) then
			buf = buf + "<option value='" + CStr(i) +"' selected>" + CStr(i) + "</option>"
		else
    		buf = buf + "<option value=" + CStr(i) + " >" + CStr(i) + "</option>"
		end if
	next
    buf = buf + "</select>"

    buf = buf + "<select name='mm1'>"
    for i=1 to 12
		if (Format00(2,i)=Format00(2,mm1)) then
			buf = buf + "<option value='" + Format00(2,i) +"' selected>" + Format00(2,i) + "</option>"
		else
    	    buf = buf + "<option value='" + Format00(2,i) +"' >" + Format00(2,i) + "</option>"
		end if
	next

    buf = buf + "</select>"

    buf = buf + "<select name='dd1' >"

    for i=1 to 31
		if (Format00(2,i)=Format00(2,dd1)) then
	    buf = buf + "<option value='" + Format00(2,i) +"' selected>" + Format00(2,i) + "</option>"
		else
        buf = buf + "<option value='" + Format00(2,i) + "' >" + Format00(2,i) + "</option>"
		end if
    next
    buf = buf + "</select>"

    buf = buf + "~"

    buf = buf + "<select name='yyyy2'>"
    for i=2001 to year(dateadd("m",1,now()))
		if (CStr(i)=CStr(yyyy2)) then
			buf = buf + "<option value='" + CStr(i) +"' selected>" + CStr(i) + "</option>"
		else
    		buf = buf + "<option value=" + CStr(i) + " >" + CStr(i) + "</option>"
		end if
	next
    buf = buf + "</select>"

    buf = buf + "<select name='mm2'>"
    for i=1 to 12
		if (Format00(2,i)=Format00(2,mm2)) then
			buf = buf + "<option value='" + Cstr(i) +"' selected>" + Format00(2,i) + "</option>"
		else
    	    buf = buf + "<option value='" + Cstr(i) +"' >" + Format00(2,i) + "</option>"
		end if
	next

    buf = buf + "</select>"

    buf = buf + "<select name='dd2' >"
    for i=1 to 31
		if (Format00(2,i)=Format00(2,dd2)) then
			buf = buf + "<option value='" + Format00(2,i) +"' selected>" + Format00(2,i) + "</option>"
		else
    	    buf = buf + "<option value='" + Format00(2,i) +"' >" + Format00(2,i) + "</option>"
		end if
    next
    buf = buf + "</select>"

    response.write buf
end Sub

function DeliverDivCd2Nm(byval divcd)
		if isNull(divcd) then
			DeliverDivCd2Nm = ""
			Exit function
		end if
		   if CStr(divcd) = "1" then
		    DeliverDivCd2Nm =  "한진택배"
		   elseif CStr(divcd) = "2" then
		    DeliverDivCd2Nm =  "현대택배"
		   elseif CStr(divcd) = "3" then
		    DeliverDivCd2Nm =  "대한통운"
		   elseif CStr(divcd) = "4" then
		    DeliverDivCd2Nm =  "CJ GLS"
		   elseif CStr(divcd) = "5" then
		    DeliverDivCd2Nm =  "이클라인"
		   elseif CStr(divcd) = "6" then
		    DeliverDivCd2Nm =  "HTH"
		   elseif CStr(divcd) = "7" then
		    DeliverDivCd2Nm =  "훼미리택배"
		   elseif CStr(divcd) = "8" then
		    DeliverDivCd2Nm =  "우체국"
		   elseif CStr(divcd) = "9" then
		    DeliverDivCd2Nm =  "(구)KGB"
		   elseif CStr(divcd) = "10" then
		    DeliverDivCd2Nm =  "아주택배"
		   elseif CStr(divcd) = "11" then
		    DeliverDivCd2Nm =  "오렌지택배"
		   elseif CStr(divcd) = "12" then
		    DeliverDivCd2Nm =  "한국택배"
		   elseif CStr(divcd) = "13" then
		    DeliverDivCd2Nm =  "옐로우캡"
		   elseif CStr(divcd) = "14" then
		    DeliverDivCd2Nm =  "나이스택배"
		   elseif CStr(divcd) = "15" then
		    DeliverDivCd2Nm =  "중앙택배"
		   elseif CStr(divcd) = "16" then
		    DeliverDivCd2Nm =  "주코택배"
		   elseif CStr(divcd) = "17" then
		    DeliverDivCd2Nm =  "트라넷택배"
		   elseif CStr(divcd) = "18" then
		    DeliverDivCd2Nm =  "로젠택배"
		   elseif CStr(divcd) = "19" then
		    DeliverDivCd2Nm =  "KGB특급택배"
		   elseif CStr(divcd) = "20" then
		    DeliverDivCd2Nm =  "KT로지스"
		   elseif CStr(divcd) = "21" then
		    DeliverDivCd2Nm =  "경동택배"
		   elseif CStr(divcd) = "99" then
		    DeliverDivCd2Nm =  "기타"
		   end if

end function

 Sub sbOptCommCd(ByVal selCd, ByVal sGroupCd)
 	IF sGroupCd = "" THEN Exit Sub
 	Dim strSql, scommCd, scommNm, intLoop, arrComm, intArrCnt
 	strSql = " SELECT commCd, commNm FROM [db_academy].[dbo].[tbl_commCd] WHERE groupCd = '"&sGroupCd&"'"
 	rsACADEMYget.Open strSql, dbACADEMYget, 1
 	If not rsACADEMYget.eof then
 		arrComm = rsACADEMYget.getRows()
 	end if
 	rsACADEMYget.close
 	intArrCnt =  UBound(arrComm,2)
 	ReDim scommCd(intArrCnt), scommNm(intArrCnt)
 		For intLoop	 = 0 To intArrCnt
 		scommCd(intLoop) = arrComm(0,intLoop)
 		scommNm(intLoop) = arrComm(1,intLoop)
%>
	<option value="<%=scommCd(intLoop)%>" <%IF selCd = scommCd(intLoop) THEN %>selected <%END IF%>><%=scommNm(intLoop)%></option>
<% 		Next
 End Sub

  Function fnGetCommNm(ByVal scommCd, ByVal sGroupCd)
 	IF (sGroupCd = "" or scommCd = "" ) THEN Exit Function
 	Dim strSql
 	strSql = " SELECT commNm FROM [db_academy].[dbo].[tbl_commCd] WHERE commCd = '"&scommCd&"' and groupCd = '"&sGroupCd&"'"
 	rsACADEMYget.Open strSql, dbACADEMYget, 1
 	If not rsACADEMYget.eof then
 		fnGetCommNm = rsACADEMYget("commNm")
	end if
 	rsACADEMYget.close
 End Function

'-----------------------------------------------------------------------
' 13.fnSetCommonCodeArr : 이벤트 공통코드 가져오기
'-----------------------------------------------------------------------
 Function fnSetCommonCodeArr(ByVal code_type, ByVal blnUse)
	Dim strSql, arrList, intLoop, strAdd
	Dim intI, intJ, arrCode(), strtype
	strAdd = ""
	IF blnUse THEN
		strAdd= " and code_using ='Y' "
	END IF
	strSql = " SELECT code_value, code_desc FROM [db_academy].[dbo].[tbl_event_commoncode] WHERE code_type='"&code_type&"'"&strAdd&_
			" Order by code_type, code_sort "
	rsACADEMYget.Open strSql,dbACADEMYget
	IF not rsACADEMYget.EOF THEN
		fnSetCommonCodeArr = rsACADEMYget.getRows()
	END IF
	rsACADEMYget.close
End Function

'-----------------------------------------------------------------------
' 2.fnSetEventCommonCode : 이벤트 공통코드 어플변수에 세팅
'-----------------------------------------------------------------------
 Function fnSetEventCommonCode
 On Error Resume Next
	Dim strSql, arrList, intLoop
	Dim intI, intJ, arrCode(), strtype
	strSql = " SELECT code_type, code_value, code_desc FROM [db_academy].dbo.tbl_event_commoncode WHERE code_using ='Y' Order by code_type, code_sort "
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
'--------------------------------------------------------------------------------
 Sub sbGetOptEventCodeValue(ByVal sType, ByVal selValue, ByVal sViewOpt, ByVal sScript)
   Dim arrList, intLoop, iValue

 	arrList = Application(sType)

 	iValue  = selValue
 	IF  isNull(selValue) THEN selValue = ""
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

 '--------------------------------------------------------------------------------
' 15.fnSetStatusDesc
' : 상태값에 따른 상태명
'--------------------------------------------------------------------------------
	Function fnSetStatusDesc(ByVal FState, ByVal FSDate, ByVal FEDate, ByVal FStateDesc)
		IF datediff("d",FSDate,date()) >= 0 and datediff("d",FEDate,date()) <=0 THEN
			fnSetStatusDesc = "오픈"
		ELSEIF datediff("d",FEDate,date()) > 0 THEN
			fnSetStatusDesc = "종료"
		ELSE
			fnSetStatusDesc = FStateDesc
		END IF
	End Function

'--------------------------------------------------------------------------------
' 14.sbGetOptCommonCodeArr(변수명, 선택값, '선택'구문 사용유무, 스크립트)
' : 특정종류의 공통코드값의 배열에서 select 처리
'--------------------------------------------------------------------------------
 Sub sbGetOptCommonCodeArr(ByVal sType, ByVal selValue, ByVal sViewOpt, ByVal blnUse, ByVal sScript)
   Dim arrCode, intLoop, iValue
   	arrCode= fnSetCommonCodeArr(sType, blnUse)
 	iValue  = selValue
 	IF  isNull(selValue) THEN 	selValue = ""
 	IF selValue = ""  THEN	iValue = 0
%>
	<select name="<%=sType%>" <%=sScript%>>
	<%IF sViewOpt THEN%>
	<option value="">선택</option>
    <%END IF%>
	<%IF sType="eventkind" THEN%>
	<option value="1,12,13,16,17,23">#관심항목</option>
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

'--------------------------------------------------------------------------------
' 10.sbGetOptGiftCodeValue(변수명, 선택값,  그룹선택구문 사용유무, 스크립트,이벤트코드)
' : 사은품공통코드 어플변수 select 구문화
' 예외사항: 그룹선택일 경우 선택적 view
'--------------------------------------------------------------------------------
 Sub sbGetOptGiftCodeValue(ByVal sType, ByVal selValue, ByVal sViewOpt, ByVal sScript,ByVal eCode)
   Dim arrList, intLoop
 	arrList = fnSetCommonCodeArr(sType, True)
 	IF  isNull(selValue) THEN selValue = ""
%>
    <select name="<%=sType%>" <%=sScript%>>
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
	<select>
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

'--------------------------------------------------------------------------------
' 11.sbGetOptStatusCodeValue(변수명, 선택값, '선택'구문 사용유무, 스크립트)
' : 이벤트 상태값 공통코드 어플변수 select 구문화 - 현재값의 상태보다 이전값으로 이동 못하도록
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
	<select>
<%
 End Sub

'// 카테고리 클래스 셀렉트 박스 생성
Function makeCateSelectBox(byval cateDiv,selCD)
	dim sql, strRet

	if cateDiv="" then Exit Function
	strRet = ""

	Select Case cateDiv
		Case "CateCD1"
			sql = "Select CateCD1, CateCD1_Name From db_academy.dbo.tbl_lec_Cate1 Order by CateCD1"
		Case "CateCD2"
			sql = "Select CateCD2, CateCD2_Name From db_academy.dbo.tbl_lec_Cate2 Where isusing='Y' Order by SortNo, CateCD2"
		Case "CateCD3"
			sql = "Select CateCD3, CateCD3_Name From db_academy.dbo.tbl_lec_Cate3 Where isusing='Y' Order by SortNo, CateCD3"
	End Select

	rsACADEMYget.Open sql, dbACADEMYget, 1
	if Not(rsACADEMYget.EOF or rsACADEMYget.BOF) then
		strRet = "<select name='" & cateDiv & "'>" &_
				"<option value=''>::선택::</option>"
		Do Until rsACADEMYget.EOF
			if selCD=rsACADEMYget(0) then
				strRet = strRet & "<option value='" & rsACADEMYget(0) & "' selected>" & rsACADEMYget(1) & "</option>"
			else
				strRet = strRet & "<option value='" & rsACADEMYget(0) & "'>" & rsACADEMYget(1) & "</option>"
			end if
		rsACADEMYget.MoveNext
		Loop
		strRet = strRet & "</select>"
	end if
	rsACADEMYget.Close

	makeCateSelectBox = strRet
End Function

'대카테고리 리스트 '/2010.11.10 한용민 추가
Sub DrawSelectBoxacademyCategoryLarge(byval selectBoxName,selectedId)
   dim tmp_str,query1   
%>
	<select class='select' name="<%=selectBoxName%>" onChange="changecontent()">
		<option value="" <% if selectedId="" then response.write " selected"%>>선택</option>
<%
   query1 = " select code_large, code_nm from [db_academy].dbo.tbl_diy_item_Cate_large"
   query1 = query1 + " where display_yn = 'Y'"
   query1 = query1 + " order by code_large Asc"

   rsACADEMYget.Open query1,dbACADEMYget,1

   if  not rsACADEMYget.EOF  then
       rsACADEMYget.Movefirst

       do until rsACADEMYget.EOF
           if Cstr(selectedId) = Cstr(rsACADEMYget("code_large")) then
               tmp_str = " selected"
           end if
           response.write("<option value='"&rsACADEMYget("code_large")&"' "&tmp_str&">"& db2html(rsACADEMYget("code_nm")) &"</option>")
           tmp_str = ""
           rsACADEMYget.MoveNext
       loop
   end if
   rsACADEMYget.close
   response.write("</select>")
end Sub

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
	<option value="<%=arrList(0,intLoop)%>" <%IF Cstr(selPartner) = arrList(0,intLoop) THEN%>selected<%END IF%>><%=arrList(1,intLoop)%></option>
<%	 	Next
	 END IF	
End Sub

'강좌 대카테고리 
Sub DrawSelectBoxLecCategoryLarge(byval selectBoxName ,  selectedId, scriptYN )

   dim tmp_str,query1
   %>
   <select class='select' name="<%=selectBoxName%>" <%=chkiif(scriptYN = "Y","onChange='changecontent()'","")%> id="<%=selectBoxName%>">
     <option value="" <% if selectedId="" Or IsNull(selectedId) then response.write " selected"%> >선택</option>
	<%

   query1 = " select code_large, code_nm from db_academy.dbo.tbl_lec_Cate_large "
   query1 = query1 + " where display_yn = 'Y'"
   query1 = query1 + " order by code_large Asc"

   rsACADEMYget.Open query1,dbACADEMYget,1

   if  not rsACADEMYget.EOF  then
       rsACADEMYget.Movefirst

       do until rsACADEMYget.EOF
           if Trim(Cstr(selectedId)) = Trim(Cstr(rsACADEMYget("code_large"))) then
               tmp_str = "selected"
           end if
           response.write("<option value='"&rsACADEMYget("code_large")&"' " &tmp_str&">"& db2html(rsACADEMYget("code_nm")) &"</option>")
           tmp_str = ""
           rsACADEMYget.MoveNext
       loop
   end if
   rsACADEMYget.close
   response.write("</select>")

end Sub
' 강좌 중카테고리 
Sub DrawSelectBoxLecCategoryMid(byval selectBoxName , largeno , selectedId , scriptYN )
   dim tmp_str,query1
   %>
   <select class='select' name="<%=selectBoxName%>" <%=chkiif(scriptYN = "Y","onChange='changecontent()'","")%> id="<%=selectBoxName%>">
     <option value="" <% if selectedId="" then response.write " selected"%>>선택</option><%
   query1 = " select code_mid, code_nm from db_academy.dbo.tbl_lec_Cate_mid "
   query1 = query1 & " where display_yn = 'Y'"
   query1 = query1 & " and code_large = '" & largeno & "'"
   query1 = query1 & " and code_mid<>0"
   query1 = query1 & " order by code_mid Asc"

   rsACADEMYget.Open query1,dbACADEMYget,1

   if  not rsACADEMYget.EOF  then
       rsACADEMYget.Movefirst

       do until rsACADEMYget.EOF
           if Not(isNull(selectedId)) then
	           if Cstr(selectedId) = Cstr(rsACADEMYget("code_mid")) then
	               tmp_str = " selected"
	           end if
	       end if
           response.write("<option value='"&rsACADEMYget("code_mid")&"' "&tmp_str&">"& db2html(rsACADEMYget("code_nm")) &"</option>")
           tmp_str = ""
           rsACADEMYget.MoveNext
       loop
   end if
   rsACADEMYget.close
   response.write("</select>")

end Sub

'/핑거스 사이트 구분	'/2016.09.20 한용민 생성
Sub drawSelectBox_academy_sitename(selectBoxName, selectedId, chplg)
%>
   <select name="<%=selectBoxName%>" <%= chplg %>>
	   <option value="" <% if selectedId="" then response.write "selected" %>>SELECT</option>
	   <option value="academy" <% if selectedId="academy" then response.write "selected" %>>강좌</option>
	   <option value="diyitem" <% if selectedId="diyitem" then response.write "selected" %>>작품</option>
   </select>
<%
end sub

'/핑거스 사이트 구분	'/2016.09.20 한용민 생성
Sub drawradio_academy_sitename(selectBoxName, selectedId, chplg, allplug)
%>
	<% if allplug="Y" then %>
		<input type="radio" value="" name="<%=selectBoxName%>" <%= chplg %> <% if selectedId="" then response.write " checked" %>>전체
	<% end if %>

	<input type="radio" value="academy" name="<%=selectBoxName%>" <%= chplg %> <% if selectedId="academy" then response.write " checked" %>>강좌
	<input type="radio" value="diyitem" name="<%=selectBoxName%>" <%= chplg %> <% if selectedId="diyitem" then response.write " checked" %>>작품
<%
end sub

'/핑거스 사이트 구분 이름	'/2016.09.20 한용민 생성
Function get_academy_sitename(v)
	if v = "academy" then
		get_academy_sitename = "강좌"
	elseif v = "diyitem" then
		get_academy_sitename = "작품"
	else
		get_academy_sitename = v
	end if
End Function

'/핑거스 판매처 구분	'/2016.09.20 한용민 생성
Sub drawSelectBox_SellChannel(selectBoxName, selectedId, chplg)
%>
   <select name="<%=selectBoxName%>" <%= chplg %>>
	   <option value="" <% if selectedId="" then response.write "selected" %>>SELECT</option>
	   <option value="WEB" <% if selectedId="WEB" then response.write "selected" %>>WWW</option>
	   <option value="MOB" <% if selectedId="MOB" then response.write "selected" %>>모바일</option>
	   <!--<option value="MOBLNK" <% 'if selectedId="MOBLNK" then response.write "selected" %>>모바일_제휴</option>-->
	   <!--<option value="APP" <% 'if selectedId="APP" then response.write "selected" %>>APP</option>-->
	   <!--<option value="APPLNK" <% 'if selectedId="APPLNK" then response.write "selected" %>>APP_제휴</option>-->
	   <!--<option value="OUT" <% 'if selectedId="OUT" then response.write "selected" %>>제휴몰</option>-->
   </select>
<%
end Sub

'/핑거스 판매처 구분 이름	'/2016.09.20 한용민 생성
Function get_SellChannel(v)
	if v = "WEB" then
		get_SellChannel = "WWW"
	elseif v = "MOB" then
		get_SellChannel = "모바일"
	elseif v = "MOBLNK" then
		get_SellChannel = "모바일_제휴"
	elseif v = "APP" then
		get_SellChannel = "APP"
	elseif v = "APPLNK" then
		get_SellChannel = "APP_제휴"
	elseif v = "OUT" then
		get_SellChannel = "제휴몰"
	else
		get_SellChannel = v
	end if
End Function
%>