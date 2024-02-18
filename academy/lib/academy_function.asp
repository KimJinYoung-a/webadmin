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
   <option value="">��ü</option>
   <option value="Y" <% if selectedId="Y" then response.write "selected" %> >�Ǹ�</option>
   <option value="S" <% if selectedId="S" then response.write "selected" %> >�Ͻ�ǰ��</option>
   <option value="N" <% if selectedId="N" then response.write "selected" %> >ǰ��</option>
   <option value="YS" <% if selectedId="YS" then response.write "selected" %> >�Ǹ�+�Ͻ�ǰ��</option>
   </select>
   <%
End Sub

Sub drawSelectBoxUsingYN(selectBoxName,selectedId)
   dim tmp_str,query1
   %>
   <select class="select" name="<%=selectBoxName%>">
   <option value="">��ü</option>
   <option value="Y" <% if selectedId="Y" then response.write "selected" %> >�����</option>
   <option value="N" <% if selectedId="N" then response.write "selected" %> >������</option>
   </select>
   <%
End Sub

Sub drawSelectBoxDanjongYN(selectBoxName,selectedId)
   dim tmp_str,query1
   %>
   <select class="select" name="<%=selectBoxName%>">
   <option value="">��ü</option>
   <option value="N" <% if selectedId="N" then response.write "selected" %> >������</option>
   <option value="S" <% if selectedId="S" then response.write "selected" %> >������</option>
   <option value="Y" <% if selectedId="Y" then response.write "selected" %> >����</option>
   <option value="M" <% if selectedId="M" then response.write "selected" %> >MDǰ��</option>
   <option value="YM" <% if selectedId="YM" then response.write "selected" %> >����+MDǰ��</option>
   <option value="SN" <% if selectedId="SN" then response.write "selected" %> >�����ƴ�</option>
   </select>
   <%
End Sub

Sub drawSelectBoxLimitYN(selectBoxName,selectedId)
   dim tmp_str,query1
   %>
   <select class="select" name="<%=selectBoxName%>">
   <option value="">��ü</option>
   <option value="N" <% if selectedId="N" then response.write "selected" %> >������</option>
   <option value="Y" <% if selectedId="Y" then response.write "selected" %> >����</option>
   <option value="Y0" <% if selectedId="Y0" then response.write "selected" %> >����(0)</option>
   </select>
   <%
End Sub

Sub drawSelectBoxMWU(selectBoxName,selectedId)
   dim tmp_str,query1
   %>
   <select class="select" name="<%=selectBoxName%>">
   <option value="">��ü</option>
   <option value="MW" <% if selectedId="MW" then response.write "selected" %> >����+Ư��</option>
   <option value="W" <% if selectedId="W" then response.write "selected" %> >Ư��</option>
   <option value="M" <% if selectedId="M" then response.write "selected" %> >����</option>
   <option value="U" <% if selectedId="U" then response.write "selected" %> >��ü</option>
   </select>
   <%
End Sub

Sub drawSelectBoxSailYN(selectBoxName,selectedId)
   dim tmp_str,query1
   %>
   <select class="select" name="<%=selectBoxName%>">
   <option value="">��ü</option>
   <option value="Y" <% if selectedId="Y" then response.write "selected" %> >����</option>
   <option value="N" <% if selectedId="N" then response.write "selected" %> >���ξ���</option>
   </select>
   <%
End Sub

Sub drawSelectBoxCouponYN(selectBoxName,selectedId)
   dim tmp_str,query1
   %>
   <select class="select" name="<%=selectBoxName%>">
   <option value="">��ü</option>
   <option value="Y" <% if selectedId="Y" then response.write "selected" %> >��������</option>
   <option value="N" <% if selectedId="N" then response.write "selected" %> >��������</option>
   </select>
   <%
End Sub

Sub drawSelectBoxVatYN(selectBoxName,selectedId)
   dim tmp_str,query1
   %>
   <select class="select" name="<%=selectBoxName%>">
   <option value="">��ü</option>
   <option value="Y" <% if selectedId="Y" then response.write "selected" %> >����</option>
   <option value="N" <% if selectedId="N" then response.write "selected" %> >�鼼</option>
   </select>
   <%
End Sub

Sub drawSelectBoxIsOverSeaYN(selectBoxName,selectedId)
   %>
   <select class="select" name="<%=selectBoxName%>">
   <option value="">��ü</option>
   <option value="Y" <% if selectedId="Y" then response.write "selected" %> >���</option>
   <option value="N" <% if selectedId="N" then response.write "selected" %> >����</option>
   </select>
   <%
End Sub

Sub drawSelectBoxIsWeightYN(selectBoxName,selectedId)
   %>
   <select class="select" name="<%=selectBoxName%>">
   <option value="">��ü</option>
   <option value="Y" <% if selectedId="Y" then response.write "selected" %> >���</option>
   <option value="N" <% if selectedId="N" then response.write "selected" %> >����</option>
   </select>
   <%
End Sub

Sub drawBeadalDiv(selectBoxName,selectedId)
   %>
   <select class="select" name="<%=selectBoxName%>" >
   	 <option value='' <%if selectedId="" then response.write " selected"%>>����</option>
     <option value='1' <%if selectedId="1" then response.write " selected"%>>�ٹ����ٹ��</option>
	 <option value='2' <%if selectedId="2" OR  selectedId="5" then response.write " selected"%>>��ü������</option>
     <option value='4' <%if selectedId="4" then response.write " selected"%>>�ٹ����ٹ�����</option>
     <!--<option value='5' <%if selectedId="5" then response.write " selected"%>>��ü������</option>-->
     <option value='7' <%if selectedId="7" then response.write " selected"%>>��ü���ҹ��</option>
     <option value='9' <%if selectedId="9" then response.write " selected"%>>��ü�������</option>
   </select>
   <%
end Sub

'// ���п� ���� ���ڿ� ���� ����
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
					fnColor = "<Font color=#F08050>����</font>"
				Case "W"
					fnColor = "<Font color=#808080>Ư��</font>"
				Case "U"
					fnColor = "<Font color=#5080F0>��ü</font>"
			end Select
		Case "tx"
			if str="Y" then
				fnColor = "<Font color=#808080>����</font>"
			else
				fnColor = "<Font color=#F08050>�鼼</font>"
			end if
		Case "dj"
			if str="Y" then
				fnColor = "<Font color=#33CC33>����</font>"
			elseif str="S" then
				fnColor = "<Font color=#3333CC>������</font>"
			elseif str="M" then
				fnColor = "<Font color=#CC3333>MDǰ��</font>"
			end if
		Case "delivery"
			IF str THEN
				fnColor = "<Font color=#F08050>��ü</font>"
			ELSE
				fnColor = "<Font color=#5080F0>10x10</font>"
			end IF
		Case "sellyn"
			IF str="N" THEN
				fnColor = "<Font color=#F08050>ǰ��</font>"
			elseif str="S" then
			    fnColor = "<Font color=#3333CC>�Ͻ�ǰ��</font>"
			end IF
		Case "cancelyn"
			IF str="N" THEN
				fnColor = "<Font color=#000000>����</font>"
			elseif str="D" then
			    fnColor = "<Font color=#FF0000>����</font>"
			elseif str="Y" then
			    fnColor = "<Font color=#FF0000>���</font>"
			elseif str="A" then
			    fnColor = "<Font color=#FF0000>�߰�</font>"
			end IF
	end Select
end Function

Sub drawSelectBoxLecturer(selectBoxName,selectedId)
   dim tmp_str,query1
   %><select name="<%=selectBoxName%>">
     <option value='' <%if selectedId="" then response.write " selected"%>>����</option><%
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
			getWeekdayStr = "<font color=red>��</font>"
		case 2
			getWeekdayStr = "��"
		case 3
			getWeekdayStr = "ȭ"
		case 4
			getWeekdayStr = "��"
		case 5
			getWeekdayStr = "��"
		case 6
			getWeekdayStr = "��"
		case 7
			getWeekdayStr = "<font color=blue>��</font>"
		case else
			getWeekdayStr = yyyymmdd
	end select

end function


Sub DrawDateBox(byval yyyy1,yyyy2,mm1,mm2,dd1,dd2)
	dim buf,i

    dim today_year,today_month,monstart,MonFirstDay,lastdaytemp,result,MonLastDay

today_year = request("Year")   '�̹� ��
	if today_year = "" then today_year = year(date) end if
today_month = request("Month")    '�̹� ��
	if today_month = "" then today_month = month(date) end if
monstart=DateSerial(today_year, today_month, 1)
MonFirstDay = day(monstart)

		for lastdaytemp = 28 to 31
			result = DateSerial(today_year, today_month, lastdaytemp)
			if int(today_month) = month(result) then
               MonLastDay = lastdaytemp  '�̹� ���� ������ ��..
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
		    DeliverDivCd2Nm =  "�����ù�"
		   elseif CStr(divcd) = "2" then
		    DeliverDivCd2Nm =  "�����ù�"
		   elseif CStr(divcd) = "3" then
		    DeliverDivCd2Nm =  "�������"
		   elseif CStr(divcd) = "4" then
		    DeliverDivCd2Nm =  "CJ GLS"
		   elseif CStr(divcd) = "5" then
		    DeliverDivCd2Nm =  "��Ŭ����"
		   elseif CStr(divcd) = "6" then
		    DeliverDivCd2Nm =  "HTH"
		   elseif CStr(divcd) = "7" then
		    DeliverDivCd2Nm =  "�ѹ̸��ù�"
		   elseif CStr(divcd) = "8" then
		    DeliverDivCd2Nm =  "��ü��"
		   elseif CStr(divcd) = "9" then
		    DeliverDivCd2Nm =  "(��)KGB"
		   elseif CStr(divcd) = "10" then
		    DeliverDivCd2Nm =  "�����ù�"
		   elseif CStr(divcd) = "11" then
		    DeliverDivCd2Nm =  "�������ù�"
		   elseif CStr(divcd) = "12" then
		    DeliverDivCd2Nm =  "�ѱ��ù�"
		   elseif CStr(divcd) = "13" then
		    DeliverDivCd2Nm =  "���ο�ĸ"
		   elseif CStr(divcd) = "14" then
		    DeliverDivCd2Nm =  "���̽��ù�"
		   elseif CStr(divcd) = "15" then
		    DeliverDivCd2Nm =  "�߾��ù�"
		   elseif CStr(divcd) = "16" then
		    DeliverDivCd2Nm =  "�����ù�"
		   elseif CStr(divcd) = "17" then
		    DeliverDivCd2Nm =  "Ʈ����ù�"
		   elseif CStr(divcd) = "18" then
		    DeliverDivCd2Nm =  "�����ù�"
		   elseif CStr(divcd) = "19" then
		    DeliverDivCd2Nm =  "KGBƯ���ù�"
		   elseif CStr(divcd) = "20" then
		    DeliverDivCd2Nm =  "KT������"
		   elseif CStr(divcd) = "21" then
		    DeliverDivCd2Nm =  "�浿�ù�"
		   elseif CStr(divcd) = "99" then
		    DeliverDivCd2Nm =  "��Ÿ"
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
' 13.fnSetCommonCodeArr : �̺�Ʈ �����ڵ� ��������
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
' 2.fnSetEventCommonCode : �̺�Ʈ �����ڵ� ���ú����� ����
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

		If strtype <> Trim(arrList(0, intI)) And Not IsEmpty(strtype) Then intJ = 0	' �ε��� �ʱ�ȭ
		ReDim Preserve arrCode(Ubound(arrList)-1, intJ)	' �迭 Ȯ��
		strtype = Trim(arrList(0, intI))  	' ���� ����
		arrCode(0, intJ) = Trim(arrList(1, intI)) ' �ڵ� ����
		arrCode(1, intJ) = Trim(arrList(2, intI)) ' �ڵ�� ����

		intJ = intJ + 1 ' �ε��� ����

		If intI = Ubound(arrList, 2) Then
			Application.Lock
			Application(Trim(strtype)) = arrCode
			Application.UnLock
		End If

	Next
End Function

'--------------------------------------------------------------------------------
' 3.sbGetOptEventCodeValue(������, ���ð�, '����'���� �������, ��ũ��Ʈ)
' : �̺�Ʈ �����ڵ� ���ú��� select ����ȭ
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
		<option value="">����</option>
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
' : ���°��� ���� ���¸�
'--------------------------------------------------------------------------------
	Function fnSetStatusDesc(ByVal FState, ByVal FSDate, ByVal FEDate, ByVal FStateDesc)
		IF datediff("d",FSDate,date()) >= 0 and datediff("d",FEDate,date()) <=0 THEN
			fnSetStatusDesc = "����"
		ELSEIF datediff("d",FEDate,date()) > 0 THEN
			fnSetStatusDesc = "����"
		ELSE
			fnSetStatusDesc = FStateDesc
		END IF
	End Function

'--------------------------------------------------------------------------------
' 14.sbGetOptCommonCodeArr(������, ���ð�, '����'���� �������, ��ũ��Ʈ)
' : Ư�������� �����ڵ尪�� �迭���� select ó��
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
	<option value="">����</option>
    <%END IF%>
	<%IF sType="eventkind" THEN%>
	<option value="1,12,13,16,17,23">#�����׸�</option>
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
' 10.sbGetOptGiftCodeValue(������, ���ð�,  �׷켱�ñ��� �������, ��ũ��Ʈ,�̺�Ʈ�ڵ�)
' : ����ǰ�����ڵ� ���ú��� select ����ȭ
' ���ܻ���: �׷켱���� ��� ������ view
'--------------------------------------------------------------------------------
 Sub sbGetOptGiftCodeValue(ByVal sType, ByVal selValue, ByVal sViewOpt, ByVal sScript,ByVal eCode)
   Dim arrList, intLoop
 	arrList = fnSetCommonCodeArr(sType, True)
 	IF  isNull(selValue) THEN selValue = ""
%>
    <select name="<%=sType%>" <%=sScript%>>
	<% if selValue="1" then %>
	        <option value="1" selected >��ü����</option>
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
' 12.fnGetCommCodeArrDesc : Ư�������� �����ڵ尪�� �迭���� Ư������ �ڵ�� ��������
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
' 11.sbGetOptStatusCodeValue(������, ���ð�, '����'���� �������, ��ũ��Ʈ)
' : �̺�Ʈ ���°� �����ڵ� ���ú��� select ����ȭ - ���簪�� ���º��� ���������� �̵� ���ϵ���
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
	<option value="">����</option>
    <%END IF%>
<%
 	For intLoop =0 To UBound(arrList,2)
 	 IF Cint(arrList(0,intLoop)) >= Cint(iValue) OR sViewOpt THEN
%>
	<option value="<%=arrList(0,intLoop)%>" <%If CStr(selValue) = CStr(arrList(0,intLoop)) THEN%>selected<%END IF%>><%=replace(arrList(1,intLoop),"���¿���","����")%></option>
<%   END IF
	Next %>
	<select>
<%
 End Sub

'// ī�װ� Ŭ���� ����Ʈ �ڽ� ����
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
				"<option value=''>::����::</option>"
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

'��ī�װ� ����Ʈ '/2010.11.10 �ѿ�� �߰�
Sub DrawSelectBoxacademyCategoryLarge(byval selectBoxName,selectedId)
   dim tmp_str,query1   
%>
	<select class='select' name="<%=selectBoxName%>" onChange="changecontent()">
		<option value="" <% if selectedId="" then response.write " selected"%>>����</option>
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

'���� ��ī�װ� 
Sub DrawSelectBoxLecCategoryLarge(byval selectBoxName ,  selectedId, scriptYN )

   dim tmp_str,query1
   %>
   <select class='select' name="<%=selectBoxName%>" <%=chkiif(scriptYN = "Y","onChange='changecontent()'","")%> id="<%=selectBoxName%>">
     <option value="" <% if selectedId="" Or IsNull(selectedId) then response.write " selected"%> >����</option>
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
' ���� ��ī�װ� 
Sub DrawSelectBoxLecCategoryMid(byval selectBoxName , largeno , selectedId , scriptYN )
   dim tmp_str,query1
   %>
   <select class='select' name="<%=selectBoxName%>" <%=chkiif(scriptYN = "Y","onChange='changecontent()'","")%> id="<%=selectBoxName%>">
     <option value="" <% if selectedId="" then response.write " selected"%>>����</option><%
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

'/�ΰŽ� ����Ʈ ����	'/2016.09.20 �ѿ�� ����
Sub drawSelectBox_academy_sitename(selectBoxName, selectedId, chplg)
%>
   <select name="<%=selectBoxName%>" <%= chplg %>>
	   <option value="" <% if selectedId="" then response.write "selected" %>>SELECT</option>
	   <option value="academy" <% if selectedId="academy" then response.write "selected" %>>����</option>
	   <option value="diyitem" <% if selectedId="diyitem" then response.write "selected" %>>��ǰ</option>
   </select>
<%
end sub

'/�ΰŽ� ����Ʈ ����	'/2016.09.20 �ѿ�� ����
Sub drawradio_academy_sitename(selectBoxName, selectedId, chplg, allplug)
%>
	<% if allplug="Y" then %>
		<input type="radio" value="" name="<%=selectBoxName%>" <%= chplg %> <% if selectedId="" then response.write " checked" %>>��ü
	<% end if %>

	<input type="radio" value="academy" name="<%=selectBoxName%>" <%= chplg %> <% if selectedId="academy" then response.write " checked" %>>����
	<input type="radio" value="diyitem" name="<%=selectBoxName%>" <%= chplg %> <% if selectedId="diyitem" then response.write " checked" %>>��ǰ
<%
end sub

'/�ΰŽ� ����Ʈ ���� �̸�	'/2016.09.20 �ѿ�� ����
Function get_academy_sitename(v)
	if v = "academy" then
		get_academy_sitename = "����"
	elseif v = "diyitem" then
		get_academy_sitename = "��ǰ"
	else
		get_academy_sitename = v
	end if
End Function

'/�ΰŽ� �Ǹ�ó ����	'/2016.09.20 �ѿ�� ����
Sub drawSelectBox_SellChannel(selectBoxName, selectedId, chplg)
%>
   <select name="<%=selectBoxName%>" <%= chplg %>>
	   <option value="" <% if selectedId="" then response.write "selected" %>>SELECT</option>
	   <option value="WEB" <% if selectedId="WEB" then response.write "selected" %>>WWW</option>
	   <option value="MOB" <% if selectedId="MOB" then response.write "selected" %>>�����</option>
	   <!--<option value="MOBLNK" <% 'if selectedId="MOBLNK" then response.write "selected" %>>�����_����</option>-->
	   <!--<option value="APP" <% 'if selectedId="APP" then response.write "selected" %>>APP</option>-->
	   <!--<option value="APPLNK" <% 'if selectedId="APPLNK" then response.write "selected" %>>APP_����</option>-->
	   <!--<option value="OUT" <% 'if selectedId="OUT" then response.write "selected" %>>���޸�</option>-->
   </select>
<%
end Sub

'/�ΰŽ� �Ǹ�ó ���� �̸�	'/2016.09.20 �ѿ�� ����
Function get_SellChannel(v)
	if v = "WEB" then
		get_SellChannel = "WWW"
	elseif v = "MOB" then
		get_SellChannel = "�����"
	elseif v = "MOBLNK" then
		get_SellChannel = "�����_����"
	elseif v = "APP" then
		get_SellChannel = "APP"
	elseif v = "APPLNK" then
		get_SellChannel = "APP_����"
	elseif v = "OUT" then
		get_SellChannel = "���޸�"
	else
		get_SellChannel = v
	end if
End Function
%>