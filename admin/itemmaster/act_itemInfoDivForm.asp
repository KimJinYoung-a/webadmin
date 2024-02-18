<%@ language=vbscript %>
<% option Explicit
Response.CharSet = "euc-kr"
%>
<%
'###########################################################
' Description : 온라인상품등록
' History : 서동석 생성
'			2023.03.02 한용민 수정(상품고시 A/S 책임자/전화번호 수정)
'###########################################################
%>
<!-- #include virtual="/common/incSessionAdminOrUpche.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<%
	dim itemid, infoDiv, oitem, strSql, i , fingerson
	itemid = request("itemid")
	infoDiv = request("ifdv")
	fingerson = request("fingerson") '//핑거스랑 같이 사용

	If fingerson = "on" Then
		if itemid<>"" and infoDiv="" then
			strSql = "Select infoDiv from db_academy.dbo.tbl_diy_item_contents where itemid=" & itemid
			rsACADEMYget.Open strSql,dbACADEMYget,1
			if Not(rsACADEMYget.EOF or rsACADEMYget.BOF) then
				infoDiv = rsACADEMYget(0)
			end if
			rsACADEMYget.Close
		end If
	Else
		if itemid<>"" and infoDiv="" then
			strSql = "Select infoDiv from db_item.dbo.tbl_item_contents where itemid=" & itemid
			rsget.Open strSql,dbget,1
			if Not(rsget.EOF or rsget.BOF) then
				infoDiv = rsget(0)
			end if
			rsget.Close
		end If
	End If 

	if infoDiv="" or isNull(infoDiv) then
		dbget.close
		response.End
	end if

	'// 항목형태별 입력폼 생성
	function getFormInfoType(icd,inm,idc,itp,irq,cdv,ctn)
		dim strRst: strRst = ""
		dim arrDv
		Select Case itp
			Case "I" '단어
				strRst = strRst & "<input type='hidden' name='infoType' value='" & itp & "' />" & vbCrLf
				strRst = strRst & "<input type='hidden' name='infoCd' value='" & icd & "' />" & vbCrLf
				strRst = strRst & "<input type='hidden' name='infoChk' value='N' />" & vbCrLf
				strRst = strRst & "<input type='text' name='infoCont' class='text' " & chkIIF(irq="Y","id='[on,off,off,off]["&inm&"]'","") & " style='width:80%;' value='" & ctn & "' />" & vbCrLf
				if Not(isNull(idc) or idc="") then
					strRst = strRst & "<br />" & idc & vbCrLf
				end if

			Case "T" '문장
				strRst = strRst & "<input type='hidden' name='infoType' value='" & itp & "' />" & vbCrLf
				strRst = strRst & "<input type='hidden' name='infoCd' value='" & icd & "' />" & vbCrLf
				strRst = strRst & "<input type='hidden' name='infoChk' value='N' />" & vbCrLf
				strRst = strRst & "<textarea name='infoCont' class='textarea' " & chkIIF(irq="Y","id='[on,off,off,off]["&inm&"]'","") & " style='width:90%;height:42px'>" & ctn & "</textarea>" & vbCrLf
				if Not(isNull(idc) or idc="") then
					strRst = strRst & "<br />" & idc & vbCrLf
				end if

			Case "C" '여부+단어
				strRst = strRst & "<input type='hidden' name='infoType' value='" & itp & "' />" & vbCrLf
				strRst = strRst & "<label><input type='radio' name='infoChk" & icd & "' value='Y' " & chkIIF(cdv="Y","checked","") & " onclick='chgInfoChk(this)' />Y</label> " & vbCrLf
				strRst = strRst & "<label><input type='radio' name='infoChk" & icd & "' value='N' " & chkIIF(cdv="N","checked","") & " onclick='chgInfoChk(this)' />N</label> " & vbCrLf
				strRst = strRst & "<input type='hidden' name='infoChk' " & chkIIF(irq="Y","id='[on,off,off,off]["&inm&" 여부]'","") & " value='"& cdv &"' />" & vbCrLf
				strRst = strRst & "<input type='hidden' name='infoCd' value='" & icd & "' />" & vbCrLf
				strRst = strRst & "<input type='text' name='infoCont' " & chkIIF(irq="Y","id='[on,off,off,off]["&inm&"]'","") & " class='text' style='width:75%;' value='" & ctn & "' />" & vbCrLf
				if Not(isNull(idc) or idc="") then
					strRst = strRst & "<br />" & idc & vbCrLf
				end if

			Case "J" '여부+선택단어
				if Not(isNull(idc) or idc="") then
					arrDv = split(idc,"|")
				else
					redim arrDv(3)
				end if
				''if isNull(ctn) or ctn="" then ctn = arrDv(1)
				if ubound(arrDv)>1 then idc=arrDv(2): else idc=""
				strRst = strRst & "<input type='hidden' name='infoType' value='" & itp & "' />" & vbCrLf
				strRst = strRst & "<label><input type='radio' name='infoChk" & icd & "' value='Y' " & chkIIF(cdv="Y","checked","") & " onclick='chgInfoSel(this)' msg='" & arrDv(0) & "' />Y</label> " & vbCrLf
				strRst = strRst & "<label><input type='radio' name='infoChk" & icd & "' value='N' " & chkIIF(cdv="N","checked","") & " onclick='chgInfoSel(this)' msg='" & arrDv(1) & "' />N</label> " & vbCrLf
				strRst = strRst & "<input type='hidden' name='infoChk' value='"& cdv &"' />" & vbCrLf
				strRst = strRst & "<input type='hidden' name='infoCd' value='" & icd & "' />" & vbCrLf
				strRst = strRst & "<input type='text' name='infoCont' " & chkIIF(irq="Y","id='[on,off,off,off]["&inm&"]'","") & " " & chkIIF(cdv="N" or isNull(cdv),"readonly","") & " class='" & chkIIF(cdv="Y","text","text_ro") & "' style='width:75%;' value='" & ctn & "' />" & vbCrLf
				if Not(isNull(idc) or idc="") then
					strRst = strRst & "<br />" & idc & vbCrLf
				end if

			Case "K" '여부+선택문장
				if Not(isNull(idc) or idc="") then
					arrDv = split(idc,"|")
				else
					redim arrDv(3)
				end if
				if ubound(arrDv)>1 then idc=arrDv(2): else idc=""
				strRst = strRst & "<input type='hidden' name='infoType' value='" & itp & "' />" & vbCrLf
				strRst = strRst & "<label><input type='radio' name='infoChk" & icd & "' value='Y' " & chkIIF(cdv="Y","checked","") & " onclick='chgInfoSel(this)' msg='" & arrDv(0) & "' />Y</label> " & vbCrLf
				strRst = strRst & "<label><input type='radio' name='infoChk" & icd & "' value='N' " & chkIIF(cdv="N","checked","") & " onclick='chgInfoSel(this)' msg='" & arrDv(1) & "' />N</label><br /> " & vbCrLf
				strRst = strRst & "<input type='hidden' name='infoChk' value='"& cdv &"' />" & vbCrLf
				strRst = strRst & "<input type='hidden' name='infoCd' value='" & icd & "' />" & vbCrLf
				strRst = strRst & "<textarea name='infoCont' " & chkIIF(irq="Y","id='[on,off,off,off]["&inm&"]'","") & " " & chkIIF(cdv="N" or isNull(cdv),"readonly","") & " class='textarea' style='width:90%;height:42px'>" & ctn & "</textarea>" & vbCrLf
				if Not(isNull(idc) or idc="") then
					strRst = strRst & "<br />" & idc & vbCrLf
				end if

			Case "M" '날짜(년월)
				strRst = strRst & "<input type='hidden' name='infoType' value='" & itp & "' />" & vbCrLf
				strRst = strRst & "<input type='hidden' name='infoCd' value='" & icd & "' />" & vbCrLf
				strRst = strRst & "<input type='hidden' name='infoChk' value='N' />" & vbCrLf
				strRst = strRst & "<input type='text' name='infoCont' " & chkIIF(irq="Y","id='[on,off,off,off]["&inm&"]'","") & " class='text' size='8' maxlength='7' value='" & ctn & "' />" & vbCrLf
				if Not(isNull(idc) or idc="") then
					strRst = strRst & idc & vbCrLf
				end if

			Case "D" '날자(년월일)
				strRst = strRst & "<input type='hidden' name='infoType' value='" & itp & "' />" & vbCrLf
				strRst = strRst & "<input type='hidden' name='infoCd' value='" & icd & "' />" & vbCrLf
				strRst = strRst & "<input type='hidden' name='infoChk' value='N' />" & vbCrLf
				strRst = strRst & "<input name='infoCont' id='[on,off,off,off]["&inm&"]' value='" & ctn & "' class='text' size='10' maxlength='10' /><img src='http://webadmin.10x10.co.kr/images/calicon.gif' id='ifCal"&icd&"_trigger' border='0' style='cursor:pointer' align='absmiddle' />" & vbCrLf
				strRst = strRst & "<script language='javascript'>" & vbCrLf
				strRst = strRst & "	var vifCal"&icd&" = new Calendar({" & vbCrLf
				strRst = strRst & "		inputField : '[on,off,off,off]["&inm&"]', trigger    : 'ifCal"&icd&"_trigger'," & vbCrLf
				strRst = strRst & "		bottomBar: true, dateFormat: '%Y.%m.%d'" & vbCrLf
				strRst = strRst & "	});" & vbCrLf
				strRst = strRst & "</script>" & vbCrLf

			Case "P" 'DESC표시
				strRst = strRst & "<input type='hidden' name='infoType' value='" & itp & "' />" & vbCrLf
				strRst = strRst & "<input type='hidden' name='infoCd' value='" & icd & "' />" & vbCrLf
				strRst = strRst & "<input type='hidden' name='infoChk' value='N' />" & vbCrLf
				strRst = strRst & "<input type='text' name='infoCont' class='text_ro' readonly style='width:80%;' value='" & idc & "' />" & vbCrLf

			Case "A" 'a/s
				strRst = strRst & "<input type='hidden' name='infoType' value='" & itp & "' />" & vbCrLf
				strRst = strRst & "<input type='hidden' name='infoCd' value='" & icd & "' />" & vbCrLf
				strRst = strRst & "<input type='hidden' name='infoChk' value='N' />" & vbCrLf
				strRst = strRst & "<input type='text' name='infoCont' class='text' " & chkIIF(irq="Y","id='[on,off,off,off]["&inm&"]'","") & " style='width:80%;' value='" & ctn & "' />" & vbCrLf
				if Not(isNull(idc) or idc="") then
					strRst = strRst & "<br />" & idc & vbCrLf
				end if

			Case "N" '입력받지 않음
		End Select

		getFormInfoType = strRst
	end Function
	
	Dim sqlname , infosqlname
	If fingerson = "on" Then
		infosqlname = "db_academy.dbo.tbl_diy_item_infoCode"
		sqlname = "db_academy.dbo.tbl_diy_item_infoCont"
	Else
		infosqlname = "db_item.dbo.tbl_item_infoCode"
		sqlname = "db_item.dbo.tbl_item_infoCont"
	End If 

	'// 해당 품목내 항목 접수
	strSql = "Select c.infoCd, c.infoItemName, isnull(c.infoDesc,'') infoDesc, c.infoType, c.infoReq "
	strSql = strSql & "	,i.chkDiv, isNull(i.infoContent,'') infoContent "
	strSql = strSql & "from "& infosqlname &" as c "
	strSql = strSql & "	left join "& sqlname &" as i "
	strSql = strSql & "		on c.infoCd=i.infoCd "
	strSql = strSql & "			and i.itemid='" & itemid & "'"
	strSql = strSql & "where c.infoDiv='" & infoDiv & "' "
	strSql = strSql & "	and c.isUsing='Y' "
	strSql = strSql & "order by c.infoSort"

	if fingerson = "on" Then '//핑거스 상세
		rsACADEMYget.Open strSql,dbACADEMYget,1
		if Not(rsACADEMYget.EOF or rsACADEMYget.BOF) then
	%>
	<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
	<% for i=1 to rsACADEMYget.recordCount %>
	<tr align="left">
		<td height="30" width="15%" bgcolor="#F4F4F4"><%=i & ". " & rsACADEMYget("infoItemName") %>:</td>
		<td bgcolor="#FFFFFF"><%=getFormInfoType(rsACADEMYget(0),rsACADEMYget(1),replace(rsACADEMYget(2),"'","&#39;"),rsACADEMYget(3),rsACADEMYget(4),rsACADEMYget(5),replace(rsACADEMYget(6),"'","&#39;"))%></td>
	</tr>
	<%
			rsACADEMYget.MoveNext
		Next
		rsACADEMYget.Close
	%>
	</table>
	<%
		end If
	Else '//텐바이텐 상세
		rsget.Open strSql,dbget,1
		if Not(rsget.EOF or rsget.BOF) then
	%>
	<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
	<% for i=1 to rsget.recordCount %>
	<tr align="left">
		<td height="30" width="15%" bgcolor="#F4F4F4"><%=i & ". " & rsget("infoItemName") %>:</td>
		<td bgcolor="#FFFFFF"><%=getFormInfoType(rsget(0),rsget(1),replace(rsget(2),"'","&#39;"),rsget(3),rsget(4),rsget(5),replace(rsget(6),"'","&#39;"))%></td>
	</tr>
	<%
			rsget.MoveNext
		next
	%>
	</table>
	<%
		end If
		rsget.Close
	End If 
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->