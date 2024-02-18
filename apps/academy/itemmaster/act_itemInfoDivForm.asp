<%@ codepage="65001" language=vbscript %>
<% option Explicit
Response.CharSet = "utf-8"
%>
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
			strSql = "Select infoDiv from db_academy.dbo.tbl_diy_wait_item where itemid=" & itemid
			rsACADEMYget.Open strSql,dbACADEMYget,1
			if Not(rsACADEMYget.EOF or rsACADEMYget.BOF) then
				infoDiv = rsACADEMYget(0)
			end if
			rsACADEMYget.Close
		end If
	End If 

	'// 항목형태별 입력폼 생성
	function getFormInfoType(icd,inm,idc,itp,irq,cdv,ctn,lnum)
		dim strRst: strRst = ""
		dim arrDv
		If cdv="" Then cdv="N"
		Select Case itp
			
			Case "I" '단어
				strRst = strRst & "<dd id='info" + CStr(lnum) + "'><input type='text' name='infoCont' " & chkIIF(irq="Y","id='[on,off,off,off]["&inm&"]'","") & " value='" & ctn & "' /><input type='hidden' name='infoCd' value='" & icd & "' /><input type='hidden' name='infoChk' value='N' /></dd>" & vbCrLf
				if Not(isNull(idc) or idc="") Then
				strRst = strRst & "<dd class='addition'>" & idc & "</dd>" & vbCrLf
				End If
			Case "T" '문장
				strRst = strRst & "<dd id='info" + CStr(lnum) + "'><textarea name='infoCont' " & chkIIF(irq="Y","id='[on,off,off,off]["&inm&"]'","") & " rows='3'>" & ctn & "</textarea><input type='hidden' name='infoCd' value='" & icd & "' /><input type='hidden' name='infoChk' value='N' /></dd>" & vbCrLf
				if Not(isNull(idc) or idc="") Then
				strRst = strRst & "<dd class='addition'>" & idc & "</dd>" & vbCrLf
				End If
			Case "C" '여부+단어
				strRst = strRst & "<dd class='sltYN selectBtn' id='info" + CStr(lnum) + "'>" & vbCrLf
				If cdv ="Y" Then
				strRst = strRst & "<div class='grid2'><button type='button' class='btnM1 btnGry selected' name='infoChk" & icd & "' value='Y' onclick=chgInfoChk(this," + CStr(lnum) + ",'Y')>Yes</button></div>" & vbCrLf
				Else
				strRst = strRst & "<div class='grid2'><button type='button' class='btnM1 btnGry' name='infoChk" & icd & "' value='Y' onclick=chgInfoChk(this," + CStr(lnum) + ",'Y')>Yes</button></div>" & vbCrLf
				End If
				If cdv ="N" Then
				strRst = strRst & "<div class='grid2'><button type='button' class='btnM1 btnGry selected' name='infoChk" & icd & "' value='N' onclick=chgInfoChk(this," + CStr(lnum) + ",'N')>No</button></div>" & vbCrLf
				Else
				strRst = strRst & "<div class='grid2'><button type='button' class='btnM1 btnGry' name='infoChk" & icd & "' value='N' onclick=chgInfoChk(this," + CStr(lnum) + ",'N')>No</button></div>" & vbCrLf
				End If
				strRst = strRst & "<input type='hidden' name='infoChk' " & chkIIF(irq="Y","id='[on,off,off,off]["&inm&" 여부]'","") & " value='"& cdv &"' />" & vbCrLf
				strRst = strRst & "<input type='hidden' name='infoCd' value='" & icd & "' />" & vbCrLf
				strRst = strRst & "</dd>" & vbCrLf
				strRst = strRst & "<dd><input type='text' name='infoCont' " & chkIIF(irq="Y","id='[on,off,off,off]["&inm&"]'","") & " value='" & ctn & "' /><dd>" & vbCrLf
				if Not(isNull(idc) or idc="") Then
				strRst = strRst & "<dd class='addition'>" & idc & "</dd>" & vbCrLf
				End If
			Case "J" '여부+선택단어
				if Not(isNull(idc) or idc="") then
					arrDv = split(idc,"|")
				else
					redim arrDv(3)
				end if
				if ubound(arrDv)>1 then idc=arrDv(2): else idc=""
				strRst = strRst & "<dd class='sltYN selectBtn' id='info" + CStr(lnum) + "'>" & vbCrLf
				If cdv ="Y" Then
				strRst = strRst & "<div class='grid2'><button type='button' class='btnM1 btnGry selected' name='infoChk" & icd & "' value='Y' onclick=chgInfoSel2(this," + CStr(lnum) + ",'Y') msg='" & arrDv(0) & "'>Yes</button></div>" & vbCrLf
				Else
				strRst = strRst & "<div class='grid2'><button type='button' class='btnM1 btnGry' name='infoChk" & icd & "' value='Y' onclick=chgInfoSel2(this," + CStr(lnum) + ",'Y') msg='" & arrDv(0) & "'>Yes</button></div>" & vbCrLf
				End If
				If cdv ="N" Then
				strRst = strRst & "<div class='grid2'><button type='button' class='btnM1 btnGry selected' name='infoChk" & icd & "' value='N' onclick=chgInfoSel2(this," + CStr(lnum) + ",'N') msg='" & arrDv(1) & "'>No</button></div>" & vbCrLf
				Else
				strRst = strRst & "<div class='grid2'><button type='button' class='btnM1 btnGry' name='infoChk" & icd & "' value='N' onclick=chgInfoSel2(this," + CStr(lnum) + ",'N') msg='" & arrDv(1) & "'>No</button></div>" & vbCrLf
				End If
				strRst = strRst & "<input type='hidden' name='infoChk' value='"& cdv &"' />" & vbCrLf
				strRst = strRst & "<input type='hidden' name='infoCd' value='" & icd & "' />" & vbCrLf
				strRst = strRst & "</dd>" & vbCrLf
				strRst = strRst & "<dd><input type='text' name='infoCont' " & chkIIF(irq="Y","id='[on,off,off,off]["&inm&"]'","") & " " & chkIIF(cdv="N" or isNull(cdv),"readonly","") & " value='" & ctn & "' /></dd>" & vbCrLf
				if Not(isNull(idc) or idc="") Then
				strRst = strRst & "<dd class='addition'>" & idc & "</dd>" & vbCrLf
				End If
			Case "K" '여부+선택문장
				if Not(isNull(idc) or idc="") then
					arrDv = split(idc,"|")
				else
					redim arrDv(3)
				end if
				if ubound(arrDv)>1 then idc=arrDv(2): else idc=""
				strRst = strRst & "<dd class='sltYN selectBtn' id='info" + CStr(lnum) + "'>" & vbCrLf
				If cdv ="Y" Then
				strRst = strRst & "<div class='grid2'><button type='button' class='btnM1 btnGry selected' name='infoChk" & icd & "' value='Y' onclick=chgInfoSel(this," + CStr(lnum) + ",'Y') msg='" & arrDv(0) & "'>Yes</button></div>" & vbCrLf
				Else
				strRst = strRst & "<div class='grid2'><button type='button' class='btnM1 btnGry' name='infoChk" & icd & "' value='Y' onclick=chgInfoSel(this," + CStr(lnum) + ",'Y') msg='" & arrDv(0) & "'>Yes</button></div>" & vbCrLf
				End If
				If cdv ="N" Then
				strRst = strRst & "<div class='grid2'><button type='button' class='btnM1 btnGry selected' name='infoChk" & icd & "' value='N' onclick=chgInfoSel(this," + CStr(lnum) + ",'N') msg='" & arrDv(1) & "'>No</button></div>" & vbCrLf
				Else
				strRst = strRst & "<div class='grid2'><button type='button' class='btnM1 btnGry' name='infoChk" & icd & "' value='N' onclick=chgInfoSel(this," + CStr(lnum) + ",'N') msg='" & arrDv(1) & "'>No</button></div>" & vbCrLf
				End If
				strRst = strRst & "<input type='hidden' name='infoChk' value='"& cdv &"' />" & vbCrLf
				strRst = strRst & "<input type='hidden' name='infoCd' value='" & icd & "' />" & vbCrLf
				strRst = strRst & "</dd>" & vbCrLf
				strRst = strRst & "<dd><textarea name='infoCont' " & chkIIF(irq="Y","id='[on,off,off,off]["&inm&"]'","") & "  rows='3'>" & ctn & "</textarea></dd>" & vbCrLf
				if Not(isNull(idc) or idc="") Then
				strRst = strRst & "<dd class='addition'>" & idc & "</dd>" & vbCrLf
				End If
			Case "M" '날짜(년월)
				strRst = strRst & "<dd id='info" + CStr(lnum) + "'><input type='text' name='infoCont' " & chkIIF(irq="Y","id='[on,off,off,off]["&inm&"]'","") & " maxlength='7' value='" & ctn & "' /><input type='hidden' name='infoCd' value='" & icd & "' /><input type='hidden' name='infoChk' value='N' /></dd>" & vbCrLf
				if Not(isNull(idc) or idc="") Then
				strRst = strRst & "<dd class='addition'>" & idc & "</dd>" & vbCrLf
				End If
			Case "D" '날자(년월일)
				strRst = strRst & "<dd id='info" + CStr(lnum) + "'><input name='infoCont' id='[on,off,off,off]["&inm&"]' value='" & ctn & "' class='text' size='10' maxlength='10' /><input type='hidden' name='infoCd' value='" & icd & "' /><input type='hidden' name='infoChk' value='N' /></dd>" & vbCrLf
				if Not(isNull(idc) or idc="") Then
				strRst = strRst & "<dd class='addition'>" & idc & "</dd>" & vbCrLf
				End If
			Case "P" 'DESC표시
				strRst = strRst & "<dd id='info" + CStr(lnum) + "'><input type='text' name='infoCont' class='text_ro' readonly style='width:80%;' value='" & idc & "' /><input type='hidden' name='infoCd' value='" & icd & "' /><input type='hidden' name='infoChk' value='N' /></dd>" & vbCrLf
				if Not(isNull(idc) or idc="") Then
				strRst = strRst & "<dd class='addition'>" & idc & "</dd>" & vbCrLf
				End If
			Case "N" '입력받지 않음
		End Select

		getFormInfoType = strRst
	end Function
	
	Dim sqlname , infosqlname
	If fingerson = "on" Then
		infosqlname = "db_academy.dbo.tbl_diy_item_infoCode"
		sqlname =     "db_academy.dbo.tbl_diy_wait_item_infoCont"
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

	rsACADEMYget.Open strSql,dbACADEMYget,1
	if Not(rsACADEMYget.EOF or rsACADEMYget.BOF) then
	%>
	<% for i=1 to rsACADEMYget.recordCount %>
	<li id="info_<%=i%>">
		<dl class="infoUnit">
			<dt><strong><%=i & ". " & rsACADEMYget("infoItemName") %></strong></dt>
			<%=getFormInfoType(rsACADEMYget(0),rsACADEMYget(1),replace(rsACADEMYget(2),"'","&#39;"),rsACADEMYget(3),rsACADEMYget(4),rsACADEMYget(5),replace(rsACADEMYget(6),"'","&#39;"),i)%>
		</dl>
	</li>
	<%
			rsACADEMYget.MoveNext
		Next
		rsACADEMYget.Close
	end If
%>
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->