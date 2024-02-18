<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/admin/etc/common/outmallCommonFunction.asp"-->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbCTopen.asp" -->
<script type="text/javascript" src="/js/jquery-1.7.2.min.js"></script>
<%
Dim mallid, itemid, linkgbn, mode , linkgbnName, poomok, stDt, edDt
Dim linkyn, valtype, intval,shortVal,textVal,regdate,theval
Dim valType_text, suntakgbn, i, qy
mallid			= request("mallid")
itemid			= request("itemid")
linkgbn	     	= request("linkgbn")
mode        	= request("mode")
poomok			= request("poomok")
linkyn			= request("linkyn")
stDt			= request("stDt")
edDt			= request("edDt")
valType_text	= request("valType_text")
suntakgbn		= request("suntakgbn")
qy				= request("qy")

Dim strSql, strSql2, cnt
Dim imat_Name, imat_percent, imat_place, imaterial
Dim linkgbnList
strSql = ""
strSql = strSql & " SELECT T.linkgbn, G.linkDesc from  db_item.dbo.tbl_OutMall_etcLink as T" & VBCRLF
strSql = strSql & " JOIN db_item.dbo.tbl_OutMall_etcLinkGubun as G on T.valtype = G.valtype "
strSql = strSql & " where T.itemid = '"&itemid&"' AND T.mallid in ('','"&mallid&"') "
rsget.Open strSql, dbget, 1
IF not rsget.EOF THEN
	linkgbnList = rsget.getRows()
END IF
rsget.Close

If mode = "I" Then
	strSql = "SELECT count(*) as cnt From db_item.dbo.tbl_OutMall_etcLink where (mallid = 'lotteCom' OR isnull(mallid,'') = '') and itemid = '"&itemid&"' and linkgbn = '"&linkgbn&"' "
	rsget.Open strSql, dbget, 1
		If rsget("cnt") = 1 Then
			response.write "<script>alert('이미 등록된 적용몰, 구분이 있습니다.\n상단의 라디오버튼을 클릭해서 수정하세요');location.replace('/admin/etc/common/popOutMallEtcLink.asp?mallid="&mallid&"&itemid="&itemid&"&poomok="&poomok&"');</script>"
			response.end
		End If
	rsget.Close

	If linkgbn = "infoDiv21Lotte" Then
		imat_Name		= request("mat_Name")
		imat_percent	= request("mat_percent")
		imat_place		= request("mat_place")
		imaterial		= imat_Name&"!!^^"&imat_percent&"!!^^"&imat_place

		strSql = ""
		strSql = strSql & " INSERT INTO db_item.dbo.tbl_OutMall_etcLink(itemid, mallid, linkgbn, linkyn, valtype, intval, shortVal, textVal, stDt, edDt, regdate) VALUES " & VBCRLF
		strSql = strSql & " ('"&itemid&"', '"&mallid&"', 'infoDiv21Lotte', '"&linkyn&"', '2', '', '"&imaterial&"', '', '"&stDt&" 00:00:00', '"&edDt&" 23:59:59', getdate()) "

		strSql2 = ""
		strSql2 = strSql2 & " INSERT INTO db_outmall.dbo.tbl_OutMall_etcLink(itemid, mallid, linkgbn, linkyn, valtype, intval, shortVal, textVal, stDt, edDt, regdate) VALUES " & VBCRLF
		strSql2 = strSql2 & " ('"&itemid&"', '"&mallid&"', 'infoDiv21Lotte', '"&linkyn&"', '2', '', '"&imaterial&"', '', '"&stDt&" 00:00:00', '"&edDt&" 23:59:59', getdate()) "
	ElseIf linkgbn = "contents" Then
		strSql = ""
		strSql = strSql & " INSERT INTO db_item.dbo.tbl_OutMall_etcLink(itemid, mallid, linkgbn, linkyn, valtype, intval, shortVal, textVal, stDt, edDt, regdate) VALUES " & VBCRLF
		strSql = strSql & " ('"&itemid&"', '"&mallid&"', 'contents', '"&linkyn&"', '3', '', '', '"&db2html(valType_text)&"', '"&stDt&" 00:00:00', '"&edDt&" 23:59:59', getdate()) "

		strSql2 = ""
		strSql2 = strSql2 & " INSERT INTO db_outmall.dbo.tbl_OutMall_etcLink(itemid, mallid, linkgbn, linkyn, valtype, intval, shortVal, textVal, stDt, edDt, regdate) VALUES " & VBCRLF
		strSql2 = strSql2 & " ('"&itemid&"', '"&mallid&"', 'contents', '"&linkyn&"', '3', '', '', '"&db2html(valType_text)&"', '"&stDt&" 00:00:00', '"&edDt&" 23:59:59', getdate()) "
	ElseIf linkgbn = "topContents" Then
		strSql = ""
		strSql = strSql & " INSERT INTO db_item.dbo.tbl_OutMall_etcLink(itemid, mallid, linkgbn, linkyn, valtype, intval, shortVal, textVal, stDt, edDt, regdate) VALUES " & VBCRLF
		strSql = strSql & " ('"&itemid&"', '"&mallid&"', 'topContents', '"&linkyn&"', '1', '', '', '"&db2html(valType_text)&"', '"&stDt&" 00:00:00', '"&edDt&" 23:59:59', getdate()) "

		strSql2 = ""
		strSql2 = strSql2 & " INSERT INTO db_outmall.dbo.tbl_OutMall_etcLink(itemid, mallid, linkgbn, linkyn, valtype, intval, shortVal, textVal, stDt, edDt, regdate) VALUES " & VBCRLF
		strSql2 = strSql2 & " ('"&itemid&"', '"&mallid&"', 'topContents', '"&linkyn&"', '1', '', '', '"&db2html(valType_text)&"', '"&stDt&" 00:00:00', '"&edDt&" 23:59:59', getdate()) "
	ElseIf linkgbn = "donotEdit" Then
		strSql = ""
		strSql = strSql & " INSERT INTO db_item.dbo.tbl_OutMall_etcLink(itemid, mallid, linkgbn, linkyn, valtype, intval, shortVal, textVal, stDt, edDt, regdate) VALUES " & VBCRLF
		strSql = strSql & " ('"&itemid&"', '"&mallid&"', 'donotEdit', '"&linkyn&"', '0', '', '', '', '"&stDt&" 00:00:00', '"&edDt&" 23:59:59', getdate()) "

		strSql2 = ""
		strSql2 = strSql2 & " INSERT INTO db_outmall.dbo.tbl_OutMall_etcLink(itemid, mallid, linkgbn, linkyn, valtype, intval, shortVal, textVal, stDt, edDt, regdate) VALUES " & VBCRLF
		strSql2 = strSql2 & " ('"&itemid&"', '"&mallid&"', 'donotEdit', '"&linkyn&"', '0', '', '', '', '"&stDt&" 00:00:00', '"&edDt&" 23:59:59', getdate()) "
	End If
	dbget.execute strSql
	dbCTget.execute strSql2
	response.write "<script>alert('등록 되었습니다');window.close();</script>"
ElseIf mode = "V" Then
	''전체 적용몰(mallid='')이 있으면 전체로 지정.
	strSql = " select top 1 L.itemid,L.mallid,L.linkgbn,L.linkyn,L.valtype,L.intval,L.shortVal,L.textVal,L.regdate,G.linkDesc, L.stDt, L.edDt" & VbCRLF
	strSql = strSql & " from db_item.dbo.tbl_OutMall_etcLink L" & VbCRLF
	strSql = strSql & "     Join db_item.dbo.tbl_OutMall_etcLinkGubun G"& VbCRLF
	strSql = strSql & "     on L.linkgbn=G.linkgbn"& VbCRLF
	strSql = strSql & " where L.itemid="&itemid & VbCRLF
	strSql = strSql & " and L.linkgbn='"&linkgbn&"'" & VbCRLF
	strSql = strSql & " and L.mallid in ('','"&mallid&"')" & VbCRLF
	strSql = strSql & " order by L.mallid"
	rsget.Open strSql, dbget, 1
	If Not(rsget.EOF or rsget.BOF) Then
	    cnt = 1
	    mallid  = rsget("mallid")
	    linkyn  = rsget("linkyn")
	    linkgbn = rsget("linkgbn")
	    valtype = rsget("valtype")
	    intval  = rsget("intval")
	    shortVal = rsget("shortVal")
	    textVal = rsget("textVal")
	    regdate = rsget("regdate")
	    linkgbnName = rsget("linkDesc")
	    stDt = rsget("stDt")
	    edDt = rsget("edDt")
	End If
	rsget.close
ElseIf mode = "U" Then
	Dim addSql
	If linkgbn = "infoDiv21Lotte" Then
		imat_Name		= request("mat_Name")
		imat_percent	= request("mat_percent")
		imat_place		= request("mat_place")
		imaterial		= imat_Name&"!!^^"&imat_percent&"!!^^"&imat_place
		addSql = " ,shortVal = '"&imaterial&"' "
	ElseIf linkgbn = "contents" OR linkgbn = "topContents" Then
		addSql = " ,textVal = '"&db2html(valType_text)&"' "
	End If

	strSql = ""
	strSql = strSql & " UPDATE db_item.dbo.tbl_OutMall_etcLink SET " & VBCRLF
	strSql = strSql & " stDt = '"&stDt&" 00:00:00' " & VBCRLF
	strSql = strSql & " ,edDt = '"&edDt&" 23:59:59' " & VBCRLF
	strSql = strSql & addSql & VBCRLF
	strSql = strSql & " where itemid = '"&itemid&"' AND mallid = '"&mallid&"' AND linkgbn = '"&linkgbn&"' "
	dbget.execute strSql

	strSql2 = ""
	strSql2 = strSql2 & " UPDATE db_outmall.dbo.tbl_OutMall_etcLink SET " & VBCRLF
	strSql2 = strSql2 & " stDt = '"&stDt&" 00:00:00' " & VBCRLF
	strSql2 = strSql2 & " ,edDt = '"&edDt&" 23:59:59' " & VBCRLF
	strSql2 = strSql2 & addSql & VBCRLF
	strSql2 = strSql2 & " where itemid = '"&itemid&"' AND mallid = '"&mallid&"' AND linkgbn = '"&linkgbn&"' "
	dbCTget.execute strSql2
	response.write "<script>alert('수정 되었습니다');window.close();</script>"
ElseIf mode = "D" Then
	strSql = ""
	strSql = strSql & " DELETE FROM db_item.dbo.tbl_OutMall_etcLink " & VBCRLF
	strSql = strSql & " where itemid = '"&itemid&"' AND mallid = '"&mallid&"' AND linkgbn = '"&linkgbn&"' "
	dbget.execute strSql

	strSql2 = ""
	strSql2 = strSql2 & " DELETE FROM db_outmall.dbo.tbl_OutMall_etcLink " & VBCRLF
	strSql2 = strSql2 & " where itemid = '"&itemid&"' AND mallid = '"&mallid&"' AND linkgbn = '"&linkgbn&"' "
	dbCTget.execute strSql2
	response.write "<script>alert('삭제 되었습니다');window.close();</script>"
End If
%>
<script>
function inputNumCom(){
	var keycode = event.keyCode;
	if( !((48 <= keycode && keycode <=57) || keycode == 13 || keycode == 46) ){
		alert("숫자만 입력 가능합니다.!");
		event.keyCode = 0;
	}
}
function lgbn(str){
	if(str == 'donotEdit'){
		$("#lkyn").hide();
		$("#ltype_1").hide();
		$("#ltype_3").hide();
		$("#ltype_0").hide();
	}else if(str == 'infoDiv21Lotte'){
		$("#lkyn").show();
		$("#ltype_1").hide();
		$("#ltype_3").hide();
		$("#ltype_0").show();
	}else{
		$("#lkyn").show();
		$("#ltype_1").hide();
		$("#ltype_3").show();
		$("#ltype_0").hide();
	}
}
function frm_Act(){
	var frm = document.frmAct;

	if("<%=cnt%>" > 0){
		document.getElementById('mode').value = 'U';
	}else{
		document.getElementById('mode').value = 'I';
	}

	if(frm.stDt.value == ''){
		alert('적용 시작일을 입력하세요');
		frm.stDt.focus();
		return;
	}

	if(frm.edDt.value == ''){
		alert('적용 마지막을 입력하세요');
		frm.edDt.focus();
		return;
	}
	if(confirm("저장 하시겠습니까?")){
// 이노디터로 저장한 값을 textarea에 할당 시작
		var strHTMLCode = fnGetEditorHTMLCode(true, 0);
		frm["valType_text"].value = strHTMLCode;
// 이노디터로 저장한 값을 textarea에 할당 끝
		frm.submit();
	}
}
function inputNumCom(){
	var keycode = event.keyCode;
	if( !((48 <= keycode && keycode <=57) || keycode == 13 || keycode == 46) ){
		alert("숫자만 입력 가능합니다.!");
		event.keyCode = 0;
	}
}
function frm_Del(){
	if(confirm("삭제 하시겠습니까?")){
		location.replace('/admin/etc/common/popOutMallEtcLink.asp?mode=D&mallid=<%=mallid%>&itemid=<%=itemid%>&poomok=<%=poomok%>&linkgbn=<%=linkgbn%>');
	}
}
</script>
<!-- 이노디터 인크루드 JS -->
<script language="javascript" type="text/javascript">
	var g_arrSetEditorArea = new Array();
	g_arrSetEditorArea[0] = "EDITOR_AREA_CONTAINER";
</script>
<script language="javascript" type="text/javascript" src="/lib/util/innoditor/js/customize.js"></script>
<script language="javascript" type="text/javascript" src="/lib/util/innoditor/js/customize_ui.js"></script>
<script language="javascript" type="text/javascript" src="/lib/util/innoditor/js/loadlayer.js"></script>
<script language="javascript" type="text/javascript">
	//이노디터에서 업로드 할 URL설정
	//Fd로 저장될 폴더를 파라메타로 넘기고 webimage에서 폴더를 만들어줘야한다.///webimage/innoditor/파라메타값
	var g_strUploadImageURL = "/lib/util/innoditor/pop_upload_img.asp?Fd=jaehyumall";

	// 크기, 높이 재정의
	g_nEditorWidth = 800;
	g_nEditorHeight = 1000;
</script>
<!-- 이노디터 인크루드 JS 끝 -->
<link rel="stylesheet" href="/css/tpl.css" type="text/css">
<% If isarray(linkgbnList) Then
		If Ubound(linkgbnList,2) = 0 AND qy = "" Then
			response.write "<script>location.replace('/admin/etc/common/popOutMallEtcLink.asp?mode=V&mallid="&mallid&"&itemid="&itemid&"&poomok="&poomok&"&linkgbn="&linkgbnList(0,0)&"&qy=Y')</script>"
		End If
%>
*등록된 데이터 수정시엔 하단 라디오버튼을 클릭하세요
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr bgcolor="#FFFFFF">
	<td width="50"></td>
	<td>구분선택</td>
</tr>
	<% For i = 0 to Ubound(linkgbnList,2) %>
<tr bgcolor="#FFFFFF">
	<td align="center"><input type="radio" name="lg" <%=chkiif(linkgbn=linkgbnList(0,i),"checked","")%>  onclick="javascript:location.replace('/admin/etc/common/popOutMallEtcLink.asp?mode=V&mallid=<%=mallid%>&itemid=<%=itemid%>&poomok=<%=poomok%>&linkgbn=<%=linkgbnList(0,i)%>');"></td>
	<td><%=linkgbnList(1,i)%></td>
</tr>
	<% Next %>
<tr bgcolor="#FFFFFF">
	<td colspan="2" align="center"><input type="button" class="button" value="새로등록" onclick="javascript:location.replace('/admin/etc/common/popOutMallEtcLink.asp?mallid=<%=mallid%>&itemid=<%=itemid%>&poomok=<%=poomok%>&qy=NEW');"></td>
</tr>
</table>
<% End If %>
<p>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frmAct" method="post" action="popOutMallEtcLink.asp">
<textarea name="valType_text" rows="0" cols="0" style="display:none"><%=textVal%></textarea> <!-- 실제 이노디터 에디터의 값이 저장되는 부분(에디터에 저장한 것이 textarea에 stlye:none으로 저장 -->
<input type= "hidden" name = "mode" id = "mode" value="<%=mode%>">
<input type= "hidden" name = "itemid" value="<%=itemid%>">
<tr bgcolor="#FFFFFF">
	<td bgcolor="<%= adminColor("tabletop") %>" width="100">상품코드</td>
	<td><%= itemid %></td>
</tr>
<tr bgcolor="#FFFFFF">
    <td bgcolor="<%= adminColor("tabletop") %>" width="100">적용몰</td>
    <td>
	<%  If (cnt > 0) Then %>
		<input type="hidden" name="mallid" value="<%=mallid%>"><%=chkiif(mallid = "", "전체", mallid)%>
	<%  Else Call drawSelectBoxXSiteAPIPartner("mallid", mallid)
		End If %>
   	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="<%= adminColor("tabletop") %>" width="100">적용구분</td>
	<td>
	<%  If (cnt > 0) Then %>
		<input type="hidden" name="linkgbn" value="<%=linkgbn%>"><%=linkgbnName%>
	<%  Else Call drawSelectBoxEtcLinkGbn("linkgbn",linkgbn,false)
		End If %>
	</td>
</tr>
<tr bgcolor="#FFFFFF">
    <td bgcolor="<%= adminColor("tabletop") %>" width="100">적용기간</td>
    <td>

		<input type="text" name="stDt" size="10" maxlength=10 readonly value="<%=Left(stDt,10)%>"> 00:00:00
		<a href="javascript:calendarOpen(frmAct.stDt);"><img src="/images/calicon.gif" border="0" align="absmiddle" height=21></a>
		&nbsp;~&nbsp;
		<input type="text" name="edDt" size="10" maxlength=10 readonly value="<%=Left(edDt,10)%>"> 23:59:59
		<a href="javascript:calendarOpen(frmAct.edDt);"><img src="/images/calicon.gif" border="0" align="absmiddle" height=21></a>
    	<input type="button" value="전기간" class="button" onClick="javascript:document.frmAct.stDt.value = '<%= Left(now(),10) %>';document.frmAct.edDt.value = '9999-12-31';">
    </td>
</tr>
<% If mode <> "V" Then %>
<tr bgcolor="#FFFFFF" id="lkyn">
	<td bgcolor="<%= adminColor("tabletop") %>" width="100">링크연동구분</td>
	<td>
	    <input type="radio" name="linkyn" value="Y" <% If linkyn = "Y" or linkyn = "" Then response.write "checked"  End If %> > 아래 적용값으로 연동
	</td>
</tr>
<tr bgcolor="#FFFFFF" id="ltype_1" style="display:none;">
    <td bgcolor="<%= adminColor("tabletop") %>" width="100">적용값(intVal)</td>
	<td><input type="text" name="valType" value="<%= theval %>" size="9" maxlength="9"></td>
</tr>
<tr bgcolor="#FFFFFF" id="ltype_3">
    <td bgcolor="<%= adminColor("tabletop") %>" width="100">적용값(textVal)</td>
	<td>
		<div id="EDITOR_AREA_CONTAINER"></div>
	</td>
</tr>
<tr bgcolor="#FFFFFF" id="ltype_0" style="display:none;">
    <td bgcolor="<%= adminColor("tabletop") %>" width="100">적용값(varcharVal)</td>
	<td>
		원재료명 : <input type="text" name="mat_Name" value="" maxlength="40">&nbsp;
		함량(%) : <input type="text" name="mat_percent" value="" maxlength="3" size="3" onkeypress="inputNumCom();" style="ime-mode:Disabled;">&nbsp;
		원료원산지 : <input type="text" name="mat_place" value="" >
	</td>
</tr>
<%
ElseIf mode = "V" Then
	 If valtype = "3" OR valtype = "1" Then
%>
<tr bgcolor="#FFFFFF" id="ltype_3">
    <td bgcolor="<%= adminColor("tabletop") %>" width="100">적용값(textVal)</td>
	<td>
		<div id="EDITOR_AREA_CONTAINER"></div>
	</td>
</tr>
<%
	ElseIf valtype = "2" Then
		Dim material, mat_Name, mat_percent, mat_place
		material	= Split(shortVal,"!!^^")
		mat_Name	= material(0)
		mat_percent	= material(1)
		mat_place	= material(2)
%>
<tr bgcolor="#FFFFFF" id="ltype_0">
    <td bgcolor="<%= adminColor("tabletop") %>" width="100">적용값(varcharVal)</td>
	<td>
		원재료명 : <input type="text" name="mat_Name" value="<%=mat_Name%>" maxlength="40">&nbsp;
		함량(%) : <input type="text" name="mat_percent" value="<%=mat_percent%>" maxlength="3" size="3" onkeypress="inputNumCom();" style="ime-mode:Disabled;">&nbsp;
		원료원산지 : <input type="text" name="mat_place" value="<%=mat_place%>" >
	</td>
</tr>
<%
	End If
End If
%>
<% If (cnt > 0) Then %>
<tr bgcolor="#FFFFFF">
    <td bgcolor="<%= adminColor("tabletop") %>" width="100">등록일</td>
	<td><%= regdate %></td>
</tr>
<% End If %>
<tr align="center" bgcolor="#FFFFFF" height="30">
	<td bgcolor="#FFFFFF" colspan="2">
		<input type="button" class="button" value="저장" onclick="javascript:frm_Act();">
	<% If (cnt>0) Then %>
		<input type="button" class="button" value="삭제" onclick="javascript:frm_Del();">
	<% End If %>
	</td>
</tr>
</form>
</table>
<!-- 글 수정시 textarea에 값 전달 시작 -->
<% If mode = "V" Then %>
<script>
	var strHTMLCode = document.frmAct["valType_text"].value;
	fnSetEditorHTMLCode(strHTMLCode, false, 0);
</script>
<% End If %>
<!-- 글 수정시 textarea에 값 전달 끝 -->
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbCTclose.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->