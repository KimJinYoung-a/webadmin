<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/admin/etc/kakaogift/kakaogiftcls.asp"-->
<%
Dim outmallOrderserial, arrRows, i
Dim okakaogift, research, isSongjang, xl
outmallOrderserial	= request("outmallOrderserial")
isSongjang			= request("isSongjang")
research			= request("research")
xl					= request("xl")
If (research = "") Then
	isSongjang = "Y"
End If

If outmallOrderserial <> "" then
	Dim iA2, arrTemp2, arroutmallOrderserial
	outmallOrderserial = replace(outmallOrderserial,",",chr(10))
	outmallOrderserial = replace(outmallOrderserial,chr(13),"")
	arrTemp2 = Split(outmallOrderserial,chr(10))
	iA2 = 0
	Do While iA2 <= ubound(arrTemp2)
		If Trim(arrTemp2(iA2))<>"" then
			arroutmallOrderserial = arroutmallOrderserial& "'"& trim(arrTemp2(iA2)) & "',"
		End If
		iA2 = iA2 + 1
	Loop
	outmallOrderserial = left(arroutmallOrderserial,len(arroutmallOrderserial)-1)
End If

If outmallOrderserial <> "" Then
	Set okakaogift = new Ckakaogift
		okakaogift.FRectoutmallorderserial = outmallOrderserial
		okakaogift.FRectIsSongjang = isSongjang
		arrRows = okakaogift.getkakaogiftSongjangList
	Set okakaogift = nothing
End If
%>
<script language="javascript">
<!--
	// 페이지 이동
	function goPage(pg) {
		frm.page.value = pg;
		frm.submit();
	}

	// 검색
	function serchItem() {
		frm.page.value = 1;
		frm.submit();
	}

	// kakaogift 카테고리 매칭 팝업
	function popkakaogiftCateMap(cdl,cdm,cds,dno) {
		var pCM = window.open("popkakaogiftCateMap.asp?cdl="+cdl+"&cdm="+cdm+"&cds="+cds+"&dspNo="+dno,"popCateMap","width=600,height=400,scrollbars=yes,resizable=yes");
		pCM.focus();
	}

	function popXL(){
		frmXL.submit();
	}	
//-->
</script>
<%
If (xl = "Y") Then
	Response.Buffer = True
	Response.ContentType = "application/vnd.ms-excel"
	Response.AddHeader "Content-Disposition", "attachment; filename=kakaogift"& FormatDate(now(), "00000000000000") &"_xl.xls"
	Response.CacheControl = "public"
Else
%>
<html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel" xmlns="http://www.w3.org/TR/REC-html40">
<head>
	<meta http-equiv="X-UA-Compatible" content="IE=edge,chrome=1" />
	<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
	<meta name="referrer" content="no-referrer-when-downgrade" />
	<script language="JavaScript" src="/js/xl.js"></script>
	<script language="JavaScript" src="/js/common.js"></script>
	<script language="JavaScript" src="/js/report.js"></script>
	<script language="JavaScript" src="/js/calendar.js"></script>
	<% If (xl <> "Y") Then %>
		<link rel="stylesheet" href="/css/scm.css" type="text/css">
	<% End If %>
</head>
<body>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="post" action="">
<input type="hidden" name="page" value="">
<input type="hidden" name="research" value="on">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="3" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
	<td align="left">
		KakaoGift 주문번호 : <textarea rows="6" cols="20" name="outmallOrderserial" id="outmallOrderserial"><%= replace(replace(outmallOrderserial,",",chr(10)), "'", "")%></textarea>
	</td>
	<td align="left">
		송장유무 : 
		<select name="isSongjang" class="select">
			<option value="">-전체-</option>
			<option value = "Y" <%= Chkiif(isSongjang = "Y", "selected", "") %>>Y</option>
			<option value = "N" <%= Chkiif(isSongjang = "N", "selected", "") %>>N</option>
		</select>
	</td>
	<td rowspan="3" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="검색" onClick="javascript:document.frm.submit();">
	</td>
</tr>
</form>
</table>
<br />
<%
	If IsArray(arrRows) Then 
		rw UBound(arrRows, 2) + 1 & " 건"
		response.write "<input type='button' class='button' value='엑셀받기' onClick='popXL()'>"
	End If
End If
%>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width="16.6%">주문번호</td>
	<td width="16.6%">배송방법</td>
	<td width="16.6%">택배사코드</td>
	<td width="16.6%">송장번호</td>
	<td width="16.6%">수령인명</td>
	<td width="16.6%">수령인연락처1</td>
</tr>
<%
	If (outmallOrderserial = "") OR NOT IsArray(arrRows) Then 
%>
<tr align="center" bgcolor="#FFFFFF" height="50">
	<td colspan="6">검색 결과가 없습니다.</td>
</tr>
<% 
	Else
		If IsArray(arrRows) Then 
%>

<%
			For i = 0 To Ubound(arrRows, 2)
%>
<tr align="center" bgcolor="#FFFFFF">
	<td width="16.6%" style="mso-number-format:\@"><%= Trim(arrRows(0, i)) %></td>
	<td width="16.6%" style="mso-number-format:\@">택배배송</td>
	<td width="16.6%" style="mso-number-format:\@"><%= Trim(arrRows(1, i)) %></td>
	<td width="16.6%" style="mso-number-format:\@"><%= Trim(arrRows(2, i)) %></td>
	<td width="16.6%" style="mso-number-format:\@"><%= Trim(arrRows(3, i)) %></td>
	<td width="16.6%" style="mso-number-format:\@">
	<%
		If Len(arrRows(5, i)) > 5 Then				'reqphone
			response.write Trim(arrRows(5, i))
		ElseIf Len(arrRows(4, i)) > 5 Then			'reqhp
			response.write Trim(arrRows(4, i))
		ElseIf Len(arrRows(6, i)) > 5 Then			'buyhp
			response.write Trim(arrRows(6, i))
		ElseIf Len(arrRows(7, i)) > 5 Then			'buyphone
			response.write Trim(arrRows(7, i))
		End If
	%>
	</td>
</tr>
<%
				If (i mod 1000) = 0 Then
					response.flush
				End If
			Next
		End If
	End If
%>
</table>
<% If (xl <> "Y") Then %>
<form name="frmXL" method="POST" style="margin:0px;">
	<input type="hidden" name="xl" value="Y">
	<input type="hidden" name="research" value="on">
	<textarea name="outmallOrderserial" style="display:none;"><%= replace(replace(outmallOrderserial,",",chr(10)), "'", "")%></textarea>
	<input type="hidden" name="isSongjang" value=<%= isSongjang %>>
</form>
<% End If %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->