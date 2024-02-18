<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/admin/etc/common/commonCls.asp"-->
<%
Dim outmallOrderserial, arrRows, i
Dim oGS25, research, isSongjang, xl
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
	Set oGS25 = new CCommon
		oGS25.FRectoutmallorderserial = outmallOrderserial
		oGS25.FRectIsSongjang = isSongjang
		arrRows = oGS25.getgs25SongjangList
	Set oGS25 = nothing
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

	function popXL(){
		frmXL.submit();
	}	
//-->
</script>
<%
If (xl = "Y") Then
	Response.Buffer = True
	Response.ContentType = "application/vnd.ms-excel"
	Response.AddHeader "Content-Disposition", "attachment; filename=gs25"& FormatDate(now(), "00000000000000") &"_xl.xls"
Else
%>
<!doctype html>
<html lang="ko">
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
		GS25 주문번호 : <textarea rows="6" cols="20" name="outmallOrderserial" id="outmallOrderserial"><%= replace(replace(outmallOrderserial,",",chr(10)), "'", "")%></textarea>
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
	<td width="12.5%">주문번호(필수)</td>
	<td width="12.5%">상품일련번호(필수)</td>
	<td width="12.5%">상품코드(필수)</td>
	<td width="16.5%">상품명</td>
	<td width="8.5%">배송요청일</td>
	<td width="12.5%">택배사코드(필수)</td>
	<td width="12.5%">택배사명(ETC일경우에만필수)</td>
	<td width="12.5%">송장번호(필수)</td>
</tr>
<%
	If (outmallOrderserial = "") OR NOT IsArray(arrRows) Then 
%>
<tr align="center" bgcolor="#FFFFFF" height="50">
	<td colspan="8">검색 결과가 없습니다.</td>
</tr>
<% 
	Else
		If IsArray(arrRows) Then 
%>

<%
			For i = 0 To Ubound(arrRows, 2)
%>
<tr align="center" bgcolor="#FFFFFF">
	<td width="12.5%" style="mso-number-format:\@"><%= Trim(arrRows(0, i)) %></td>
	<td width="12.5%" style="mso-number-format:\@"><%= Trim(arrRows(1, i)) %></td>
	<td width="12.5%" style="mso-number-format:\@"><%= Trim(arrRows(2, i)) %></td>
	<td width="16.5%" style="mso-number-format:\@"><%= Trim(arrRows(3, i)) %></td>
	<td width="8.5%" style="mso-number-format:\@"></td>
	<td width="12.5%" style="mso-number-format:\@"><%= Trim(arrRows(4, i)) %></td>
	<td width="12.5%" style="mso-number-format:\@"></td>
	<td width="12.5%" style="mso-number-format:\@"><%= Trim(arrRows(5, i)) %></td>
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