<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/etc/shoplinker/shoplinkercls.asp"-->
<%
Dim mallgubun, makerid, page, action, idx, isusing
Dim sqlStr, in_makerid, i
mallgubun	= request("mallgubun")
makerid		= request("makerid")
action		= request("action")
in_makerid	= request("in_makerid")
idx			= request("idx")
isusing		= request("isusing")
page		= request("page")
If page = "" Then page = 1
If action = "insert" Then
	sqlStr = ""
	sqlStr = sqlStr & " SELECT count(*) as cnt FROM db_user.dbo.tbl_user_c WHERE userid = '"&in_makerid&"' "
	rsget.Open sqlStr,dbget,1
	If rsget("cnt") = 0 Then
		response.write "<script>alert('브랜드ID가 잘 못 되었습니다');location.href='/admin/etc/shoplinker/JaehyuMall_Not_In_Makerid.asp?mallgubun="&mallgubun&"';</script>"
		response.end
	End If
	rsget.Close

	sqlStr = ""
	sqlStr = sqlStr & " SELECT count(*) as cnt FROM db_temp.dbo.tbl_shoplinker_not_in_makerid WHERE mallgubun = '"&mallgubun&"' AND makerid = '"&in_makerid&"' "
	rsget.Open sqlStr,dbget,1
	If rsget("cnt") <> 0 Then
		response.write "<script>alert('이미 등록된 브랜드 입니다');history.back(-1);</script>"
	Else
		sqlStr = ""
		sqlStr = sqlStr & " INSERT INTO db_temp.dbo.tbl_shoplinker_not_in_makerid (makerid, mallgubun, regdate, reguserid) VALUES "
		sqlStr = sqlStr & " ('"&in_makerid&"', '"&mallgubun&"', getdate(), '"&session("ssBctID")&"') "
		dbget.Execute sqlStr
		response.write "<script language='javascript'>alert('등록 되었습니다');location.href='/admin/etc/shoplinker/JaehyuMall_Not_In_Makerid.asp?mallgubun="&mallgubun&"';</script>"
	End If
	rsget.Close
ElseIf action = "update" Then
	sqlStr = ""
	sqlStr = sqlStr & " UPDATE db_temp.dbo.tbl_shoplinker_not_in_makerid SET "
	sqlStr = sqlStr & " isusing = '"&isusing&"', lastupdate =getdate(), lastuserid = '"&session("ssBctID")&"' "
	sqlStr = sqlStr & " WHERE idx = '"&idx&"' "
	dbget.Execute sqlStr
	response.write "<script language='javascript'>alert('수정 되었습니다');location.href='/admin/etc/shoplinker/JaehyuMall_Not_In_Makerid.asp?mallgubun="&mallgubun&"';</script>"
End If

Dim oshoplinker
SET oshoplinker = new CShoplinker
	oshoplinker.FPageSize 					= 20
	oshoplinker.FCurrPage					= page
	oshoplinker.FRectMakerid				= makerid
	oshoplinker.getNotInMakeridList
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<script language="JavaScript" src="/js/xl.js"></script>
<script language="JavaScript" src="/js/common.js"></script>
<script language="JavaScript" src="/js/report.js"></script>
<link rel="stylesheet" href="/css/scm.css" type="text/css">
</head>
<script language='javascript'>
function insert_makerid(){
	if(document.frm.in_makerid.value == "")
	{
		alert("브랜드ID를 입력하세요.");
		document.frm.in_makerid.focus();
		return;
	}
	document.frm.action.value = "insert";
	document.frm.submit();
}
function update_makerid(idx, ckyn){
	if(confirm(ckyn+"으로 변경하시겠습니까?")){
		document.frm.action.value = "update";
		document.frm.idx.value = idx;
		document.frm.isusing.value = ckyn;
		document.frm.submit();
	}
}
function goPage(pg){
    document.frm.page.value = pg;
    document.frm.submit();
}
</script>
<body onload="javascript:window.resizeTo(1200, 770);">
<div id="calendarPopup" style="position: absolute; visibility: hidden; z-index: 2;"></div>
<link rel="stylesheet" href="/css/scm.css" type="text/css">
<form name="frmsearch" method="get" action="<%=CurrURL()%>" style="margin:0px;">
<input type="hidden" name="mallgubun" value="<%=mallgubun%>">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr bgcolor="#FFFFFF">
	<td height="50">
		<table width="100%" class="a">
		<tr>
		    <td width="90%">Mall 구분 : <%= mallgubun %></td>
		    <td rowspan="4" width="10%"><input type="submit" value="검 색" style="width:50px;height:50px;"></td>
		</tr>
		<tr>
			<td >
			브랜드ID : <input type="text" class="text" name="makerid" value="<%=makerid%>" size="20"> <input type="button" class="button" value="ID검색" onclick="jsSearchBrandID(this.form.name,'makerid');" >
			</td>
		</tr>
		</table>
	</td>
</tr>
</table>
</form>
<br>
<form name="frm" action="<%=CurrURL()%>" methd="post" style="margin:0px;">
<input type="hidden" name="action" value="">
<input type="hidden" name="idx" value="">
<input type="hidden" name="isusing" value="">
<input type="hidden" name="mallgubun" value="<%=mallgubun%>">
<input type="hidden" name="page" value="<%=Page%>">
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a">
<tr>
	<td>
		<table width="100%" class="a">
		<tr>
			<td width="80%">
				예외 브랜드 : 
				<input type="text" name="in_makerid" value="" size="20" AUTOCOMPLETE="off" onKeyPress="if (event.keyCode == 13){ insert_makerid(); return false;}">
				<input type="button" class="button" value="저 장" onClick="insert_makerid()">
			</td>
			<td width="20%" align="right">총 : <b><%=oshoplinker.FTotalCount%></b></td>
		</tr>
		</table>
	</td>
</tr>
</table>
</form>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="#DDDDDD">
    <td>몰구분</td>
    <td>브랜드ID</td>
    <td>등록일</td>
    <td>등록자</td>
    <td>최종수정일</td>
    <td>최종수정자</td>
    <td>수정</td>
</tr>
<%
If oshoplinker.FResultCount > 0 Then
	For i=0 to oshoplinker.FResultCount - 1
%>
<tr align="center" bgcolor="#FFFFFF" height="20">
    <td><%= oshoplinker.FItemList(i).FMallgubun %></td>
    <td><%= oshoplinker.FItemList(i).FMakerid %></td>
    <td><%= oshoplinker.FItemList(i).FRegdate %></td>
    <td><%= oshoplinker.FItemList(i).FReguserid %></td>
    <td><%= oshoplinker.FItemList(i).FLastupdate %></td>
    <td><%= oshoplinker.FItemList(i).FLastuserid %></td>
    <td>
    	<input type="radio" value="Y" <%= CHkiif(oshoplinker.FItemList(i).FIsusing ="Y","checked","") %> onclick="update_makerid('<%= oshoplinker.FItemList(i).FIdx %>','Y')">Y
    	<input type="radio" value="N" <%= CHkiif(oshoplinker.FItemList(i).FIsusing ="N","checked","") %> onclick="update_makerid('<%= oshoplinker.FItemList(i).FIdx %>','N')">N
    </td>
</tr>
<%
	Next
%>
<tr height="20">
    <td colspan="7" align="center" bgcolor="#FFFFFF">
        <% If oshoplinker.HasPreScroll then %>
		<a href="javascript:goPage('<%= oshoplinker.StartScrollPage-1 %>');">[pre]</a>
    	<% Else %>
    		[pre]
    	<% End If %>

    	<% For i = 0 + oshoplinker.StartScrollPage to oshoplinker.FScrollCount + oshoplinker.StartScrollPage - 1 %>
    		<% If i>oshoplinker.FTotalpage Then Exit For %>
    		<% If CStr(page) = CStr(i) Then %>
    		<font color="red">[<%= i %>]</font>
    		<% Else %>
    		<a href="javascript:goPage('<%= i %>');">[<%= i %>]</a>
    		<% End If %>
    	<% Next %>

    	<% If oshoplinker.HasNextScroll Then %>
    		<a href="javascript:goPage('<%= i %>');">[next]</a>
    	<% Else %>
    		[next]
    	<% End If %>
    </td>
</tr>
<%
Else
%>
<tr bgcolor="#FFFFFF" height="50" align="center">
    <td colspan="7">등록된 브랜드가 없습니다.</td>
</tr>
<%
End If
%>
</table>
<% Set oshoplinker = nothing%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
