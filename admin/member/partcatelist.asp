<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  파트관리자관리
' History : 2011.01.25 김진영 생성
'####################################################
%>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/admin/partpersonCls.asp"-->
<%
Dim clist, arlist, i, idx, arlist3, isusing, research
isusing		= request("isusing")
research	= request("research")
''기본조건 등록예정이상
If (research = "") Then
	isusing = "Y"
End If

	Set clist = new Partlist
		clist.FGubun = isusing
		arlist = clist.fnGetlist
	Set clist = nothing
%>
<script language="javascript">
function cmodify(k){
	var popwin = window.open('partcate_pop.asp?mode=modify&idx=' + k,'pop','width=500,height=200,scrollbars=yes,resizable=yes');
	popwin.focus();
}
function cinsert(){
	var popwin = window.open('partcate_pop.asp?mode=insert','pop','width=500,height=200,scrollbars=yes,resizable=yes');
	popwin.focus();
}
function lpop(k){
	var popwin = window.open('partcate2_pop.asp?idx='+k ,'pop','width=1200,height=600,scrollbars=yes,resizable=yes');
	popwin.focus();
}
function pop_10x10_person(){
	var popwin = window.open('/common/pop_10x10_person.asp','op2','width=700,height=600,scrollbars=yes,resizable=no');
	popwin.focus();
}
</script>
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
	<tr>
		<td align="left">
			<input type="button" class="button" value="신규카테고리 등록" onclick="cinsert();">
		</td>
		<td align="right"><a href="javascript:pop_10x10_person();">전체보기</a></td>
	</tr>
</table>
<br />
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get" action="">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="research" value="on">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="3" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
	<td align="left">
		상태 :
		<select name="isusing" class="select">
			<option value="">전체</option>
			<option value="Y" <%= Chkiif(isusing="Y", "selected", "") %>>사용중</option>
			<option value="N" <%= Chkiif(isusing="N", "selected", "") %>>비사용중</option>
		</select>
	</td>
	<td rowspan="3" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="검색" onClick="javascript:document.frm.submit();">
	</td>
</tr>
</form>
</table>
<br />
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
<form name="cform" method="post">

<tr align="center" bgcolor="#DDDDFF">
	<td align="center" width="10%">카테고리 번호</td>
	<td align="center" width="50%">상위 카테고리 이름</td>
	<td align="center" width="30%">하위 카테고리</td>
	<td align="center" width="10%">상태</td>
</tr>
<%
	For i = 0 to Ubound(arlist,2)
%>
<tr align="center" bgcolor="#FFFFFF" height="30">
	<td><%= i+1 %></td>
	<td><a href="javascript:cmodify('<%= arlist(0,i) %>')"><%= arlist(1,i) %></a></td>
	<td><a href="javascript:lpop('<%= arlist(0,i) %>')">목록보기</a></td>
	<td>
		<%If arlist(3,i) = "Y" Then response.write "<font color='blue'>사용중</font>" End If%>
		<%If arlist(3,i) = "N" Then response.write "<font color='red'>비사용중</font>" End If%>
	</td>
</tr>
<%
	Next
%>
</form>
</table>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->