<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 감성모모 공지사항
' Hieditor : 2009.11.27 한용민 생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/momo/momo_cls.asp"-->
<%
'// 변수 선언 //
dim ntcId , oNotice, i, lp , commcd
dim page, searchDiv, searchKey, searchString, param
	'// 파라메터 접수 //
	ntcId = request("ntcId")
	page = request("page")
	searchDiv = request("searchDiv")
	searchKey = request("searchKey")
	searchString = request("searchString")	

	if page="" then page=1
	if searchKey="" then searchKey="titleLong"

	param = "&page=" & page & "&searchDiv=" & searchDiv & "&searchKey=" & searchKey & "&searchString=" & searchString '페이지 변수

'// 내용 접수
set oNotice = new CNotice
	oNotice.FRectntcId = ntcId
	oNotice.GetNoitceRead()
%>

<script language="javascript">

	// 글삭제
	function GotoNoticeDel(){
		if (confirm('삭제 하시겠습니까?')){
			document.frm_trans.submit();
		}
	}
	
</script>

<table width="100%" border="0" align="center" class="a" cellpadding="0" cellspacing="1" bgcolor="#BABABA">
<tr align="center" bgcolor="#FFFFFF">
	<td colspan="2">
		<table width="100%" border="0" class="a" cellpadding="0" cellspacing="0">
		<tr>
			<td height="26" align="left"><b>공지사항 상세 정보</b></td>
			<td height="26" align="right"><%=oNotice.FNoticeList(0).Fregdate%>&nbsp;</td>
		</tr>
		</table>
	</td>
</tr>
<tr>
	<td align="center"  bgcolor="#FFFFFF">구분</td>
	<td bgcolor="#FFFFFF"><%= getnotics_gubun(oNotice.FNoticeList(0).fcommcd)%></td>
</tr>
<tr>
	<td align="center"  bgcolor="#FFFFFF">작성자</td>
	<td bgcolor="#FFFFFF"><%=oNotice.FNoticeList(0).Fusername & "(" & oNotice.FNoticeList(0).Fuserid & ")" %></td>
</tr>
<tr>
	<td align="center"  bgcolor="#FFFFFF">사용여부</td>
	<td bgcolor="#FFFFFF"><%=oNotice.FNoticeList(0).fisusing%></td>
</tr>
<tr>
	<td align="center"  bgcolor="#FFFFFF">제목</td>
	<td bgcolor="#F8F8FF"><%=db2html(oNotice.FNoticeList(0).Ftitle)%></td>
</tr>
<tr>
	<td align="center"  bgcolor="#FFFFFF">내용</td>
	<td bgcolor="#FFFFFF"><%=nl2br(db2html(oNotice.FNoticeList(0).Fcontents))%></td>
</tr>
<tr>
	<td colspan="2" height="32" bgcolor="#FAFAFA" align="center">
		<img src="/images/icon_modify.jpg" onClick="self.location='notice_modi.asp?menupos=<%=menupos%>&ntcId=<%=ntcId & param%>'" style="cursor:pointer" align="absmiddle"> &nbsp;
		<img src="/images/icon_delete.gif" onClick="GotoNoticeDel()" style="cursor:pointer" align="absmiddle"> &nbsp;
		<img src="/images/icon_list.gif" onClick="self.location='notice_list.asp?menupos=<%=menupos & param %>'" style="cursor:pointer" align="absmiddle">
	</td>
</tr>
<form name="frm_trans" method="POST" action="doNotice.asp">
<input type="hidden" name="ntcId" value="<%=ntcId%>">
<input type="hidden" name="mode" value="delete">
</form>
</table>

<%
	set oNotice = nothing
%>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
