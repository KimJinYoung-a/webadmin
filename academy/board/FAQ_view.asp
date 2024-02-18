<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/academy/lib/classes/faq_cls.asp"-->
<% 
	'// 변수 선언 //
	dim faqid
	dim page, searchDiv, searchKey, searchString, param

	dim ofaq, i, lp

	'// 파라메터 접수 //
	faqid = RequestCheckvar(request("faqid"),10)
	page = RequestCheckvar(request("page"),10)
	searchDiv = RequestCheckvar(request("searchDiv"),16)
	searchKey = RequestCheckvar(request("searchKey"),16)
	searchString = RequestCheckvar(request("searchString"),128)

	if page="" then page=1
	if searchKey="" then searchKey="titleLong"

	param = "&page=" & page & "&searchDiv=" & searchDiv & "&searchKey=" & searchKey & "&searchString=" & searchString	'페이지 변수

	'// 내용 접수
	set ofaq = new Cfaq
	ofaq.FRectfaqid = faqid

	ofaq.GetFAQRead

%>
<script language="javascript">
<!--
	// 글삭제
	function GotofaqDel(){
		if (confirm('삭제 하시겠습니까?')){
			document.frm_trans.submit();
		}
	}
//-->
</script>
<!-- 보기 화면 시작 -->
<table width="100%" border="0" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
<tr align="center" bgcolor="#F0F0FD">
	<td colspan="2">
		<table width="100%" border="0" class="a" cellpadding="0" cellspacing="0">
		<tr>
			<td height="26" align="left"><b>FAQ 상세 정보</b></td>
			<td height="26" align="right"><%=ofaq.FfaqList(0).Fregdate%>&nbsp;</td>
		</tr>
		</table>
	</td>
</tr>
<tr>
	<td align="center" width="120" bgcolor="#DDDDFF">작성자</td>
	<td bgcolor="#FFFFFF"><%=ofaq.FfaqList(0).Fusername & "(" & ofaq.FfaqList(0).Fuserid & ")" %></td>
</tr>
<tr>
	<td align="center" width="120" bgcolor="#DDDDFF">구분</td>
	<td bgcolor="#FFFFFF"><%=ofaq.FfaqList(0).FcommNm%></td>
</tr>
<tr>
	<td align="center" width="120" bgcolor="#DDDDFF">제목</td>
	<td bgcolor="#F8F8FF"><%=db2html(ofaq.FfaqList(0).Ftitle)%></td>
</tr>
<tr>
	<td align="center" width="120" bgcolor="#DDDDFF">내용</td>
	<td bgcolor="#FFFFFF"><%=nl2br(db2html(ofaq.FfaqList(0).Fcontents))%></td>
</tr>
<tr><td height="1" colspan="2" bgcolor="#D0D0D0"></td></tr>
<tr>
	<td colspan="2" height="32" bgcolor="#FAFAFA" align="center">
		<img src="/images/icon_modify.jpg" onClick="self.location='faq_modi.asp?menupos=<%=menupos%>&faqid=<%=faqid & param%>'" style="cursor:pointer" align="absmiddle"> &nbsp;
		<img src="/images/icon_delete.gif" onClick="GotofaqDel()" style="cursor:pointer" align="absmiddle"> &nbsp;
		<img src="/images/icon_list.gif" onClick="self.location='faq_list.asp?menupos=<%=menupos & param %>'" style="cursor:pointer" align="absmiddle">
	</td>
</tr>
<form name="frm_trans" method="POST" action="dofaq.asp">
<input type="hidden" name="faqid" value="<%=faqid%>">
<input type="hidden" name="mode" value="delete">
<input type="hidden" name="menupos" value="<%=menupos%>">
<input type="hidden" name="page" value="<%=page%>">
<input type="hidden" name="searchDiv" value="<%=searchDiv%>">
<input type="hidden" name="searchKey" value="<%=searchKey%>">
<input type="hidden" name="searchString" value="<%=searchString%>">
</form>
</table>
<!-- 쓰기 화면 끝 -->
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->
