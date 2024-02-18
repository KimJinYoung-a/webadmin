<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : [업체]공지사항
' Hieditor : 서동석 생성
'			 2023.10.23 한용민 수정(이메일발송 cdo->메일러로 변경)
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/board/boardcls.asp"-->
<!-- #include virtual="/lib/classes/partners/partnerusercls.asp"-->
<%
dim i, brdid, mduserid, catecode, targt, mode, page, nboardmail
	brdid		= requestcheckvar(getNumeric(request("id")),10)
	mduserid    = requestcheckvar(request("mduserid"),32)
	catecode    = request("catecode")
	targt		= request("targt")
	mode		= requestcheckvar(request("mode"),32)
	page		= requestcheckvar(getNumeric(request("page")),10)

if page="" then page=1
if targt="" then targt="basic"

%>
<script type='text/javascript'>

function SendEmail(frm){
	if (confirm('업체 전체메일을 발송 하시겠습니까?')){
		frm.action="dodesignernoticemail.asp";
		frm.method="POST";
		frm.submit();
	}
}

function previewTarget(frm) {
	frm.action="";
	frm.method="GET";
	frm.submit();
}

function goPage(pg) {
	frm.page.value=pg;
	frm.action="";
	frm.method="GET";
	frm.submit();
}

</script>
<span style="font-size:13px; font-weight:bold;">⊙ 업체 공지메일 발송</span>

<form name="frm" method="post" action="" style="margin:0px;">
<input type="hidden" name="id" value="<%=brdid%>">
<input type="hidden" name="mode" value="upcheall">
<input type="hidden" name="page" value="1">
<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr bgcolor="FFFFFF">
	<td>
		* 담당자 : <% drawSelectBoxCoWorker "mduserid", mduserid %>
		&nbsp;
		* 카테고리 : <% SelectBoxBrandCategory "catecode", catecode %>
		&nbsp;
		<br>
		* 발송대상 :
			<label><input type="radio" name="targt" value="basic" <% if targt="basic" then Response.Write "checked" %> onfocus="this.blur()">기본담당자</label> &nbsp;
			<label><input type="radio" name="targt" value="deliver" <% if targt="deliver" then Response.Write "checked" %> onfocus="this.blur()">배송담당자</label> &nbsp;
			<label><input type="radio" name="targt" value="account" <% if targt="account" then Response.Write "checked" %> onfocus="this.blur()">정산담당자</label>
	</td>
	<td align="center" width=60>
		<input type="button" value="메일발송" onClick="SendEmail(frm);" class="button">
	</td>
</tr>

</table>
</form>

<br>
<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a">
<tr>
	<td align="left">
		<input type="button" value="대상자 보기" onClick="previewTarget(frm)" class="button">
	</td>
	<td align="right"></td>
</tr>
</table>
<!-- 액션 끝 -->

<%
'// 대상자 보기일경우 목록 출력
if mode="upcheall" then
	set nboardmail = new CBoard
		nboardmail.FCurrPage		= page
		nboardmail.FPageSize		= 15
		nboardmail.FRectMDid		= mduserid
		nboardmail.FRectCatCD		= catecode
		nboardmail.FRectTarget		= targt
		nboardmail.design_notice_mail_preview
%>
	<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
	<tr bgcolor="FFFFFF">
		<td colspan="3">
			검색결과 : <b><%= nboardmail.FTotalCount %></b>
			&nbsp;
			페이지 : <b><%= page %>/ <%= nboardmail.FTotalPage %></b>
			&nbsp;
			대상메일(중복제외) : <b><%= nboardmail.Fint_total %></b>
		</td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td><b>브랜드ID</b></td>
		<td><b>브랜드명</b></td>
		<td><b>이메일</b></td>
	</tr>

	<% if nboardmail.FResultCount>0 then %>
		<% for i=0 to nboardmail.FResultCount-1 %>	
		<tr align='center' bgcolor='#FFFFFF'>
			<td><%= nboardmail.BoardItem(i).FRectDesignerID %></td>
			<td><%= nboardmail.BoardItem(i).FRectName %></td>
			<td><%= nboardmail.BoardItem(i).FRectEmail %></td>
		</tr>	
		<% next %>

		<tr height="25" bgcolor="FFFFFF">
			<td colspan="3" align="center">
				<% if nboardmail.HasPreScroll then %>
					<span class="list_link"><a href="javascript:goPage(<%= nboardmail.StartScrollPage-1 %>)">[pre]</a></span>
				<% else %>
				[pre]
				<% end if %>
				<% for i = 0 + nboardmail.StartScrollPage to nboardmail.StartScrollPage + nboardmail.FScrollCount - 1 %>
					<% if (i > nboardmail.FTotalpage) then Exit for %>
					<% if CStr(i) = CStr(nboardmail.FCurrPage) then %>
					<span class="page_link"><font color="red"><b><%= i %></b></font></span>
					<% else %>
					<a href="javascript:goPage(<%= i %>)" class="list_link"><font color="#000000"><%= i %></font></a>
					<% end if %>
				<% next %>
				<% if nboardmail.HasNextScroll then %>
					<span class="list_link"><a href="javascript:goPage(<%= i %>)">[next]</a></span>
				<% else %>
				[next]
				<% end if %>
			</td>
		</tr>
	<% else %>
		<tr bgcolor="#FFFFFF">
			<td colspan="3" align="center" class="page_link">[검색결과가 없습니다.]</td>
		</tr>
	<% end if %>
	</table>
<%
	set nboardmail = Nothing
end if
%>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->