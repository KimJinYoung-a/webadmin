<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/classes/board/companyrequestcls.asp" -->
<%
'###########################################################
' Description : 사업제휴/광고 문의
' History : 2013.07.25 허진원 생성
'###########################################################
%>
<%
dim i, j
dim commmode
	commmode=request("commmode")
dim page,gubun, onlymifinish
dim research, searchkey,catevalue
dim ipjumYN
	page = request("pg")
	gubun = request("gubun")
	onlymifinish = request("onlymifinish")
	research = request("research")
	searchkey = request("searchkey")
	catevalue=request("catevalue")
	ipjumYN=request("ipjumYN")
	if research="" and onlymifinish="" then onlymifinish="on"

	'// 기본값으로 사업제휴
	if gubun="" then gubun="02"
	if (page = "") then page = "1"

dim companyrequest
	set companyrequest = New CCompanyRequest
	companyrequest.read(request("id"))

%>
<STYLE TYPE="text/css">
<!--
	A:link, A:visited, A:active { text-decoration: none; }
	A:hover { text-decoration:underline; }
	BODY, TD, UL, OL, PRE { font-size: 10pt; }
	INPUT,SELECT,TEXTAREA { border:1 solid #666666; background-color: #ffffff; color: #000000; }
-->
</STYLE>
게시판관리 - 사업제휴 상세 보기<br><br>
<script type="text/javascript">
function SubmitForm() {
	if (confirm("처리상태를 완료로 전환합니까?") == true) { document.f.submit(); }
}

function sendmail(){
	if(confirm("메일을 보내시겠습니까?.") ==true)
	frmmail.submit();
}
function MovePage(page){
	frm.pg.value=page;
	frm.research.value="<%=research %>";
	frm.gubun.value="<%=gubun%>";
	frm.onlymifinish.value="<%=onlymifinish%>";
	frm.catevalue.value="<%=catevalue%>";
	frm.ipjumYNvalue="<%=ipjumYN%>";
	frm.searchkey.value="<%=searchkey%>";
	frm.action="/admin/board/upche/req_list.asp";
	frm.submit();
}
function editcomm(){
	frm.commmode.value="edit";
	frm.id.value="<%= companyrequest.results(0).id %>";
	frm.user.value="<%= session("ssBctCname") %>";
	frm.action="/admin/board/upche/req_view2.asp";
	frm.submit();
}
function savecomm(){
	frm.mode.value="comm";
	frm.id.value="<%= companyrequest.results(0).id %>";
	frm.user.value="<%= session("ssBctCname") %>";
	frm.comment.value=commfrm.comment.value;
	frm.action="/admin/board/upche/req_act.asp";
	frm.submit();
	}

function upcheworkerlist(id)
{
	var upWorker = null;
	upWorker = window.open('/admin/board/upche/upchePopWorkerList.asp?id='+id+'&team=14','openWorker','width=570,height=570,scrollbars=yes');
	upWorker.focus();
}
function upcheworkerDel(id)
{
	frm.mode.value="delworkid";
	frm.action="/admin/board/upche/req_act.asp";
	frm.submit();
}
</script>

<!-- 업체정보 시작 -->
<form method="post" name="f" action="/admin/board/upche/req_act.asp" onsubmit="return false">
<input type="hidden" name="mode" value="finish">
<input type="hidden" name="gubun" value="<%=gubun%>">
<input type="hidden" name="id" value="<%= companyrequest.results(0).id %>">
<table width="100%" align="center" cellpadding="3" cellspacing="1" bgcolor="black">
<tr bgcolor="FFFFFF" align="center">
	<td colspan="4"><b><font size=3 color="blue"><%= companyrequest.results(0).getAllianceGubun %> 문의</font></b></td>
</tr>
<tr bgcolor="FFFFFF" align="center">
	<td bgcolor="<%= adminColor("gray") %>">텐바이텐 담당자</td>
	<td colspan="3" align="left">
		<% sbGetwork "workid",companyrequest.results(0).Fworkid, "" %>
	</td>
</tr>
<tr bgcolor="FFFFFF" align="center">
	<td width="100" bgcolor="<%= adminColor("gray") %>">회사명</td>
	<td colspan="3" align="left"><%= db2html(companyrequest.results(0).companyname) %></td>
</tr>
<tr bgcolor="FFFFFF" align="center">
	<td bgcolor="<%= adminColor("gray") %>">회사주소</td>
	<td colspan="3" align="left"><%= db2html(companyrequest.results(0).address) %></td>
</tr>
<tr bgcolor="FFFFFF">
	<td bgcolor="<%= adminColor("gray") %>" align="center">사업자등록번호</td>	
	<td colspan="3" align="left"><%= db2html(companyrequest.results(0).license_no) %></td>
</tr>
<tr bgcolor="FFFFFF" align="center">
	<td width="100" bgcolor="<%= adminColor("gray") %>">담당자명</td>
	<td align="left"><%= db2html(companyrequest.results(0).chargename) %></td>
	<td width="100" bgcolor="<%= adminColor("gray") %>">부서명</td>
	<td align="left"><%= db2html(companyrequest.results(0).chargeposition) %></td>
</tr>
<tr bgcolor="FFFFFF" align="center">
	<td bgcolor="<%= adminColor("gray") %>">전화번호#1</td>
	<td align="left"><%= db2html(companyrequest.results(0).phone) %></td>
	<td bgcolor="<%= adminColor("gray") %>">전화번호#2</td>
	<td align="left"><%= db2html(companyrequest.results(0).hp) %></td>
</tr>
	<tr bgcolor="FFFFFF" align="center">
	<td bgcolor="<%= adminColor("gray") %>">이메일</td>
	<td colspan="3" align="left"><a href="mailto:<%= db2html(companyrequest.results(0).email) %>"><%= db2html(companyrequest.results(0).email) %></a></td>
</tr>
<tr bgcolor="FFFFFF" align="center">
	<td bgcolor="<%= adminColor("gray") %>">회사URL</td>
	<td colspan="3" align="left">
		<%
			dim arrUrl
			arrUrl = split(companyrequest.results(0).companyurl,",")
			if ubound(arrUrl)>0 then
				Response.Write "<a href='"
				if Left(arrUrl(0),7) <> "http://" then Response.Write "http://"
				Response.Write arrUrl(0) & "' target='_blank'>" & arrUrl(0) & "</a>"
				Response.Write "<br><br><b>입점쇼핑몰</b> : " & arrUrl(1)
			else
				Response.Write "<a href='"
				if Left(companyrequest.results(0).companyurl,7) <> "http://" then Response.Write "http://"
				Response.Write companyrequest.results(0).companyurl & "' target='_blank'>" & companyrequest.results(0).companyurl & "</a>"
			end if
		%>
	</td>
</tr>
<tr bgcolor="FFFFFF" align="center">
	<td bgcolor="<%= adminColor("gray") %>">회사소개</td>
	<td colspan="3" align="left">
		<%= nl2br(db2html(companyrequest.results(0).companycomments)) %>
	</td>
</tr>
<tr bgcolor="FFFFFF" align="center">
	<td bgcolor="<%= adminColor("gray") %>">문의내용</td>
	<td colspan="3" align="left">
		<%= nl2br(db2html(companyrequest.results(0).reqcomment)) %>
	</td>
</tr>
<tr bgcolor="FFFFFF" align="center">
	<td bgcolor="<%= adminColor("gray") %>">첨부파일</td>
	<td align="left">
		<% if (companyrequest.results(0).attachfile <> "") then %>
			<% if (Left(companyrequest.results(0).attachfile,4) = "http") then %>
				<a href="<%= companyrequest.results(0).attachfile %>" target="_blank">다운받기</a>
			<% else %>
				<a href="http://imgstatic.10x10.co.kr<%= companyrequest.results(0).attachfile %>" target="_blank">다운받기</a>
			<% end if %>
		<% else %>
			없음
		<% end if %>
	</td>
	<td bgcolor="<%= adminColor("gray") %>">처리상태</td>
	<td align="left">
		<% if (IsNull(companyrequest.results(0).finishdate) = true) then %>
			미완료
			&nbsp;
			<input type="button" value=" 완료처리 " onclick="SubmitForm()">
		<% else %>
			<%= FormatDate(companyrequest.results(0).finishdate, "0000-00-00") %>
		<% end if %>
	</td>
</tr>
</table>
</form>
<div style="text-align:right;padding-bottom:16px;"><a href="javascript:MovePage(<%=page%>);">목록으로</a></div>

<!-- 코멘트 부분 -->
<form name="commfrm" method=post action="" onsubmit="return false">
<table width="100%" align="center" cellpadding="0" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr bgcolor="FFFFFF" align="center">
	<td colspan=3><b><font size=3 color="blue">업체에게 메일보내기</font></b></td>
</tr>

<% if commmode="" and companyrequest.results(0).replyuser <>"" then %>
<tr bgcolor="FFFFFF" align="center">
	<td width="10%" valign="top">
		작성: <%= db2html(companyrequest.results(0).replyuser) %>
	</td>
	<td width="75%" valign="top">
		<%= nl2br(db2html(companyrequest.results(0).replycomment)) %>
	</td>
	<td width="15%">
		<input type="button" value="수정" onclick="javascript:editcomm();">
	</td>
</tr>
<tr bgcolor="FFFFFF" align="left">
	<td colspan=3><input type="button" value="mail보내기" onclick="javascript:sendmail();">	</td>
</tr>

<%
	'//수정모드
	elseif commmode="edit" then %>
<tr bgcolor="FFFFFF" align="center">
	<td width="10%" valign="top">
		작성: <%= session("ssBctCname") %>
	</td>
	<td valign="top">
		<textarea name="comment" rows=10 cols=95><%= db2html(companyrequest.results(0).replycomment) %></textarea>
	</td>
	<td>
		<input type="button" value="저장" onclick="javascript:savecomm();">
	</td>
</tr>

<%
	'//작성모드
	elseif companyrequest.results(0).replyuser ="" then %>
<tr bgcolor="FFFFFF" align="center">
	<td valign="top">
		작성: <%= session("ssBctCname") %>
	</td>
	<td valign="top">
		<textarea name="comment" rows=10 cols=95></textarea>
	</td>
	<td>
		<input type="button" value="저장" onclick="javascript:savecomm();">
	</td>
</tr>
<% end if %>
</table>
</form>

<form name="frm" method="post" action="" onsubmit="return false">
	<input type="hidden" name="id" value="<%= companyrequest.results(0).id %>">
	<input type="hidden" name="pg" value="<%= page %>">
	<input type="hidden" name="mode" value="">
	<input type="hidden" name="cd1" value="">
	<input type="hidden" name="cd2" value="">
	<input type="hidden" name="cd3" value="">
	<input type="hidden" name="sellgubun" value="">
	<input type="hidden" name="ipjumYN" value="">
	<input type="hidden" name="research" value="on">
	<input type="hidden" name="gubun" value="<%= gubun%>" >
	<input type="hidden" name="onlymifinish" value="<%=onlymifinish%>">
	<input type="hidden" name="catevalue" value="<%=catevalue%>">
	<input type="hidden" name="searchkey" value="<%=searchkey%>">
	<input type="hidden" name="commmode" value="">
	<input type="hidden" name="user" value="">
	<input type="hidden" name="comment" value="">
</form>

<form name="frmmail" method="post" action="/admin/board/upche/req_mail.asp" onsubmit="return false">
	<input type="hidden" name="user" value="<%= session("ssBctCname") %>">
	<input type="hidden" name="mailname" value="<%= companyrequest.results(0).chargename %>">
	<input type="hidden" name="mailto" value="<%= companyrequest.results(0).email %>">
	<input type="hidden" name="content" value="<%= companyrequest.results(0).replycomment %>">
	<input type="hidden" name="id" value="<%= companyrequest.results(0).id %>">
</form>
<!-- #include virtual="/lib/db/dbclose.asp" -->
