<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/classes/board/companyrequestcls.asp" -->
<%

dim i, j

'==============================================================================
'업체상담게시판
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

if gubun="" then gubun="01"

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
고객센타 - 업체상담게시판<br><br>
<script>
function SubmitForm()
{
		if (confirm("처리상태를 완료로 전환합니까?") == true) { document.f.submit(); }
}
function catesubmit(){

	if (confirm("카테고리를 변경 합니다.") ==true)
		frm.mode.value="chcate";
		frm.categubun.value=f.categubun.value;
		frm.action="cscenter_req_board_act.asp";
		frm.submit();
}
function sellsubmit(){

	if (confirm("판매형식을 변경합니다.") ==true)
		frm.mode.value="chsell";
		frm.sellgubun.value=f.sellgubun.value;
		frm.action="cscenter_req_board_act.asp";
		frm.submit();
}
function ipjumYNsubmit(){

	if(confirm("입점여부 선택합니다.") ==true)
		frm.mode.value="ipjum";
		frm.ipjumYN.value=f.ipjumYN.value;
		frm.action="cscenter_req_board_act.asp";
		frm.submit();
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
	frm.action="cscenter_req_board_list.asp";
	frm.submit();
}
function editcomm(){
	frm.commmode.value="edit";
	frm.id.value="<%= companyrequest.results(0).id %>";
	frm.user.value="<%= session("ssBctCname") %>";
	frm.action="cscenter_req_board_view.asp";
	frm.submit();
}
function savecomm(){
	frm.mode.value="comm";
	frm.id.value="<%= companyrequest.results(0).id %>";
	frm.user.value="<%= session("ssBctCname") %>";
	frm.comment.value=commfrm.comment.value;
	frm.action="cscenter_req_board_act.asp";
	frm.submit();
	}

function changecontent() {}
</script>
<form method="post" name="f" action="cscenter_req_board_act.asp" onsubmit="return false">
<input type="hidden" name="mode" value="finish">
<input type="hidden" name="id" value="<%= companyrequest.results(0).id %>">

<b>구분</b> : <%= companyrequest.code2name(companyrequest.results(0).reqcd) %><br><br>

<b>작성일</b> : <%= FormatDate(companyrequest.results(0).regdate, "0000-00-00") %><br><br>

<b>회사명</b> : <%= db2html(companyrequest.results(0).companyname) %><br><br>

<b>담당자명</b> : <%= db2html(companyrequest.results(0).chargename) %>(<%= db2html(companyrequest.results(0).chargeposition) %>)<br><br>

<b>회사주소</b> : <%= db2html(companyrequest.results(0).address) %><br><br>

<b>연락처</b> : TEL <%= db2html(companyrequest.results(0).phone) %> / HP <%= companyrequest.results(0).hp %><br><br>

<b>이메일</b> : <a href="mailto:<%= db2html(companyrequest.results(0).email) %>"><%= db2html(companyrequest.results(0).email) %></a><br><br>

<b>회사URL</b> : <%
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
<br><br>

<b>내용/회사설명</b> : <%= nl2br(db2html(companyrequest.results(0).companycomments)) %><br><br>

<b>첨부화일</b> : <% if (companyrequest.results(0).attachfile <> "") then %><a href="http://imgstatic.10x10.co.kr<%= companyrequest.results(0).attachfile %>" target="_blank">다운받기</a><% end if %>
( 첨부화일 안될때 <a href="http://www.10x10.co.kr<%= replace(companyrequest.results(0).attachfile,"uploadimg","uploadimage") %>" target="_blank">다운받기</a> )
<br><br>

<b>처리상태</b> : <% if (IsNull(companyrequest.results(0).finishdate) = true) then %>미완료<% else %><%= FormatDate(companyrequest.results(0).finishdate, "0000-00-00") %><% end if %><br><br>

<hr>
<b>카테고리 분류</b> : <b><font color=blue><%= GetCategoryName(companyrequest.results(0).categubun) %></font></b>&nbsp;&nbsp;&nbsp;-->&nbsp;&nbsp;
<b>카테고리 변경</b> :
		<% call DrawSelectBoxCategoryLarge("categubun",catevalue) %>
		<input type=button value="변경" onclick="catesubmit();">
<br><br>
<b>판매 형태</b> :
<% if companyrequest.results(0).sellgubun="Y" then %>
<b><font color=blue>ON-Line</font></b>
<% elseif companyrequest.results(0).sellgubun="N" then%>
<b><font color=blue>OFF-Line</font></b>
<% else %>
<b><font color=blue>기타</font></b>
<% end if %>&nbsp;&nbsp;&nbsp;-->&nbsp;&nbsp;
<b>판매형태 변경</b> :
		<select name="sellgubun" class="a">
			<option value="Y">ON-Line</option>
			<option value="N">OFF-Line</option>
		</select>
		<input type=button value="변경" onclick="sellsubmit();">
<br><br>
<b>입점여부</b> :
		<%if companyrequest.results(i).ipjumYN="Y" then response.write "<b><font color=blue>입점완료</font></b>" %>
		<%if companyrequest.results(i).ipjumYN="N" then response.write "<b><font color=blue>미입점</font></b>" %>
		&nbsp;&nbsp;&nbsp;-->&nbsp;&nbsp;
<b>입점여부 변경</b> :
		<select name="ipjumYN" class="a">
			<option value="Y">입점 완료</option>
			<option value="N">미 입점</option>
		</select>
		<input type=button value="변경" onclick="ipjumYNsubmit();">
<br><br>
<b>세부사항</b> : <%= db2html(nl2br(companyrequest.results(0).reqcomment)) %><br><br>

<hr>
<input type="button" value=" 완료처리 " onclick="SubmitForm()">
<a href="javascript:MovePage(<%=page%>);">목록으로 이동</a>
</form>

<hr>

<!-- 코멘트 부분 -->
<table width="100%" cellspacing=0 cellpadding=0 border=0>
<tr>
	<td></td>
	<td><b>업체에게 메일보내기</b></td>
</tr>
	<form name="commfrm" method=post action="" onsubmit="return false">
	<% if commmode="" and companyrequest.results(0).replyuser <>"" then %>

	<tr>
		<td width="10%" valign="top">
			<%= db2html(companyrequest.results(0).replyuser) %>
		</td>
		<td width="75%" valign="top">
			<%= nl2br(db2html(companyrequest.results(0).replycomment)) %>
		</td>
		<td width="15%">
			<input type="button" value="수정" onclick="javascript:editcomm();">
		</td>
	</tr>
	<tr>
		<td height=40><input type="button" value="mail보내기" onclick="javascript:sendmail();">	</td>
	</tr>


	<% elseif commmode="edit" then %>

	<tr>
		<td width="10%" valign="top">
			<%= session("ssBctCname") %>
		</td>
		<td width="75%" valign="top">
			<textarea name="comment" rows=10 cols=100><%= db2html(companyrequest.results(0).replycomment) %></textarea>
		</td>
		<td width="15%">
			<input type="button" value="저장" onclick="javascript:savecomm();">
		</td>
	</tr>


	<% elseif companyrequest.results(0).replyuser ="" then %>


	<tr>
		<td width="10%" valign="top">
			<%= session("ssBctCname") %>
		</td>
		<td width="75%" valign="top">
			<textarea name="comment" rows=10 cols=100></textarea>
		</td>
		<td width="15%">
			<input type="button" value="저장" onclick="javascript:savecomm();">
		</td>
	</tr>
	<% end if %>
	</form>
</table>

<form name="frm" method="post" action="" onsubmit="return false">
	<input type="hidden" name="id" value="<%= companyrequest.results(0).id %>">
	<input type="hidden" name="pg" value="<%= page %>">
	<input type="hidden" name="mode" value="">
	<input type="hidden" name="categubun" value="">
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

<form name="frmmail" method="post" action="cscenter_req_board_mail.asp" onsubmit="return false">
	<input type="hidden" name="user" value="<%= session("ssBctCname") %>">
	<input type="hidden" name="mailname" value="<%= companyrequest.results(0).chargename %>">
	<input type="hidden" name="mailto" value="<%= companyrequest.results(0).email %>">
	<input type="hidden" name="content" value="<%= companyrequest.results(0).replycomment %>">
	<input type="hidden" name="id" value="<%= companyrequest.results(0).id %>">
</form>

<!-- #include virtual="/lib/db/dbclose.asp" -->