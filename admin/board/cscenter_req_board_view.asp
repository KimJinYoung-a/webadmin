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
'��ü���Խ���
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
����Ÿ - ��ü���Խ���<br><br>
<script>
function SubmitForm()
{
		if (confirm("ó�����¸� �Ϸ�� ��ȯ�մϱ�?") == true) { document.f.submit(); }
}
function catesubmit(){

	if (confirm("ī�װ��� ���� �մϴ�.") ==true)
		frm.mode.value="chcate";
		frm.categubun.value=f.categubun.value;
		frm.action="cscenter_req_board_act.asp";
		frm.submit();
}
function sellsubmit(){

	if (confirm("�Ǹ������� �����մϴ�.") ==true)
		frm.mode.value="chsell";
		frm.sellgubun.value=f.sellgubun.value;
		frm.action="cscenter_req_board_act.asp";
		frm.submit();
}
function ipjumYNsubmit(){

	if(confirm("�������� �����մϴ�.") ==true)
		frm.mode.value="ipjum";
		frm.ipjumYN.value=f.ipjumYN.value;
		frm.action="cscenter_req_board_act.asp";
		frm.submit();
}
function sendmail(){
	if(confirm("������ �����ðڽ��ϱ�?.") ==true)
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

<b>����</b> : <%= companyrequest.code2name(companyrequest.results(0).reqcd) %><br><br>

<b>�ۼ���</b> : <%= FormatDate(companyrequest.results(0).regdate, "0000-00-00") %><br><br>

<b>ȸ���</b> : <%= db2html(companyrequest.results(0).companyname) %><br><br>

<b>����ڸ�</b> : <%= db2html(companyrequest.results(0).chargename) %>(<%= db2html(companyrequest.results(0).chargeposition) %>)<br><br>

<b>ȸ���ּ�</b> : <%= db2html(companyrequest.results(0).address) %><br><br>

<b>����ó</b> : TEL <%= db2html(companyrequest.results(0).phone) %> / HP <%= companyrequest.results(0).hp %><br><br>

<b>�̸���</b> : <a href="mailto:<%= db2html(companyrequest.results(0).email) %>"><%= db2html(companyrequest.results(0).email) %></a><br><br>

<b>ȸ��URL</b> : <%
	dim arrUrl
	arrUrl = split(companyrequest.results(0).companyurl,",")
	if ubound(arrUrl)>0 then
		Response.Write "<a href='"
		if Left(arrUrl(0),7) <> "http://" then Response.Write "http://"
		Response.Write arrUrl(0) & "' target='_blank'>" & arrUrl(0) & "</a>"
		Response.Write "<br><br><b>�������θ�</b> : " & arrUrl(1)
	else
		Response.Write "<a href='"
		if Left(companyrequest.results(0).companyurl,7) <> "http://" then Response.Write "http://"
		Response.Write companyrequest.results(0).companyurl & "' target='_blank'>" & companyrequest.results(0).companyurl & "</a>"
	end if
%>
<br><br>

<b>����/ȸ�缳��</b> : <%= nl2br(db2html(companyrequest.results(0).companycomments)) %><br><br>

<b>÷��ȭ��</b> : <% if (companyrequest.results(0).attachfile <> "") then %><a href="http://imgstatic.10x10.co.kr<%= companyrequest.results(0).attachfile %>" target="_blank">�ٿ�ޱ�</a><% end if %>
( ÷��ȭ�� �ȵɶ� <a href="http://www.10x10.co.kr<%= replace(companyrequest.results(0).attachfile,"uploadimg","uploadimage") %>" target="_blank">�ٿ�ޱ�</a> )
<br><br>

<b>ó������</b> : <% if (IsNull(companyrequest.results(0).finishdate) = true) then %>�̿Ϸ�<% else %><%= FormatDate(companyrequest.results(0).finishdate, "0000-00-00") %><% end if %><br><br>

<hr>
<b>ī�װ� �з�</b> : <b><font color=blue><%= GetCategoryName(companyrequest.results(0).categubun) %></font></b>&nbsp;&nbsp;&nbsp;-->&nbsp;&nbsp;
<b>ī�װ� ����</b> :
		<% call DrawSelectBoxCategoryLarge("categubun",catevalue) %>
		<input type=button value="����" onclick="catesubmit();">
<br><br>
<b>�Ǹ� ����</b> :
<% if companyrequest.results(0).sellgubun="Y" then %>
<b><font color=blue>ON-Line</font></b>
<% elseif companyrequest.results(0).sellgubun="N" then%>
<b><font color=blue>OFF-Line</font></b>
<% else %>
<b><font color=blue>��Ÿ</font></b>
<% end if %>&nbsp;&nbsp;&nbsp;-->&nbsp;&nbsp;
<b>�Ǹ����� ����</b> :
		<select name="sellgubun" class="a">
			<option value="Y">ON-Line</option>
			<option value="N">OFF-Line</option>
		</select>
		<input type=button value="����" onclick="sellsubmit();">
<br><br>
<b>��������</b> :
		<%if companyrequest.results(i).ipjumYN="Y" then response.write "<b><font color=blue>�����Ϸ�</font></b>" %>
		<%if companyrequest.results(i).ipjumYN="N" then response.write "<b><font color=blue>������</font></b>" %>
		&nbsp;&nbsp;&nbsp;-->&nbsp;&nbsp;
<b>�������� ����</b> :
		<select name="ipjumYN" class="a">
			<option value="Y">���� �Ϸ�</option>
			<option value="N">�� ����</option>
		</select>
		<input type=button value="����" onclick="ipjumYNsubmit();">
<br><br>
<b>���λ���</b> : <%= db2html(nl2br(companyrequest.results(0).reqcomment)) %><br><br>

<hr>
<input type="button" value=" �Ϸ�ó�� " onclick="SubmitForm()">
<a href="javascript:MovePage(<%=page%>);">������� �̵�</a>
</form>

<hr>

<!-- �ڸ�Ʈ �κ� -->
<table width="100%" cellspacing=0 cellpadding=0 border=0>
<tr>
	<td></td>
	<td><b>��ü���� ���Ϻ�����</b></td>
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
			<input type="button" value="����" onclick="javascript:editcomm();">
		</td>
	</tr>
	<tr>
		<td height=40><input type="button" value="mail������" onclick="javascript:sendmail();">	</td>
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
			<input type="button" value="����" onclick="javascript:savecomm();">
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
			<input type="button" value="����" onclick="javascript:savecomm();">
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