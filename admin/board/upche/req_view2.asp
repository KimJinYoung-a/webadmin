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
' Description : �������/���� ����
' History : 2013.07.25 ������ ����
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

	'// �⺻������ �������
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
�Խ��ǰ��� - ������� �� ����<br><br>
<script type="text/javascript">
function SubmitForm() {
	if (confirm("ó�����¸� �Ϸ�� ��ȯ�մϱ�?") == true) { document.f.submit(); }
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

<!-- ��ü���� ���� -->
<form method="post" name="f" action="/admin/board/upche/req_act.asp" onsubmit="return false">
<input type="hidden" name="mode" value="finish">
<input type="hidden" name="gubun" value="<%=gubun%>">
<input type="hidden" name="id" value="<%= companyrequest.results(0).id %>">
<table width="100%" align="center" cellpadding="3" cellspacing="1" bgcolor="black">
<tr bgcolor="FFFFFF" align="center">
	<td colspan="4"><b><font size=3 color="blue"><%= companyrequest.results(0).getAllianceGubun %> ����</font></b></td>
</tr>
<tr bgcolor="FFFFFF" align="center">
	<td bgcolor="<%= adminColor("gray") %>">�ٹ����� �����</td>
	<td colspan="3" align="left">
		<% sbGetwork "workid",companyrequest.results(0).Fworkid, "" %>
	</td>
</tr>
<tr bgcolor="FFFFFF" align="center">
	<td width="100" bgcolor="<%= adminColor("gray") %>">ȸ���</td>
	<td colspan="3" align="left"><%= db2html(companyrequest.results(0).companyname) %></td>
</tr>
<tr bgcolor="FFFFFF" align="center">
	<td bgcolor="<%= adminColor("gray") %>">ȸ���ּ�</td>
	<td colspan="3" align="left"><%= db2html(companyrequest.results(0).address) %></td>
</tr>
<tr bgcolor="FFFFFF">
	<td bgcolor="<%= adminColor("gray") %>" align="center">����ڵ�Ϲ�ȣ</td>	
	<td colspan="3" align="left"><%= db2html(companyrequest.results(0).license_no) %></td>
</tr>
<tr bgcolor="FFFFFF" align="center">
	<td width="100" bgcolor="<%= adminColor("gray") %>">����ڸ�</td>
	<td align="left"><%= db2html(companyrequest.results(0).chargename) %></td>
	<td width="100" bgcolor="<%= adminColor("gray") %>">�μ���</td>
	<td align="left"><%= db2html(companyrequest.results(0).chargeposition) %></td>
</tr>
<tr bgcolor="FFFFFF" align="center">
	<td bgcolor="<%= adminColor("gray") %>">��ȭ��ȣ#1</td>
	<td align="left"><%= db2html(companyrequest.results(0).phone) %></td>
	<td bgcolor="<%= adminColor("gray") %>">��ȭ��ȣ#2</td>
	<td align="left"><%= db2html(companyrequest.results(0).hp) %></td>
</tr>
	<tr bgcolor="FFFFFF" align="center">
	<td bgcolor="<%= adminColor("gray") %>">�̸���</td>
	<td colspan="3" align="left"><a href="mailto:<%= db2html(companyrequest.results(0).email) %>"><%= db2html(companyrequest.results(0).email) %></a></td>
</tr>
<tr bgcolor="FFFFFF" align="center">
	<td bgcolor="<%= adminColor("gray") %>">ȸ��URL</td>
	<td colspan="3" align="left">
		<%
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
	</td>
</tr>
<tr bgcolor="FFFFFF" align="center">
	<td bgcolor="<%= adminColor("gray") %>">ȸ��Ұ�</td>
	<td colspan="3" align="left">
		<%= nl2br(db2html(companyrequest.results(0).companycomments)) %>
	</td>
</tr>
<tr bgcolor="FFFFFF" align="center">
	<td bgcolor="<%= adminColor("gray") %>">���ǳ���</td>
	<td colspan="3" align="left">
		<%= nl2br(db2html(companyrequest.results(0).reqcomment)) %>
	</td>
</tr>
<tr bgcolor="FFFFFF" align="center">
	<td bgcolor="<%= adminColor("gray") %>">÷������</td>
	<td align="left">
		<% if (companyrequest.results(0).attachfile <> "") then %>
			<% if (Left(companyrequest.results(0).attachfile,4) = "http") then %>
				<a href="<%= companyrequest.results(0).attachfile %>" target="_blank">�ٿ�ޱ�</a>
			<% else %>
				<a href="http://imgstatic.10x10.co.kr<%= companyrequest.results(0).attachfile %>" target="_blank">�ٿ�ޱ�</a>
			<% end if %>
		<% else %>
			����
		<% end if %>
	</td>
	<td bgcolor="<%= adminColor("gray") %>">ó������</td>
	<td align="left">
		<% if (IsNull(companyrequest.results(0).finishdate) = true) then %>
			�̿Ϸ�
			&nbsp;
			<input type="button" value=" �Ϸ�ó�� " onclick="SubmitForm()">
		<% else %>
			<%= FormatDate(companyrequest.results(0).finishdate, "0000-00-00") %>
		<% end if %>
	</td>
</tr>
</table>
</form>
<div style="text-align:right;padding-bottom:16px;"><a href="javascript:MovePage(<%=page%>);">�������</a></div>

<!-- �ڸ�Ʈ �κ� -->
<form name="commfrm" method=post action="" onsubmit="return false">
<table width="100%" align="center" cellpadding="0" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr bgcolor="FFFFFF" align="center">
	<td colspan=3><b><font size=3 color="blue">��ü���� ���Ϻ�����</font></b></td>
</tr>

<% if commmode="" and companyrequest.results(0).replyuser <>"" then %>
<tr bgcolor="FFFFFF" align="center">
	<td width="10%" valign="top">
		�ۼ�: <%= db2html(companyrequest.results(0).replyuser) %>
	</td>
	<td width="75%" valign="top">
		<%= nl2br(db2html(companyrequest.results(0).replycomment)) %>
	</td>
	<td width="15%">
		<input type="button" value="����" onclick="javascript:editcomm();">
	</td>
</tr>
<tr bgcolor="FFFFFF" align="left">
	<td colspan=3><input type="button" value="mail������" onclick="javascript:sendmail();">	</td>
</tr>

<%
	'//�������
	elseif commmode="edit" then %>
<tr bgcolor="FFFFFF" align="center">
	<td width="10%" valign="top">
		�ۼ�: <%= session("ssBctCname") %>
	</td>
	<td valign="top">
		<textarea name="comment" rows=10 cols=95><%= db2html(companyrequest.results(0).replycomment) %></textarea>
	</td>
	<td>
		<input type="button" value="����" onclick="javascript:savecomm();">
	</td>
</tr>

<%
	'//�ۼ����
	elseif companyrequest.results(0).replyuser ="" then %>
<tr bgcolor="FFFFFF" align="center">
	<td valign="top">
		�ۼ�: <%= session("ssBctCname") %>
	</td>
	<td valign="top">
		<textarea name="comment" rows=10 cols=95></textarea>
	</td>
	<td>
		<input type="button" value="����" onclick="javascript:savecomm();">
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
