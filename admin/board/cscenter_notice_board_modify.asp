<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' History : 2008.04.29 �ѿ�� �߰�
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/classes/board/boardnoticecls.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<%

dim i, j

'==============================================================================
'��������
dim boardnotice

dim SearchKey, SearchString, param, page, noticetype, menupos,oldyn
page = Request("page")
noticetype = Request("noticetype")
SearchKey = Request("SearchKey")
SearchString = Request("SearchString")
menupos = Request("menupos")
oldyn = request("oldyn")
param = "&SearchKey=" & SearchKey & "&SearchString=" & Server.URLencode(SearchString) & "&noticetype=" & noticetype & "&menupos=" & menupos

param = "&SearchKey=" & SearchKey & "&SearchString=" & Server.URLencode(SearchString) & "&oldyn="& oldyn &"&noticetype=" & noticetype & "&menupos=" & menupos



set boardnotice = New CBoardNotice

boardnotice.read(request("id"))

%>
<STYLE TYPE="text/css">
<!--
    A:link, A:visited, A:active { text-decoration: none; }
    A:hover { text-decoration:underline; }
    BODY, TD, UL, OL, PRE { font-size: 10pt; }
-->
</STYLE>
<link rel=stylesheet type="text/css" href="/bct.css">
����Ÿ - ��������<br><br>
<script>
function jsPopCal(fName,sName)
{
	var fd = eval("document."+fName+"."+sName);

	if(fd.readOnly==false)
	{
		var winCal;
		winCal = window.open('/lib/common_cal.asp?FN='+fName+'&DN='+sName,'pCal','width=250, height=200');
		winCal.focus();
	}
}

function SubmitForm()
{
        if (document.f.title.value == "") {
                alert("������ �Է��ϼ���.");
                return;
        }
        if (document.f.contents.value == "") {
                alert("������ �Է��ϼ���.");
                return;
        }
        if (document.f.yuhyostart.value == "") {
                alert("��ȿ�������� �Է��ϼ���.");
                return;
        }
        if (document.f.yuhyoend.value == "") {
                alert("��ȿ�������� �Է��ϼ���.");
                return;
        }

		if (confirm("�����Ͻðڽ��ϱ�?") == true) {
			document.f.submit();
		}
}

function SubmitDelete()
{
        if (confirm("�����Ͻðڽ��ϱ�?") == true) {
                document.f.mode.value = "delete";
                document.f.submit();
        }
}
</script>


<table border="0" cellpadding="0" cellspacing="1" bgcolor="#808080" class="a">
<form method="post" name="f" action="cscenter_notice_board_act.asp" onsubmit="return false">
<input type="hidden" name="id" value="<%= request("id") %>">
<input type="hidden" name="mode" value="modify">
<input type="hidden" name="page" value="<%=page%>">
<input type="hidden" name="SearchKey" value="<%=SearchKey%>">
<input type="hidden" name="SearchString" value="<%=SearchString%>">
<input type="hidden" name="listtype" value="<%=noticetype%>">
<input type="hidden" name="oldyn" value="<%=oldyn%>">
<input type="hidden" name="menupos" value="<%=menupos%>">

<tr bgcolor="#FFFFFF">
	<td>��������</td>
	<td>
		  <select name="noticetype">
				<option value="" <% if boardnotice.results(0).Fnoticetype = "" then response.write "selected" %>>����</option>
				<!--<option value="01" <% if boardnotice.results(0).Fnoticetype = "01" then response.write "selected" %>>��ü����</option> 2015�����󿡼� ����. �̻��ش븮.//-->
				<option value="02" <% if boardnotice.results(0).Fnoticetype = "02" then response.write "selected" %>>�ȳ�</option>
				<option value="03" <% if boardnotice.results(0).Fnoticetype = "03" then response.write "selected" %>>�̺�Ʈ����</option>
				<option value="04" <% if boardnotice.results(0).Fnoticetype = "04" then response.write "selected" %>>��۰���</option>
				<option value="05" <% if boardnotice.results(0).Fnoticetype = "05" then response.write "selected" %>>��÷�ڰ���</option>
				<option value="06" <% if boardnotice.results(0).Fnoticetype = "06" then response.write "selected" %>>CultureStation</option>
		  </select>
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td>ī�װ�</td>
	<td>
		<% DrawSelectBoxCategoryOnlyLarge "malltype", boardnotice.results(0).Fmalltype, ""%>
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td>����</td>
	<td><input type="text" name="title" size="60" value="<%= boardnotice.results(0).Ftitle %>" class="input_b"></td>
</tr>
<tr bgcolor="#FFFFFF">
	<td>����</td>
	<td><textarea name="contents" cols="70" rows="15" class="textarea2"><%= db2html(boardnotice.results(0).Fcontents) %></textarea><br><font color="red">(������������� �Դϴ�. ������ ����Ű�� �ٸ������ּ���!!)</font></td>
</tr>
<tr bgcolor="#FFFFFF">
	<td>��ȿ������</td>
	<td><input type="text" size="10" name="yuhyostart" value="<%= boardnotice.results(0).Fyuhyostart %>" onClick="jsPopCal('f','yuhyostart');" style="cursor:hand;" class="input_b"></td>
</tr>
<tr bgcolor="#FFFFFF">
	<td>��ȿ������</td>
	<td><input type="text" size="10" name="yuhyoend" value="<%= boardnotice.results(0).Fyuhyoend %>" onClick="jsPopCal('f','yuhyoend');" style="cursor:hand;" class="input_b"></td>
</tr>
<tr bgcolor="#FFFFFF">
	<td>����������</td>
	<td><input type="radio" name="fixyn" value="Y" <% if boardnotice.results(0).Ffixyn = "Y" then response.write "checked" %>>��� <input type="radio" name="fixyn" value="N" <% if boardnotice.results(0).Ffixyn = "N" then response.write "checked" %>>������</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td>����� �߿���� ����</td>
	<td><input type="radio" name="importantnotice" value="Y" <% if boardnotice.results(0).FImportantNotice = "Y" then response.write "checked" %>>��� <input type="radio" name="importantnotice" value="N" <% if boardnotice.results(0).FImportantNotice = "N" then response.write "checked" %>>������</td>
</tr>
</form>
</table>
<br>
<a href="javascript:SubmitForm()" onfocus="this.blur()"><img src="/images/icon_modify.gif" border="0" align="absmiddle"></a>
<a href="javascript:SubmitDelete()" onfocus="this.blur()"><img src="/images/icon_delete.gif" border="0" align="absmiddle"></a>
<a href="cscenter_notice_board_list.asp?page=<%=page & param%>" onfocus="this.blur()"><img src="/images/icon_list.gif" border="0" align="absmiddle"></a>
<br><br>
<table cellpadding="5" cellspacing="0" border="0" class="a">
<tr>
	<td bgcolor="#F8F8FA" style="border:1px solid #D8D8DA">
		<b>(��ũ ���)</b><br>
		&lt;a href="javascript:GoParent('http://www.10x10.co.kr/event/eventmain.asp?eventid=2607')"&gt;�ҳ� �̺�Ʈ �ٷΰ���&lt;/a&gt;
	</td>
</tr>
</table>
<!-- #include virtual="/lib/db/dbclose.asp" -->