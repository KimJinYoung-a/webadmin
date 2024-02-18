<%@ language=vbscript %>
<% option explicit %>
<%
	Response.AddHeader "Cache-Control","no-cache"
	Response.AddHeader "Expires","0"
	Response.AddHeader "Pragma","no-cache"
%>
<%
'###########################################################
' Description : �÷��� �������� �±װ���
' Hieditor : 2013-09-03 ����ȭ ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/play/play_moCls.asp" -->
<%
Dim idx , subidx , playcate , oPlayTag , i
	idx  = requestCheckVar(request("idx"),10)
	subidx  = requestCheckVar(request("subidx"),10)
	playcate = requestCheckVar(request("playcate"),10)

IF idx = "" THEN
	Response.Write "<script>alert('�߸��� ����Դϴ�.\nNo. ��ȣ�� �־�� �մϴ�.');</script>"
	dbget.close()
	Response.End
END IF
IF IsNumeric(idx) = False THEN
	Response.Write "<script>alert('�߸��� ����Դϴ�.\nNo. ��ȣ�� �־�� �մϴ�.');</script>"
	dbget.close()
	Response.End
END If

set oPlayTag = new CPlayMoContents
	oPlayTag.FRectIdx = idx
	oPlayTag.FRectType = playcate
	oPlayTag.GetRowTagContent()

%>
<script src="/js/jquery-1.7.1.min.js" type="text/javascript"></script>
<script type="text/javascript">
	$(document).ready(function(){
		// �ɼ��߰� ��ư Ŭ����
		$("#addItemBtn").click(function(){
			// item �� �ִ��ȣ ���ϱ�
			var lastItemNo = $("#imgIn tr:last").attr("class").replace("item", "");

			var newitem = $("#imgIn tr:eq(1)").clone();
			newitem.removeClass();
			newitem.find("td:eq(0)").attr("rowspan", "1");
			newitem.find("#tagname").attr("value", "");
			newitem.find("#tagurl").attr("value", "");
			newitem.addClass("item"+(parseInt(lastItemNo)+1));

			$("#imgIn").append(newitem);
		});
	});
</script>
<div style="padding: 0 5 5 5"> <img src="/images/icon_arrow_link.gif" align="absmiddle"><font size="5" color="red"><strong> �����</strong></font> �±װ���<br/>���±� ���Է½� �ڵ� ���� �˴ϴ�. (URL �Է� ���� ��� ����)��<br/>��URL ���Է½� �˻��������� �̵��մϴ�.��<br/>������ ���̴� ������� �������� �ѷ����ϴ�.��</div>
<form name="frmtag" method="post" action="tagProc.asp" >
<input type="hidden" name="mode" value="tag"/>
<input type="hidden" name="idx" value="<%=idx%>"/>
<input type="hidden" name="subidx" value="<%=subidx%>"/>
<input type="hidden" name="playcate" value="<%=playcate%>"/>
<table width="450" border="0" cellpadding="3" cellspacing="1" class="a">
<tr>
	<td colspan="3">
		<table width="450" border="0" align="left" class="a" cellpadding="3" cellspacing="1" bgcolor="a" id="imgIn">
		<tr>
			<td bgcolor="<%= adminColor("tabletop") %>">&nbsp;</td>
			<td bgcolor="<%= adminColor("tabletop") %>">�±��Է�</td>
			<td bgcolor="<%= adminColor("tabletop") %>">URL�Է�</td>
		</tr>
		<% If oPlayTag.FTotalCount > 0  Then %>
		<% for i=0 to oPlayTag.FTotalCount - 1 %>
		<tr class="item<%=i+1%>">
			<td bgcolor="<%= adminColor("tabletop") %>">�±׵��</td>
			<td bgcolor="#FFFFFF" width="50"><input type="text" name="tagname" value="<%=oPlayTag.FItemList(i).Ftagname%>" size="15" id="tagname" /></td>
			<td bgcolor="#FFFFFF" width="250"><input type="text" name="tagurl" value="<%=oPlayTag.FItemList(i).Ftagurl%>" size="35" id="tagurl"/></td>
		</tr>
		<% next%>
		<% Else %>
		<tr class="item1">
			<td bgcolor="<%= adminColor("tabletop") %>">�±׵��</td>
			<td bgcolor="#FFFFFF" width="50"><input type="text" name="tagname" value="" size="15" id="tagname" /></td>
			<td bgcolor="#FFFFFF" width="250"><input type="text" name="tagurl" value="" size="35" id="tagurl"/></td>
		</tr>
		<% End If %>
		</table>
	</td>
</tr>
<tr>
	<td align="left" colspan="1">
		<INPUT TYPE="button" id="addItemBtn" value="�±� �߰�"/>
	</td>
	<td align="right">
		<input type="image" src="/images/icon_confirm.gif"/>
		<a href="javascript:window.close();"><img src="/images/icon_cancel.gif" border="0"></a>
	</td>
</tr>
</table>
</form>
<%
	set oPlayTag = nothing
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->