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
<!-- #include virtual="/lib/classes/play/playCls.asp" -->
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

set oPlayTag = new CPlayContents
	oPlayTag.FRectIdx = idx
	oPlayTag.FRectsubIdx = subidx
	oPlayTag.FRectPlaycate = playcate
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
			newitem.find("#tagurl2").attr("value", "");
			newitem.find("#tagurl3").attr("value", "");
			newitem.find("#tagurl4").attr("value", "");
			newitem.addClass("item"+(parseInt(lastItemNo)+1));

			$("#imgIn").append(newitem);
		});
	});

	function chgextxt(v) {
	var urllink = document.getElementById("extxt");
		switch(v) {
			case "1":
				urllink.value='�˻����ڵ��Է�(�̱���)';
				break;
			case "2":
				urllink.value='55073 <--�̺�Ʈ��ȣ �Է�';
				break;
			case "3":
				urllink.value='392832 <--��ǰ�ڵ� �Է�';
				break;
			case "4":
				urllink.value='102102104 <--ī�װ���ȣ �Է�';
				break;
			case "5":
				urllink.value='ithinkso <--�귣����̵� �Է�';
				break;
		}
	}
</script>
<div style="padding: 0 5 5 5"> <img src="/images/icon_arrow_link.gif" align="absmiddle"> �±װ���<br/>���±� ���Է½� �ڵ� ���� �˴ϴ�. (URL �Է� ���� ��� ����)��<br/>��URL ���Է½� �˻��������� �̵��մϴ�.��<br/>������ ���̴� ������� �������� �ѷ����ϴ�.��<br/>��<span style="color:blue;font-weight:800;">�ϴ� ����) ���� ���ǻ����� �ý��������� �����ٶ��ϴ�.</span>��
<br/>��<span style="color:red;font-weight:300;">PC-URL : ���� URL �Է� ��)/event/eventmain.asp?eventid=�̺�Ʈ�ڵ�</span>��
<br/>��<span style="color:red;font-weight:300;">MO-URL : ����� URL�Է� ��)/category/category_itemprd.asp?itemid=��ǰ�ڵ�</span>��
<br/>��<span style="color:red;font-weight:300;">APP-URL : SELECT������ �ش� �ڵ� �Է� (�̺�Ʈ�ڵ�,��ǰ�ڵ�,�귣���,ī�װ���ȣ �� 1��)</span>��
</div>
<table width="100%" border="0" cellpadding="3" cellspacing="1" class="a" bgcolor="red">
<tr>
	<td colspan="3">
		<table width="100%" border="0" align="left" class="a" cellpadding="3" cellspacing="1">
		<tr>
			<td bgcolor="<%= adminColor("tabletop") %>" width="50">����)</td>
			<td bgcolor="<%= adminColor("tabletop") %>" width="50"><input type="text" value="����" size="15" readonly/></td>
			<td bgcolor="<%= adminColor("tabletop") %>" width="250"><input type="text" value="/event/eventmain.asp?eventid=�̺�Ʈ�ڵ�" size="35" readonly/></td>
			<td bgcolor="<%= adminColor("tabletop") %>" width="250"><input type="text" value="/category/category_itemprd.asp?itemid=��ǰ�ڵ�" size="35" readonly/></td>
			<td bgcolor="<%= adminColor("tabletop") %>" width="300">
				<select onchange="chgextxt(this.value);">
					<option value="">=����=</option>
					<option value="1">��ǰ��</option>
					<option value="2">�̺�Ʈ</option>
					<option value="3">�귣��</option>
					<option value="4">ī�װ�</option>
				</select>
				<input type="text" id="extxt" value="&lt;-- ������ �ش� ��ȣ �Է�" size="30" readonly/>
			</td>
		</tr>
		</table>
	</td>
</tr>
</table>

<form name="frmtag" method="post" action="/admin/sitemaster/play/lib/tagProc.asp" >
<input type="hidden" name="mode" value="tag"/>
<input type="hidden" name="idx" value="<%=idx%>"/>
<input type="hidden" name="subidx" value="<%=subidx%>"/>
<input type="hidden" name="playcate" value="<%=playcate%>"/>
<table width="100%" border="0" cellpadding="3" cellspacing="1" class="a">
<tr>
	<td colspan="3">
		<table width="100%" border="0" align="left" class="a" cellpadding="3" cellspacing="1" bgcolor="a" id="imgIn">
		<tr>
			<td bgcolor="<%= adminColor("tabletop") %>" width="50">&nbsp;</td>
			<td bgcolor="<%= adminColor("tabletop") %>">�±��Է�</td>
			<td bgcolor="<%= adminColor("tabletop") %>">PC-URL�Է�</td>
			<td bgcolor="<%= adminColor("tabletop") %>">MO-URL�Է�</td>
			<td bgcolor="<%= adminColor("tabletop") %>">APP-URL�Է�</td>
		</tr>
		<% If oPlayTag.FTotalCount > 0  Then %>
		<% for i=0 to oPlayTag.FTotalCount - 1 %>
		<tr class="item<%=i+1%>">
			<td bgcolor="<%= adminColor("tabletop") %>" width="50">�±׵��</td>
			<td bgcolor="#FFFFFF" width="50"><input type="text" name="tagname" value="<%=oPlayTag.FItemList(i).Ftagname%>" size="15" id="tagname" /></td>
			<td bgcolor="#FFFFFF" width="250"><input type="text" name="tagurl" value="<%=oPlayTag.FItemList(i).Ftagurl%>" size="35" id="tagurl"/></td>
			<td bgcolor="#FFFFFF" width="250"><input type="text" name="tagurl2" value="<%=oPlayTag.FItemList(i).Ftagurl2%>" size="35" id="tagurl2"/></td>
			<td bgcolor="#FFFFFF" width="300">
				<select name="tagurl3" id="tagurl3">
					<option value="" <%= chkiif(oPlayTag.FItemList(i).Ftagurl3="","selected","")%>>==����==</option>
					<option value="1" <%= chkiif(oPlayTag.FItemList(i).Ftagurl3="1","selected","")%>>��ǰ��</option>
					<option value="2" <%= chkiif(oPlayTag.FItemList(i).Ftagurl3="2","selected","")%>>�̺�Ʈ</option>
					<option value="3" <%= chkiif(oPlayTag.FItemList(i).Ftagurl3="3","selected","")%>>�귣��</option>
					<option value="4" <%= chkiif(oPlayTag.FItemList(i).Ftagurl3="4","selected","")%>>ī�װ�</option>
				</select>
				<input type="text" name="tagurl4" value="<%=oPlayTag.FItemList(i).Ftagurl4%>" size="30" id="tagurl4"/>
			</td>
		</tr>
		<% next%>
		<% Else %>
		<tr class="item1">
			<td bgcolor="<%= adminColor("tabletop") %>">�±׵��</td>
			<td bgcolor="#FFFFFF" width="50"><input type="text" name="tagname" value="" size="15" id="tagname" /></td>
			<td bgcolor="#FFFFFF" width="250"><input type="text" name="tagurl" value="" size="35" id="tagurl"/></td>
			<td bgcolor="#FFFFFF" width="250"><input type="text" name="tagurl2" value="" size="35" id="tagurl2"/></td>
			<td bgcolor="#FFFFFF" width="300">
				<select name="tagurl3" id="tagurl3">
					<option value="">=����=</option>
					<option value="1">��ǰ��</option>
					<option value="2">�̺�Ʈ</option>
					<option value="3">�귣��</option>
					<option value="4">ī�װ�</option>
				</select>
				<input type="text" name="tagurl4" value="" size="30" id="tagurl4"/>
			</td>
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