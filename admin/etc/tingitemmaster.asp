<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/classes/tingcls.asp"-->

<%
dim page
page=request("page")
if page="" then page=1

dim iting
set iting = new CTingItemList
iting.FPageSize = 100
iting.FCurrPage = page
iting.GetAllTingItemList

dim ix
%>
<script language="javascript">
function CheckNEditTing(frm){
	for (var i=0;i<frm.elements.length;i++){
		var e = frm.elements[i];

		if ((e.name=="itemid") || (e.name=="tingpoint") || (e.name=="tingpoint_b") || (e.name=="limitea") || (e.name=="limitsell")){
				if (e.value.length<1){
					alert('�ʼ� �Է� �����Դϴ�.');
					e.focus();
					return;
				}
			}

		if ((e.name=="itemid") || (e.name=="tingpoint") || (e.name=="tingpoint_b") || (e.name=="limitea") || (e.name=="limitsell")){
			if (!IsDigit(e.value)){
				alert('���ڸ� �����մϴ�.');
				e.focus();
				return;
			}
		}
	}

	var ret = confirm('�����Ͻðڽ��ϱ�?');
	if (ret){
		frm.submit();
	}
}

function checkNAddting(frm){
	for (var i=0;i<frm.elements.length;i++){
		var e = frm.elements[i];

		if ((e.name=="itemid") || (e.name=="tingpoint") || (e.name=="tingpoint_b") || (e.name=="limitea") || (e.name=="limitsell")){
				if (e.value.length<1){
					alert('�ʼ� �Է� �����Դϴ�.');
					e.focus();
					return;
				}
			}

		if ((e.name=="itemid") || (e.name=="tingpoint") || (e.name=="tingpoint_b") || (e.name=="limitea") || (e.name=="limitsell")){
			if (!IsDigit(e.value)){
				alert('���ڸ� �����մϴ�.');
				e.focus();
				return;
			}
		}
	}
	var ret = confirm('�߰��Ͻðڽ��ϱ�?');
	if (ret){
		frm.submit();
	}
}
</script>
<table width="960" border="1" cellpadding="0" cellspacing="0" class="a">
<tr >
	<td width="50" align="center">������ID</td>
	<td width="50" align="center">�̹���</td>
	<td width="120" align="center">�����۸�</td>
	<td width="70" align="center">������Ʈ(��)</td>
	<td width="70" align="center">������Ʈ(��)</td>
	<td width="70" align="center">���ű���</td>
	<td width="70" align="center">�����Ǹ�</td>
	<td width="64" align="center">����������</td>
	<td width="64" align="center">�Ǹż���</td>
	<td width="60" align="center">��������</td>
	<td width="70" align="center">����(���)����</td>
	<td width="70" align="center">�Ǹſ���</td>
	<td width="70" align="center">�̺�Ʈ����</td>
	<td width="70" align="center">Evt_CPCode</td>
	<td width="50" align="center">����</td>
</tr>
<% for ix=0 to iting.FResultCount-1 %>
<form name="frm_<%= iting.FTingList(ix).FID %>" method="post" action="dotingedit.asp">
<input type="hidden" name="mode" value="edit">
<input type="hidden" name="id" value="<%= iting.FTingList(ix).FID %>">
<tr>
	<td align="center">
		<input type="text" name="itemid" value="<%= iting.FTingList(ix).FItemID %>" size="7" maxlength="7" readonly >
	</td>
	<td align="center"><img src=<%= iting.FTingList(ix).FImageSmall %>></td>
	<td align="center"><%= iting.FTingList(ix).FItemName %></td>
	<td align="center">
		<input type="text" name="tingpoint" value="<%= iting.FTingList(ix).FTingPoint %>" size=6>
	</td>
	<td align="center">
		<input type="text" name="tingpoint_b" value="<%= iting.FTingList(ix).FTingPoint_B %>" size=6>
	</td>
	<td align="center">
		<select name="userclass">
			<option value="A" <% if iting.FTingList(ix).FUserClass="A" then response.write "selected" %> >����,����</option>
			<option value="Y" <% if iting.FTingList(ix).FUserClass="Y" then response.write "selected" %> >����</option>
			<option value="N" <% if iting.FTingList(ix).FUserClass="N" then response.write "selected" %> >����,����,����</option>
		</select>
	</td>
	<td align="center">
		<select name="limitdiv">
			<option value="0" <% if iting.FTingList(ix).FLimitDiv="0" then response.write "selected" %> >�������Ǹ�</option>
			<option value="1" <% if iting.FTingList(ix).FLimitDiv="1" then response.write "selected" %> >��������</option>
			<option value="2" <% if iting.FTingList(ix).FLimitDiv="2" then response.write "selected" %> >�Ϻ�����</option>
			<option value="3" <% if iting.FTingList(ix).FLimitDiv="3" then response.write "selected" %> >��������</option>
		</select>
	</td>
	<td align="center">
		<input type="text" name="limitea" value="<%= iting.FTingList(ix).FLimitea %>" size=6>
	</td>
	<td align="center">
		<input type="text" name="limitsell" value="<%= iting.FTingList(ix).FLimitSell %>" size=6>
	</td>
	<td align="center"><font color="#FF0000"><%= iting.FTingList(ix).FLimitea-iting.FTingList(ix).FLimitSell %></font></td>
	<td align="center">
		<select name="isusing">
			<option value="Y" <% if iting.FTingList(ix).Fisusing="Y" then response.write "selected" %> >������</option>
			<option value="N" <% if iting.FTingList(ix).Fisusing="N" then response.write "selected" %> >���þ���</option>
		</select>
	</td>
	<td align="center">
		<select name="sellyn">
			<option value="Y" <% if iting.FTingList(ix).Fsellyn="Y" then response.write "selected" %> >�Ǹ���</option>
			<option value="N" <% if iting.FTingList(ix).Fsellyn="N" then response.write "selected" %> >�Ǹž���</option>
		</select>
	</td>
	<td align="center">
		<select name="eventdiv">
			<option value="0" <% if iting.FTingList(ix).Feventdiv="0" then response.write "selected" %> >-</option>
			<option value="1" <% if iting.FTingList(ix).Feventdiv="1" then response.write "selected" %> >�̺�Ʈ1(��ǰ)</option>
			<option value="2" <% if iting.FTingList(ix).Feventdiv="2" then response.write "selected" %> >�̺�Ʈ2(��Ÿ)</option>
		</select>
	</td>
	<td align="center">
		<input type="text" name="eventcpcode" value="<%= iting.FTingList(ix).FEventCpCode %>" size=7>
	</td>
	<td align="center">
		<input type="button" value="����" onclick="CheckNEditTing(frm_<%= iting.FTingList(ix).FID %>)">
	</td>
</tr>
</form>
<% next %>
</table>
<%
set iting = Nothing
%>
<br>
<table width="400" border="1" cellpadding="0" cellspacing="0" class="a">
<tr>
	<td colspan="2">�� ��ǰ�߰�</td>
</tr>
<form name="frmting" method="post" action="dotingedit.asp">
<input type="hidden" name="mode" value="add">
<tr>
	<td width="100">������ID</td>
	<td ><input type="text" name="itemid" value="" size=6></td>
</tr>
<tr>
	<td width="100">���ű���</td>
	<td >
		<select name="userclass">
			<option value="A" >����,����</option>
			<option value="Y" >����</option>
			<option value="N" >����,����,����</option>
		</select>
	</td>
</tr>
<tr>
	<td width="100">�����Ǹ�</td>
	<td >
		<select name="limitdiv">
			<option value="0" >�������Ǹ�</option>
			<option value="1" >��������</option>
			<option value="2" >�Ϻ�����</option>
			<option value="3" >��������</option>
		</select>
	</td>
</tr>
<tr>
	<td width="100">������Ʈ(��)</td>
	<td ><input type="text" name="tingpoint" value="" size=7></td>
</tr>
<tr>
	<td width="100">������Ʈ(��)</td>
	<td ><input type="text" name="tingpoint_b" value="" size=7></td>
</tr>
<tr>
	<td width="100">�����Ǹż���</td>
	<td ><input type="text" name="limitea" value="0" size=7></td>
</tr>
<tr>
	<td width="100">���������Ǹż���</td>
	<td ><input type="text" name="limitsell" value="0" size=7></td>
</tr>
<tr>
	<td width="100">���ÿ���</td>
	<td>
		<select name="isusing">
			<option value="Y">������</option>
			<option value="N">���þ���</option>
		</select>
	</td>
</tr>
<tr>
	<td width="100">�Ǹſ���</td>
	<td>
		<select name="sellyn">
			<option value="Y">�Ǹ���</option>
			<option value="N">�Ǹž���</option>
		</select>
	</td>
</tr>
<tr>
	<td colspan="2" align="center"><input type="button" value="�߰�" onclick="checkNAddting(frmting)"></td>
</tr>
</form>
</table>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->