<%@ language=vbscript %>
<% option explicit %>
<%
	Response.AddHeader "Cache-Control","no-cache"
	Response.AddHeader "Expires","0"
	Response.AddHeader "Pragma","no-cache"
%>
<%
'###########################################################
' Description : �Ű��� ī�װ� ��� ���� ������
' Hieditor : 2016-03-08 ���¿� ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/academy/magazineCls.asp" -->
<%
Dim oMagaZinecatecode , i

set oMagaZinecatecode = new CMagazineContents
	oMagaZinecatecode.GetRowcatecodeContent()
%>
<script src="/js/jquery-1.7.1.min.js" type="text/javascript"></script>
<script type="text/javascript">

function jscatecode()	{
	var frm=document.frmcatecode;

	if (!frm.catecodename.value){
		alert('ī�װ����� �Է����ּ���');
		frm.catecodename.focus();
		return;
	}

	frm.submit();
}

function jsDelcatecode(cidx)	{
	if(confirm("�����Ͻðڽ��ϱ�?")){
		document.frmcatecodedel.Cidx.value = cidx;
   		document.frmcatecodedel.submit();
	}
}
</script>
<div style="padding: 0 5 5 5">
</div>

<form name="frmcatecode" method="post" action="/academy/magazine/lib/catecodeProc.asp" >
<input type="hidden" name="mode" value="catecode"/>
<table width="100%" border="0" cellpadding="3" cellspacing="1" class="a">
<!-- ī�װ� �߰� -->
<tr>
	<td colspan="2">
		<table width="100%" border="0" align="left" class="a" cellpadding="3" cellspacing="1" bgcolor="a" id="tagadd">
		<tr>
			<td align="center" bgcolor="<%= adminColor("tabletop") %>">ī�װ��� �Է�</td>
			<td align="center" bgcolor="<%= adminColor("tabletop") %>">�߰�</td>
		</tr>
			<tr class="item<%= i+1 %>">
				<td bgcolor="#FFFFFF"><input type="text" name="catecodename"  value="" size="15" id="catecodename" /></td>
				<td bgcolor="<%= adminColor("tabletop") %>">
					<INPUT TYPE="button" onclick="jscatecode(); return false;" value="�߰�"/>
				</td>
			</tr>
		</table>
	</td>
</tr>
<!--// ī�װ� �߰� -->

<!-- ī�װ� ����Ʈ -->
<tr>
	<td colspan="2">
		<table width="100%" border="0" align="left" class="a" cellpadding="3" cellspacing="1" bgcolor="a" id="tagadd">
		<tr>
			<td align="center" bgcolor="<%= adminColor("tabletop") %>">ī�װ���</td>
			<td align="center" bgcolor="<%= adminColor("tabletop") %>">����</td>
		</tr>
		<% If oMagaZinecatecode.FTotalCount > 0  Then %>
			<% for i=0 to oMagaZinecatecode.FTotalCount - 1 %>
				<tr class="item<%= i+1 %>">
					<td bgcolor="#FFFFFF"><%= oMagaZinecatecode.FItemList(i).Fcatename %></td>
					<td bgcolor="<%= adminColor("tabletop") %>" width="60">
						<INPUT TYPE="button" onclick="jsDelcatecode('<%= oMagaZinecatecode.FItemList(i).Fidx %>'); return false;" value="����"/>
					</td>
				</tr>
			<% next %>
		<% end if %>
		</table>
	</td>
</tr>
<!--// ī�װ� ����Ʈ -->
<td align="left">
	<font color="red"> �� ī�װ� ������ ������ ī�װ��� ���õ� �Ű����� M.A �Ű��� ����Ʈ���� ������ϴ�.</font>		
</td>
<tr>
	<td align="right">
		<input type="button" value="Ȯ�� " class="button" onclick="window.close();"/>
	</td>
</tr>
</table>
</form>

<form name="frmcatecodedel" method="post" action="/academy/magazine/lib/catecodeProc.asp" >
	<input type="hidden" name="mode" value="catecodedel"/>
	<input type="hidden" name="Cidx" value="">
</form>
<% set oMagaZinecatecode = nothing %>
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->