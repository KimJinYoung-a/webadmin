<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/admin/etc/wetoo1300k/wetoo1300kcls.asp"-->
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script language="javascript">
function fngobrd(v){
	location.replace('/admin/etc/wetoo1300k/popwetoo1300kBrandList.asp?brandcode='+v);
}
function savebrandcode(v){
    if (confirm('���� �Ͻðڽ��ϱ�?')){
		document.frmSvArr.brandcode.value = v;
		document.frmSvArr.makerid.value = $("#"+v+"").val();
		document.frmSvArr.action = "procwetoo1300k.asp";
		document.frmSvArr.submit();
    }
}
function delbrandcode(v, b){
    if (confirm('���� �Ͻðڽ��ϱ�?')){
		document.frmSvArr.makerid.value = v;
		document.frmSvArr.brandcode.value = b;
		document.frmSvArr.mode.value = "delbrandcode";
		document.frmSvArr.action = "procwetoo1300k.asp";
		document.frmSvArr.submit();
    }
}
</script>
<form name="frmSvArr" method="post" onSubmit="return false;" action="" style="margin:0px;">
<input type="hidden" name="mode" value="savebrandcode">
<input type="hidden" name="categbn" value="brand">
<input type="hidden" name="brandcode" value="">
<input type="hidden" name="makerid" value="">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width="80">�귣���ڵ�</td>
	<td>�귣���</td>
	<td>10x10 �귣��ID</td>
	<td>���</td>
</tr>
<tr align="center" bgcolor="#FFFFFF">
	<td width="80">001</td>
	<td>PEANUTS[�ǳ���]_�ٹ�����</td>
	<td>
		<input type="text" class="text" id="001" value="">
		<input type="button" class="button" onclick="savebrandcode('001');" value="����">
	</td>
	<td><input type="button" class="button" onclick="fngobrd('001');" value="����"></td>
</tr>
<tr align="center" bgcolor="#FFFFFF">
	<td width="80">002</td>
	<td>SANRIO[�긮��]_�ٹ�����</td>
	<td>
		<input type="text" class="text" id="002" value="">
		<input type="button" class="button" onclick="savebrandcode('002');" value="����">
	</td>
	<td><input type="button" class="button" onclick="fngobrd('002');" value="����"></td>
</tr>
<tr align="center" bgcolor="#FFFFFF">
	<td width="80">003</td>
	<td>Disney[�����]_�ٹ�����</td>
	<td>
		<input type="text" class="text" id="003" value="">
		<input type="button" class="button" onclick="savebrandcode('003');" value="����">
	</td>
	<td><input type="button" class="button" onclick="fngobrd('003');" value="����"></td>
</tr>
<tr align="center" bgcolor="#FFFFFF">
	<td width="80">004</td>
	<td>tenbyten[�ٹ�����]</td>
	<td>
		<input type="text" class="text" id="004" value="">
		<input type="button" class="button" onclick="savebrandcode('004');" value="����">
	</td>
	<td><input type="button" class="button" onclick="fngobrd('004');" value="����"></td>
</tr>
</table>
</form>
<%
Dim brandcode, o1300k, i, page
brandcode	= request("brandcode")
page		= request("page")
If page = ""	Then page = 1
If brandcode <> "" Then
	Set o1300k = new C1300k
		o1300k.FPageSize 		= 50
		o1300k.FCurrPage		= page
		o1300k.FRectBrandCode = brandcode
		o1300k.getTen1300kBrandCodeList
%>
<br />
<hr style="border:solid 3px;" />
<br />
<table width="50%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width="80">�귣���ڵ�</td>
	<td><%= brandCode %></td>
	<td width="80">���</td>
</tr>
<% For i=0 to o1300k.FResultCount - 1 %>
<tr align="center" bgcolor="#FFFFFF">
	<td width="80">�귣��ID</td>
	<td><%= o1300k.FItemList(i).FMakerid %></td>
	<td width="80">
		<input type="button" class="button" onclick="delbrandcode('<%= o1300k.FItemList(i).FMakerid %>', '<%= brandCode %>');" value="����">
	</td>
</tr>
<% Next %>
</table>
<%
	Set o1300k = nothing
End If
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->