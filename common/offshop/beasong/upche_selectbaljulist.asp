<%@ language=vbscript %>
<%
option explicit
Response.Expires = -1
%>
<%
'###########################################################
' Description : �������� ���
' Hieditor : 2011.02.26 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/common/incSessionAdminOrUpche.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/common/lib/commonbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/offshop/upche/upchebeasong_cls.asp" -->

<%
dim ojumun ,i ,ix,sql ,j ,detailidxarr
	detailidxarr =  request("detailidx")

set ojumun = new cupchebeasong_list
	ojumun.frectdetailidxarr = detailidxarr
	ojumun.FRectDesignerID = session("ssBctID")
	ojumun.fDesignerSelectBaljuList()
%>

<SCRIPT LANGUAGE="JavaScript">

function winPrint() {
	window.print();
}

function ExcelPrint(iSheetType) {
    xlfrm.SheetType.value = iSheetType;
	xlfrm.target="iiframeXL";
	xlfrm.action="/common/offshop/beasong/upche_dobeasonglistexcel.asp";
	xlfrm.submit();
}

function CsvPrint(iSheetType){
    xlfrm.SheetType.value = iSheetType;
	xlfrm.target="iiframeXL";
	xlfrm.action="/common/offshop/beasong/upche_dobeasonglistCSV.asp";
	xlfrm.submit();
}

</SCRIPT>

<STYLE TYPE="text/css">

.print {page-break-before: always;font-size: 12px; color:red;}
.no {font-size: 12px; color:red;}
body {background-color:"#FFFFFF"}

</STYLE>
<input type="hidden" name="menupos" value="<%= menupos %>">

<!-- �׼� ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr align="center" bgcolor="<%= adminColor("topbar") %>">
		<td width="50" bgcolor="<%= adminColor("gray") %>">�׼�</td>
		<td align="left">
			<input type="button" class="button" onclick="winPrint()" value="����Ʈ�ϱ�">
			&nbsp;
			<input type=button class="button" onclick="ExcelPrint('')" value="����(�ּҺи�)">
			&nbsp;
			<input type=button class="button" onclick="ExcelPrint('V2')" value="����(�ּ�����)">
			&nbsp;
			<input type=button class="button" onclick="ExcelPrint('V4')" value="����(�Ϸù�ȣ �߰�)">
			&nbsp;
			<input type=button class="button" onclick="CsvPrint()" value="CSV�� ����">
		</td>
		<td width="100" bgcolor="<%= adminColor("gray") %>">
			�� �Ǽ� : <font color="red"><span id="totalno"></span>��</font>
		</td>
	</tr>
	<!--
	<tr bgcolor="<%= adminColor("topbar") %>">
		<td colspan="10">
			�������Ϸ� ����(1)�� ����� �ּҰ� 1,2�� �������� ��µ˴ϴ�.<br>
			�������Ϸ� ����(2)�� ����� �ּҰ� 1,2�� �ϳ��� �������� ��µ˴ϴ�.<br>
			����Ͻô� ��Ŀ� �°� (1) �Ǵ� (2)�� �����ϼż� ����Ͻʽÿ�.
		</td>
	</tr>
	-->
</table>
<!-- �׼� �� -->

<% for ix=0 to ojumun.FResultCount - 1 %>
<table class="no">
<tr>
	<td><% = ix +1 %></td>
</tr>
</table>

<table width="100%" border="1" cellspacing="0" cellpadding="0" class="a">
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td height="25">�ֹ���ȣ</td>
	<td>�ֹ���</td>
	<td>������</td>
	<td>������ ��ȭ</td>
	<td>������ �ڵ���</td>
	<td>������ email</td>
</tr>
<tr align="center">
	<td height="25"><%= ojumun.FItemList(ix).Forderno %></td>
	<td><%= FormatDateTime(ojumun.FItemList(ix).FRegDate,2) %></td>
	<td><%= ojumun.FItemList(ix).FReqName %></td>
	<td><%= ojumun.FItemList(ix).FReqPhone %></td>
	<td><%= ojumun.FItemList(ix).FReqHp %></td>
	<td><%= ojumun.FItemList(ix).FBuyemail %></td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td colspan="6">������ �ּ�</td>
</tr>
<tr align="center">
	<td colspan="6"><%= ojumun.FItemList(ix).FReqZipCode %>&nbsp;<%= ojumun.FItemList(ix).FReqZipAddr %>&nbsp;<%= ojumun.FItemList(ix).FReqAddress %></td>
</tr>
<tr>
	<td align="center" height="25" bgcolor="<%= adminColor("tabletop") %>">��Ÿ����</td>
	<td colspan="5" align="center">&nbsp;<%= nl2br(db2html(ojumun.FItemList(ix).FComment)) %></td>
</tr>
</table>

<br>

<table width="100%" border="1" cellspacing="0" cellpadding="0" class="a">
<tr align="center" height="25" bgcolor="<%= adminColor("tabletop") %>">
	<td>��ǰID</td>
	<td>��ǰ��</td>
	<td>�ɼǸ�</td>
	<td>�ǸŰ�</td>
	<td>����</td>
</tr>
<tr align="center" height="25">
	<td><%= ojumun.fitemlist(ix).fitemgubun %>-<%= FormatCode(ojumun.fitemlist(ix).FitemID) %>-<%= ojumun.fitemlist(ix).fitemoption %></td>
	<td><%= ojumun.FItemList(ix).FItemName %></td>
	<td><%= ojumun.FItemList(ix).FItemoptionName %></td>
	<td><%= FormatNumber(ojumun.FItemList(ix).fsellprice,0) %></td>
	<td><%= ojumun.FItemList(ix).FItemNo %></td>
</tr>
</table>

<br>
<% if ((ix+1) mod 4) = 0 then %><div class="print">&nbsp;</div><% end if %>
<% next %>
<%
set ojumun = Nothing
%>
<iframe name="iiframeXL" name="iiframeXL" width=0 height=0 frameborder=0></iframe>

<form name="xlfrm" method="post" action="">
	<input type="hidden" name="detailidxarr" value="<%= detailidxarr %>">
	<input type="hidden" name="isall" value="">
	<input type="hidden" name="SheetType" value="">
</form>

<script language='javascript'>
	totalno.innerText = "<%= ix %>";
</script>

<!-- #include virtual="/common/lib/commonbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->