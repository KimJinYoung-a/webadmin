<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/designer/incSessionDesigner.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/designer/lib/designerbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lectureadmin/lib/classes/jumun/baljucls.asp"-->

<SCRIPT LANGUAGE="JavaScript">
<!--
function winPrint() {
window.print();
}
//-->
</SCRIPT>
<STYLE TYPE="text/css">
<!--
.print {page-break-before: always;font-size: 12px; color:red;}
.no {font-size: 12px; color:red;}
body {background-color:"#FFFFFF"}
-->
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
			<input type=button class="button" onclick="ExcelPrint('V3')" value="����(��ü�ڵ�)">
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



<%
dim i
dim ojumun
dim ix,sql
Dim listitemlist,listitem,listitemcount


listitem =  request("orderserial")  '' orderserial is Index of Order
if listitem <> "" then
	if checkNotValidHTML(listitem) then
	response.write "<script type='text/javascript'>"
	response.write "	alert('��ȿ���� ���� ���ڰ� ���ԵǾ� �ֽ��ϴ�. �ٽ� �ۼ� ���ּ���');"
	response.write "</script>"
	response.End
	end if
end if
'''response.write listitem
set ojumun = new CJumunMaster

ojumun.FRectOrderSerial = listitem
ojumun.FRectDesignerID = session("ssBctID")
ojumun.DesignerSelectBaljuList

%>
<script language="JavaScript">
<!--
function ExcelPrint(iSheetType) {
    xlfrm.SheetType.value = iSheetType;
	xlfrm.target="iiframeXL";
	xlfrm.action="dobeasonglistexcel.asp";
	xlfrm.submit();
}

function CsvPrint(iSheetType){
    xlfrm.SheetType.value = iSheetType;
	xlfrm.target="iiframeXL";
	xlfrm.action="dobeasonglistCSV.asp";
	xlfrm.submit();
}


//OLD function
function ExcelGo1() {
	//var popwin = window.open('','popexcel','width=800, height=600, scrollbars=1,resizable=1');
	//xlfrm.target="popexcel";
	//popwin.location="beasonglistexcel_process.asp?orderserial=<%= listitem %>";


	xlfrm.target="_blank";
	xlfrm.action="beasonglistexcel_process.asp";
	xlfrm.submit();

}

//OLD function
function ExcelGo2() {
	//var popwin = window.open('','popexcel','width=800, height=600, scrollbars=1,resizable=1');
	//xlfrm.target="popexcel";
	//popwin.location="beasonglistexcel_process.asp?orderserial=<%= listitem %>";

	xlfrm.target="_blank";
	xlfrm.action="beasonglistexcel2_process.asp";
	xlfrm.submit();
}
//-->
</script>

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
		<td>������ ����</td>
		<td>������ ��ȭ</td>
		<td>������ �ڵ���</td>
		<td>������ email</td>
	</tr>
	<tr align="center">
		<td height="25"><%= ojumun.FMasterItemList(ix).FOrderSerial %></td>
		<td><%= FormatDateTime(ojumun.FMasterItemList(ix).FRegDate,2) %></td>
		<td><%= ojumun.FMasterItemList(ix).FBuyName %></td>
		<td><%= ojumun.FMasterItemList(ix).FBuyPhone %></td>
		<td><%= ojumun.FMasterItemList(ix).FBuyHp %></td>
		<td><%= ojumun.FMasterItemList(ix).FBuyemail %></td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td height="25">������</td>
		<td>������ ��ȭ</td>
		<td>������ �ڵ���</td>
		<td colspan="3">������ �ּ�</td>
	</tr>
	<tr align="center">
		<td height="25"><%= ojumun.FMasterItemList(ix).FReqName %></td>
		<td><%= ojumun.FMasterItemList(ix).FReqPhone %></td>
		<td><%= ojumun.FMasterItemList(ix).FReqHp %></td>
		<td colspan="3"><%= ojumun.FMasterItemList(ix).FReqZipCode %>&nbsp;<%= ojumun.FMasterItemList(ix).FReqZipAddr %>&nbsp;<%= ojumun.FMasterItemList(ix).FReqAddress %></td>
	</tr>
<% if Not IsNULL(ojumun.FMasterItemList(ix).Freqdate) then %>
	<tr>
		<td align="center" height="25">�޼���<br>����</td>
		<td colspan="5" align="left">
			<table border="0" cellspacing="5" cellpadding="0" class="a">
				<tr>
					<td>�������� : </td>
					<td> <%= Left(CStr(ojumun.FMasterItemList(ix).Freqdate),10) %>�� <%= (ojumun.FMasterItemList(ix).Freqtime) %>�� </td>
				</tr>
				<tr>
					<td>ī��/���� : </td>
					<td> <%= (ojumun.FMasterItemList(ix).getCardribbonName) %></td>
				</tr>
				<tr>
					<td>�޼��� : </td>
					<td><%= nl2br(db2html(ojumun.FMasterItemList(ix).Fmessage)) %></td>
				</tr>
				<tr>
					<td>������ ��� : </td>
					<td><%= (db2html(ojumun.FMasterItemList(ix).Ffromname)) %></td>
				</tr>
			</table>
		</td>
	</tr>
<% end if %>
	<tr>
		<td align="center" height="25" bgcolor="<%= adminColor("tabletop") %>">��Ÿ����</td>
		<td colspan="5" align="center">&nbsp;<%= nl2br(db2html(ojumun.FMasterItemList(ix).FComment)) %></td>
	</tr>
</table>

<p>

<table width="100%" border="1" cellspacing="0" cellpadding="0" class="a">
	<tr align="center" height="25" bgcolor="<%= adminColor("tabletop") %>">
		<td width="60" height="25">��ǰID</td>
		<td>��ǰ��</td>
		<td>�ɼǸ�</td>
		<td width="70">�ǸŰ�</td>
		<td width="50">����</td>
	</tr>
	<tr align="center" height="25">
		<td><a href="http://www.10x10.co.kr/street/designershop.asp?itemid=<%= ojumun.Fitemid %>" target="_blank"><%= ojumun.FMasterItemList(ix).Fitemid %></a></td>
		<td><%= ojumun.FMasterItemList(ix).FItemName %></td>
		<td><%= ojumun.FMasterItemList(ix).FItemoptionName %></td>
		<td><%= FormatNumber(ojumun.FMasterItemList(ix).FItemCost,0) %></td>
		<td><%= ojumun.FMasterItemList(ix).FItemNo %></td>
	</tr>
	<tr align="center">
		<td>�ֹ�����<br>�޼���</td>
		<td colspan="4" align="left">&nbsp;
		<% if (Not IsNULL(ojumun.FMasterItemList(ix).Frequiredetail)) and ojumun.FMasterItemList(ix).Frequiredetail<>"" then %>

			<% if (ojumun.FMasterItemList(ix).FItemNo>1) then %>
			<% for i=0 to ojumun.FMasterItemList(ix).FItemNo-1 %>
			    [<%= i+ 1 %>�� ��ǰ ����]
			    <%= nl2Br(splitValue(ojumun.FMasterItemList(ix).Frequiredetail,CAddDetailSpliter,i)) %>
			    <br>
			<% next %>
			<% else %>
			<%= nl2Br(ojumun.FMasterItemList(ix).Frequiredetail) %>
			<% end if %>

		<% end if %>
		</td>
	</tr>
</table>

<br>
<% if ((ix+1) mod 4) = 0 then %><div class="print">&nbsp;</div><% end if %>
<% next %>
<%
set ojumun = Nothing
%>
<iframe name="iiframeXL" name="iiframeXL" width="0" height="0" frameborder=0 scrolling=no marginheight=0 marginwidth=0 align=center></iframe>

<form name=xlfrm method=post action="">
<input type="hidden" name="orderserial" value="<%= listitem %>">
<input type="hidden" name="isall" value="">
<input type="hidden" name="SheetType" value="">
</form>
<script language='javascript'>
	totalno.innerText = "<%= ix %>";
</script>
<!-- #include virtual="/designer/lib/designerbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->