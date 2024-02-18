<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  �������� ���������ֹ�������
' History : 2010.06.03 �ѿ�� ����
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/stock/offinvoicecls.asp"-->
<%

menupos = request("menupos")



dim page, shopid
dim research, i

page = request("page")
shopid = request("shopid")
research = request("research")

if (page = "") then
	page = 1
end if



'================================================================================
dim ocoffinvoice

set ocoffinvoice = new COffInvoice

ocoffinvoice.FRectShopid = shopid

ocoffinvoice.FCurrPage = page
ocoffinvoice.Fpagesize = 25

ocoffinvoice.GetMasterList

%>

<script language='javascript'>

function PopDownloadExportDeclareFile(masteridx,ino) {
	var popwin;

	popwin = window.open('<%= uploadImgUrl %>/linkweb/offinvoice/offinvoice_download.asp?idx=' + masteridx+'&ino='+ino,'PopDownloadExportDeclareFile','width=100,height=100,scrollbars=yes,resizable=yes');
	popwin.focus();
}

function popJungsanMaster(iid){
	var popwin = window.open('/admin/offshop/franmeaippopsubmaster.asp?idx=' + iid,'popsubmaster','width=800, height=600, scrollbars=yes, resizable=yes');
	popwin.focus();
}

function PopExportSheet(v){
	var popwin;
	popwin = window.open('/admin/fran/cartoonbox_modify.asp?menupos=1357&idx=' + v ,'PopExportSheet','width=740,height=600,scrollbars=yes,resizable=yes');
	popwin.focus();
}
</script>

<!-- �˻� ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frm" method="get" action="">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<input type="hidden" name="research" value="on">
	<tr align="center" bgcolor="#FFFFFF" >
		<td width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
		<td align="left">
			ShopID : 
			<% 'drawSelectBoxOffShop "shopid",shopid %>
			<% Call NewDrawSelectBoxDesignerwithNameAndUserDIV("shopid",shopid, "21") %>
		</td>
		<td width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button_s" value="�˻�" onClick="javascript:document.frm.submit();">
		</td>
	</tr>
	</form>
</table>
<!-- �˻� �� -->

<p>

<!-- �׼� ���� -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
<form name="offinvoiceaction" action="offinvoice_modify.asp">
<form name="mode" value="new">
<input type="hidden" name="menupos" value="<%= menupos %>">
<tr>
	<td align="right">
		<!--
		<input type="button" value="���κ��̽����" onclick="javascript:document.offinvoiceaction.submit();" class="button">
		-->
		* �κ��̽� �ۼ��� �������������(����) ���� �� �� �ֽ��ϴ�.
	</td>
</tr>
</form>
</table>
<!-- �׼� �� -->

<p>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="18">
			�˻���� : <b><%= ocoffinvoice.FTotalCount %></b>
			&nbsp;
			������ : <b><%= page %> / <%= ocoffinvoice.FTotalpage %></b>
		</td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td width="40">IDX</td>
		<td>�����̵�</td>
		<td>����IDX</td>
		<td>�۾�IDX</td>
		<td>�κ��̽�<br>NO</td>
		<td>���<br>���</td>
		<td>����<br>�δ�</td>
		<td>����<br>�ñ�</td>
		<td>�ڽ�<br>����</td>
		<td>�ѻ�ǰ�ݾ�<br>(��)</td>
		<td>�ѿ���<br>(��)</td>
		<td>�ۼ�ȭ��</td>
		<td>����ȯ��</td>
		<td>�ѻ�ǰ�ݾ�<br>(��ȯ)</td>
		<td>�ѿ���<br>(��ȯ)</td>
		<td width="80">�����</td>
		<td>����</td>
		<td>���</td>
	</tr>
	<% if ocoffinvoice.FResultCount >0 then %>
	<% for i=0 to ocoffinvoice.FResultcount-1 %>
	<tr bgcolor="#FFFFFF">
		<td align="center"><%= ocoffinvoice.FItemList(i).Fidx %></td>
		<td align="center"><a href="offinvoice_modify.asp?menupos=<%= menupos %>&idx=<%= ocoffinvoice.FItemList(i).Fidx %>"><%= ocoffinvoice.FItemList(i).Fshopid %><br><%= ocoffinvoice.FItemList(i).Fshopname %></a></td>
		<td align="center">
			<a href="javascript:popJungsanMaster(<%= ocoffinvoice.FItemList(i).Fjungsanidx %>)"><%= ocoffinvoice.FItemList(i).Fjungsanidx %></a>
		</td>
		<td align="center">
			<a href="javascript:PopExportSheet(<%= ocoffinvoice.FItemList(i).Fworkidx %>)"><%= ocoffinvoice.FItemList(i).Fworkidx %></a>
		</td>
		<td align="center"><%= ocoffinvoice.FItemList(i).Finvoiceno %></td>
		<td align="center"><%= ocoffinvoice.FItemList(i).GetDeliverMethodName %></td>
		<td align="center"><%= ocoffinvoice.FItemList(i).GetExportMethodName %></td>
		<td align="center"><%= ocoffinvoice.FItemList(i).GetJungsanTypeName %></td>
		<td align="center"><%= ocoffinvoice.FItemList(i).Ftotalboxno %></td>
		<td align="right">
			<%= FormatNumber(ocoffinvoice.FItemList(i).Ftotalgoodsprice, 0) %>&nbsp;
		</td>
		<td align="right">
			<%= FormatNumber(ocoffinvoice.FItemList(i).Ftotalboxprice, 0) %>&nbsp;
		</td>
		<td align="center"><%= ocoffinvoice.FItemList(i).Fpriceunit %></td>
		<td align="center"><%= FormatNumber(ocoffinvoice.FItemList(i).Fexchangerate, 0) %> ��</td>
		<td align="right">
			<% if (ocoffinvoice.FItemList(i).Fexchangerate <> "") and (Not IsNull(ocoffinvoice.FItemList(i).Fexchangerate)) and (ocoffinvoice.FItemList(i).Fexchangerate <> "0") then %>
				<% if (ocoffinvoice.FItemList(i).Fpriceunit = "JPY") then %>
					<%= FormatNumber(Round((ocoffinvoice.FItemList(i).Ftotalgoodsprice/(ocoffinvoice.FItemList(i).Fexchangerate/100)), 0), 0) %>&nbsp;
				<% else %>
					<%= FormatNumber(Round((ocoffinvoice.FItemList(i).Ftotalgoodsprice/ocoffinvoice.FItemList(i).Fexchangerate), 2), 2) %>&nbsp;
				<% end if %>
			<% end if %>
		</td>
		<td align="right">
			<% if (ocoffinvoice.FItemList(i).Fexchangerate <> "") and (Not IsNull(ocoffinvoice.FItemList(i).Fexchangerate)) and (ocoffinvoice.FItemList(i).Fexchangerate <> "0") then %>
				<% if (ocoffinvoice.FItemList(i).Fpriceunit = "JPY") then %>
					<%= FormatNumber(Round((ocoffinvoice.FItemList(i).Ftotalboxprice/(ocoffinvoice.FItemList(i).Fexchangerate/100)), 0), 0) %>&nbsp;
				<% else %>
					<%= FormatNumber(Round((ocoffinvoice.FItemList(i).Ftotalboxprice/ocoffinvoice.FItemList(i).Fexchangerate), 2), 2) %>&nbsp;
				<% end if %>
			<% end if %>
		</td>
		<td align="center"><%= Left(ocoffinvoice.FItemList(i).Fregdate, 10) %></td>
		<td align="center"><%= ocoffinvoice.FItemList(i).GetStateCDName %></td>
		<td align="center">
			<% if (ocoffinvoice.FItemList(i).Fexportdeclarefilename <> "") then %>
			<input type="button" class="button" value="����Ű�����1" onClick="PopDownloadExportDeclareFile(<%= ocoffinvoice.FItemList(i).Fidx %>,1)">
			<% end if %>
			<% if (ocoffinvoice.FItemList(i).Fexportdeclarefilename2 <> "") then %>
			<input type="button" class="button" value="����Ű�����2" onClick="PopDownloadExportDeclareFile(<%= ocoffinvoice.FItemList(i).Fidx %>,2)">
			<% end if %>
			<% if (ocoffinvoice.FItemList(i).Fexportdeclarefilename3 <> "") then %>
			<input type="button" class="button" value="����Ű�����3" onClick="PopDownloadExportDeclareFile(<%= ocoffinvoice.FItemList(i).Fidx %>,3)">
			<% end if %>
		</td>
	</tr>
	<% next %>
	<% else %>
<tr bgcolor="#FFFFFF">
		<td colspan=18 align=center>[ �˻������ �����ϴ�. ]</td>
	</tr>
	<% end if %>
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="18" align="center">
			<%
			dim strparam
			strparam = "&shopid=" + CStr(shopid)

			strparam = strparam + "&menupos=" + CStr(menupos)

			%>
			<% if ocoffinvoice.HasPreScroll then %>
				<a href="?page=<%= ocoffinvoice.StartScrollPage-1 %>&research=on<%= strparam %>">[pre]</a>
			<% else %>
				[pre]
			<% end if %>

			<% for i=0 + ocoffinvoice.StartScrollPage to ocoffinvoice.FScrollCount + ocoffinvoice.StartScrollPage - 1 %>
				<% if i>ocoffinvoice.FTotalpage then Exit for %>
				<% if CStr(page)=CStr(i) then %>
				<font color="red">[<%= i %>]</font>
				<% else %>
				<a href="?page=<%= i %>&research=on<%= strparam %>">[<%= i %>]</a>
				<% end if %>
			<% next %>

			<% if ocoffinvoice.HasNextScroll then %>
				<a href="?page=<%= i %>&research=on<%= strparam %>">[next]</a>
			<% else %>
				[next]
			<% end if %>
		</td>
	</tr>
</table>


<%
set ocoffinvoice = Nothing
%>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
