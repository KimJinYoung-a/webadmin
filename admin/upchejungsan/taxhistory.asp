<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/jungsan/electaxhistorycls.asp"-->
<%
dim makerid
dim onlyavail

dim page
dim otaxhistory
dim onlyComuniErr, onoffgubun, biz_no
dim taxcorp

set otaxhistory = new CElecTaxHistory

page = request("page")
if (page="") then page=1

makerid         = request("makerid")
onlyavail       = request("onlyavail")
onlyComuniErr   = request("onlyComuniErr")
onoffgubun      = request("onoffgubun")
biz_no          = request("biz_no")
taxcorp         = request("taxcorp")

otaxhistory.Fcomp = makerid
otaxhistory.Fright = onlyavail
otaxhistory.FRectonoffgubun    = onoffgubun
otaxhistory.FRectOnlyComuniErr = onlyComuniErr

otaxhistory.FPageSize = 100
otaxhistory.FCurrPage = page
otaxhistory.FRectbiz_no = biz_no
otaxhistory.FRectTaxCorp = taxcorp
otaxhistory.datalist()

dim ix

%>

<SCRIPT LANGUAGE="JavaScript">
<!--
function NextPage(ipage){
	document.frm.page.value= ipage;
	document.frm.submit();
}

function deleteTax(idx){
	if (confirm('��꼭���೻���� ����Ͻðڽ��ϱ�?')){
		window.open('do_taxhistory.asp?idx=' + idx + '&mode=delhistory','deleteTax','width=100, height=100')
	}
}
//-->
</SCRIPT>


<!-- ǥ ��ܹ� ����-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
	<form name="frm" method="get" action="">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<input type="hidden" name="research" value="on">
	<input type="hidden" name="page" value="">
   	<tr height="10" valign="bottom">
	        <td width="10" align="right"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
	        <td background="/images/tbl_blue_round_02.gif"></td>
	        <td background="/images/tbl_blue_round_02.gif"></td>
	        <td width="10" align="left" ><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
	</tr>
	<tr height="25" valign="bottom">
	        <td background="/images/tbl_blue_round_04.gif"></td>
	        <td valign="top">
	        	�귣�� : <% drawSelectBoxDesignerwithName "makerid", makerid %>
	        	<select name="onoffgubun">
	        	<option value="">�¿�������
	        	<option value="ON"  <%= chkIIF(onoffgubun="ON","selected","") %> >ON
	        	<option value="OF" <%= chkIIF(onoffgubun="OF","selected","") %> >OFF
	        	</select>
	        	<select name="taxcorp">
	        	<option value="">����籸��
	        	<option value="N"  <%= chkIIF(taxcorp="N","selected","") %> >NeoPort
	        	<option value="B" <%= chkIIF(taxcorp="B","selected","") %> >Bill36524
	        	</select>
	        	����ڹ�ȣ : <input type="text" name="biz_no" value="<%= biz_no %>" size="10" maxlength="10">
				&nbsp;
				<input type=checkbox name="onlyavail" value="Y" <% if onlyavail="Y" then response.write "checked" %> >����Ǹ��˻�
				&nbsp;
				<input type=checkbox name="onlyComuniErr" value="Y" <% if onlyComuniErr="Y" then response.write "checked" %> >��ſ������˻�(���������˻�����)
	        </td>
	        <td valign="top" align="right">
	        	<input type="image" src="/admin/images/search2.gif" width="74" height="22" border="0"></a>
	        </td>
	        <td background="/images/tbl_blue_round_05.gif"></td>
	</tr>
	</form>
</table>
<!-- ǥ ��ܹ� ��-->

<!-- ǥ �߰��� ����-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
	<form name="frmshortage" method=post action="doshortagestock.asp">
	<input type="hidden" name="mode" value="maxsellday">
	<tr height="5" valign="top">
		<td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
		<td align="right">�˻���� : �� <font color="red"><% = otaxhistory.FTotalCount %></font>��</td>
		<td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
	</tr>
	</form>
</table>
<!-- ǥ �߰��� ��-->

<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
	<tr align="center" bgcolor="#DDDDFF">
	    <td width="50">idx</td>
	    <td width="50">���걸��</td>
		<td width="50">�귣��</td>
		<td width="50">����</td>
		<td width="30">����</td>
		<td width="70">��꼭��ȣ</td>
		<td width="100">��꼭��</td>
		<td width="70">������</td>
		<td width="70">���ް�</td>
		<td width="70">�ΰ���</td>
		<td width="70">�ѹ���ݾ�</td>
		<td width="80">ȸ���</td>
		<td width="60">����ڹ�ȣ</td>
		<!-- td width="40">�����</td -->
		<td width="60">������</td>
		<td width="30">����</td>
		<td width="100">�����</td>
		<td width="30">����</td>
	</tr>
<% if otaxhistory.FresultCount<1 then %>
<tr>
	<td colspan="15" align="center">[�˻������ �����ϴ�.]</td>
</tr>
<% else %>
	<% for ix=0 to otaxhistory.FresultCount-1 %>
	<% if otaxhistory.FMasterItemList(ix).F_resultmsg<>"OK" then %>
	<tr bgcolor="#EEEEEE">
	<% else %>
	<tr bgcolor="#FFFFFF">
	<% end if %>
	    <td><%= otaxhistory.FMasterItemList(ix).F_idx %></td>
	    <td><%= otaxhistory.FMasterItemList(ix).getJGubunName %></td>
		<td><%= otaxhistory.FMasterItemList(ix).F_makerid %></td>
		<td><%= otaxhistory.FMasterItemList(ix).F_jungsangubun %></td>
		<td align=center><%= otaxhistory.FMasterItemList(ix).getTaxTypeName %></td>
		<td><%= otaxhistory.FMasterItemList(ix).F_uniq_id %></td>
		<td><%= otaxhistory.FMasterItemList(ix).F_jungsanname %></td>
		<td align=center><%= otaxhistory.FMasterItemList(ix).F_write_date %></td>
		<td align=right>
		<% IF Not IsNULL(otaxhistory.FMasterItemList(ix).F_item_amt) then %>
		<%= FormatNumber(otaxhistory.FMasterItemList(ix).F_item_amt,0) %>
		<% end if %>
		</td>
		<td align=right>
		<% IF Not IsNULL(otaxhistory.FMasterItemList(ix).F_item_vat) then %>
		<%= FormatNumber(otaxhistory.FMasterItemList(ix).F_item_vat,0) %>
		<% end if %>
		</td>
		<td align=right>
		<% IF Not IsNULL(otaxhistory.FMasterItemList(ix).F_item_amt) then %>
		<%= FormatNumber(otaxhistory.FMasterItemList(ix).F_item_amt+otaxhistory.FMasterItemList(ix).F_item_vat,0) %>
		<% end if %>
		</td>
		<td align=center><%= otaxhistory.FMasterItemList(ix).F_corp_nm %></td>
		<td align=center><%= otaxhistory.FMasterItemList(ix).F_biz_no %></td>
		<!-- td align=center><%= otaxhistory.FMasterItemList(ix).F_cur_dam_nm %></td -->
		<td align=center><acronym title="<%= otaxhistory.FMasterItemList(ix).F_resultmsg %>"><%= otaxhistory.FMasterItemList(ix).F_tax_no %></acronym></td>
		<% if LEFT(otaxhistory.FMasterItemList(ix).F_tax_no,2)="TX" then %>
		<td align=center style="cursor:hand"><img src="/images/icon_print02.gif" onclick="window.open('http://www.bill36524.com/popupBillTax.jsp?NO_TAX=<%= otaxhistory.FMasterItemList(ix).F_tax_no %>&NO_BIZ_NO=2118700620')"></td>
		<% else %>
		<td align=center style="cursor:hand"><img src="/images/icon_print02.gif" onclick="window.open('http://www.neoport.net/jsp/dti/tx/dti_get_pin.jsp?tax_no=<%= otaxhistory.FMasterItemList(ix).F_tax_no %>&cur_biz_no=2118700620&s_biz_no=<%= otaxhistory.FMasterItemList(ix).F_biz_no %>&b_biz_no=2118700620')"></td>
	<% end if %>
		<td align=center ><%= otaxhistory.FMasterItemList(ix).F_regdate %></td>
		<td align=center>
		<% if otaxhistory.FMasterItemList(ix).F_deleteyn="Y" then %>
		<font color="red"><%= otaxhistory.FMasterItemList(ix).F_deleteyn %></font>
		<% else %>
    		<% if (IsNULL(otaxhistory.FMasterItemList(ix).F_tax_no)) and (otaxhistory.FMasterItemList(ix).F_regdate>"2010-01-01") then %>
    		<a href="javascript:deleteTax('<%= otaxhistory.FMasterItemList(ix).F_idx %>')"><strong><%= otaxhistory.FMasterItemList(ix).F_deleteyn %></strong></a>
    		<% else %>
    		<a href="javascript:deleteTax('<%= otaxhistory.FMasterItemList(ix).F_idx %>')"><%= otaxhistory.FMasterItemList(ix).F_deleteyn %></a>
    		<% end if %>
		<% end if %>
		</td>
	</tr>
	<% next %>
<% end if %>

<!-- ǥ �ϴܹ� ����-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
    <tr valign="top" height="25">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td valign="bottom" align="center">
		<% if otaxhistory.HasPreScroll then %>
			<a href="javascript:NextPage('<%= otaxhistory.StarScrollPage-1 %>')">[pre]</a>
		<% else %>
			[pre]
		<% end if %>
		<% for ix=0 + otaxhistory.StarScrollPage to otaxhistory.FScrollCount + otaxhistory.StarScrollPage - 1 %>
			<% if ix>otaxhistory.FTotalpage then Exit for %>
			<% if CStr(page)=CStr(ix) then %>
			<font color="red">[<%= ix %>]</font>
			<% else %>
			<a href="javascript:NextPage('<%= ix %>')">[<%= ix %>]</a>
			<% end if %>
		<% next %>

		<% if otaxhistory.HasNextScroll then %>
			<a href="javascript:NextPage('<%= ix %>')">[next]</a>
		<% else %>
			[next]
		<% end if %>
        </td>
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
    </tr>
    <tr valign="bottom" height="10">
        <td width="10" align="right"><img src="/images/tbl_blue_round_07.gif" width="10" height="10"></td>
        <td background="/images/tbl_blue_round_08.gif"></td>
        <td width="10" align="left"><img src="/images/tbl_blue_round_09.gif" width="10" height="10"></td>
    </tr>
</table>
<!-- ǥ �ϴܹ� ��-->

<%
set otaxhistory = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->