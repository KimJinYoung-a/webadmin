<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/classes/items/extsiteitemcls.asp"-->

<%
dim notmatch, research, page, cdl
notmatch = request("notmatch")
research = request("research")
page     = request("page")
cdl      = RequestCheckVar(request("cdl"),3)

if ((research="") and (notmatch="")) then notmatch="on"
if (page="") then page=1

dim oDnshopitem
set oDnshopitem = new CExtSiteItem
oDnshopitem.FRectNotMatchCategory = notmatch
oDnshopitem.FRectCate_large = cdl

'if (cdl<>"") then
    oDnshopitem.GetDnshopCategoryMachingList
'end if

dim i
%>
<script language='javascript'>
function MatchCateSubmit()
{
	var j = document.getElementsByName("ten_code").length;
	var m = 0;
	var tmp_mng = "";
	var tmp_disp = "";
	var tmp_stor = "";
	var tmp_eca = "";
	var tmp_rca = "";
	var tmp_spk = "";
	var tmp_sec = "";
	for(var i=0; i < j ; i++){
	    if (document.getElementsByName("ten_code")[i].checked == true)
	    {
	    	m = m+1;
	    	if (document.getElementsByName("mng")[i].value == "" || document.getElementsByName("disp")[i].value == "" || document.getElementsByName("stor")[i].value == "" || document.getElementsByName("eca")[i].value == "" || document.getElementsByName("rca")[i].value == "" || document.getElementsByName("spk")[i].value == "")
	    	{
	    		alert("������ �Է°� 3�� ��� �Է��� �ּ���.");
	    		document.getElementsByName("ten_code")[i].focus();
	    		return false;
	    	}
	    	else
	    	{
	    		tmp_mng = tmp_mng + document.getElementsByName("mng")[i].value + ",";
	    		tmp_disp = tmp_disp + document.getElementsByName("disp")[i].value + ",";
	    		tmp_stor = tmp_stor + document.getElementsByName("stor")[i].value + ",";
	    		tmp_eca = tmp_eca + document.getElementsByName("eca")[i].value + ",";
	    		tmp_rca = tmp_rca + document.getElementsByName("rca")[i].value + ",";
	    		tmp_spk = tmp_spk + document.getElementsByName("spk")[i].value + ",";
	    		tmp_sec = tmp_sec + document.getElementsByName("sec")[i].value + ",";
	    	}
	    }
	}

	if (m == 0)
	{
		alert("�����Ͻ� ī�װ��� �����ϴ�.");
		return false;
	}
	else
	{
		if(confirm("�����Ͻ� ī�װ������� "+m+" ���Դϱ�?") == true) {
			frmCate.mngcate.value = tmp_mng;
			frmCate.dispcate.value = tmp_disp;
			frmCate.storcate.value = tmp_stor;
			frmCate.ecate.value = tmp_eca;
			frmCate.rcate.value = tmp_rca;
			frmCate.spkey.value = tmp_spk;
			frmCate.secate.value = tmp_sec;
			frmCate.submit();
			return true;
		} else {
			return false;
		}
	}
}

</script>
<table width="100%" border="0" cellpadding="5" cellspacing="1" bgcolor="#EEEEEE">
	<form name="frm" method="get" action="">
	<input type="hidden" name="page" value="1">
	<input type="hidden" name="research" value="on">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<tr >
		<td class="a">
    		<input type="checkbox" name="notmatch" <%= ChkIIF(notmatch="on","checked","") %> >��Ī �ȵ� ������
    		&nbsp;
    		ī�װ� : <% call DrawSelectBoxCategoryLarge("cdl",cdl) %>
		</td>
		<td class="a" align="right">
			<a href="exceldownload.asp?notmatch=<%=notmatch%>&cdl=<%=cdl%>"><img src="http://webadmin.10x10.co.kr/images/btn_excel.gif" border="0"></a>&nbsp;
			<a href="javascript:document.frm.submit();"><img src="/admin/images/search2.gif" width="74" height="22" border="0"></a>
		</td>
	</tr>
	</form>
</table>

<form name="frmCate" method="post" action="DnshopCate_Process.asp">
<input type="hidden" name="mngcate" value="">
<input type="hidden" name="dispcate" value="">
<input type="hidden" name="storcate" value="">
<input type="hidden" name="ecate" value="">
<input type="hidden" name="rcate" value="">
<input type="hidden" name="spkey" value="">
<input type="hidden" name="secate" value="">
<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="#CCCCCC">
<tr bgcolor="#FFFFFF">
	<td colspan="13">
		<table width="100%" cellpadding="0" cellspacing="0" border="0" class="a">
		<tr>
			<td></td>
			<td align="right" height="30">
				page: <%= FormatNumber(page,0) %> / <%= FormatNumber(oDnshopitem.FTotalPage,0) %> �ѰǼ�: <%= FormatNumber(oDnshopitem.FTotalCount,0) %>
				<br>
				<input type="button" value="�����Ѱ� ����" onClick="MatchCateSubmit()">
			</td>
		</tr>
		</table>
	</td>
</tr>
<tr align="center" bgcolor="#F3F3FF" height="20">
	<td width="100">Ten ī���ڵ�</td>
	<td width="100">��з�</td>
	<td width="100">�ߺз�</td>
	<td width="100">�Һз�</td>
	<td width="100">��ǰ��</td>
	<td width="100">���� cate</td>
	<td width="100">disp cate</td>
	<td width="100">store cate</td>
	<td width="100">���� cate</td>
	<td width="100">�̼� cate</td>
	<td width="100">�� cate</td>
	<td width="100">������Ű</td>
	<td></td>
</tr>
<% for i=0 to oDnshopitem.FResultCount-1 %>
<tr align="center" bgcolor="#FFFFFF">
    <td><%= oDnshopitem.FItemList(i).FCate_Large %><%= oDnshopitem.FItemList(i).FCate_Mid %><%= oDnshopitem.FItemList(i).FCate_Small %></td>
    <td><%= oDnshopitem.FItemList(i).Fnmlarge %></td>
    <td><%= oDnshopitem.FItemList(i).FnmMid %></td>
    <td><%= oDnshopitem.FItemList(i).FnmSmall %></td>
    <td><%= oDnshopitem.FItemList(i).FItemCnt %></td>
    <td><input type="text" name="mng" value="<%= oDnshopitem.FItemList(i).Fdnshopmngcategory %>" size="10"></td>
    <td><input type="text" name="disp" value="<%= oDnshopitem.FItemList(i).Fdnshopdispcategory%>" size="10"></td>
    <td><input type="text" name="stor" value="<%= oDnshopitem.FItemList(i).Fdnshopstorecategory%>" size="10"></td>
    <td><input type="text" name="eca" value="<%= oDnshopitem.FItemList(i).FdnshopEcategory%>" size="10"></td>
    <td><input type="text" name="rca" value="<%= oDnshopitem.FItemList(i).FdnshopRcategory%>" size="10"></td>
    <td><input type="text" name="sec" value="<%= oDnshopitem.FItemList(i).FdnshopSeCategory%>" size="10"></td>
    <td><input type="text" name="spk" value="<%= oDnshopitem.FItemList(i).FdnshopSpkey%>" size="10"></td>
    <td><input type="checkbox" name="ten_code" value="<%= oDnshopitem.FItemList(i).FCate_Large %>|<%= oDnshopitem.FItemList(i).FCate_Mid %>|<%= oDnshopitem.FItemList(i).FCate_Small %>"></td>
</tr>
<% next %>
<tr bgcolor="#FFFFFF">
	<td colspan="13">
		<table width="100%" cellpadding="0" cellspacing="0" border="0" class="a">
		<tr>
			<td></td>
			<td align="right" height="30">
				<input type="button" value="�����Ѱ� ����" onClick="MatchCateSubmit()">
			</td>
		</tr>
		</table>
	</td>
</tr>
</table>
</form>

<%
set oDnshopitem = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
