<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/items/barcodeCls.asp"-->
<%
Dim itemid, page, itemgubun, useYN
Dim i, obarcode
page    				= request("page")
itemid  				= request("itemid")
useYN					= request("useYN")
'itemgubun				= request("itemgubun")

If page = "" Then page = 1

'�ٹ����� ��ǰ�ڵ� ����Ű�� �˻��ǰ�
If itemid<>"" then
	Dim iA, arrTemp, arrItemid
	itemid = replace(itemid,",",chr(10))
	itemid = replace(itemid,chr(13),"")
	arrTemp = Split(itemid,chr(10))
	iA = 0
	Do While iA <= ubound(arrTemp)
		If Trim(arrTemp(iA))<>"" then
			If Not(isNumeric(trim(arrTemp(iA)))) then
				Response.Write "<script language=javascript>alert('[" & arrTemp(iA) & "]��(��) ��ȿ�� ��ǰ�ڵ尡 �ƴմϴ�.');history.back();</script>"
				dbget.close()	:	response.End
			Else
				arrItemid = arrItemid & trim(arrTemp(iA)) & ","
			End If
		End If
		iA = iA + 1
	Loop
	itemid = left(arrItemid,len(arrItemid)-1)
End If

SET obarcode = new CBarcode
	obarcode.FCurrPage			= page
	obarcode.FPageSize			= 20
	obarcode.FRectItemID		= itemid
	obarcode.FRectUseYN			= useYN
	obarcode.FRectItemGubun		= itemgubun
	obarcode.getBarcodelist
%>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script language='javascript'>
function goPage(pg){
    frm.page.value = pg;
    frm.submit();
}

function pop_BarcodeCont(idx){
	var pCM = window.open("/admin/itemmaster/pop_barcode.asp?idx="+idx,"pop_BarcodeCont","width=800,height=300,scrollbars=yes,resizable=yes");
	pCM.focus();
}

function popBarcodeMulti() {
	var pop = window.open("/admin/itemmaster/pop_barcode_multi.asp","popBarcodeMulti","width=500,height=500,scrollbars=yes,resizable=yes");
	pop.focus();
}
</script>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get" action="">
<input type="hidden" name="page" value="">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="research" value="on">
<tr align="center" bgcolor="#FFFFFF">
	<td rowspan="3" width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
	<td align="left">
		��ǰ�ڵ� : <textarea rows="2" cols="20" name="itemid" id="itemid"><%=replace(itemid,",",chr(10))%></textarea>
		&nbsp;
	<!--
		���� :
		<select name="itemgubun" class="select">
			<option value="">��ü</option>
		</select>
	-->
		&nbsp;
		��Ͽ��� :
		<select name="useYN" class="select">
			<option value="">��ü</option>
			<option value="Y" <%= CHkIIF(useYN="Y","selected","") %>>��ϿϷ�</option>
			<option value="N" <%= CHkIIF(useYN="N","selected","") %>>�������</option>
		</select>
	</td>
	<td rowspan="3" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="�˻�" onClick="javascript:document.frm.submit();">
	</td>
</tr>
</form>
</table>
<!-- �׼� �� -->
<div style="height:5px;"></div>

<input type="button" class="button" value="�ϰ����" onClick="popBarcodeMulti()">

<div style="height:5px;"></div>

<!-- ����Ʈ ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frmSvArr" method="post" onSubmit="return false;" action="" style="margin:0px;">
<input type="hidden" name="cmdparam" value="">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="18">
		�˻���� : <b><%= FormatNumber(obarcode.FTotalCount,0) %></b>
		&nbsp;
		������ : <b> <%= FormatNumber(page,0) %> / <%= FormatNumber(obarcode.FTotalPage,0) %></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width="50">idx</td>
	<td width="100"><b>������ڵ�</b></td>
	<td width="30"><b>����</b></td>
	<td width="80"><b>��ǰ�ڵ�</b></td>
	<td width="40"><b>�ɼ�<br />�ڵ�</b></td>
	<td>�귣��</td>
	<td>��ǰ��</td>
	<td>�ɼǸ�</td>
	<td width="80">�Է���</td>
	<td width="80">�����</td>
	<td>��ϻ�ǰ��</td>
	<td width="100">�����</td>
</tr>

<% For i=0 to obarcode.FResultCount - 1 %>
<tr align="center" bgcolor="#FFFFFF" onclick="pop_BarcodeCont('<%= obarcode.FItemList(i).FIdx %>')" style="cursor:pointer;">
	<td align="center"><%= obarcode.FItemList(i).FIdx %></td>
	<td align="center"><%= obarcode.FItemList(i).FBarcode %></td>
	<td align="center"><%= obarcode.FItemList(i).FItemgubun %></td>
	<td align="center"><%= obarcode.FItemList(i).FItemid %></td>
	<td align="center"><%= obarcode.FItemList(i).FItemoption %></td>
	<td align="center"><%= obarcode.FItemList(i).Fmakerid %></td>
	<td align="left"><%= obarcode.FItemList(i).Fitemname %></td>
	<td align="left"><%= obarcode.FItemList(i).Fitemoptionname %></td>
	<td align="center"><%= Left(obarcode.FItemList(i).FRegdate, 10) %></td>
	<td align="center"><%= Left(obarcode.FItemList(i).FReservedDate, 10) %></td>
	<td align="left"><%= nl2br(db2html(obarcode.FItemList(i).FReservedCont)) %></td>
	<td align="center"><%= obarcode.FItemList(i).Freguserid %></td>
</tr>
<% Next %>
<tr height="20">
    <td colspan="18" align="center" bgcolor="#FFFFFF">
        <% if obarcode.HasPreScroll then %>
		<a href="javascript:goPage('<%= obarcode.StartScrollPage-1 %>');">[pre]</a>
    	<% else %>
    		[pre]
    	<% end if %>

    	<% for i=0 + obarcode.StartScrollPage to obarcode.FScrollCount + obarcode.StartScrollPage - 1 %>
    		<% if i>obarcode.FTotalpage then Exit for %>
    		<% if CStr(page)=CStr(i) then %>
    		<font color="red">[<%= i %>]</font>
    		<% else %>
    		<a href="javascript:goPage('<%= i %>');">[<%= i %>]</a>
    		<% end if %>
    	<% next %>

    	<% if obarcode.HasNextScroll then %>
    		<a href="javascript:goPage('<%= i %>');">[next]</a>
    	<% else %>
    		[next]
    	<% end if %>
    </td>
</tr>
</form>
</table>
<% SET obarcode = nothing %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
