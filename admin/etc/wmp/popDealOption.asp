<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/admin/etc/wmp/wmpCls.asp"-->
<%
Dim oDealItem, itemid, i, arrRows, availCnt
itemid = Trim(request("itemid"))

SET oDealItem = new CWmp
	oDealItem.FRectItemID				= itemid
    arrRows = oDealItem.getDealOption
%>

<script language='javascript'>
function frm_check(frm){
	var obj;
	frm.itemoptionarr.value = "";
	frm.optionCountArr.value = "";

	for (i = 0; i < frm.elements.length; i++) {
		var e = frm.elements[i];
		if ((e.type == 'text')||(e.type == 'radio')) {
		  	if ((e.name.substring(0,"optionCount".length)) == "optionCount"){
				curritemoption = e.id;
		  	    //숫자만 가능
		  	    if (!IsDigit(e.value)){
		  	        alert('한정 수량은 숫자만 가능합니다.');
		  	        e.select();
		  	        e.focus();
		  	        return;
		  	    }
				frm.itemoptionarr.value = frm.itemoptionarr.value + curritemoption + "," ;
				frm.optionCountArr.value = frm.optionCountArr.value + e.value + "," ;
		  	}
		}
  	}
    if (confirm('저장 하시겠습니까?')){
        frm.submit();
    }
}
</script>

<form name="frm" method="post" action="procDealItem.asp" onsubmit="return false;">
<input type="hidden" name="mode" value="O">
<input type="hidden" name="itemid" value="<%= itemid %>">
<input type="hidden" name="itemoptionarr" value="">
<input type="hidden" name="optionCountArr" value="">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="<%= adminColor("gray") %>" align="center">
	<td>상품코드</td>
	<td colspan="2"><%= itemid %></td>
</tr>
<tr height="3" bgcolor="black" align="center">
    <td colspan=3></td>
</tr>
<tr height="25" bgcolor="<%= adminColor("gray") %>" align="center">
    <td width="20%">옵션코드</td>
	<td>옵션명</td>
	<td width="20%">재고</td>
</tr>
<%
If IsArray(arrRows) Then
	For i = 0 To Ubound(arrRows, 2)
		availCnt = 0
		If arrRows(7,i) = "" Then
			If arrRows(1,i) = "Y" Then
				availCnt = arrRows(4,i) - arrRows(5,i) - 5
			Else
				availCnt = 999
			End If

			If availCnt < 1 Then
				availCnt = 0
			End If
		Else
			availCnt = arrRows(6,i)
		End If
%>
<tr align="center" bgcolor="#FFFFFF">
    <td width="20%"><%= arrRows(2,i) %></td>
	<td width="20%"><%= arrRows(3,i) %></td>
	<td width="20%">
		<input type="text" id="<%= arrRows(2,i) %>" name="optionCount<%= arrRows(2,i) %>" size=5 value="<%= availCnt %>">
		<%= CHKIIF(arrRows(7,i) = "", "(미저장상태)", "")%>
	</td>
</tr>
<%
	Next
%>
<tr height="25" bgcolor="<%= adminColor("gray") %>" align="center">
    <td bgcolor="#FFFFFF" colspan="3">
        <input type="button" value="저장" class="button" onclick="frm_check(this.form);" />
    </td>
</tr>
<% Else %>
<tr height="25" bgcolor="<%= adminColor("gray") %>" align="center">
    <td bgcolor="#FFFFFF" colspan="3">
        옵션없음
    </td>
</tr>
<% End If %>
</table>
</form>
<% SET oDealItem = nothing %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->