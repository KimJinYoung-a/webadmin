<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/admin/etc/wmp/wmpCls.asp"-->
<%
Dim makerid, itemid, isGetDate
Dim page, i
Dim oDealItem
page                = request("page")
makerid				= requestCheckVar(request("makerid"), 32)
itemid  			= request("itemid")
isGetDate           = requestCheckVar(request("isGetDate"), 1)

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

SET oDealItem = new CWmp
	oDealItem.FCurrPage					= page
	oDealItem.FPageSize					= 50
    oDealItem.FRectMakerid				= makerid
	oDealItem.FRectItemID				= itemid
    oDealItem.FRectIsGetDate		   	= isGetDate
    oDealItem.getDealItemList
%>
<script language='javascript'>
function goPage(pg){
	frm.page.value = pg;
	frm.submit();
}
function popDealItem(){
	var popDealItem = window.open("/admin/etc/wmp/popDealItem.asp","popDealItem","width=700,height=400,scrollbars=yes,resizable=yes");
	popDealItem.focus();
}
function fnModifyMustPrice(iidx){
	var popMustPrice = window.open("/admin/etc/wmp/popDealItem.asp?idx="+iidx+"&isModify=Y","popMustPrice","width=700,height=400,scrollbars=yes,resizable=yes");
	popMustPrice.focus();
}
function popOption(iitemid){
	var popOption = window.open("/admin/etc/wmp/popDealOption.asp?itemid="+iitemid,"popOption","width=700,height=400,scrollbars=yes,resizable=yes");
	popOption.focus();
}
function fnDelItems(){
	var chkSel=0;
	try {
		if(frmSvArr.cksel.length>1) {
			for(var i=0;i<frmSvArr.cksel.length;i++) {
				if(frmSvArr.cksel[i].checked) chkSel++;
			}
		} else {
			if(frmSvArr.cksel.checked) chkSel++;
		}
		if(chkSel<=0) {
			alert("������ ��ǰ�� �����ϴ�.");
			return;
		}
	}
	catch(e) {
		alert("��ǰ�� �����ϴ�.");
		return;
	}
	if (confirm('�����Ͻ� ' + chkSel + '�� ���� �Ͻðڽ��ϱ�?')){
		document.frmSvArr.target = "xLink";
		document.frmSvArr.mode.value = "D";
		document.frmSvArr.action = "/admin/etc/wmp/procDealItem.asp"
		document.frmSvArr.submit();
    }
}
</script>
<!-- �˻� ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get" action="">
<input type="hidden" name="page" value="">
<input type="hidden" name="menupos" value="<%= menupos %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="3" width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
	<td align="left">
		�귣��&nbsp;&nbsp;&nbsp; : <% drawSelectBoxDesignerwithName "makerid",makerid %>&nbsp;
		&nbsp;
		<br /><br />
		��ǰ�ڵ� : <textarea rows="2" cols="20" name="itemid" id="itemid"><%=replace(itemid,",",chr(10))%></textarea>
        &nbsp;
        �����࿩��(���糯¥����) :
        <select name="isGetDate" class="select">
            <option value="" >-Choice-</option>
            <option value="Y" <%= CHKiif(isGetDate="Y","selected","") %> >������</option>
        </select>
    </td>
	<td rowspan="3" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="�˻�" onClick="javascript:document.frm.submit();">
	</td>
</tr>
</form>
</table>

<br />
<!-- ����Ʈ ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frmSvArr" method="post" onSubmit="return false;" action="" style="margin:0px;">
<input type="hidden" name="mode">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="10">
		�˻���� : <b><%= FormatNumber(oDealItem.FTotalCount,0) %></b>
		&nbsp;
		������ : <b> <%= FormatNumber(page,0) %> / <%= FormatNumber(oDealItem.FTotalPage,0) %></b>
	</td>
	<td align="right">
        <input type="button" class="button" value="����" onclick="popDealItem();" />
        &nbsp;
        <input type="button" class="button" value="����" onclick="fnDelItems();" />
    </td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
    <td width="20"><input type="checkbox" name="chkAll" onClick="fnCheckAll(this.checked,frmSvArr.cksel);"></td>
	<td width="50">�̹���</td>
    <td width="60">�ٹ�����<br>��ǰ��ȣ</td>
	<td>�귣��<br>��ǰ��<br><font color="blue">�����ǰ��</font></td>
    <td width="300">���Ⱓ</td>
	<td width="70">�ٹ�����<br>�ǸŰ�</td>
	<td width="70">�ٹ�����<br>����</td>
	<td width="70">ǰ������</td>
	<td width="70">�ֹ�����<br>����</td>
	<td width="70">�ɼǰ���</td>
	<td width="80">������ID</td>
</tr>
<% For i = 0 To oDealItem.FResultCount - 1 %>
<tr align="center" bgcolor="#FFFFFF">
    <td><input type="checkbox" name="cksel" onClick="AnCheckClick(this);"  value="<%= oDealItem.FItemList(i).FItemId %>"></td>
	<td><img src="<%= oDealItem.FItemList(i).Fsmallimage %>" width="50"></td>
	<td align="center">
		<a href="<%=wwwURL%>/<%=oDealItem.FItemList(i).FItemID%>" target="_blank"><%= oDealItem.FItemList(i).FItemID %></a>
	</td>
	<td align="left" style="cursor:pointer;" onclick="fnModifyMustPrice('<%= oDealItem.FItemList(i).FIdx %>');">
        <%= oDealItem.FItemList(i).FMakerid %><%= oDealItem.FItemList(i).getDeliverytypeName %><br><%= oDealItem.FItemList(i).FItemName %>
		<br/>
		<font color="blue"><%= oDealItem.FItemList(i).FNewItemName %></font>
    </td>
	<td>
		<%= FormatDate(oDealItem.FItemList(i).FStartDate,"0000-00-00 00:00:00") %> <br />~ <%= FormatDate(oDealItem.FItemList(i).FEndDate,"0000-00-00 00:00:00") %>
	</td>
	<td align="right">
	<% If oDealItem.FItemList(i).FSaleYn="Y" Then %>
		<strike><%= FormatNumber(oDealItem.FItemList(i).FOrgPrice,0) %></strike><br>
		<font color="#CC3333"><%= FormatNumber(oDealItem.FItemList(i).FSellcash,0) %></font>
	<% Else %>
		<%= FormatNumber(oDealItem.FItemList(i).FSellcash,0) %>
	<% End If %>
	</td>
	<td align="center">
	<%
		If oDealItem.FItemList(i).Fsellcash<>0 Then
			response.write CLng(10000-oDealItem.FItemList(i).Fbuycash/oDealItem.FItemList(i).Fsellcash*100*100)/100 & "%"
		End If
	%>
	</td>
	<td align="center">
	<%
		If oDealItem.FItemList(i).IsSoldOut Then
			If oDealItem.FItemList(i).FSellyn = "N" Then
	%>
			<font color="red">ǰ��</font>
	<%
			Else
	%>
			<font color="red">�Ͻ�<br>ǰ��</font>
	<%
			End If
		End If
	%>
	</td>
	<td align="center">
	<%
		If oDealItem.FItemList(i).FItemdiv = "06" OR oDealItem.FItemList(i).FItemdiv = "16" Then
			response.write "<font color='green'>�ֹ�����</font>"
		End If
	%>
	</td>
    <td align="center">
		<input type="button" class="button" value="�ɼ�" onclick="popOption('<%= oDealItem.FItemList(i).FItemId %>');">
	</td>
	<td align="center"><%= Chkiif(oDealItem.FItemList(i).Freguserid <> "", oDealItem.FItemList(i).Freguserid, oDealItem.FItemList(i).FLastUpdateUserId ) %></td>
</tr>
<% Next %>
<tr height="25" bgcolor="FFFFFF">
	<td colspan="18" align="center">
	<% If oDealItem.HasPreScroll Then %>
		<a href="javascript:goPage('<%= oDealItem.StartScrollPage-1 %>');">[pre]</a>
	<% Else %>
		[pre]
	<% End If %>
	<% For i=0 + oDealItem.StartScrollPage To oDealItem.FScrollCount + oDealItem.StartScrollPage - 1 %>
		<% If i>oDealItem.FTotalpage Then Exit For %>
		<% If CStr(page)=CStr(i) Then %>
		<font color="red">[<%= i %>]</font>
		<% Else %>
		<a href="javascript:goPage('<%= i %>');">[<%= i %>]</a>
		<% End If %>
	<% Next %>
	<% If oDealItem.HasNextScroll Then %>
		<a href="javascript:goPage('<%= i %>');">[next]</a>
	<% Else %>
	[next]
	<% End If %>
	</td>
</tr>
</form>
</table>
<iframe name="xLink" id="xLink" frameborder="0" width="100%" height="300"></iframe>
<% SET oDealItem = nothing %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->