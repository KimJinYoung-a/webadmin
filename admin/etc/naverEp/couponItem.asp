<%@ language=vbscript %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbCTopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/admin/etc/potal/potalCls.asp"-->
<%
Dim mallid, mode, couponitemid, sqlStr, itemid, SavearrCnt, mallName
Dim nItem, page, itemidarr, makerid, bigo
mallid		= requestCheckvar(request("mallid"),32)
page		= request("page")
mode		= request("mode")
couponitemid	= Trim(request("couponitemid"))
makerid		= requestCheckvar(request("makerid"), 32)

Select Case mallid
	Case "ggshop"		mallName = "���ۼ���"
	Case "naverEP"		mallName = "���̹�EP"
	Case "daumEP"		mallName = "����EP"
End Select

If Right(couponitemid,1) = "," Then couponitemid = Left(couponitemid, Len(couponitemid) - 1)

itemidarr	= request("itemidarr")
itemid		= request("itemid")
bigo 		= NullFillWith(Trim(requestCheckVar(request("bigo"),300)),"")
SavearrCnt 	= Ubound(Split(couponitemid,",")) + 1

If page = "" Then page = 1

Dim iA2, tmpItemID, arrTemp2, arrItemid2, j
If mode = "I" Then
	If couponitemid<>"" then
		tmpItemID = couponitemid
		tmpItemID = replace(tmpItemID,",",chr(10))
		tmpItemID = replace(tmpItemID,chr(13),"")
		arrTemp2 = Split(tmpItemID,chr(10))
		iA2 = 0
		Do While iA2 <= ubound(arrTemp2)
			If Trim(arrTemp2(iA2))<>"" then
				If Not(isNumeric(trim(arrTemp2(iA2)))) then
					Response.Write "<script language=javascript>alert('[" & arrTemp2(iA2) & "]��(��) ��ȿ�� ��ǰ�ڵ尡 �ƴմϴ�.');history.back();</script>"
					dbget.close()	:	response.End
				Else
					arrItemid2 = arrItemid2 & trim(arrTemp2(iA2)) & ","
				End If
			End If
			iA2 = iA2 + 1
		Loop
		arrItemid2 = left(arrItemid2,len(arrItemid2)-1)
	End If

	arrItemid2 = Split(arrItemid2, ",")
	for j = 0 to UBound(arrItemid2)
		if Trim(arrItemid2(j)) <> "" then
			couponitemid = Trim(arrItemid2(j))
			strSql = "	DECLARE @Temp CHAR(1) " & _
						"	If NOT EXISTS(SELECT * FROM db_item.[dbo].[tbl_nvs_item_force_coupon_by_item] Where itemid = '" & couponitemid & "') " & _
						"		BEGIN " & _
						"			INSERT INTO db_item.[dbo].[tbl_nvs_item_force_coupon_by_item] (itemid, regdate, adminid, comment) VALUES('" & couponitemid & "', getdate(),  '"&session("ssBctID")&"', '"& bigo &"') " & _
						"		END	"
			dbget.execute strSql
		end if
	Next
	couponitemid = Request("couponitemid")
 	response.write "<script language='javascript'>alert('�����Ͽ����ϴ�.');location.href='/admin/etc/naverEp/couponItem.asp?mallid="&mallid&"&menupos="&menupos&"';</script>"
ElseIf mode = "U" Then
	Dim cnt
	itemidarr = split(itemidarr,",")
	cnt = ubound(itemidarr)
	For i = 0 to cnt
		sqlStr = "DELETE db_item.[dbo].[tbl_nvs_item_force_coupon_by_item] WHERE itemid =" & itemidarr(i)
		dbget.execute sqlStr
	Next
	response.write "<script language='javascript'>alert('���� �Ͽ����ϴ�.');location.href='/admin/etc/naverEp/couponItem.asp?mallid="&mallid&"&menupos="&menupos&"';</script>"
End If

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

SET nItem = new CPotal
	nItem.FCurrPage					= page
	nItem.FPageSize					= 100
	nItem.FRectItemid				= itemid
	nItem.FMakerId					= makerid
    nItem.getPotalCouponItemidList
%>
<script language='javascript'>
var ichk = 1;
function jsChkAll(){
	var frm, blnChk;
	frm = document.fitem;
	if(!frm.chkI) return;
	if ( ichk == 1 ){
		blnChk = true;
		ichk = 0;
	}else{
		blnChk = false;
		ichk = 1;
	}
	for (var i=0;i<frm.elements.length;i++){
		var e = frm.elements[i];
		if ((e.type=="checkbox")) {
			e.checked = blnChk ;
		}
	}
}

//���� �귣�� �����ϱ�
function jsIsusing() {
	var frm;
	var sValue;
	frm = document.fitem;
	sValue = "";
	chkSel	= 0;

	if (frm.chkI.length > 1){
		for (var i=0;i<frm.chkI.length;i++){
			if(frm.chkI[i].checked) chkSel++;
			if (frm.chkI[i].checked){
				if (sValue==""){
					sValue = frm.chkI[i].value;
				}else{
					sValue =sValue+","+frm.chkI[i].value;
				}
			}
		}
	}else{
		if(frm.chkI.checked) chkSel++;
		if(frm.chkI.checked){
			sValue = frm.chkI.value;
		}
	}
	if(chkSel<=0) {
		alert("������ ��ǰ�� �����ϴ�.");
		return;
	}

	if(confirm("���� �Ͻðڽ��ϱ�?")){
		document.frmIsusing.itemidarr.value = sValue;
		document.frmIsusing.mode.value = "U";
		document.frmIsusing.submit();
	}
}

function insert_itemid()
{
	if(document.frm.couponitemid.value == "")
	{
		alert("��ǰ�ڵ带 �Է��ϼ���.");
		document.frm.couponitemid.focus();
		return;
	}
	if(confirm("���� �Ͻðڽ��ϱ�?")){
		document.frm.mode.value = "I";
		document.frm.submit();
	}
}
function goPage(pg){
    var frm = document.frmsearch;
    frm.page.value=pg;
	frm.submit();
}
</script>
<% If mallid = "ggshop" Then %>
<!-- #include virtual="/admin/etc/potal/inc_googleHead.asp" -->
<% ElseIf mallid = "naverEP" Then %>
<!-- #include virtual="/admin/etc/potal/inc_naverHead.asp" -->
<% End If %>
<!-- �˻� ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frmsearch" method="get" action="couponItem.asp" style="margin:0px;">
<input type="hidden" name="page" value="<%=page%>">
<input type="hidden" name="research" value="on">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="mallid" value="<%= mallid %>">
<tr bgcolor="#FFFFFF">
	<td height="50">
		<table width="100%" class="a">
		<tr>
		    <td width="90%">Mall ���� : <%= mallName %></td>
		    <td rowspan="4" width="10%"><input type="button" value="�� ��" onClick="goPage(1)" style="width:50px;height:50px;"></td>
		</tr>
		<tr>
			<td >
			��ǰ�ڵ� : <textarea rows="2" cols="20" name="itemid" id="itemid"><%=replace(itemid,",",chr(10))%></textarea>
			&nbsp;
			�귣��ID : <input type="text" class="text" name="makerid" value="<%=makerid%>" size="20"> <input type="button" class="button" value="ID�˻�" onclick="jsSearchBrandID(this.form.name,'makerid');" >&nbsp;
			</td>
		</tr>
		</table>
	</td>
</tr>
</form>
</table>

<form name="frmIsusing" method="post" action="couponItem.asp" style="margin:0px;">
	<input type="hidden" name="itemidarr" value="">
	<input type="hidden" name="mode">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<input type="hidden" name="mallid" value="<%= mallid %>">
</form>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a">
<form name="frm" action="couponItem.asp" method="post" style="margin:0px;">
<input type="hidden" name="mode">
<input type="hidden" name="mallid" value="<%= mallid %>">
<tr>
	<td>
		���� ���� ��ǰ�ڵ� : <textarea class="textarea" name="couponitemid" rows="2" cols="16"></textarea>
		&nbsp;&nbsp;
		�ڸ�Ʈ : <input type="text" class="text" name="bigo" size="40">
		<input type="button" class="button" value="����" onClick="insert_itemid()">
	</td>
	<td align="right">
		<% If nItem.fresultcount >0 then %>
			<input class="button" type="button" id="btnEditSel" value="�����������" onClick="jsIsusing();">
	    <% End If %>
	</td>
</tr>
</form>
</table>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr bgcolor="#FFFFFF" height="30" align="LEFT" height="25">
	<td colspan="10">
		�˻���� : <b><%= FormatNumber(nItem.FTotalCount,0) %></b>
		&nbsp;
		������ : <b> <%= FormatNumber(page,0) %> / <%= FormatNumber(nItem.FTotalPage,0) %></b>
	</td>
</tr>
<form name="fitem" method="post" style="margin:0px;">
<input type="hidden" name="sortarr" value="">
<input type="hidden" name="mallid" value="<%= mallid %>">
<tr align="center" bgcolor="<%= adminColor("tabletop") %>" height="30">
	<td><input type="checkbox" name="chkA" onClick="jsChkAll();"></td>
	<td>��ǰ�ڵ�</td>
	<td>�귣��ID</td>
	<td>�����</td>
	<td>�����</td>
	<td>�ڸ�Ʈ</td>
</tr>
<% If nItem.FResultCount > 0 Then %>
<% For i = 0 To nItem.FResultCount - 1 %>
<tr bgcolor="#FFFFFF" height="30" align="center" height="25">
	<td><input type="checkbox" name="chkI" onClick="AnCheckClick(this);"  value="<%= nItem.FItemlist(i).FItemid %>"></td>
	<td><%=nItem.FItemList(i).FItemid%></td>
	<td><%=nItem.FItemList(i).FMakerid%></td>
	<td><%=nItem.FItemList(i).FRegdate%></td>
	<td><%=nItem.FItemList(i).FRegid%></td>
	<td><%=nItem.FItemList(i).Fbigo%></td>
</tr>
<% Next %>
<tr height="30">
	<td colspan="16" align="center" bgcolor="#FFFFFF">
	<% If nItem.HasPreScroll Then %>
		<a href="javascript:goPage('<%= nItem.StartScrollPage-1 %>');">[pre]</a>
	<% Else %>
		[pre]
	<% End If %>
	<% For i=0 + nItem.StartScrollPage To nItem.FScrollCount + nItem.StartScrollPage - 1 %>
		<% If i>nItem.FTotalpage Then Exit For %>
		<% If CStr(page)=CStr(i) Then %>
		<font color="red">[<%= i %>]</font>
		<% Else %>
		<a href="javascript:goPage('<%= i %>');">[<%= i %>]</a>
		<% End If %>
	<% Next %>
	<% If nItem.HasNextScroll Then %>
		<a href="javascript:goPage('<%= i %>');">[next]</a>
	<% Else %>
	[next]
	<% End If %>
	</td>
</tr>
<% Else %>
<tr height="50">
	<td colspan="16" align="center" bgcolor="#FFFFFF">
		��ϵ� ��ǰ�ڵ尡 �����ϴ�
	</td>
</tr>
<% End If %>
</form>
</table>
<% SET nItem = nothing %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbCTclose.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->