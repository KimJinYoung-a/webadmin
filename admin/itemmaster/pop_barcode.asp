<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : ���ڵ�
' Hieditor : 2015.12.28 ������ ����
'			 2016.08.08 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/items/barcodeCls.asp"-->
<%
Dim idx, obarcode, mode, sqlStr
Dim itemid, barcode, itemoption, reservedCont, itemgubun
dim errMsg

idx				= request("idx")
mode			= request("mode")
itemid			= Trim(request("itemid"))
barcode			= request("barcode")
itemoption		= Trim(request("itemoption"))
reservedCont	= request("reservedCont")
itemgubun		= request("itemgubun")

If mode = "I" Then
	If len(itemgubun) <> 2 Then
		Call Alert_Return("�߸��� ��ǰ���� �Դϴ�")
		response.end
	End If

	If Not(isNumeric(itemid)) Then
		Call Alert_Return("�߸��� ��ǰ��ȣ �Դϴ�")
		response.end
	End If

	If len(itemoption) <> 4 Then
		Call Alert_Return("�߸��� �ɼ��ڵ� �Դϴ�")
		response.end
	End If

	'// �����ڵ� ����
	sqlStr = " select top 1 r.reservedDate, IsNull(s.barcode, '') as barcode, s.itemgubun, s.itemid, s.itemoption "
	sqlStr = sqlStr + " from "
	sqlStr = sqlStr + "		db_item.dbo.tbl_public_Barcode_reserved r "
	sqlStr = sqlStr + "		left join [db_item].[dbo].[tbl_item_option_stock] s "
	sqlStr = sqlStr + "		on "
	sqlStr = sqlStr + "			1 = 1 "
	sqlStr = sqlStr + "			and r.barcode = s.barcode "
	sqlStr = sqlStr + " where "
	sqlStr = sqlStr + "		1 = 1 "
	sqlStr = sqlStr + "		and r.barcode = '" & barcode & "' "
	rsget.Open sqlStr,dbget,1
	If not rsget.EOF Then
		if Not IsNull(rsget("reservedDate")) then
			errMsg = "�̹� ��ϵ� ������ڵ��Դϴ�. : " & barcode
		elseif rsget("barcode") <> "" then
			if (rsget("itemgubun") = itemgubun) and (CStr(rsget("itemid")) = CStr(itemid)) and (rsget("itemoption") = itemoption) then
				'// skip
			else
				errMsg = "�̹� ������� ������ڵ��Դϴ�. : " & rsget("itemgubun") & "-" & rsget("itemid") & "-" & rsget("itemoption")
			end if
		end if
	else
		errMsg = "�߸��� ������ڵ��Դϴ�. : " & barcode
	end if
	rsget.Close

	'// ��ǰ�ڵ� ����
	sqlStr = " select top 1 r.reservedDate "
	sqlStr = sqlStr + " from "
	sqlStr = sqlStr + "		db_item.dbo.tbl_public_Barcode_reserved r "
	sqlStr = sqlStr + " where "
	sqlStr = sqlStr + "		1 = 1 "
	sqlStr = sqlStr + "		and r.itemgubun = '" & itemgubun & "' "
	sqlStr = sqlStr + "		and r.itemid = " & itemid & " "
	sqlStr = sqlStr + "		and r.itemoption = '" & itemoption & "' "
	rsget.Open sqlStr,dbget,1
	If not rsget.EOF and errMsg = "" Then
		errMsg = "�̹� ��ϵ� ��ǰ�ڵ��Դϴ�. : " & itemgubun & "-" & itemid & "-" & itemoption
	end if
	rsget.Close

	if (itemgubun = "10") and errMsg = "" then
		sqlStr = " select top 1 i.itemid, IsNull(o.itemoption, '0000') as itemoption "
		sqlStr = sqlStr & "		from [db_item].[dbo].[tbl_item] i "
		sqlStr = sqlStr & "		left join [db_item].[dbo].[tbl_item_option] o "
		sqlStr = sqlStr & "		on "
		sqlStr = sqlStr & "			1 = 1 "
		sqlStr = sqlStr & "			and i.itemid = o.itemid "
		sqlStr = sqlStr & "			and o.itemoption = '" & itemoption & "' "
		sqlStr = sqlStr & "	where i.itemid = '" & itemid & "' "
		rsget.Open sqlStr,dbget,1
		If not rsget.EOF Then
			if (rsget("itemoption") <> itemoption) then
				errMsg = "�߸��� �ɼ��ڵ��Դϴ�. : " & itemgubun & "-" & itemid & "-" & itemoption
			end if
		else
			errMsg = "�߸��� ��ǰ�ڵ��Դϴ�. : " & itemgubun & "-" & itemid & "-" & itemoption
		end if
		rsget.Close
	elseif errMsg = "" then
		sqlStr = " select top 1 i.shopitemid as itemid "
		sqlStr = sqlStr & "		from [db_shop].[dbo].[tbl_shop_item] i "
		sqlStr = sqlStr & "	where i.shopitemid = '" & itemid & "' "
		sqlStr = sqlStr & "	and i.itemoption = '" & itemoption & "' "
		sqlStr = sqlStr & "	and i.itemgubun = '" & itemgubun & "' "
		rsget.Open sqlStr,dbget,1
		If rsget.EOF Then
			errMsg = "�߸��� ��ǰ�ڵ��Դϴ�. : " & itemgubun & "-" & itemid & "-" & itemoption
		end if
		rsget.Close
	end if

	If errMsg <> "" Then
		Call Alert_Return(errMsg)
		dbget.close() : response.end
	End If

	sqlStr = "EXECUTE [db_item].[dbo].[sp_Ten_itemBarCode_Reg] '" & itemgubun & "', '" & itemid & "', '" & itemoption & "', '" & barcode & "' "
	dbget.execute sqlStr

	sqlStr = ""
	sqlStr = sqlStr & " UPDATE db_item.dbo.tbl_public_Barcode_reserved " & VBCRLF
	sqlStr = sqlStr & " SET itemgubun = '"&itemgubun&"', itemid = '"&itemid&"',itemoption = '"&itemoption&"' " & VBCRLF
	sqlStr = sqlStr & " , reguserid = '" & session("ssBctId") & "' " & VBCRLF
	sqlStr = sqlStr & " WHERE barcode = '" & barcode & "' and reservedDate is NULL "
	dbget.Execute sqlStr

	sqlStr = ""
	sqlStr = sqlStr & " update r "
	sqlStr = sqlStr & " set "
	sqlStr = sqlStr & "		r.reservedDate = getdate() "
	sqlStr = sqlStr & "		, r.reservedCont = (case "
	sqlStr = sqlStr & "			when r.itemgubun = '10' and IsNull(o.optionname, '') = '' then i.itemname "
	sqlStr = sqlStr & "			when r.itemgubun = '10' and IsNull(o.optionname, '') <> '' then i.itemname + ' _ ' + o.optionname "
	sqlStr = sqlStr & "			when r.itemgubun <> '10' and IsNull(si.shopitemoptionname, '') = '' then si.shopitemname "
	sqlStr = sqlStr & "			when r.itemgubun <> '10' and IsNull(si.shopitemoptionname, '') <> '' then si.shopitemname + ' _ ' + si.shopitemoptionname "
	sqlStr = sqlStr & "			else '' end) "
	sqlStr = sqlStr & " from "
	sqlStr = sqlStr & "		db_item.dbo.tbl_public_Barcode_reserved r "
	sqlStr = sqlStr & "		left join [db_item].[dbo].[tbl_item] i "
	sqlStr = sqlStr & "		on "
	sqlStr = sqlStr & "			1 = 1 "
	sqlStr = sqlStr & "			and r.itemgubun = '10' "
	sqlStr = sqlStr & "			and i.itemid = r.itemid "
	sqlStr = sqlStr & "		left join [db_item].[dbo].[tbl_item_option] o "
	sqlStr = sqlStr & "		on "
	sqlStr = sqlStr & "			1 = 1 "
	sqlStr = sqlStr & "			and r.itemgubun = '10' "
	sqlStr = sqlStr & "			and r.itemid = o.itemid "
	sqlStr = sqlStr & "			and r.itemoption = o.itemoption "
	sqlStr = sqlStr & "		left join [db_shop].[dbo].[tbl_shop_item] si "
	sqlStr = sqlStr & "		on "
	sqlStr = sqlStr & "			1 = 1 "
	sqlStr = sqlStr & "			and r.itemgubun = si.itemgubun "
	sqlStr = sqlStr & "			and r.itemid = si.shopitemid "
	sqlStr = sqlStr & "			and r.itemoption = si.itemoption "
	sqlStr = sqlStr & " where "
	sqlStr = sqlStr & "		1 = 1 "
	sqlStr = sqlStr & "		and r.barcode = '" & barcode & "' "
	sqlStr = sqlStr & "		and r.reservedDate is NULL "
	dbget.Execute sqlStr

	response.write "<script language='javascript'>alert('�����Ͽ����ϴ�.');opener.location.reload();window.close();</script>"

'elseif mode = "D" Then
	' ����� ���� �ȵ�	' 2019.10.23 �ѿ��
	' sqlStr = "EXECUTE [db_item].[dbo].[sp_Ten_itemBarCode_Reg] '" & itemgubun & "', '" & itemid & "', '" & itemoption & "', '' "
	' dbget.execute sqlStr

	' sqlStr = "UPDATE db_item.dbo.tbl_public_Barcode_reserved " & VBCRLF
	' sqlStr = sqlStr & " SET itemid = NULL" & VBCRLF
	' sqlStr = sqlStr & " ,itemoption = NULL" & VBCRLF
	' sqlStr = sqlStr & " ,reservedDate = NULL" & VBCRLF
	' sqlStr = sqlStr & " ,reservedCont = NULL" & VBCRLF
	' sqlStr = sqlStr & " ,itemgubun = NULL" & VBCRLF
	' sqlStr = sqlStr & " WHERE idx = '"&idx&"' "

	' 'response.write sqlStr & "<Br>"
	' dbget.Execute sqlStr

	' response.write "<script language='javascript'>alert('�����Ͽ����ϴ�.');opener.location.reload();window.close();</script>"
End If

SET obarcode = new CBarcode
	obarcode.FRectIdx = idx
	obarcode.getBarcodeOneItem

%>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script language='javascript'>

function deletereg(){
	var frm = document.frmcontents;
	frm.mode.value='D';
	frm.submit();
}

function form_check(){
	var frm = document.frmcontents;
	if ($("#itemid").val() == '') {
		alert('��ǰ�ڵ带 �Է��ϼ���');
		$("#itemid").focus();
		 return;
	}
	if ($("#reservedCont").val() == '') {
		alert('������ �Է��ϼ���');
		$("#reservedCont").focus();
		 return;
	}

	frm.submit();
}

</script>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frmcontents" method="post" action="pop_barcode.asp">
<input type="hidden" name="idx" value="<%= obarcode.FOneItem.Fidx %>">
<input type="hidden" name="barcode" value="<%= obarcode.FOneItem.FBarcode %>">
<input type="hidden" name="mode" value="<%= CHKIIF(IsNull(obarcode.FOneItem.FItemid), "I", "U")%>">

	<tr bgcolor="#FFFFFF">
	    <td width="150" align="center">Idx :</td>
	    <td><%= obarcode.FOneItem.Fidx %></td>
	</tr>
	<tr bgcolor="#FFFFFF">
	    <td width="150" align="center">���ڵ� :</td>
	    <td><%= obarcode.FOneItem.FBarcode %></td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td width="150" align="center">��ǰ�ڵ�</td>
		<td>
			<input type="text" name="itemgubun" id="itemgubun" value="<%= obarcode.FOneItem.FItemgubun %>" maxlength="2" size="2" <%= CHKIIF(Not IsNull(obarcode.FOneItem.FItemid), "readonly", "") %> >
			<input type="text" name="itemid" id="itemid" value="<%= obarcode.FOneItem.FItemid %>" maxlength="8" size="8" <%= CHKIIF(Not IsNull(obarcode.FOneItem.FItemid), "readonly", "") %> >
			<input type="text" name="itemoption" value="<%= obarcode.FOneItem.FItemoption %>" maxlength="4" size="4" <%= CHKIIF(Not IsNull(obarcode.FOneItem.FItemid), "readonly", "") %> >
		</td>
	</tr>
	<tr bgcolor="#FFFFFF">
	    <td width="150" align="center">����</td>
	    <td>
			<%= obarcode.FOneItem.FReservedCont %>
	    </td>
	</tr>
	<tr bgcolor="#FFFFFF">
	    <td align="center" colspan=2>
			<% if IsNull(obarcode.FOneItem.FItemid) then %>
			<input type="button" value=" �� �� " onClick="form_check();" class="button">
			<% else %>
			<% '<input type="button" value=" �� �� " onClick="deletereg();" class="button"> %>
			<% end if %>
	    </td>
	</tr>
</form>
</table>
<% SET obarcode = nothing %>
<!-- #include virtual="/common/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
