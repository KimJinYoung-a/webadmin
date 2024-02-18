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

dim mode, orgdata, oneline, onelineItems
dim barcode, itemgubun, itemid, itemoption
dim i, j, k
dim sqlStr, errMsg

mode			= request("mode")
orgdata			= request("orgdata")

if (mode = "ins") then
	orgdata = Split(Trim(orgdata), vbCrLf)
	for i = 0 to UBound(orgdata) - 1
		oneline = Trim(orgdata(i))
		if (oneline <> "") then
			onelineItems = Split(oneline, vbTab)
			barcode = onelineItems(0)
			itemgubun = onelineItems(1)
			itemid = onelineItems(2)
			itemoption = onelineItems(3)

			if (Len(barcode) <> 13) or (Len(itemgubun) <> 2) or (Not IsNumeric(barcode)) or (Len(itemoption) <> 4) then
				response.write "�߸��� ����Ÿ�Դϴ�." & oneline & "<br />"
			else
				errMsg = ""

				'// �����ڵ� ����
				sqlStr = " select top 1 r.reservedDate, IsNull(s.barcode, '') as barcode, s.itemgubun, s.itemid, s.itemoption "
				sqlStr = sqlStr + " from "
				sqlStr = sqlStr + " 	db_item.dbo.tbl_public_Barcode_reserved r "
				sqlStr = sqlStr + " 	left join [db_item].[dbo].[tbl_item_option_stock] s "
				sqlStr = sqlStr + " 	on "
				sqlStr = sqlStr + " 		1 = 1 "
				sqlStr = sqlStr + " 		and r.barcode = s.barcode "
				sqlStr = sqlStr + " where "
				sqlStr = sqlStr + " 	1 = 1 "
				sqlStr = sqlStr + " 	and r.barcode = '" & barcode & "' "
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
				sqlStr = sqlStr + " 	db_item.dbo.tbl_public_Barcode_reserved r "
				sqlStr = sqlStr + " where "
				sqlStr = sqlStr + " 	1 = 1 "
				sqlStr = sqlStr + " 	and r.itemgubun = '" & itemgubun & "' "
				sqlStr = sqlStr + " 	and r.itemid = " & itemid & " "
				sqlStr = sqlStr + " 	and r.itemoption = '" & itemoption & "' "
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

				if errMsg = "" then
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
					sqlStr = sqlStr & " 	r.reservedDate = getdate() "
					sqlStr = sqlStr & " 	, r.reservedCont = (case "
					sqlStr = sqlStr & " 		when r.itemgubun = '10' and IsNull(o.optionname, '') = '' then i.itemname "
					sqlStr = sqlStr & " 		when r.itemgubun = '10' and IsNull(o.optionname, '') <> '' then i.itemname + ' _ ' + o.optionname "
					sqlStr = sqlStr & " 		when r.itemgubun <> '10' and IsNull(si.shopitemoptionname, '') = '' then si.shopitemname "
					sqlStr = sqlStr & " 		when r.itemgubun <> '10' and IsNull(si.shopitemoptionname, '') <> '' then si.shopitemname + ' _ ' + si.shopitemoptionname "
					sqlStr = sqlStr & " 		else '' end) "
					sqlStr = sqlStr & " from "
					sqlStr = sqlStr & " 	db_item.dbo.tbl_public_Barcode_reserved r "
					sqlStr = sqlStr & " 	left join [db_item].[dbo].[tbl_item] i "
					sqlStr = sqlStr & " 	on "
					sqlStr = sqlStr & " 		1 = 1 "
					sqlStr = sqlStr & " 		and r.itemgubun = '10' "
					sqlStr = sqlStr & " 		and i.itemid = r.itemid "
					sqlStr = sqlStr & " 	left join [db_item].[dbo].[tbl_item_option] o "
					sqlStr = sqlStr & " 	on "
					sqlStr = sqlStr & " 		1 = 1 "
					sqlStr = sqlStr & " 		and r.itemgubun = '10' "
					sqlStr = sqlStr & " 		and r.itemid = o.itemid "
					sqlStr = sqlStr & " 		and r.itemoption = o.itemoption "
					sqlStr = sqlStr & " 	left join [db_shop].[dbo].[tbl_shop_item] si "
					sqlStr = sqlStr & " 	on "
					sqlStr = sqlStr & " 		1 = 1 "
					sqlStr = sqlStr & " 		and r.itemgubun = si.itemgubun "
					sqlStr = sqlStr & " 		and r.itemid = si.shopitemid "
					sqlStr = sqlStr & " 		and r.itemoption = si.itemoption "
					sqlStr = sqlStr & " where "
					sqlStr = sqlStr & " 	1 = 1 "
					sqlStr = sqlStr & " 	and r.barcode = '" & barcode & "' "
					sqlStr = sqlStr & " 	and r.reservedDate is NULL "
					dbget.Execute sqlStr

					response.write "OK : " & barcode & "<br />"
				else
					response.write "<font color=red>" & errMsg & "</font><br />"
				end if
			end if
		end if
	next

	dbget.close() : response.end
end if

%>
<script language='javascript'>

String.prototype.trim = function() {
    return this.replace(/(^\s*)|(\s*$)/gi, "");
}

function jsSubmit() {
	var frm = document.frm;
	var orgdata = frm.orgdata.value;
	var oneline, onelineItems, validitemcount;

	validitemcount = 0;
	orgdata = orgdata.split("\n");
	for (var i = 0; i < orgdata.length; i++) {
		oneline = orgdata[i];
		onelineItems = oneline.split("\t");

		if (oneline.trim() == "") {
			continue;
		}

		if (onelineItems.length != 4) {
			alert("������ ������ 4���� �÷��� �Ǿ�� �մϴ�.\n\n" + oneline);
			return false;
		}

		if (onelineItems[0].trim().length != 13) {
			alert("�߸��� ������ڵ��Դϴ�.\n\n" + oneline);
			return false;
		}

		if (onelineItems[1].trim().length != 2) {
			alert("�߸��� �����ڵ��Դϴ�.\n\n" + oneline);
			return false;
		}

		if (onelineItems[2].trim()*0 != 0) {
			alert("�߸��� ��ǰ�ڵ��Դϴ�.\n\n" + oneline);
			return false;
		}

		if (onelineItems[3].trim().length != 4) {
			alert("�߸��� �ɼ��ڵ��Դϴ�.\n\n" + oneline);
			return false;
		}

		validitemcount = validitemcount + 1;
	}

	if (validitemcount > 0) {
		frm.mode.value = "ins";
		frm.submit();
	} else {
		alert("��ϰ����� ��ǰ�� �����ϴ�.");
		return false;
	}

}

</script>
<table border=0 cellspacing=0 cellpadding=0 class="a">
<form name="frm" method="post" action="pop_barcode_multi.asp">
<input type="hidden" name="mode" value="">
<tr>
	<td>
		<font color="red">������ �и�</font><br>
		������ڵ�, ����, ��ǰ�ڵ�, �ɼ��ڵ�<br>
		<font color="red">��簪���� ������ ������ ����� �ȵ˴ϴ�.</font>
	</td>
	<td align="right" valign="bottom">
	</td>
</tr>
<tr>
	<td colspan=2>
	<textarea name="orgdata" cols=70 rows=5></textarea>
	</td>
</tr>
<tr>
	<td>
	</td>
	<td align="right">
		<div style="height:5px;"></div>
		<input type= button class="button" value=" �� �� �� " onclick="jsSubmit()">
	</td>
</tr>
</form>
</table>

<p>

</form>
</table>
<!-- #include virtual="/common/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
