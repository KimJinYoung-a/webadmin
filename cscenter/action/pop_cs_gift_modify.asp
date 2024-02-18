<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/order/ordergiftcls.asp"-->
<%

function removeSpace(inputstr)
	removeSpace = inputstr

	removeSpace = Replace(removeSpace, " ", "")
	removeSpace = Replace(removeSpace, Chr(9), "")
	removeSpace = Replace(removeSpace, vbCr, "")
	removeSpace = Replace(removeSpace, vbLf, "")
end function


dim orderserial, mode, itemlist

orderserial		= requestCheckVar(request("orderserial"), 32)
mode 			= requestCheckVar(request("mode"), 32)
itemlist 		= removeSpace(request("itemlist"))


dim oGift, oGiftModi
set oGift = new COrderGift

select case mode
	case "chk"
		oGift.FRectOrderSerial = orderserial
		oGift.GetOneOrderGiftlist
	case else
		response.end
end select

dim i, j, k
dim IsGiftOK

dim sqlStr

%>
<script>
function jsDelGift(frm, gift_code) {
	if (confirm("����ǰ �����Ͻðڽ��ϱ�?") != true) { return false; }

	frm.mode.value = "del";
	frm.gift_code.value = gift_code;
	frm.submit();
}

function jsModiGift(frm, gift_code, modi_gift_code, modi_giftkind_code) {
	if (confirm("����ǰ�� �����մϴ�.\n\n�����Ͻðڽ��ϱ�?") != true) { return false; }

	frm.mode.value = "modi";
	frm.gift_code.value = gift_code;
	frm.modi_gift_code.value = modi_gift_code;
	frm.modi_giftkind_code.value = modi_giftkind_code;
	frm.submit();
}
</script>
<!-- ǥ ��ܹ� ����-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
   	<tr height="10" valign="bottom">
        <td width="10" align="right"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
        <td background="/images/tbl_blue_round_02.gif"></td>
        <td width="10" align="left" ><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
	</tr>
	<tr height="25" valign="top">
        <td background="/images/tbl_blue_round_04.gif"></td>
        <td>
        	<img src="/images/icon_star.gif" align="absbottom">&nbsp;<b>����ǰ �����ϱ�</b>
        </td>
        <td background="/images/tbl_blue_round_05.gif"></td>
	</tr>
</table>
<!-- ǥ ��ܹ� ��-->

<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
    <form name="frm" method="post" onSubmit="return false;">
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td width="50">���<br />����</td>
		<td>����Ʈ<br />�ڵ�</td>
		<td>�̺�Ʈ<br />�ڵ�</td>
		<td>�̺�Ʈ��</td>
		<td width="70">�̺�Ʈ<br />������</td>
		<td width="70">�̺�Ʈ<br />������</td>
		<td>���� ���</td>
		<td>���� ����</td>
		<td>����ǰ</td>
		<td>����</td>
		<td>��������</td>
		<td>�����<br />����ǰ����<br />��������</td>
		<td>���</td>
	</tr>
	<%
	for i=0 to oGift.FResultCount -1
		'// 1:��ü��
		'// 2:�̺�Ʈ ��ϻ�ǰ���Ű�
		'// 3:Ư�� �귣���ǰ ���Ű�
		'// 4:�̺�Ʈ �׷��ǰ ���Ű�
		'// 5:Ư����ǰ ���Ű�
		'// 9:���̾ ����ǰ ���Ű�
		IsGiftOK = False
		if (oGift.FItemList(i).Fgift_scope <> "1") and (oGift.FItemList(i).Fgift_scope <> "2") and (oGift.FItemList(i).Fgift_scope <> "3") and (oGift.FItemList(i).Fgift_scope <> "4") and (oGift.FItemList(i).Fgift_scope <> "5") and (oGift.FItemList(i).Fgift_scope <> "9") then
			response.write "�ý��� ����"
			response.end
		end if

		sqlStr = " exec [db_order].[dbo].[sp_Ten_order_gift_chkValid_CS] '" & orderserial & "', " & oGift.FItemList(i).Fgift_scope & ", " & oGift.FItemList(i).Fgift_code & ", '" & itemlist & "' "
        if (orderserial = "20100547132") then
            response.write "<!-- " & sqlStr & " -->"
        end if
		rsget.Open sqlStr,dbget,1
		if Not rsget.Eof then
			if (rsget("validYN") = 1) then
				IsGiftOK = True
			end if
		end if
		rsget.close

		set oGiftModi = Nothing
		set oGiftModi = new COrderGift
		if oGift.FItemList(i).Fevt_code <> 0 then
			oGiftModi.FRectOrderSerial = orderserial
			oGiftModi.FRectGiftScope = oGift.FItemList(i).Fgift_scope
			oGiftModi.FRectGiftCode = oGift.FItemList(i).Fgift_code
			oGiftModi.FRectEvtCode = oGift.FItemList(i).Fevt_code
			oGiftModi.FRectItemListArr = itemlist
			oGiftModi.GetOneOrderValidGiftlist
		end if
	%>
	<tr height="60" align="center" bgcolor="#FFFFFF">
		<td>
			<% if oGift.FItemList(i).Fisupchebeasong="Y" then %>
			<font color="red">��ü</font>
			<% else %>
			<font color="blue">�ٹ�</font>
			<% end if %>
		</td>
		<td><%= oGift.FItemList(i).Fgift_code %></td>
		<td><%= oGift.FItemList(i).Fevt_code %></td>
		<td>
			<% if (oGift.FItemList(i).Fevt_code<>0) then %>
			<a target="_blank" href="http://www.10x10.co.kr/event/eventmain.asp?eventid=<%= oGift.FItemList(i).Fevt_code %>"><font color="blue"><%= oGift.FItemList(i).Fevt_name %></font></a>
			<% end if %>
		</td>
		<td><%= oGift.FItemList(i).Fevt_startdate %></td>
		<td><%= oGift.FItemList(i).Fevt_enddate %></td>
		<td>
			<%
			select case oGift.FItemList(i).Fgift_scope
				case "1"
					response.write "��ü��"
				case "2"
					response.write "�̺�Ʈ<br />��ϻ�ǰ<br />���Ű�"
				case "3"
					response.write "Ư��<br />�귣���ǰ<br />���Ű�"
				case "4"
					response.write "�̺�Ʈ<br />�׷��ǰ<br />���Ű�"
				case "5"
					response.write "Ư����ǰ<br />���Ű�"
				case "9"
					response.write "���̾<br />����ǰ<br />���Ű�"
				case else
					response.write "ERR"
			end select
			%>
		</td>
		<td>
			<%
			select case oGift.FItemList(i).Fgift_type
				case "1"
					response.write "����"
				case "2"
					response.write "���űݾ�<br />"
					if (oGift.FItemList(i).Fgift_range2 = 0) then
						response.write CStr(FormatNumber(oGift.FItemList(i).Fgift_range1,0)) + " �� �̻�"
					else
						response.write CStr(FormatNumber(oGift.FItemList(i).Fgift_range1,0)) + "~" + CStr(FormatNumber(oGift.FItemList(i).Fgift_range2,0)) + " ��"
					end if
				case "3"
					response.write "���ż���<br />"
					if (oGift.FItemList(i).Fgift_range2 = 0) then
						response.write CStr(oGift.FItemList(i).Fgift_range1) + " �� �̻�"
					else
						response.write CStr(oGift.FItemList(i).Fgift_range1) + "~" + CStr(oGift.FItemList(i).Fgift_range2) + " ��"
					end if
				case else
					response.write "ERR"
			end select
			%>
		</td>
		<td>
			<%
			if Not IsNull(oGift.FItemList(i).Fchg_giftSTR) then
				if oGift.FItemList(i).Fchg_giftSTR <> "" then
					response.write oGift.FItemList(i).Fchg_giftSTR
				else
					response.write oGift.FItemList(i).getGiftName()
				end if
			else
				response.write oGift.FItemList(i).getGiftName()
			end if
			%>
		</td>
		<td>
			<%= oGift.FItemList(i).Fgiftkind_cnt %> ��
			<%
			select case oGift.FItemList(i).Fgiftkind_type
				case "2"
					response.write "<br />[1+1]"
				case "3"
					response.write "<br />[1:1]"
				case else
					'//
			end select
			%>
		</td>
		<td>
			<%
			if (oGift.FItemList(i).Fgiftkind_limit <> 0) and ((oGift.FItemList(i).Fgiftkind_limit - oGift.FItemList(i).Fgiftkind_givecnt) <= 100) then
				response.write (oGift.FItemList(i).Fgiftkind_limit - oGift.FItemList(i).Fgiftkind_givecnt) & " / " & oGift.FItemList(i).Fgiftkind_limit
			end if
			%>
		</td>
		<td>
			<%= CHKIIF(IsGiftOK=True, "����", "<font color='red'>��������</font>") %>
		</td>
		<td>
			<% if (IsGiftOK<>True) then %>
			<input type="button" class="button" value="�����ϱ�" onclick="jsDelGift(frmAct, <%= oGift.FItemList(i).Fgift_code %>);">
			<% end if %>
		</td>
	</tr>
	<% if (oGiftModi.FResultCount>0) then %>
	<% for j=0 to oGiftModi.FResultCount - 1 %>
	<tr height="60" align="center" bgcolor="#FFFFFF">
		<td></td>
		<td></td>
		<td></td>
		<td></td>
		<td></td>
		<td></td>
		<td></td>
		<td>
			<%
			select case oGiftModi.FItemList(j).Fgift_type
				case "1"
					response.write "����"
				case "2"
					response.write "���űݾ�<br />"
					if (oGiftModi.FItemList(j).Fgift_range2 = 0) then
						response.write CStr(FormatNumber(oGiftModi.FItemList(j).Fgift_range1,0)) + " �� �̻�"
					else
						response.write CStr(FormatNumber(oGiftModi.FItemList(j).Fgift_range1,0)) + "~" + CStr(FormatNumber(oGiftModi.FItemList(j).Fgift_range2,0)) + " ��"
					end if
				case "3"
					response.write "���ż���<br />"
					if (oGiftModi.FItemList(j).Fgift_range2 = 0) then
						response.write CStr(oGiftModi.FItemList(j).Fgift_range1) + " �� �̻�"
					else
						response.write CStr(oGiftModi.FItemList(j).Fgift_range1) + "~" + CStr(oGiftModi.FItemList(j).Fgift_range2) + " ��"
					end if
				case else
					response.write "ERR"
			end select
			%>
		</td>
		<td>
			<%= oGiftModi.FItemList(j).Fgiftkind_name %>
		</td>
		<td></td>
		<td></td>
		<td><%= oGiftModi.FItemList(j).FvalidStr %></td>
		<td>
			<% if (oGiftModi.FItemList(j).FvalidStr = "OK") then %>
			<input type="button" class="button" value="�����ϱ�" onclick="jsModiGift(frmAct, <%= oGift.FItemList(i).Fgift_code %>, <%= oGiftModi.FItemList(j).Fgift_code %>, <%= oGiftModi.FItemList(j).Fgiftkind_code %>);">
			<% end if %>
		</td>
	</tr>
	<% next %>
	<% end if %>
	<% next %>
</table>

<!-- ǥ �ϴܹ� ����-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
    <tr valign="bottom" height="25">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td valign="bottom" align="center">
        </td>
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
    </tr>
    <tr valign="top" height="10">
        <td width="10" align="right"><img src="/images/tbl_blue_round_07.gif" width="10" height="10"></td>
        <td background="/images/tbl_blue_round_08.gif"></td>
        <td width="10" align="left"><img src="/images/tbl_blue_round_09.gif" width="10" height="10"></td>
    </tr>
    </form>
</table>
<!-- ǥ �ϴܹ� ��-->

<form name="frmAct" method="post" onSubmit="return false;" action="pop_cs_gift_modify_process.asp">
	<input type="hidden" name="mode" value="">
	<input type="hidden" name="orderserial" value="<%= orderserial %>">
	<input type="hidden" name="gift_code" value="">
	<input type="hidden" name="modi_gift_code" value="">
	<input type="hidden" name="modi_giftkind_code" value="">
	<input type="hidden" name="itemlist" value="<%= itemlist %>">
</form>

<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
