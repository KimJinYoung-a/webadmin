<%@ language=vbscript %>
<% option explicit %>
<%
session.codePage = 949
Response.CharSet = "EUC-KR"
%>
<%
'###########################################################
' Description : cs���� �ɼǱ�ȯ
' History : �̻� ����
'			2023.09.05 �ѿ�� ����(6�������� �ֹ��� ��ȯ �����ϰ� ó��)
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbHelper.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/order/jumuncls.asp"-->
<!-- #include virtual="/lib/classes/order/new_ordercls.asp"-->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/classes/cscenter/cs_couponcls.asp" -->
<!-- #include virtual="/lib/classes/cscenter/sp_itemcouponcls.asp" -->
<!-- #include virtual="/cscenter/lib/csOrderFunction.asp" -->
<%
dim i, idx, orderserial, result, ojumunDetail, ojumun, IsOrderCanceled, IsChangeOrder
dim itemoption, optionname, optsellyn, optlimityn, optlimitno, optlimitsold, issameoptaddprice, isusing, optaddprice
Dim sqlStr, rsOption, k, optionText, itemStatus, sqlsub, changedindex
dim prevregno, contents_jupsu, title, divcd, oupchebeasongpay, upchebeasongpay, isupchebeasong, requiremakerid
	idx = requestCheckVar(getNumeric(request("idx")),10)

set ojumunDetail = new CJumunMaster
ojumunDetail.SearchOneJumunDetail idx

if (ojumunDetail.FResultCount < 1) then
	ojumunDetail.FRectOldJumun = "on"
	ojumunDetail.SearchOneJumunDetail idx
end if

orderserial = ojumunDetail.FJumunDetail.FOrderSerial

set ojumun = new COrderMaster

if (orderserial <> "") then
    ojumun.FRectOrderSerial = orderserial
    ojumun.QuickSearchOrderMaster
end if

if (ojumun.FResultCount<1) and (Len(orderserial)=11) and (IsNumeric(orderserial)) then
    ojumun.FRectOldOrder = "on"
    ojumun.QuickSearchOrderMaster
end if

if ojumun.FTotalCount < 1 then
	response.write "�ش�Ǵ� �ֹ����� �����ϴ�."
	dbget.close() : response.end
end if

IsOrderCanceled = (ojumun.FOneItem.Fcancelyn = "Y")
IsChangeOrder   = (ojumun.FOneItem.FjumunDiv="6")

If ojumunDetail.FJumunDetail.Fitemoption <> "0000" Then
	'* �ɼǺ����� <font color=red>�ɼǰ�</font>�� ������ �ɼǻ�ǰ�� �����մϴ�.<br>
	'* �ֹ���� �ɼǰ��ݿ� ������� ���� ��ǰ���� ���� �ɼǰ������� ���մϴ�.<br>
	'* ��ǰ��������(�ǸŰ�,���԰� ��)�� �ֹ���� ������ ����˴ϴ�.<br>
	' �ֹ��� ������ ó���� �Ǿ ǥ��

	sqlsub = "select top 1 optaddprice "
	sqlsub = sqlsub + "from [db_item].[dbo].tbl_item_option "
	sqlsub = sqlsub + "where 1 = 1 "
	sqlsub = sqlsub + "and itemid = " & CStr(ojumunDetail.FJumunDetail.Fitemid) & " "
	sqlsub = sqlsub + "and itemoption = '" & CStr(ojumunDetail.FJumunDetail.Fitemoption) & "' "

	sqlStr = " select "
	sqlStr = sqlStr + " v.itemoption "
	sqlStr = sqlStr + " , v.optionname "
	sqlStr = sqlStr + " , v.optsellyn "
	sqlStr = sqlStr + " , v.optlimityn "
	sqlStr = sqlStr + " , v.optlimitno "
	sqlStr = sqlStr + " , v.optlimitsold "
	sqlStr = sqlStr + " , case when v.optaddprice=IsNULL((" & sqlsub & "),0) " & " then 'T' else 'F' end "
	sqlStr = sqlStr + " , v.isusing "
	sqlStr = sqlStr + " , IsNull(P.regno, 0) as prevregno "
	sqlStr = sqlStr + " from [db_item].[dbo].tbl_item i "
	sqlStr = sqlStr + " left join [db_item].[dbo].tbl_item_option v "
	sqlStr = sqlStr + " on i.itemid=v.itemid "

	'���� CS��ǰ����(����+�Ϸ᳻��, ��ǰ�����������)
	sqlStr = sqlStr + "		LEFT JOIN (" + VbCrlf
	sqlStr = sqlStr + "		    select d.itemid, d.itemoption, sum(confirmitemno) as regno, max(a.id) asId " + VbCrlf
    sqlStr = sqlStr + "		    from" + VbCrlf
    sqlStr = sqlStr + "		    	[db_cs].[dbo].tbl_new_as_list a" + VbCrlf
    sqlStr = sqlStr + "		    	Join [db_cs].[dbo].tbl_new_as_detail d" + VbCrlf
    sqlStr = sqlStr + "		    on a.id=d.masterid" + VbCrlf
    sqlStr = sqlStr + "		    where a.orderserial='" + CStr(orderserial) + "'" + VbCrlf
    sqlStr = sqlStr + "		    and a.divcd in ('A004','A010', 'A111', 'A112')" + VbCrlf                ''��ǰ / ȸ�� / ��ǰ���� �±�ȯȸ��(�ٹ����ٹ��) / ��ǰ���� �±�ȯ��ǰ(��ü���).
    sqlStr = sqlStr + "		    and a.deleteyn='N'" + VbCrlf
    'sqlStr = sqlStr + "		    	and a.currstate='B007'" + VbCrlf					'����+�Ϸ� ��� ���
    sqlStr = sqlStr + "			group by d.itemid, d.itemoption" + VbCrlf
    sqlStr = sqlStr + " ) P " + VbCrlf
    sqlStr = sqlStr + "     ON i.itemid=P.itemid and v.itemoption=P.itemoption" + VbCrlf

	sqlStr = sqlStr + " WHERE 1=1 "
	sqlStr = sqlStr + " and i.itemid=" & ojumunDetail.FJumunDetail.Fitemid & ""
	sqlStr = sqlStr + " order by i.itemid desc, v.itemoption"

	rsget.Open sqlStr,dbget,1
	If Not rsget.EOF Then
		rsOption = rsget.getrows
	End If
	rsget.close()

	'response.write sqlStr
End If

prevregno = 0

'// �⺻���� ����
if Not IsNull(session("ssBctCname")) then
	contents_jupsu = "�ٹ����� ������ " + CStr(session("ssBctCname")) + " �Դϴ�"
end if

if (ojumunDetail.FJumunDetail.FcurrState = "7") or IsChangeOrder then

	'==============================================================================
	'�������

	isupchebeasong = ojumunDetail.FJumunDetail.Fisupchebeasong

	if (isupchebeasong = "Y") then
		requiremakerid = ojumunDetail.FJumunDetail.Fmakerid
	end if

	divCd = "A100"	' ��ǰ���� �±�ȯ���
	title = "��ȯ���(�ɼǺ���)"

	'// �ɼǺ��� �±�ȯ�� ��� ������ǰ����
	For i = 0 To UBound(rsOption,2)
		itemoption = rsOption(0,i)
		prevregno = rsOption(8,i)

		if (ojumunDetail.FJumunDetail.Fitemoption = itemoption) then
			Exit For
		end if
	Next

	set oupchebeasongpay = new COrderMaster
	upchebeasongpay = getDefaultBeasongPayByDate(Left(Now, 10))		' ��ۺ�

	if (orderserial <> "") and (isupchebeasong = "Y") then
		oupchebeasongpay.FRectOrderSerial = orderserial
		oupchebeasongpay.getUpcheBeasongPayList

		for i = 0 to oupchebeasongpay.FResultCount - 1
			if (oupchebeasongpay.FItemList(i).Fmakerid = requiremakerid) then
				'// ��ü����̸� ��ü �⺻��ۺ� ��������
				upchebeasongpay = oupchebeasongpay.FItemList(i).Fdefaultdeliverpay
			end if
		next

		if (upchebeasongpay = 0) then
			'// XXXX ��ü�������̸� ���ٹ�ۺ�� ����
			'�⺻��ۺ� ���� �ʵǾ� ������ 2500��(since 2012-06-18)
			upchebeasongpay = 2500
		end if
	end if

else
	'==============================================================================
	' ��� ����

	divCd = "A900"	' �ֹ���������
	title = "��ǰ�ɼǺ���"
end if

%>
<script type="text/javascript" src="/cscenter/js/cscenter.js"></script>
<script type="text/javascript" SRC="/js/ajax.js"></script>
<script type="text/javascript" SRC="/cscenter/js/newcsas.js"></script>
<script type='text/javascript'>

// ��������(ajax) �� ����ϱ� ���� �ʿ�
var IsPossibleModifyCSMaster = true;
var IsPossibleModifyItemList = true;

// ============================================================================
// �ɼǺ� ���� �ڵ�����(���̳ʽ� �ԷºҰ�)
// ============================================================================
function CheckItemOptionNo(changedindex) {
    var frm = document.frm;
    var i;

	var orgitemno = parseInt(frm.orgitemoptionno.value);
	var regitemno = 0;

	if (frm.itemoptionno[changedindex].value*1 < 0) {
		alert('������ ���̳ʽ��� �Է��� �� �����ϴ�.');
		return;
	}

	if ((frm.itemoptionno[changedindex].value.length < 1) || (frm.itemoptionno[changedindex].value*0 != 0)) {
		alert('������ ���ڸ� �Է��ϼ���.');
		return;
	}

	for (i = 1; i < parseInt(frm.itemoptionno.length); i++) {
		regitemno = regitemno + parseInt(frm.itemoptionno[i].value);
	}

	if (regitemno > orgitemno) {
		alert('���氡���� ������ �ʰ��Ͽ����ϴ�.');
		frm.itemoptionno[changedindex].value = frm.itemoptionno[changedindex].value - (regitemno - orgitemno);
		regitemno = orgitemno;
	}

	frm.itemoptionno[0].value = orgitemno - regitemno;
}

function SetAddBeasongPay() {
    var frm = document.frm;

	if (!frm.isupchebeasong) {
		return;
	}

	if ((frm.gubun01.value == "C004") && (frm.gubun02.value == "CD01")) {
		// �ܼ�����
		frm.add_customeraddbeasongpay.value = frm.upchebeasongpay.value*2;
		frm.add_customeraddmethod.value = "1";
	} else {
		frm.add_customeraddbeasongpay.value = 0;
		frm.add_customeraddmethod.value = "";
	}
}

// ============================================================================
// �ɼǺ��� �ֹ�����
// ============================================================================
function SaveItemOptionNo() {
    var frm = document.frm;

	var orgitemno = parseInt(frm.orgitemoptionno.value);
	var remainitemno = parseInt(frm.itemoptionno[0].value);

	var isupchebeasong = "<%= ojumunDetail.FJumunDetail.Fisupchebeasong %>";
	var itemstate = "<%= ojumunDetail.FJumunDetail.FcurrState %>";

	if (<%= LCase(IsChangeOrder) %> == true) {
		alert('��ȯ�ֹ��� ��ǰ���� �� �� �����ϴ�.');
		return;
	}

	if (orgitemno == remainitemno) {
		alert('������ ������ 0�Դϴ�.');
		return;
	}

	if (frm.gubun01.value == "") {
		alert("���������� �����ϼ���.");
		return;
	}

	if (frm.title.value == "") {
		alert("������ �Է��ϼ���.");
		return;
	}

	if ((isupchebeasong == "Y") && (itemstate >= "3") && (itemstate < "7")) {
		// ��ü���, ��ǰ�غ� ����
		if (confirm('��ü����̸鼭 ��ǰ�غ� �����Դϴ�.\n\���� �Ͻðڽ��ϱ�?') != true) {
			return;
		}
	}

	if (itemstate == "7") {
		// ��ǰ��� ����
		alert('��ǰ��� �����Դϴ�. ����� �� �����ϴ�.');
		return;
	}


	if (confirm('���� �Ͻðڽ��ϱ�?')) {
		frm.mode.value="EditItemNoPart";
		frm.submit();
	}
}

// ============================================================================
// �ɼǺ��� �±�ȯ
// ============================================================================
function SaveChangeItemOptionNo(){
    var frm = document.frm;

	<% 'if (ojumun.FRectOldOrder = "on") then %>
		//alert("6���� �����ֹ� ó���Ұ�!!");
		//return;
	<% 'end if %>

	var orgitemno = parseInt(frm.orgitemoptionno.value);
	var remainitemno = parseInt(frm.itemoptionno[0].value);

	var itemstate = "<%= ojumunDetail.FJumunDetail.FcurrState %>";

	if (orgitemno < 1) {
		alert('���ֹ� ������ �����ϴ�.');
		return;
	}

	if (orgitemno == remainitemno) {
		alert('������ ������ 0�Դϴ�.');
		return;
	}

	if (frm.gubun01.value == "") {
		alert("���������� �����ϼ���.");
		return;
	}

	if (frm.title.value == "") {
		alert("������ �Է��ϼ���.");
		return;
	}

	if ((itemstate < "7") && (<%= LCase(IsChangeOrder) %> != true)) {
		// ��ǰ��� ����
		alert('��ǰ��� ���� ��ǰ�Դϴ�. ��ȯ���(�ɼǺ���)�� �� �����ϴ�.');
		return;
	}

	if (confirm('��ȯ(�ɼǺ���) �����Ͻðڽ��ϱ�?')){
		frm.mode.value="ChangeEditItemNoPart";
		frm.submit();
	}
}

</script>

<form name="frm" method="post" action="/cscenter/ordermaster/orderdetail_simple_editoption_process.asp" style="margin:0px;">
<input type="hidden" name="mode" value="">
<input type="hidden" name="detailidx" value="<%= ojumunDetail.FJumunDetail.Fdetailidx %>">
<input type="hidden" name="itemid" value="<%= ojumunDetail.FJumunDetail.Fitemid %>">
<input type="hidden" name="orderserial" value="<%= ojumunDetail.FJumunDetail.FOrderSerial %>">
<input type="hidden" name="divcd" value="<%= divcd %>">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="15">
			<img src="/images/icon_star.gif" align="absbottom">&nbsp;<b>�ֹ����������� ����</b>
		</td>
	</tr>
	<tr bgcolor="#FFFFFF" height="25">
		<td width="100" bgcolor="<%= adminColor("tabletop") %>">IDX</td>
		<td><%= ojumunDetail.FJumunDetail.Fdetailidx %></td>
		<td width="110" rowspan="4"><img src="<%= ojumunDetail.FJumunDetail.FImageList %>"></td>
	</tr>
	<tr bgcolor="#FFFFFF" height="25">
		<td bgcolor="<%= adminColor("tabletop") %>">�귣�� ID</td>
		<td><%= ojumunDetail.FJumunDetail.Fmakerid %></td>
	</tr>
	<tr bgcolor="#FFFFFF" height="25">
		<td bgcolor="<%= adminColor("tabletop") %>">��ǰ�ڵ�</td>
		<td><%= ojumunDetail.FJumunDetail.Fitemid %></td>
	</tr>
	<tr bgcolor="#FFFFFF" height="25">
		<td bgcolor="<%= adminColor("tabletop") %>">��ǰ��</td>
		<td><%= ojumunDetail.FJumunDetail.Fitemname %></td>
	</tr>
	<tr bgcolor="#FFFFFF" height="25">
		<td bgcolor="<%= adminColor("tabletop") %>">��һ���</td>
		<td><%= ojumunDetail.FJumunDetail.Fcancelyn %></td>
		<td>

		</td>
	</tr>
</table>

<br>

<%

if ojumunDetail.FJumunDetail.Fitemoption = "0000" Then
	response.write "�ɼ��� �����ϴ�. ����� �� �����ϴ�."
	dbget.close() : response.end
end If

%>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr bgcolor="#FFFFFF">
		<td width="100" bgcolor="<%= adminColor("tabletop") %>">
			����ɼ�
		</td>
		<td>
			<%= "[" & ojumunDetail.FJumunDetail.FitemOption & "] " & ojumunDetail.FJumunDetail.FitemOptionName %>
		</td>
		<input type=hidden name=orgitemoptionno value="<%= ojumunDetail.FJumunDetail.Fitemno - prevregno %>">
		<input type=hidden name=itemoptioncode value="<%= ojumunDetail.FJumunDetail.FitemOption %>">
		<td>
			<input type="text" class="text_ro" name="itemoptionno" value="<%= (ojumunDetail.FJumunDetail.Fitemno - prevregno) %>" size="3" maxlength="9" readonly> ��(<%= ojumunDetail.FJumunDetail.Fitemno %>��)

			<% if (prevregno <> 0) then %>
				<font color=red>(������ǰ : <%= prevregno %> ��)</font>
			<% end if %>
		</td>
	</tr>
	<%
	changedindex = 0
	%>
	<% For i = 0 To UBound(rsOption,2) %>
		<%
		itemoption 			= rsOption(0,i)
		optionname 			= rsOption(1,i)
		optsellyn 			= rsOption(2,i)
		optlimityn 			= rsOption(3,i)
		optlimitno 			= rsOption(4,i)
		optlimitsold 		= rsOption(5,i)
		issameoptaddprice 	= rsOption(6,i)
		isusing 			= rsOption(7,i)

		itemStatus = ""

		if (itemoption < ojumunDetail.FJumunDetail.Fitemoption) then
			changedindex = i + 1
		else
			changedindex = i
		end if

		%>
		<% if (itemoption = ojumunDetail.FJumunDetail.Fitemoption) then %>
			<!-- ����ɼ� ��ŵ -->
		<% else %>
			<%

			if (optsellyn = "N") then
				itemStatus = itemStatus + "<font color=red>�Ǹž���</font>,"
			end if

			if (optlimityn = "Y") then
				if ((optlimitno - optlimitsold) < 1) then
					itemStatus = itemStatus + "<font color=red>����:0</font>,"
				else
					itemStatus = itemStatus + "����:" & ( optlimitno - optlimitsold ) & ","
				end if
			end if

			if (isusing = "N") then
				itemStatus = itemStatus + "<font color=red>������</font>,"
			end if

			If itemStatus <> "" Then
				itemStatus = " ( " & Mid(itemStatus, 1, Len(itemStatus) - 1) & " )"
			End If

			optionText = "[" & itemoption & "] " & optionname

			%>
			<% if (issameoptaddprice = "F") then %>
			<tr bgcolor="#FFFFFF">
				<td width="100" bgcolor="<%= adminColor("tabletop") %>">
					����Ұ��ɼ�(<%= changedindex %>)
				</td>
				<td>
					<%=optionText%><font color=red>(�ɼǰ� �ٸ�)</font>
				</td>
				<input type=hidden name=itemoptioncode value="<%= itemoption %>">
				<td>
					<input type="text" class="text_ro" name="itemoptionno" value="0" size="3" maxlength="9" onKeyUp="CheckItemOptionNo(<%= changedindex %>)" readonly> ��
					<%= itemStatus %>
				</td>
			</tr>
			<% else %>
			<tr bgcolor="#FFFFFF">
				<td width="100" bgcolor="<%= adminColor("tabletop") %>">
					���氡�ɿɼ�(<%= changedindex %>)
				</td>
				<td>
					<%=optionText%>
				</td>
				<input type=hidden name=itemoptioncode value="<%= itemoption %>">
				<td>
					<input type="text" class="text" name="itemoptionno" value="0" size="3" maxlength="9" onKeyUp="CheckItemOptionNo(<%= changedindex %>)"> ��
					<%= itemStatus %>
				</td>
			</tr>
			<% end if %>
		<% end If %>
	<% Next %>

	<tr bgcolor="#FFFFFF">
		<td width="100" height=30 bgcolor="<%= adminColor("tabletop") %>">
			��������
		</td>
		<td colspan=2>
                <input type="hidden" name="gubun01" value="">
                <input type="hidden" name="gubun02" value="">
                <input class="text_ro" type="text" name="gubun01name" value="" size="16" Readonly >
                &gt;
                <input class="text_ro" type="text" name="gubun02name" value="" size="16" Readonly >
                <input class="csbutton" type="button" value="����" onClick="divCsAsGubunSelect(frm.gubun01.value, frm.gubun02.value, frm.gubun01.name, frm.gubun02.name, frm.gubun01name.name, frm.gubun02name.name,'frm','causepop');">
                <div id="causepop" style="position:absolute;"></div>

                <!-- �Ϻ� ���� �̸� ǥ�� -->
                <%
                '��������
				'select top 100 m.comm_cd, m.comm_name, d.comm_cd, d.comm_name
				'from
				'	db_cs.dbo.tbl_cs_comm_code m
				'	left join db_cs.dbo.tbl_cs_comm_code d
				'	on
				'		m.comm_cd = d.comm_group
				'where
				'	1 = 1
				'	and m.comm_group = 'Z020'
				'	and m.comm_isdel <> 'Y'
				'	and d.comm_isdel <> 'Y'
				'order by m.comm_cd, d.comm_cd
                %>
                [<a href="javascript:selectGubun('C004','CD01','����','�ܼ�����','gubun01','gubun02','gubun01name','gubun02name','frm','causepop'); SetAddBeasongPay();">�ܼ�����</a>]
                [<a href="javascript:selectGubun('C004','CD05','����','ǰ��','gubun01','gubun02','gubun01name','gubun02name','frm','causepop'); SetAddBeasongPay();">ǰ��</a>]
                [<a href="javascript:selectGubun('C005','CE01','��ǰ����','��ǰ�ҷ�','gubun01','gubun02','gubun01name','gubun02name','frm','causepop'); SetAddBeasongPay();">��ǰ�ҷ�</a>]
                [<a href="javascript:selectGubun('C004','CD99','����','��Ÿ','gubun01','gubun02','gubun01name','gubun02name','frm','causepop'); SetAddBeasongPay();">��Ÿ</a>]
            	&nbsp; &nbsp; &nbsp;
            	<div id="chkmodifyitemstockoutyn" style="display: inline;"><input type="checkbox" name="modifyitemstockoutyn" value="Y" checked> ǰ������ ����(�����ǰ)</div>
		</td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td width="100" height=30 bgcolor="<%= adminColor("tabletop") %>">
			��������
		</td>
		<td colspan=2>
                <input class='text' type="text" name="title" value="<%= title %>" size="56" maxlength="56">
		</td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td width="100" height=30 bgcolor="<%= adminColor("tabletop") %>">
			��������
		</td>
		<td colspan=2>
                <textarea class='textarea' name="contents_jupsu" cols="68" rows="6"><%= contents_jupsu %></textarea>
		</td>
	</tr>

	<% if (divcd = "A100") then %>
		<input type="hidden" name="isupchebeasong" value="<%= isupchebeasong %>">
		<input type="hidden" name="requiremakerid" value="<%= requiremakerid %>">
		<input type="hidden" name="upchebeasongpay" value="<%= upchebeasongpay %>">
		<tr bgcolor="#FFFFFF">
			<td width="100" height=30 bgcolor="<%= adminColor("tabletop") %>">
				��۱���
			</td>
			<td colspan=2>
		    	<% if (isupchebeasong = "Y") then %>
		    		<font color=red><%= requiremakerid %></font> (�⺻��ۺ� : <%= FormatNumber(upchebeasongpay, 0) %>��)
		    	<% else %>
		    		�ٹ����ٹ��
		    	<% end if %>
			</td>
		</tr>
		<tr bgcolor="#FFFFFF">
			<td width="100" height=30 bgcolor="<%= adminColor("tabletop") %>">
				�߰���ۺ�
			</td>
			<td colspan=2>
		    	<input type="text" class="text" name="add_customeraddbeasongpay" value="0" size="20">
		    	&nbsp;
	    	    <select class="select" name="add_customeraddmethod" class="text">
		    	    <option value="">����
		    	    <option value="1">�ڽ�����
		    	    <option value="2">�ù�� ���δ�
		    	    <option value="5">��Ÿ
	    	    </select>
			</td>
		</tr>
	<% end if %>

	<tr bgcolor="#FFFFFF" height=35>
		<td colspan="3" align="center">
			<% if Not IsOrderCanceled then %>
				<input type="button" class="button" value="�ɼǺ���" onclick="javascript:SaveItemOptionNo()" <% if (ojumunDetail.FJumunDetail.FcurrState = "7") then %>disabled<% end if %>>
				<input type="button" class="button" value="�ɼǺ��� �±�ȯ" onclick="javascript:SaveChangeItemOptionNo()" <% if (ojumunDetail.FJumunDetail.FcurrState <> "7") and Not IsChangeOrder then %>disabled<% end if %>>
			<% else %>
				��ҵ� ��ǰ�� �������� �Ұ�
			<% end if %>
		</td>
	</tr>
</table>
</form>
<div>
* �ɼǺ����� <font color=red>�ɼǰ�</font>�� ������ �ɼǻ�ǰ�� �����մϴ�.<br>
* �ֹ���� �ɼǰ��ݿ� ������� ���� ��ǰ���� ���� �ɼǰ������� ���մϴ�.<br>
* ��ǰ��������(�ǸŰ�,���԰� ��)�� �ֹ���� ������ ����˴ϴ�.<br>
* <font color=red>�߰��� ��ǰ�� ���Ϸ�</font> �����̸� �ֹ����� �Ұ��մϴ�.<br>
</div>
<%
set ojumun       = Nothing
set ojumunDetail = Nothing
%>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
