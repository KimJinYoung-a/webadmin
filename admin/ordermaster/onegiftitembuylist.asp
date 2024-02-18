<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/items/new_itemcls.asp"-->
<!-- #include virtual="/lib/classes/offshop/stock/offitemstock_cls.asp"-->
<!-- #include virtual="/lib/classes/maechul/managementSupport/maechulCls.asp" -->
<%

function IpkumDivName(byval v )
	if v="0" then
		IpkumDivName="�ֹ����"
	elseif v="1" then
		IpkumDivName="�ֹ�����"
	elseif v="2" then
		IpkumDivName="�ֹ�����"
	elseif v="3" then
		IpkumDivName="�ֹ�����"
	elseif v="4" then
		IpkumDivName="�����Ϸ�"
	elseif v="5" then
		IpkumDivName="�ֹ��뺸"
	elseif v="6" then
		IpkumDivName="��ǰ�غ�"
	elseif v="7" then
		IpkumDivName="�Ϻ����"
	elseif v="8" then
		IpkumDivName="���Ϸ�"
	elseif v="9" then
		IpkumDivName="���̳ʽ�"
	end if
end function

function getCurrstateName(byval v1, byval v)
    if (v=0) then
        if (v1>3) and (v1<8) then
           getCurrstateName = "�����Ϸ�"
        else
            getCurrstateName = IpkumDivName(v1)
        end if
    else
        if v=2 then
            getCurrstateName = "�ֹ��뺸"
        elseif v=3 then
            getCurrstateName = "��ǰ�غ�"
        elseif v=7 then
            getCurrstateName = "���Ϸ�"
        else
            getCurrstateName = v
        end if
    end if
end function

function getCurrstateNameColor(byval v1, byval v)
    if (v=0) then
        if (v1>3) and (v1<8) then
            getCurrstateNameColor = IpkumDivColor(4)
        else
            getCurrstateNameColor = IpkumDivName(v1)
        end if
    else
        if v=2 then
            getCurrstateNameColor = IpkumDivColor(v)
        elseif v=3 then
            getCurrstateNameColor = IpkumDivColor(v)
        elseif v=7 then
            getCurrstateNameColor = IpkumDivColor(v)
        else
            getCurrstateNameColor = "#000000"
        end if
    end if
end function

function getJumundivName(byval ijumundiv)
    if (isNULL(ijumundiv)) then
        getJumundivName = ""
        Exit function
    end if

    if ijumundiv="1" then
		getJumundivName="���"
    elseif ijumundiv="9" then
        getJumundivName="<font color='red'>��ǰ</font>"
    elseif ijumundiv="6" then
        getJumundivName="<font color='blue'>��ȯ</font>"
    else
        getJumundivName=ijumundiv
    end if
end function

Const MaxRowSize = 1000
dim itemgubun, itemid, itemoption
dim itemstate
dim sitename

dim yyyy1,yyyy2,mm1,mm2,dd1,dd2
dim nowdate,searchnextdate,oldlist
dim premonthdate
dim datetype

nowdate         = Left(CStr(now()),10)
premonthdate    = DateAdd("d",-14,nowdate)

itemgubun = requestCheckvar(request("itemgubun"),10)
itemid = requestCheckvar(request("itemid"),10)
itemoption = requestCheckvar(request("itemoption"),10)
itemstate = requestCheckvar(request("itemstate"),10)
oldlist = request("oldlist")

yyyy1   = request("yyyy1")
mm1     = request("mm1")
dd1     = request("dd1")
yyyy2   = request("yyyy2")
mm2     = request("mm2")
dd2     = request("dd2")
datetype = request("datetype")
sitename = requestCheckvar(request("sitename"),32)

if (itemstate="5") then itemstate="6"
if (itemgubun = "") then itemgubun = "85"


if (yyyy1="") then
	yyyy1 = Left(premonthdate,4)
	mm1   = Mid(premonthdate,6,2)
	dd1   = Mid(premonthdate,9,2)

	nowdate = Left(CStr(now()),10)
	yyyy2 = Left(nowdate,4)
	mm2   = Mid(nowdate,6,2)
	dd2   = Mid(nowdate,9,2)
else
	nowdate = Left(CStr(DateSerial(yyyy1 , mm1 , dd1)),10)
	yyyy1 = Left(nowdate,4)
	mm1   = Mid(nowdate,6,2)
	dd1   = Mid(nowdate,9,2)
end if

searchnextdate = Left(CStr(DateAdd("d",DateSerial(yyyy2 , mm2 , dd2),1)),10)

if (datetype="") then datetype="reg"

'��ǰ�ڵ� ��ȿ�� �˻�(2008.08.05;������)
if itemid<>"" then
	if Not(isNumeric(itemid)) then
		Response.Write "<script language=javascript>alert('[" & itemid & "]��(��) ��ȿ�� ��ǰ�ڵ尡 �ƴմϴ�.');history.back();</script>"
		dbget.close()	:	response.End
	end if
end if

dim sqlStr, RowArr


sqlStr = " select top " & CStr(MaxRowSize)
sqlStr = sqlStr + " 	m.orderserial "
sqlStr = sqlStr + " 	, m.ipkumdiv "
sqlStr = sqlStr + " 	, g.giftkind_cnt as sm "
sqlStr = sqlStr + " 	, m.buyname "
sqlStr = sqlStr + " 	, m.buyemail "
sqlStr = sqlStr + " 	, m.buyhp "
sqlStr = sqlStr + " 	, m.buyphone "
sqlStr = sqlStr + " 	, m.reqname "
sqlStr = sqlStr + " 	, m.reqhp "
sqlStr = sqlStr + " 	, m.reqphone "
sqlStr = sqlStr + " 	, i.shopitemoptionname "
sqlStr = sqlStr + " 	, isnull(m.ipkumdiv, 0) as currstate "
sqlStr = sqlStr + " 	, m.sitename "
sqlStr = sqlStr + " 	, g.chulgodate as beasongdate "
sqlStr = sqlStr + " 	, m.userid "
sqlStr = sqlStr + " 	, m.jumundiv "
sqlStr = sqlStr + " 	, i.centermwdiv as omwdiv "

sqlStr = sqlStr + " from "
if oldlist="on" then
	sqlStr = sqlStr + " [db_log].[dbo].tbl_old_order_master_2003 m "
else
	sqlStr = sqlStr + " [db_order].[dbo].tbl_order_master m"
end if

sqlStr = sqlStr + " 	inner join db_order.dbo.tbl_order_gift g "
sqlStr = sqlStr + " 	on "
sqlStr = sqlStr + " 		m.orderserial = g.orderserial "
sqlStr = sqlStr + " 	inner join db_shop.dbo.tbl_shop_item i "
sqlStr = sqlStr + " 	on "
sqlStr = sqlStr + " 		1 = 1 "
sqlStr = sqlStr + " 		and g.prd_itemgubun = i.itemgubun "
sqlStr = sqlStr + " 		and g.prd_itemid = i.shopitemid "
sqlStr = sqlStr + " 		and g.prd_itemoption = i.itemoption "

sqlStr = sqlStr + " where "
sqlStr = sqlStr + " 	1 = 1 "
    if (datetype="ipkum") then
        sqlStr = sqlStr + " and m.ipkumdate >= '" + yyyy1 + "-" + mm1 + "-" + dd1 + "'"
        sqlStr = sqlStr + " and m.ipkumdate < '" + searchnextdate + "'"
    elseif (datetype="beasong") then
        sqlStr = sqlStr + " and g.chulgodate >= '" + yyyy1 + "-" + mm1 + "-" + dd1 + "'"
        sqlStr = sqlStr + " and g.chulgodate < '" + searchnextdate + "'"
    else
        sqlStr = sqlStr + " and m.regdate >= '" + yyyy1 + "-" + mm1 + "-" + dd1 + "'"
        sqlStr = sqlStr + " and m.regdate < '" + searchnextdate + "'"
    end if


sqlStr = sqlStr + " 	and m.cancelyn = 'N' "
sqlStr = sqlStr + " 	and m.ipkumdiv > 1 "
sqlStr = sqlStr + " 	and m.cancelyn <> 'Y' "
sqlStr = sqlStr + " 	and g.prd_itemgubun = '85' "		'// ON����ǰ��
sqlStr = sqlStr + " 	and g.prd_itemid = " + CStr(itemid)

if itemoption<>"" then
    sqlStr = sqlStr + " and i.itemoption='" + CStr(itemoption) + "'"
end if

if itemstate="2" then   '�ֹ�����
	sqlStr = sqlStr + " and m.ipkumdiv=2"
elseif itemstate="4" then	'�����Ϸ�
	sqlStr = sqlStr + " and m.ipkumdiv>=4 and m.ipkumdiv<8 "
elseif itemstate="8" then	'���Ϸ�
	sqlStr = sqlStr + " and g.chulgodate is not NULL "
elseif itemstate="ipkumfinishall" then	'�����Ϸ��̻�
	sqlStr = sqlStr + " and m.ipkumdiv>=4"
end if

if sitename <> "" then
	sqlStr = sqlStr + " and m.sitename = '" + CStr(sitename) + "' "
end if

sqlStr = sqlStr + " order by m.ipkumdiv , m.orderserial"
''response.write sqlStr
if (itemid<>"") then
    rsget.Open sqlStr,dbget,1
    if not rsget.Eof then
        RowArr = rsget.getRows
    end if
    rsget.Close
end if

dim RowCount, jumuncnt
RowCount = 0
jumuncnt = 0
if IsArray(RowArr) then
    RowCount = Ubound(RowArr,2)
    jumuncnt = RowCount + 1
end if

dim totno, i
totno = 0


dim oitem
set oitem = new CoffstockItemlist	'//�¶��� ��ũ������� Ŭ������ �浹, �������� ���� ����
	oitem.frectitemgubun = itemgubun
	oitem.FRectItemID = itemid
	oitem.frectitemoption = itemoption

	if itemid<>"" then
		oitem.GetoffItemDefaultData
	end if

dim oitemoption
set oitemoption = new CItemOption
oitemoption.FRectItemID = itemid
if itemid<>"" then
    oitemoption.FRectItemGubun = itemgubun
	oitemoption.GetItemOptionInfoByOffItemTable
end if
%>

<script language='javascript'>
function PopItemSellEdit(iitemid){
	var popwin = window.open('/admin/lib/popitemsellinfo.asp?itemid=' + iitemid,'itemselledit','width=500,height=600,scrollbars=yes,resizable=yes')
	popwin.focus();
}
</script>

<!-- �˻� ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frm" method="get" action="">
	<input type="hidden" name="page" value="1">
	<input type="hidden" name="menupos" value="<%= menupos %>">

	<tr height="25" bgcolor="#FFFFFF">
	    <td align="center" width="50" bgcolor="#EEEEEE">�˻�<br>����</td>
        <td>
        	������ID :
            <input type="text_ro" class="text" name="itemgubun" value="<%= itemgubun %>" size="2" maxlength="2" readonly>
            <input type="text" class="text" name="itemid" value="<%= itemid %>" size="11" maxlength="16">
            <% if oitemoption.FResultCount = 0 then %>
            <input type="text_ro" class="text" name="itemoption" value="<%= itemoption %>" size="4" maxlength="4" readonly>
            <% end if %>
            &nbsp;&nbsp;

        	<% if oitemoption.FResultCount>0 then %>
			&nbsp;
			�ɼǼ��� :
			<select class="select" name="itemoption">
				<option  value="">----
				<% for i=0 to oitemoption.FResultCount-1 %>
				<option value="<%= oitemoption.FITemList(i).FItemOption %>" <% if itemoption=oitemoption.FITemList(i).FItemOption then response.write "selected" %> ><%= oitemoption.FITemList(i).FOptionName %>
				<% next %>
				</select>
			<% end if %>

			&nbsp;
			�˻��Ⱓ :
			<select class="select" name="datetype">
			    <option value="reg" <%= chkIIF(datetype="reg","selected","") %> >�ֹ���
			    <option value="ipkum" <%= chkIIF(datetype="ipkum","selected","") %> >������
			    <option value="beasong" <%= chkIIF(datetype="beasong","selected","") %> >�����
			</select>
			<% DrawDateBox yyyy1,yyyy2,mm1,mm2,dd1,dd2 %>
			<input type="checkbox" name="oldlist" <% if oldlist="on" then response.write "checked" %> >6������������

			<br>

			�ֹ����� :
			<select class="select" name="itemstate">
				<option value="availall" <% if itemstate="availall" then response.write "selected" %>>����� ��ü
				<option value="ipkumfinishall" <% if itemstate="ipkumfinishall" then response.write "selected" %>>�����Ϸ��̻�
				<option value="2" <% if itemstate="2" then response.write "selected" %>>�ֹ�����
				<option value="4" <% if itemstate="4" then response.write "selected" %>>�����Ϸ�
				<option value="8" <% if itemstate="8" then response.write "selected" %>>���Ϸ�
			</select>
			&nbsp;
			����Ʈ :
			<% Drawsitename "sitename",sitename %>
        </td>
        <td align="center" width="50" bgcolor="#EEEEEE">
			<input type="button" class="button_s" value="�˻�" onClick="javascript:document.frm.submit();">
		</td>
	</tr>
	</form>
</table>
<!-- �˻� �� -->

<p />

* �ִ� <%= MaxRowSize %>�� ������ �˻��˴ϴ�.

<p />

<% if oitem.FResultCount>0 then %>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr bgcolor="#FFFFFF">
    	<td rowspan=<%= 6 + oitemoption.FResultCount -1 %> width="110" valign=top align=center><img src="<%= oitem.FOneItem.FImageList %>" width="100" height="100"></td>
      	<td width="60" bgcolor="<%= adminColor("tabletop") %>">��ǰ�ڵ�</td>
      	<td width="300">
      		10 <b><%= CHKIIF(oitem.FOneItem.FItemID>=1000000,Format00(8,oitem.FOneItem.FItemID),Format00(6,oitem.FOneItem.FItemID)) %></b> <%= itemoption %>
      		&nbsp;
      		<!--
      		<input type="button" value="����" onclick="PopItemSellEdit('<%= itemid %>');">
      		-->
      	</td>
      	<td width="60" bgcolor="<%= adminColor("tabletop") %>">���ÿ���</td>
      	<td colspan=2></td>
    </tr>
    <tr bgcolor="#FFFFFF">
      	<td bgcolor="<%= adminColor("tabletop") %>">�귣��ID</td>
      	<td><%= oitem.FOneItem.FMakerid %></td>
      	<td bgcolor="<%= adminColor("tabletop") %>">�Ǹſ���</td>
      	<td colspan=2></td>
    </tr>
    <tr bgcolor="#FFFFFF">
      	<td bgcolor="<%= adminColor("tabletop") %>">��ǰ��</td>
      	<td><%= oitem.FOneItem.FItemName %></td>
      	<td bgcolor="<%= adminColor("tabletop") %>">��뿩��</td>
      	<td colspan=2><font color="<%= ynColor(oitem.FOneItem.FIsUsing) %>"><%= oitem.FOneItem.FIsUsing %></font></td>
    </tr>
    <tr bgcolor="#FFFFFF">
      	<td bgcolor="<%= adminColor("tabletop") %>">�ǸŰ�</td>
      	<td>
      		<%= FormatNumber(oitem.FOneItem.FSellcash,0) %> / <%= FormatNumber(oitem.FOneItem.FBuycash,0) %>
      		&nbsp;&nbsp;
      		<font color="<%= oitem.FOneItem.getMwDivColor %>"><%= oitem.FOneItem.getMwDivName %></font>
      	    <% if oitem.FOneItem.FSellcash<>0 then %>
			<%= CLng((1- oitem.FOneItem.FBuycash/oitem.FOneItem.FSellcash)*100) %> %
			<% end if %>
			&nbsp;&nbsp;
      	</td>
      	<td bgcolor="<%= adminColor("tabletop") %>">��������</td>
      	<td colspan=2>
      		<% if oitem.FOneItem.Fdanjongyn="Y" then %>
			<font color="#33CC33">����</font>
			<% elseif oitem.FOneItem.Fdanjongyn="S" then %>
			<font color="#33CC33">�Ͻ�ǰ��</font>
			<% else %>
			������
			<% end if %>
		</td>
    </tr>
</table>
<% end if %>
<p>

<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td width="70">�ֹ���ȣ</td>
		<td width="40">����</td>
		<td width="40">����</td>
		<td width="70">Site</td>
		<td width="60">�ֹ�����</td>
		<td width="60">��ǰ����</td>
		<td width="40">����</td>
		<td>�ɼǸ�</td>
		<td>ȸ��ID</td>
		<td width="70">������</td>
		<td width="140">�����</td>
	</tr>
<%
if IsArray(RowArr) then
	for i=0 to RowCount
%>

	<tr align="center" bgcolor="#FFFFFF">
		<td><%= RowArr(0,i) %></td>
		<td><%= getJumundivName(RowArr(15,i)) %></td>
		<td><%= (RowArr(16,i)) %></td>
		<td><%= RowArr(12,i) %></td>
		<td><font color="<%= IpkumDivColor(RowArr(1,i)) %>"><%= IpkumDivName(RowArr(1,i)) %></font></td>
		<td><font color="<%= getCurrstateNameColor(RowArr(1,i),RowArr(11,i)) %>"><%= getCurrstateName(RowArr(1,i),RowArr(11,i)) %></font></td>
		<td><%= RowArr(2,i) %></td>
		<td><%= DdotFormat(RowArr(10,i),10) %></td>
		<td><%= printUserId(RowArr(14,i),2,"*") %></td>
		<td><%= RowArr(7,i) %></td>
		<td><%= RowArr(13,i) %></td>
	</tr>
<%
			totno = totno + RowArr(2,i)
    next
end if

%>
    <tr height="25" bgcolor="#FFFFFF">
        <td align="right" colspan="13">�ѻ�ǰ�� <%= totno %> �� / ���ֹ��Ǽ� <%= jumuncnt %> ��</td>
    </tr>
</table>

<%
set oitem = Nothing
set oitemoption = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
