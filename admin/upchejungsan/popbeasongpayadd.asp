<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->

<%
dim gubun, orderserial
dim yyyy1, mm1

gubun       = "witakchulgo"
orderserial = request("orderserial")
yyyy1       = request("yyyy1")
mm1         = request("mm1")

dim sqlStr
dim jungsanDataExists
dim orderRows, jungsanRows
sqlStr = " select distinct m.buyname, m.reqname, d.makerid "
sqlStr = sqlStr + " from [db_order].[dbo].tbl_order_master m "
sqlStr = sqlStr + " left join [db_order].[dbo].tbl_order_detail d on m.orderserial=d.orderserial"
sqlStr = sqlStr + " where m.orderserial='" + CStr(orderserial) + "'"

if (orderserial<>"") then
    rsget.Open sqlStr,dbget,1
    If Not rsget.Eof then
        orderRows = rsget.getRows()
    end if
    rsget.Close
end if

sqlStr = " select top 10 m.yyyymm, m.designerid, d.gubuncd, d.mastercode, d.itemid, d.itemname, d.itemno, d.sellcash, d.suplycash "
sqlStr = sqlStr + " from [db_jungsan].[dbo].tbl_designer_jungsan_master m,"
sqlStr = sqlStr + " [db_jungsan].[dbo].tbl_designer_jungsan_detail d"
sqlStr = sqlStr + " where m.id=d.masteridx"
sqlStr = sqlStr + " and d.itemid=0"
sqlStr = sqlStr + " and d.gubuncd='witakchulgo'"
sqlStr = sqlStr + " and d.mastercode='" + CStr(orderserial) + "'"

if (orderserial<>"") then
    rsget.Open sqlStr,dbget,1
    If Not rsget.Eof then
        jungsanRows = rsget.getRows()
    end if
    rsget.Close
end if

dim i
%>

<script language='javascript'>
function searchOrder(frm){
    if (frm.orderserial.value.length!=11){
        alert('�ֹ���ȣ 11 �ڸ��� �Է��ϼ���.');
        frm.orderserial.focus();
        return;
    }
    frm.method="get";
    frm.action="/admin/upchejungsan/popbeasongpayadd.asp";
    frm.submit();
}

function adddata(frm){
    if (frm.orderserial.value.length!=11){
        alert('�ֹ���ȣ�� �Է��ϼ���.');
		frm.orderserial.focus();
		return;
    }
    
    if (frm.makerid.value.length<1){
        alert('�귣��ID�� �����ϼ���.');
		frm.makerid.focus();
		return;
    }
    
	if (frm.itemname.value.length<1){
		alert('������ �Է��ϼ���.');
		frm.itemname.focus();
		return;
	}

	if (frm.itemno.value.length<1){
		alert('������ �Է��ϼ���.');
		frm.itemno.focus();
		return;
	}

	if (!IsDigit(frm.itemno.value)){
		alert('������ ���ڸ� �����մϴ�.');
		frm.itemno.focus();
		return;
	}

	if (frm.sellcash.value.length<1){
		alert('�ǸŰ��� �Է��ϼ���.');
		frm.sellcash.focus();
		return;
	}

	if (!IsDigit(frm.sellcash.value)){
		alert('�ǸŰ��� ���ڸ� �����մϴ�.');
		frm.sellcash.focus();
		return;
	}

	if (frm.suplycash.value.length<1){
		alert('���԰��� �Է��ϼ���.');
		frm.suplycash.focus();
		return;
	}

	if (!IsDigit(frm.suplycash.value)){
		alert('���԰��� ���ڸ� �����մϴ�.');
		frm.suplycash.focus();
		return;
	}

	var ret = confirm('���� �Ͻðڽ��ϱ�?');
	if (ret){
	    
		frm.submit();
	}
}
</script>


<!-- ǥ ��ܹ� ����-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
   	<tr height="10" valign="bottom">
        <td width="10" align="right"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
        <td background="/images/tbl_blue_round_02.gif"></td>
        <td background="/images/tbl_blue_round_02.gif"></td>
        <td width="10" align="left" ><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
	</tr>
	<tr height="25" valign="top">
        <td background="/images/tbl_blue_round_04.gif"></td>
        <td>
        	<img src="/images/icon_star.gif" align="absbottom">
			<font color="red"><strong>��Ÿ�����߰�</strong></font>
        </td>
        <td align="right">��ۺ�����
        </td>
        <td background="/images/tbl_blue_round_05.gif"></td>
	</tr>
</table>
<!-- ǥ ��ܹ� ��-->
<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frmadd" method="post" action="dodesignerjungsan.asp">
    <input type="hidden" name="mode" value="etcbeasongpayadd">
    <input type="hidden" name="gubun" value="<%= gubun %>">
    <input type="hidden" name="yyyy1" value="<%= yyyy1 %>">
    <input type="hidden" name="mm1" value="<%= mm1 %>">
    <tr align="center" bgcolor="<%= adminColor("tabletop") %>">
        <td width="80">������</td>
        <td width="100">�ֹ���ȣ</td>
        <td width="100">�귣��ID</td>
        <td width="80">������</td>
        <td width="80">������</td>
		<td>����</td>
		<td width="40">����</td>
		<td width="80">�ǸŰ�</td>
		<td width="80">���ް�</td>
    </tr>
    <tr bgcolor="#FFFFFF">
        <td><%= yyyy1 %>-<%= mm1 %></td>
        <td><input type="text" name="orderserial" value="<%= orderserial %>" size="12" maxlength="11"><input type="button" value="�˻�" onClick="searchOrder(frmadd);" onFocus="this.blur"></td>
		<td>
		    <select name="makerid">
            <% if IsArray(orderRows) then %>
            <% for i=0 to UBound(orderRows,2) %>
            <option value="<%= orderRows(2,i) %>"><%= orderRows(2,i) %>
            <% next %>
            <% end if %>
            </select>
        </td>
		<td>
		    <% if IsArray(orderRows) then %>
		    <input type="text" name="buyname" value="<%= db2html(orderRows(0,0)) %>" size="8">
		    <% else %>
		    <input type="text" name="buyname" value="" size="8">
		    <% end if %>
		</td>
		<td>
		    <% if IsArray(orderRows) then %>
		    <input type="text" name="reqname" value="<%= db2html(orderRows(1,0)) %>" size="8">
		    <% else %>
		    <input type="text" name="reqname" value="" size="8">
		    <% end if %>
		</td>
		<td><input type="text" name="itemname" value="" size="40"></td>
		<td><input type="text" name="itemno" value="1" size="3"></td>
		<td><input type="text" name="sellcash" value="" size="8"></td>
		<td><input type="text" name="suplycash" value="" size="8"></td>
    </tr>
</table>

<!-- ǥ �ϴܹ� ����-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
    <tr valign="bottom" height="25">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td valign="bottom" align="center"><input type="button" value="���� �߰�" onclick="adddata(frmadd)"></td>
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
    </tr>
    <tr valign="top" height="10">
        <td width="10" align="right"><img src="/images/tbl_blue_round_07.gif" width="10" height="10"></td>
        <td background="/images/tbl_blue_round_08.gif"></td>
        <td width="10" align="left"><img src="/images/tbl_blue_round_09.gif" width="10" height="10"></td>
    </tr>
    </form>
</table>
<br>
<% if IsArray(jungsanRows) then %>
<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<tr bgcolor="#FFFFFF">
    <td colspan="9">��ϵ� ���� ����</td>
</tr>
<tr bgcolor="#DDDDFF">
    <td>�����</td>
    <td>�귣��</td>
    <td>����</td>
    <td>�ֹ���ȣ</td>
    <td>��ǰ�ڵ�</td>
    <td>��ǰ��</td>
    <td>����</td>
    <td>�ǸŰ�</td>
    <td>����ݿ���</td>
</tr>
<% for i=0 to UBound(jungsanRows,2) %>
<tr bgcolor="#FFFFFF">
    <td><%= jungsanRows(0,i) %></td>
    <td><%= jungsanRows(1,i) %></td>
    <td><%= jungsanRows(2,i) %></td>
    <td><%= jungsanRows(3,i) %></td>
    <td><%= jungsanRows(4,i) %></td>
    <td><%= jungsanRows(5,i) %></td>
    <td><%= jungsanRows(6,i) %></td>
    <td><%= jungsanRows(7,i) %></td>
    <td><%= jungsanRows(8,i) %></td>
</tr>
<% next %>
</table>

<% end if %>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->