<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<%
dim gubun, detailid
gubun       = "witakchulgo"
detailid = request("detailid")

dim jungsanRows
dim sqlStr
sqlStr = " select m.yyyymm, m.designerid, d.gubuncd, d.mastercode, d.buyname, d.reqname, d.itemid, d.itemname, d.itemno, d.sellcash, d.suplycash "
sqlStr = sqlStr + " from [db_jungsan].[dbo].tbl_designer_jungsan_master m,"
sqlStr = sqlStr + " [db_jungsan].[dbo].tbl_designer_jungsan_detail d"
sqlStr = sqlStr + " where m.id=d.masteridx"
sqlStr = sqlStr + " and d.id=" + CStr(detailid)

rsget.Open sqlStr,dbget,1
if (detailid<>"") then
    If Not rsget.Eof then
        jungsanRows = rsget.getRows()
    end if
end if
rsget.Close

dim yyyymm, yyyy1, mm1, makerid, orderserial, buyname, reqname
dim itemid, itemname, itemno, sellcash, suplycash

if IsArray(jungsanRows) then
    yyyymm = jungsanRows(0,0)
    yyyy1  = Left(yyyymm,4)
    mm1  = Right(yyyymm,2)
    
    makerid     = jungsanRows(1,0)
    orderserial = jungsanRows(3,0)
    
    buyname     = db2html(jungsanRows(4,0))
    reqname     = db2html(jungsanRows(5,0))
    
    itemid     = jungsanRows(6,0)
    itemname   = db2html(jungsanRows(7,0))
    itemno     = jungsanRows(8,0)
    sellcash   = jungsanRows(9,0)
    suplycash  = jungsanRows(10,0)
end if
%>

<script language='javascript'>
function Editdata(frm){
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
			<font color="red"><strong>��Ÿ��ۺ� ���� ����</strong></font>
        </td>
        <td align="right">
        </td>
        <td background="/images/tbl_blue_round_05.gif"></td>
	</tr>
</table>
<!-- ǥ ��ܹ� ��-->
<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frmedit" method="post" action="dodesignerjungsan.asp">
    <input type="hidden" name="mode" value="etcbeasongpayedit">
    <input type="hidden" name="detailid" value="<%= detailid %>">
    <input type="hidden" name="gubun" value="<%= gubun %>">
    <input type="hidden" name="yyyy1" value="<%= yyyy1 %>">
    <input type="hidden" name="mm1" value="<%= mm1 %>">
    <input type="hidden" name="orderserial" value="<%= orderserial %>">
    <input type="hidden" name="makerid" value="<%= makerid %>">
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
        <td><%= orderserial %></td>
		<td>
		    <%= makerid %>
        </td>
		<td>
		    <input type="text" name="buyname" value="<%= buyname %>" size="8">
		</td>
		<td>
		    <input type="text" name="reqname" value="<%= reqname %>" size="8">
		</td>
		<td><input type="text" name="itemname" value="<%= itemname %>" size="40"></td>
		<td><input type="text" name="itemno" value="<%= itemno %>" size="3"></td>
		<td><input type="text" name="sellcash" value="<%= sellcash %>" size="8"></td>
		<td><input type="text" name="suplycash" value="<%= suplycash %>" size="8"></td>
    </tr>
</table>

<!-- ǥ �ϴܹ� ����-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
    <tr valign="bottom" height="25">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td valign="bottom" align="center"><input type="button" value="���� ����" onclick="Editdata(frmedit)"></td>
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
    </tr>
    <tr valign="top" height="10">
        <td width="10" align="right"><img src="/images/tbl_blue_round_07.gif" width="10" height="10"></td>
        <td background="/images/tbl_blue_round_08.gif"></td>
        <td width="10" align="left"><img src="/images/tbl_blue_round_09.gif" width="10" height="10"></td>
    </tr>
    </form>
</table>

<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
