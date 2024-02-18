<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/partners/partnerusercls.asp"-->
<!-- #include virtual="/lib/classes/jungsan/new_upchejungsancls.asp"-->
<!-- #include virtual="/lib/classes/jungsan/offjungsancls.asp"-->
<%
dim opartner, designer, i
	designer = requestCheckvar(request("designer"),40)

set opartner = new CPartnerUser
	opartner.FRectDesignerID = designer
	opartner.GetOnePartnerNUser

if opartner.FResultCount< 1 then
	response.write "<script type='text/javascript'>"
	response.write "	alert('�ش�Ǵ� �귣�尡 �����ϴ�.');"
	response.write "</script>"
	dbget.close() : response.end
end if

''On line ����ǰ����
dim sqlStr
dim totalitemcount , totalitemcount_m , totalitemcount_w , totalitemcount_u
dim usingitemcount , usingitemcount_m , usingitemcount_w , usingitemcount_u

sqlStr = " select count(itemid) as totalitemcount, "
sqlStr = sqlStr + " sum(case when mwdiv='M' then 1 else 0 end) as totalitemcount_m, "
sqlStr = sqlStr + " sum(case when mwdiv='W' then 1 else 0 end) as totalitemcount_w, "
sqlStr = sqlStr + " sum(case when mwdiv='U' then 1 else 0 end) as totalitemcount_u, "

sqlStr = sqlStr + " sum(case when isusing='Y' then 1 else 0 end) as usingitemcount, "
sqlStr = sqlStr + " sum(case when isusing='Y' and mwdiv='M' then 1 else 0 end) as usingitemcount_m, "
sqlStr = sqlStr + " sum(case when isusing='Y' and mwdiv='W' then 1 else 0 end) as usingitemcount_w, "
sqlStr = sqlStr + " sum(case when isusing='Y' and mwdiv='U' then 1 else 0 end) as usingitemcount_u "

sqlStr = sqlStr & " from [db_item].[dbo].tbl_item with (nolock)"
sqlStr = sqlStr + " where makerid='" + designer + "'"
rsget.Open sqlStr,dbget,1

totalitemcount = rsget("totalitemcount")
totalitemcount_m = rsget("totalitemcount_m")
totalitemcount_w = rsget("totalitemcount_w")
totalitemcount_u = rsget("totalitemcount_u")

usingitemcount = rsget("usingitemcount")
usingitemcount_m = rsget("usingitemcount_m")
usingitemcount_w = rsget("usingitemcount_w")
usingitemcount_u = rsget("usingitemcount_u")

if IsNULL(usingitemcount) then usingitemcount=0

if IsNULL(totalitemcount) then totalitemcount = 0 end if
if IsNULL(totalitemcount_m) then totalitemcount_m = 0 end if
if IsNULL(totalitemcount_w) then totalitemcount_w = 0 end if
if IsNULL(totalitemcount_u) then totalitemcount_u = 0 end if
if IsNULL(usingitemcount) then usingitemcount = 0 end if
if IsNULL(usingitemcount_m) then usingitemcount_m = 0 end if
if IsNULL(usingitemcount_w) then usingitemcount_w = 0 end if
if IsNULL(usingitemcount_u) then usingitemcount_u = 0 end if
rsget.Close

''Off line ����ǰ����
dim offtotalitemcount , offtotalitemcount_00 , offtotalitemcount_10 , offtotalitemcount_70 , offtotalitemcount_80 , offtotalitemcount_90 , offtotalitemcount_95
dim offusingitemcount , offusingitemcount_00 , offusingitemcount_10 , offusingitemcount_70 , offusingitemcount_80 , offusingitemcount_90 , offusingitemcount_95
sqlStr = " select count(shopitemid) as offtotalitemcount, "
sqlStr = sqlStr + " sum(case when itemgubun='00' then 1 else 0 end) as offtotalitemcount_00, "
sqlStr = sqlStr + " sum(case when itemgubun='10' then 1 else 0 end) as offtotalitemcount_10, "
sqlStr = sqlStr + " sum(case when itemgubun='70' then 1 else 0 end) as offtotalitemcount_70, "
sqlStr = sqlStr + " sum(case when itemgubun='80' then 1 else 0 end) as offtotalitemcount_80, "
sqlStr = sqlStr + " sum(case when itemgubun='90' then 1 else 0 end) as offtotalitemcount_90, "
sqlStr = sqlStr + " sum(case when itemgubun='95' then 1 else 0 end) as offtotalitemcount_95, "
sqlStr = sqlStr + " sum(case when isusing='Y' then 1 else 0 end) as offusingitemcount, "
sqlStr = sqlStr + " sum(case when isusing='Y' and itemgubun='00' then 1 else 0 end) as offusingitemcount_00, "
sqlStr = sqlStr + " sum(case when isusing='Y' and itemgubun='10' then 1 else 0 end) as offusingitemcount_10, "
sqlStr = sqlStr + " sum(case when isusing='Y' and itemgubun='70' then 1 else 0 end) as offusingitemcount_70, "
sqlStr = sqlStr + " sum(case when isusing='Y' and itemgubun='80' then 1 else 0 end) as offusingitemcount_80, "
sqlStr = sqlStr + " sum(case when isusing='Y' and itemgubun='90' then 1 else 0 end) as offusingitemcount_90, "
sqlStr = sqlStr + " sum(case when isusing='Y' and itemgubun='95' then 1 else 0 end) as offusingitemcount_95 "
sqlStr = sqlStr & " from [db_shop].[dbo].tbl_shop_item with (nolock)"
sqlStr = sqlStr + " where makerid='" + designer + "'"

rsget.Open sqlStr,dbget,1

offtotalitemcount = rsget("offtotalitemcount")
offtotalitemcount_00 = rsget("offtotalitemcount_00")
offtotalitemcount_10 = rsget("offtotalitemcount_10")
offtotalitemcount_70 = rsget("offtotalitemcount_70")
offtotalitemcount_80 = rsget("offtotalitemcount_80")
offtotalitemcount_90 = rsget("offtotalitemcount_90")
offtotalitemcount_95 = rsget("offtotalitemcount_95")
offusingitemcount = rsget("offusingitemcount")
offusingitemcount_00 = rsget("offusingitemcount_00")
offusingitemcount_10 = rsget("offusingitemcount_10")
offusingitemcount_70 = rsget("offusingitemcount_70")
offusingitemcount_80 = rsget("offusingitemcount_80")
offusingitemcount_90 = rsget("offusingitemcount_90")
offusingitemcount_95 = rsget("offusingitemcount_95")

if IsNULL(offtotalitemcount) then offtotalitemcount = 0 end if
if IsNULL(offtotalitemcount_00) then offtotalitemcount_00 = 0 end if
if IsNULL(offtotalitemcount_10) then offtotalitemcount_10 = 0 end if
if IsNULL(offtotalitemcount_70) then offtotalitemcount_70 = 0 end if
if IsNULL(offtotalitemcount_80) then offtotalitemcount_80 = 0 end if
if IsNULL(offtotalitemcount_90) then offtotalitemcount_90 = 0 end if
if IsNULL(offtotalitemcount_95) then offtotalitemcount_95 = 0 end if
if IsNULL(offusingitemcount) then offusingitemcount = 0 end if
if IsNULL(offusingitemcount_00) then offusingitemcount_00 = 0 end if
if IsNULL(offusingitemcount_10) then offusingitemcount_10 = 0 end if
if IsNULL(offusingitemcount_70) then offusingitemcount_70 = 0 end if
if IsNULL(offusingitemcount_80) then offusingitemcount_80 = 0 end if
if IsNULL(offusingitemcount_90) then offusingitemcount_90 = 0 end if
if IsNULL(offusingitemcount_95) then offusingitemcount_95 = 0 end if

rsget.Close

dim ojungsan
set ojungsan = new CUpcheJungsan
	ojungsan.FRectDesigner = designer
	ojungsan.JungsanMasterList

dim oshopjungsan
set oshopjungsan = new COffJungsan
	oshopjungsan.FPageSize = 100
	oshopjungsan.FRectMakerid = designer
	oshopjungsan.GetOffJungsanMasterListBrandView

dim notfinishedjungsancount
notfinishedjungsancount = 0

for i=0 to ojungsan.FResultCount - 1
	if (ojungsan.FItemList(i).Ffinishflag<>"7") then
		notfinishedjungsancount = notfinishedjungsancount + 1
	end if
next

for i=0 to oshopjungsan.FResultCount - 1
	if (oshopjungsan.FItemList(i).Ffinishflag<>"7") then
		notfinishedjungsancount = notfinishedjungsancount + 1
	end if
next
%>
<script type='text/javascript'>

var usingitemcount = <%= usingitemcount %>;
var	notfinishedjungsancount = <%= notfinishedjungsancount %>;
function popItemSellEdit(designerid,mwdiv,usingyn){
	var popwin = window.open('/admin/shopmaster/itemviewset.asp?menupos=24&makerid=' + designerid + '&mwdiv=' + mwdiv + '&usingyn=' + usingyn  ,'popItemSellEdit','width=1000,height=800,scrollbars=yes,resizable=yes')
	popwin.focus();
}

function popOffItemSellEdit(designerid,itemgubun,usingyn){
	var popwin = window.open('/admin/offshop/shopitemlist.asp?menupos=184&research=on&page=1&ckonlyusing=on&designer=' + designerid + '&itemgubun=' + itemgubun + '&usingyn=' + usingyn ,'popOffItemSellEdit','width=1000,height=800,scrollbars=yes,resizable=yes')
	popwin.focus();
}


function saveForm(frm){

	if (frm.partnerusing[1].checked){
		if (usingitemcount>0){
			alert('��� �ϰ� �ִ� ��ǰ�� �����մϴ�. ������ ������ ���� �������� ���� �� �� �ֽ��ϴ�.');
			return;
		}
        <% if NOT(C_ADMIN_AUTH) then %>
		if (notfinishedjungsancount>0){
			alert('���� �Ϸ���� ���� ������ �ֽ��ϴ�. ���� �Ϸ��� ���� �������� ���� �� �� �ֽ��ϴ�.');
			return;
		}
	    <% end if %>
	}else{

	}

	if (confirm('���� �Ͻðڽ��ϱ�?')){
		frm.submit();
	}

}

function savePolicyForm(frm){
	<% if (usingitemcount_m <> 0) or (usingitemcount_w <> 0) then %>
		// ���� ��ۺ� ��å ���尪�� �⺻��å�ΰ��		// 2019.02.21 �ѿ�� ����(�̹��� �̻�� ����)
		if (frm.orgdefaultdeliveryType.value==""){
			// �⺻��å�� �ƴѰ� ���ý� �ðܳ�
			if (frm.defaultdeliveryType[0].checked != true){
				alert('���� �Ǵ� ��Ź ��ǰ�� �ִ°��  �⺻��å�� �����մϴ�.');
				return;
			}
		}
	<% end if %>

	if ((frm.defaultFreeBeasongLimit.value.length < 1) || (frm.defaultFreeBeasongLimit.value*0 != 0)){
		alert('�����۱��رݾ��� ��Ȯ�� �Է��ϼ���.');
		return;
	}

	if ((frm.defaultDeliverPay.value.length < 1) || (frm.defaultDeliverPay.value*0 != 0)){
		alert('��ۺ� ��Ȯ�� �Է��ϼ���.');
		return;
	}

//�߰� ���� üũ
    if ((frm.defaultdeliveryType[1].checked)&&(frm.defaultFreeBeasongLimit.value*1<1)){
        alert('���� ����� ��� �����۱��رݾ��� 0�� �̻��̾�� �մϴ�.');
		return;
    }

    if ((frm.defaultdeliveryType[1].checked) && ((frm.defaultDeliverPay.value*0 != 0) || (frm.defaultDeliverPay.value*1 < 1))) {
        alert('���� ����� ��� ��ۺ� �Է��ؾ� �մϴ�.');
		return;
    }

    //�⺻��� ��å�� �����۱��� ���� ����
    if ((frm.defaultdeliveryType[0].checked)&&(frm.defaultFreeBeasongLimit.value*1!=0)){
        alert('��ü������ ��å�� ��� �����۱��رݾ��� 0���̾�� �մϴ�.');
		return;
    }

    //����
    if ((frm.defaultdeliveryType[2].checked)&&(frm.defaultFreeBeasongLimit.value*1!=0)){
        alert('���ҹ�� ��å�� ��� �����۱��رݾ��� 0���̾�� �մϴ�.');
		return;
    }

    if ((frm.defaultdeliveryType[2].checked)&&(frm.defaultDeliverPay.value*1!=0)){
        alert('���ҹ�� ��å�� ��� ��ۺ�� 0���̾�� �մϴ�.');
		return;
    }


	if (confirm('���� �Ͻðڽ��ϱ�?')){
		frm.submit();
	}
}

</script>

<table width="600" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frm" method="get" action="">
	<input type="hidden" name="page" value="1">
	<tr height="25" bgcolor="FFFFFF">
		<td height="25" colspan="15">
			�귣�� ID : <input type="text" class="text" name="designer" value="<%= designer %>" Maxlength="32" size="16">
			<input type="button" class="button" value="�˻�" onclick="document.frm.submit()">
		</td>
	</tr>
	</form>

	<form name="frmedit" method="post" action="dobrandadminusing.asp">
	<input type="hidden" name="designer" value="<%= designer %>">
	<input type="hidden" name="mode" value="using">
	<tr>
		<td height="25" bgcolor="#FFFFFF" colspan="3">**�귣�� ��뿩�� ����**</td>
	</tr>
	<tr>
		<td rowspan="3" width="110" bgcolor="<%= adminColor("pink") %>">�귣��<br>��뿩��</td>
		<td bgcolor="#FFFFFF">�ٹ�����</td>
		<td bgcolor="#FFFFFF"><input type=radio name="isusing" value="Y" <% if opartner.FOneItem.Fisusing="Y" then response.write "checked" %>  >��� <input type=radio name="isusing" value="N" <% if opartner.FOneItem.Fisusing="N" then response.write "checked" %>  >������</td>
	</tr>
	<tr>
		<td bgcolor="#FFFFFF">�ٹ�����OFF</td>
		<td bgcolor="#FFFFFF"><input type=radio name="isoffusing" value="Y" <% if opartner.FOneItem.Fisoffusing="Y" then response.write "checked" %>  >��� <input type=radio name="isoffusing" value="N" <% if opartner.FOneItem.Fisoffusing="N" then response.write "checked" %>  >������	</td>
	</tr>
	<tr>
		<td bgcolor="#FFFFFF">���޸�</td>
		<td bgcolor="#FFFFFF"><input type=radio name="isextusing" value="Y" <% if opartner.FOneItem.Fisextusing="Y" then response.write "checked" %>  >��� <input type=radio name="isextusing" value="N" <% if opartner.FOneItem.Fisextusing="N" then response.write "checked" %>  >������	</td>
	</tr>
	<tr>
		<td rowspan="2" width="110" bgcolor="<%= adminColor("pink") %>">��Ʈ��Ʈ<br>ǥ�ÿ���<br>(�귣������)</td>
		<td bgcolor="#FFFFFF">�ٹ�����</td>
		<td bgcolor="#FFFFFF"><input type=radio name="streetusing" value="Y" <% if opartner.FOneItem.Fstreetusing="Y" then response.write "checked" %>  >��� <input type=radio name="streetusing" value="N" <% if opartner.FOneItem.Fstreetusing="N" then response.write "checked" %>  >������</td>
	</tr>
	<tr>
		<td bgcolor="#FFFFFF">���޸�</td>
		<td bgcolor="#FFFFFF"><input type=radio name="extstreetusing" value="Y" <% if opartner.FOneItem.Fextstreetusing="Y" then response.write "checked" %>  >��� <input type=radio name="extstreetusing" value="N" <% if opartner.FOneItem.Fextstreetusing="N" then response.write "checked" %>  >������	</td>
	</tr>
	<tr >
		<td bgcolor="<%= adminColor("pink") %>">Ŀ�´�Ƽ </td>
		<td bgcolor="#FFFFFF">��ǰQ/A</td>
		<td bgcolor="#FFFFFF"><input type=radio name="specialbrand" value="Y" <% if opartner.FOneItem.Fspecialbrand="Y" then response.write "checked" %>  >��� <input type=radio name="specialbrand" value="N" <% if opartner.FOneItem.Fspecialbrand="N" then response.write "checked" %>  >������</td>
	</tr>
	<tr >
		<td bgcolor="#DDDDFF">��ü����</td>
		<td bgcolor="#FFFFFF">���¿���</td>
		<td bgcolor="#FFFFFF"><input type=radio name="partnerusing" value="Y" <% if opartner.FOneItem.Fpartnerusing="Y" then response.write "checked" %>  >��� <input type=radio name="partnerusing" value="N" <% if opartner.FOneItem.Fpartnerusing="N" then response.write "checked" %>  >������</td>
	</tr>
	<tr>
		<td colspan="3" align="center" bgcolor="#FFFFFF"><input type="button" class="button" value="���� " onclick="saveForm(frmedit)"></td>
	</tr>
	</form>
</table>

<p>

<table width="600" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr>
		<td height="25" bgcolor="#FFFFFF" colspan="15">**��ǰ ����** [��ǰ������ Ŭ���Ͻø� �󼼳��� Ȯ�� �����մϴ�.]</td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("pink") %>">
		<td rowspan="2">����</td>
		<td colspan="4">�¶���</td>
		<td colspan="3">��������</td>
		<td colspan="4">��Ÿ</td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("pink") %>">
		<td>����</td>
		<td>��Ź</td>
		<td>��ü</td>
		<td>�հ�</td>

		<td>10</td>
		<td>90</td>
		<td>�հ�</td>

		<td>00</td>
		<td>70</td>
		<td>80</td>
		<td>95</td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF">
		<td>��ü</td>
		<td><a href="javascript:popItemSellEdit('<%= designer %>','M','');"><%= (totalitemcount_m) %></a></td>
		<td><a href="javascript:popItemSellEdit('<%= designer %>','W','');"><%= (totalitemcount_w) %></a></td>
		<td><a href="javascript:popItemSellEdit('<%= designer %>','U','');"><%= (totalitemcount_u) %></a></td>
		<td><a href="javascript:popItemSellEdit('<%= designer %>','','');"><b><%= (totalitemcount) %><b></a></td>

		<td><a href="javascript:popOffItemSellEdit('<%= designer %>','10','');"><%= (offtotalitemcount_10) %></a></td>
		<td><a href="javascript:popOffItemSellEdit('<%= designer %>','90','');"><%= (offtotalitemcount_90) %></a></td>
		<td><a href="javascript:popOffItemSellEdit('<%= designer %>','','');"><b><%= (offtotalitemcount) %></b></a></td>

		<td><a href="javascript:popOffItemSellEdit('<%= designer %>','00','');"><%= (offtotalitemcount_00) %></a></td>
		<td><a href="javascript:popOffItemSellEdit('<%= designer %>','70','');"><%= (offtotalitemcount_70) %></a></td>
		<td><a href="javascript:popOffItemSellEdit('<%= designer %>','80','');"><%= (offtotalitemcount_80) %></a></td>
		<td><a href="javascript:popOffItemSellEdit('<%= designer %>','95','');"><%= (offtotalitemcount_95) %></a></td>

	</tr>
	<tr align="center" bgcolor="#FFFFFF">
		<td>���</td>
		<td><a href="javascript:popItemSellEdit('<%= designer %>','M','Y');"><%= FormatNumber(usingitemcount_m,0) %></a></td>
		<td><a href="javascript:popItemSellEdit('<%= designer %>','W','Y');"><%= FormatNumber(usingitemcount_w,0) %></a></td>
		<td><a href="javascript:popItemSellEdit('<%= designer %>','U','Y');"><%= FormatNumber(usingitemcount_u,0) %></a></td>
		<td><a href="javascript:popItemSellEdit('<%= designer %>','','Y');"><font color="Red"><b><%= FormatNumber(usingitemcount,0) %><b></font></a></td>

		<td><a href="javascript:popOffItemSellEdit('<%= designer %>','10','Y');"><%= (offusingitemcount_10) %></a></td>
		<td><a href="javascript:popOffItemSellEdit('<%= designer %>','90','Y');"><%= (offusingitemcount_90) %></a></td>
		<td><a href="javascript:popOffItemSellEdit('<%= designer %>','','Y');"><font color="Red"><b><%= (offusingitemcount) %></b></font></a></td>

		<td><a href="javascript:popOffItemSellEdit('<%= designer %>','00','Y');"><%= (offusingitemcount_00) %></a></td>
		<td><a href="javascript:popOffItemSellEdit('<%= designer %>','70','Y');"><%= (offusingitemcount_70) %></a></td>
		<td><a href="javascript:popOffItemSellEdit('<%= designer %>','80','Y');"><%= (offusingitemcount_80) %></a></td>
		<td><a href="javascript:popOffItemSellEdit('<%= designer %>','95','Y');"><%= (offusingitemcount_95) %></a></td>
	</tr>
</table>

<p>

<table width="600" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frmeditpolicy" method="post" action="dobrandadminusing.asp">
	<input type="hidden" name="designer" value="<%= designer %>">
	<input type="hidden" name="mode" value="policy">
	<input type="hidden" name="orgdefaultdeliveryType" value="<%= opartner.FOneItem.FdefaultDeliveryType %>">
	<tr>
		<td height="25" bgcolor="#FFFFFF" colspan="15">**�����å ����**</td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td width="110" bgcolor="<%= adminColor("pink") %>">���ǹ�ۿ���</td>
		<td>
			<input type="radio" name="defaultdeliveryType" value="" checked>��ü������
			<input type="radio" name="defaultdeliveryType" value="9" <% if (opartner.FOneItem.FdefaultDeliveryType = "9") then %>checked<% end if %>>��ü���ǹ��
			<input type="radio" name="defaultdeliveryType" value="7" <% if (opartner.FOneItem.FdefaultDeliveryType = "7") then %>checked<% end if %>>��ü���ҹ��
		</td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td bgcolor="<%= adminColor("pink") %>">�����۱��رݾ�</td>
		<td><input type="text" class="text" name="defaultFreeBeasongLimit" value="<%= opartner.FOneItem.FdefaultFreeBeasongLimit %>" size="10"> ��</td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td bgcolor="<%= adminColor("pink") %>">��ۺ�</td>
		<td>
		    <input type="text" class="text" name="defaultDeliverPay" value="<%= opartner.FOneItem.FdefaultDeliverPay %>" size="10"> ��
		    (��ü�������� ��� ��ۺ� ��������/<font color=red><b>��ü ȸ��/��ǰ</b></font> ��ۺ񿡼� ���)
		</td>
	</tr>
	<tr>
		<td colspan="3" align="center" bgcolor="#FFFFFF">
<!-- �����̻� + �ٹ����ٻ���� - MD��Ʈ(�/�ҽ�) �����̻� ��������(�������� ����:2011.09.01) -->
<% if ((session("ssAdminLsn") = "1") or (session("ssAdminLsn") = "2") or ((session("ssAdminLsn") <= "4") and (session("ssAdminPsn")="11" or session("ssAdminPsn")="21"))) then %>
			<input type="button" class="button" value="�����å �������� " onclick="savePolicyForm(frmeditpolicy)">
<% else %>
		���������� ���MD���� �����ϼ���.
<% end if %>
		</td>
	</tr>
	</form>
</table>

<p>

<table width="600" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr>
		<td height="25" bgcolor="#FFFFFF" colspan="15">**���� ����**</td>
	</tr>
	<tr>
		<td height="25" bgcolor="#FFFFFF" colspan="15">- online ����</td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("pink") %>">
		<td width="60">�����</td>
		<td width="50">����</td>
		<td width="90">�������</td>
		<td width="90">�������</td>
		<td width="70">������</td>
		<td width="70">�Ա���</td>
	    <td>������</td>
	</tr>
	<% for i=0 to ojungsan.FResultCount-1 %>
	<tr align="center" bgcolor="#FFFFFF" >
		<td align="center"><%= ojungsan.FItemList(i).FYYYYMM %></td>
		<td><%= ojungsan.FItemList(i).GetSimpleTaxtypeName %></td>
		<td align="right"><%= FormatNumber(ojungsan.FItemList(i).GetTotalSuplycash,0) %></td>
		<td><font color="<%= ojungsan.FItemList(i).GetStateColor %>"><%= ojungsan.FItemList(i).GetStateName %></font></td>
		<td><% if not isNull(ojungsan.FItemList(i).Ftaxregdate) then %><%= Left(Cstr(ojungsan.FItemList(i).Ftaxregdate),10) %><% end if %></td>
		<td><% if not isNull(ojungsan.FItemList(i).Fipkumdate) then %><%= Left(Cstr(ojungsan.FItemList(i).Fipkumdate),10) %><% end if %></td>
	    <td>�Ϳ� <%= ojungsan.FItemList(i).Fjungsan_date %></td>
	</tr>
	<% next %>
	<tr>
		<td height="25" bgcolor="#FFFFFF" colspan="15">- offline ����</td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("pink") %>">
		<td>�����</td>
		<td>����</td>
		<td>�������</td>
		<td>�������</td>
		<td>������</td>
		<td>�Ա���</td>
	    <td>������</td>
	</tr>
	<% for i=0 to oshopjungsan.FResultCount-1  %>
	<tr align="center" bgcolor="#FFFFFF" >
		<td><%= oshopjungsan.FItemList(i).FYYYYMM %></td>
		<td><%= oshopjungsan.FItemList(i).GetSimpleTaxtypeName %></td>
		<td align="right"><%= FormatNumber(oshopjungsan.FItemList(i).Ftot_jungsanprice,0) %></td>
		<td align="center"><font color="<%= oshopjungsan.FItemList(i).GetStateColor %>"><%= oshopjungsan.FItemList(i).GetStateName %></font></td>
		<td><% if not isNull(oshopjungsan.FItemList(i).Ftaxregdate) then %><%= Left(Cstr(oshopjungsan.FItemList(i).Ftaxregdate),10) %><% end if %></td>
		<td><% if not isNull(oshopjungsan.FItemList(i).Fipkumdate) then %><%= Left(Cstr(oshopjungsan.FItemList(i).Fipkumdate),10) %><% end if %></td>
	    <td>�Ϳ� <%= oshopjungsan.FItemList(i).Fjungsan_date_off %></td>
	</tr>
	<% next %>

</table>


<%
set opartner = Nothing
set ojungsan = Nothing
set oshopjungsan = Nothing
%>

<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
