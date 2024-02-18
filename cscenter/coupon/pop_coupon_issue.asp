<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : cs���� ��������
' History : �̻󱸻���
'			2023.05.23 �ѿ�� ����(���� ������Ŷ ���� üũ �߰�)
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbHelper.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/order/new_ordercls.asp"-->
<!-- #include virtual="/lib/classes/cscenter/cs_aslistcls.asp" -->
<%
' ������, cs�� �� ��밡��
if not(C_ADMIN_AUTH or C_CSUser) then
	response.write "�ش�Ŵ��� ������ �̰ų� cs���� ��밡���մϴ�."
	dbget.close() : response.end
end if

dim userid, orderserial, jukyo, i, buf, page, ojumun, ix,iy
	userid = requestCheckvar(request("userid"),32)
	orderserial = requestCheckvar(request("orderserial"),32)
	jukyo = requestCheckvar(request("jukyo"),32)

if (userid = "") then
    response.write "<script>alert('�߸��� �����Դϴ�.'); history.back();</script>"
    dbget.close()	:	response.End
end if

page = 1

set ojumun = new COrderMaster
	ojumun.FPageSize = 5
	ojumun.FCurrPage = page
	ojumun.FRectUserID = userid
	ojumun.FRectOrderSerial = orderserial
	ojumun.QuickSearchOrderList

'' ���� 6���� ���� ���� �˻�
if (ojumun.FResultCount<1) and (Len(orderserial)=11) and (IsNumeric(orderserial)) then
    ojumun.FRectOldOrder = "on"
    ojumun.QuickSearchOrderList

    if (ojumun.FResultCount>0) then
        response.write "<script>alert('6���� ���� �ֹ��Դϴ�.');</script>"
    end if
end if

'' �˻������ 1��
dim ResultOneOrderserial
ResultOneOrderserial = ""
if (ojumun.FResultCount=1) then
    ResultOneOrderserial = ojumun.FItemList(0).FOrderSerial
end if

if ((orderserial <> "") and (ojumun.FResultCount <> 1)) then
	response.write "<script>alert('�߸��� �ֹ���ȣ�Դϴ�.');</script>"

	orderserial = ""
end if

dim Coupon3000IssueAllow, Coupon5perIssueAllow, CouponDeliverIssueAllow, CouponBirthday
Coupon3000IssueAllow = False
Coupon5perIssueAllow = False
CouponDeliverIssueAllow = True
CouponBirthday = False

' �������̰ų� cs�� ������(��� �̻�) �̰�� ���డ��
if C_ADMIN_AUTH or C_CSpermanentUser then
	Coupon3000IssueAllow = True
	Coupon5perIssueAllow = True
end if

if C_ADMIN_AUTH or C_CSUser then
	CouponBirthday = True
end if

%>
<script type="text/javascript">

// ��������
function IssueCouponBirthday(frm){
	<% if (CouponBirthday <> True) then %>
		alert("�������� ��������� �����ϴ�.");
		return;
	<% end if %>

	//if (CheckForm(frm) != true) {
	//	return;
	//}

	if (confirm("���������� �����Ͻðڽ��ϱ�?") == true) {
		frm.submode.value = "IssueCouponBirthday"
		frm.submit();
	}
}

function IssueCoupon3000(frm)
{
	<% if (Coupon3000IssueAllow <> True) then %>
		alert("��������� �����ϴ�.");
		return;
	<% end if %>

	if (CheckForm(frm) != true) {
		return;
	}

	if (confirm("3000�� ���������� �����Ͻðڽ��ϱ�?") == true) {
		frm.submode.value = "issuecoupon3000"
		frm.submit();
	}
}

function IssueCoupon5per(frm)
{
	<% if (Coupon5perIssueAllow <> True) then %>
		alert("��������� �����ϴ�.");
		return;
	<% end if %>

	if (CheckForm(frm) != true) {
		return;
	}

	if (confirm("5% ���������� �����Ͻðڽ��ϱ�?") == true) {
		frm.submode.value = "issuecoupon5per"
		frm.submit();
	}
}

function IssueCouponDeliver(frm)
{
	<% if (CouponDeliverIssueAllow <> True) then %>
		alert("��������� �����ϴ�.");
		return;
	<% end if %>

	if (CheckForm(frm) != true) {
		return;
	}

	if (confirm("�����ۺ� ����(<%=Cstr(getDefaultBeasongPayByDate(now()))%>��)�� �����Ͻðڽ��ϱ�?") == true) {
		frm.submode.value = "issuecoupondeliver"
		frm.submit();
	}
}

function CheckForm(frm)
{
	if (frm.orderserial.value == "") {
		alert("���� �ֹ���ȣ�� �Է��ϼ���.");
		return false;
	}

	if (frm.jukyo.value == "") {
		alert("�߱޻����� �����ϼ���.");
		return false;
	}

	return true;
}

function SearchOrderSerial()
{
	document.frmsearch.orderserial.value = document.frm.orderserial.value;
	document.frmsearch.jukyo.value = document.frm.jukyo.value;
	document.frmsearch.submit();
}

function SetOrderSerial(orderserial)
{
	document.frm.orderserial.value = orderserial;
}

</script>

<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
<tr>
	<td align="left">
		<b>��������</b> &nbsp; ���������� �ϸ� CS�޸� ��ϵǸ�, ��� ����˴ϴ�.
		<br>*
		<% if (Coupon3000IssueAllow = True) then %>
		3000�� �������� ���డ��,
		<% end if %>
		<% if (Coupon5perIssueAllow = True) then %>
		5% �������� ���డ��,
		<% end if %>
		<% if (CouponDeliverIssueAllow = True) then %>
		�����ۺ� ����(<%=Cstr(getDefaultBeasongPayByDate(now()))%>��) ���డ��.
		<% end if %>
	</td>
</tr>
</table>

<form name="frmsearch" method="get" action="" onsubmit="return false;" style="margin:0px;">
<input type="hidden" name="userid" value="<%= userid %>">
<input type="hidden" name="orderserial" value="<%= orderserial %>">
<input type="hidden" name="jukyo" value="<%= jukyo %>">
</form>
<form name="frm" method="post" action="/cscenter/coupon/domodifycoupon.asp" onsubmit="return false;" style="margin:0px;">
<input type="hidden" name="mode" value="issuecoupon">
<input type="hidden" name="submode" value="">
<input type="hidden" name="userid" value="<%= userid %>">
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
  <tr align="left">
  	<td height="30" width="20%" bgcolor="#f1f1f1">���̵� :</td>
  	<td bgcolor="#FFFFFF" width="25%" >
  	  <b><%= userid %></b>
  	</td>
  	<td height="30" width="20%" bgcolor="#f1f1f1">�����ֹ���ȣ :</td>
  	<td bgcolor="#FFFFFF"  >
  	  <input type=text name=orderserial value="<%= ResultOneOrderserial %>">
  	  <input type=button class="button" value="�˻�" onclick="SearchOrderSerial()">
  	</td>
  </tr>
  <tr align="left">
  	<td height="30" bgcolor="#f1f1f1">�߱޻��� :</td>
  	<td bgcolor="#FFFFFF"  colspan=3>
		<select class="select" name="jukyo">
			<option value=''></option>
     		<option value='�������' <% if (jukyo = "�������") then %>selected<% end if %>>�������</option>
     		<option value='CS����' <% if (jukyo = "CS����") then %>selected<% end if %>>CS����</option>
			<option value='ǰ��' <% if (jukyo = "ǰ��") then %>selected<% end if %>>ǰ��</option>
			<option value='���ݿ���' <% if (jukyo = "���ݿ���") then %>selected<% end if %>>���ݿ���</option>
     		<option value='��Ÿ' <% if (jukyo = "��Ÿ") then %>selected<% end if %>>��Ÿ</option>
     	</select>
  	</td>
  </tr>
</table>
</form>

<p><br><p>

<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
    <tr align="center" bgcolor="<%= adminColor("tabletop") %>">
    	<td height=25>����</td>
    	<td>�ֹ���ȣ</td>
      	<td>������</td>
      	<td>�����Ѿ�</td>
      	<td>�������</td>
      	<td>�ŷ�����</td>
      	<td>�ֹ���</td>
      	<td>���</td>
    </tr>
	<% if (ojumun.FresultCount > 0) then %>
        <% for i=0 to ojumun.FResultCount - 1 %>
    <tr align="center" bgcolor="#FFFFFF">
    	<td height=25><%= ojumun.FItemList(ix).CancelYnName %></td>
    	<td><%= ojumun.FItemList(i).FOrderSerial %></td>
    	<td><%= ojumun.FItemList(i).FBuyName %></td>
    	<td><%= FormatNumber(ojumun.FItemList(i).FTotalSum,0) %></td>
    	<td><%= ojumun.FItemList(i).JumunMethodName %></td>
    	<td><%= ojumun.FItemList(i).IpkumDivName %></td>
    	<td><acronym title="<%= ojumun.FItemList(i).FRegDate %>"><%= Left(ojumun.FItemList(i).FRegDate,10) %></acronym></td>
    	<td><input type=button class="button" value="����" onclick="SetOrderSerial('<%= ojumun.FItemList(i).FOrderSerial %>')"></td>
    </tr>
        <% next %>
	<% else %>
    <tr align="center" bgcolor="#FFFFFF">
      	<td colspan="8"> �˻��� ����� �����ϴ�.</td>
    </tr>
	<% end if %>
</table>
<br>
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
<tr>
	<td align="center">
		<input type="button" class="button" value="3000�� ���� ��������" onClick="IssueCoupon3000(document.frm)">
		<input type="button" class="button" value="5% ���� ��������" onClick="IssueCoupon5per(document.frm)">
		<input type="button" class="button" value="������(<%=Cstr(getDefaultBeasongPayByDate(now()))%>��) ��������" onClick="IssueCouponDeliver(document.frm)">
		<input type="button" class="button" value="���� ��������" onClick="IssueCouponBirthday(document.frm)">
	</td>
</tr>
<tr>
	<td align="center">
		<br>
		<input type="button" class="button" value=" â �� �� " onClick="self.close()">
	</td>
</tr>
</table>

<%
set ojumun = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
