<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : �ΰŽ� ������ ���ϸ���
' Hieditor : 2015.05.27 �̻� ����
'			 2017.07.07 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/cscenterv2/lib/incSessionAdminCS.asp" -->
<!-- #include virtual="/cscenterv2/lib/db/dbopen.asp" -->
<!-- #include virtual="/cscenterv2/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/cscenterv2/lib/function.asp"-->
<!-- #include virtual="/cscenterv2/lib/classes/order/ordercls.asp"-->
<%
dim userid, orderserial, mileage, jukyo
dim i, buf
	userid = requestCheckvar(request("userid"),32)
	orderserial = requestCheckvar(request("orderserial"),32)
	mileage = requestCheckvar(request("mileage"),32)
	jukyo = requestCheckvar(request("jukyo"),32)

if (userid = "") then
    response.write "<script>alert('�߸��� �����Դϴ�.'); history.back();</script>"
    dbget.close()	:	response.End
end if

dim page
dim ojumun

page = 1

set ojumun = new COrderMaster
ojumun.FPageSize = 10
ojumun.FCurrPage = page

ojumun.FRectUserID = userid
ojumun.FRectOrderSerial = orderserial

if (Left(orderserial,1) = "B") then
	EXCLUDE_SITENAME = "diyitem"
end if

ojumun.QuickSearchOrderList

dim ix,iy


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

%>
<script language="javascript">

function SubmitForm()
{
	if (document.frm.orderserial.value == "") {
		alert("���� �ֹ���ȣ�� �Է��ϼ���.");
		return;
	}

	/*
	if (document.frmsearch.orderserial.value != document.frm.orderserial.value) {
		alert("���� �˻��� �ϼ���.");
		return;
	}
	*/

	if (document.frm.mileage.value == "") {
		alert("�������� ��Ȯ�� �Է��ϼ���.");
		return;
	}

	if (document.frm.mileage.value*0 != 0) {
		alert("�������� ��Ȯ�� �Է��ϼ���.");
		return;
	}

	if (document.frm.mileage.value == 0) {
		alert("�������� 0 �� �� �� �����ϴ�.");
		return;
	}

	if (document.frm.jukyo.value == "") {
		alert("���������� �����ϴ�.");
		return;
	}

	if (confirm("������û �Ͻðڽ��ϱ�?") == true) {
		document.frm.submit();
	}
}

function SearchOrderSerial(orderserial)
{
	if (document.frm.orderserial.value == "") {
		alert("�ֹ���ȣ�� �Է��ϼ���.");
		return;
	}

	document.frmsearch.orderserial.value = document.frm.orderserial.value;
	document.frmsearch.submit();
}

</script>

<!-- ǥ ��ܹ� ����-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
   	<tr height="10" valign="bottom">
	        <td width="10" align="right"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
	        <td background="/images/tbl_blue_round_02.gif" colspan="2"></td>
	        <td background="/images/tbl_blue_round_02.gif" colspan="2"></td>
	        <td width="10" align="left" ><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
	</tr>

</table>
<!-- ǥ ��ܹ� ��-->


<!-- ǥ �߰��� ����-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
    <tr height="5" valign="top">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td align="left">
          <br><b>���ϸ��� ������û</b> ������û�� �Ͻø�, CSó������Ʈ�� ��ϵǸ�, ������ ���ΰ� �Բ� �����˴ϴ�.
        </td>
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
    </tr>
</table>
<!-- ǥ �߰��� ��-->


<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
  <form name="frmsearch" method="get" action="" onsubmit="return false;">
  	<input type="hidden" name="userid" value="<%= userid %>">
  	<input type="hidden" name="orderserial" value="<%= orderserial %>">
  </form>
  <form name="frm" method="post" action="domodifymileage.asp" onsubmit="return false;">
  <input type="hidden" name="mode" value="request">
  <input type="hidden" name="userid" value="<%= userid %>">
  <tr align="left">
  	<td height="30" width="20%" bgcolor="#DDDDFF">���̵� :</td>
  	<td bgcolor="#FFFFFF" width="25%" >
  	  <b><%= userid %></b>
  	</td>
  	<td height="30" width="20%" bgcolor="#DDDDFF">�����ֹ���ȣ :</td>
  	<td bgcolor="#FFFFFF"  >
  	  <input type=text name=orderserial value="<%= ResultOneOrderserial %>">
  	  <input type=button value="�˻�" onclick="SearchOrderSerial()">
  	</td>

  </tr>
  <tr align="left">
  	<td height="30" bgcolor="#DDDDFF">������ :</td>
  	<td bgcolor="#FFFFFF" >
	  <input type=text name=mileage value="<%= mileage %>">
  	</td>
  	<td height="30" bgcolor="#DDDDFF">�������� :</td>
  	<td bgcolor="#FFFFFF" >
		<select class="select" name="jukyo">
			<option value='' selected>��Ͼ���</option>
     		<option value='�Ա�����' <% if (jukyo = "�Ա�����") then %>selected<% end if %>>�Ա�����</option>
     		<option value='��ǰ����' <% if (jukyo = "��ǰ����") then %>selected<% end if %>>��ǰ����</option>
     		<option value='�������' <% if (jukyo = "�������") then %>selected<% end if %>>�������</option>
     		<option value='CS����' <% if (jukyo = "CS����") then %>selected<% end if %>>CS����</option>
     		<option value='��ǰ���ȯ��' <% if (jukyo = "��ǰ���ȯ��") then %>selected<% end if %>>��ǰ���ȯ��</option>
     		<option value='��Ÿ' <% if (jukyo = "��Ÿ") then %>selected<% end if %>>��Ÿ</option>
     	</select>
  	</td>
  </tr>
</form>
</table>

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
    </tr>
<% if (orderserial <> "") then %>
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
    </tr>
        <% next %>
	<% else %>
    <tr align="center" bgcolor="#FFFFFF">
      	<td colspan="6"> �˻��� ����� �����ϴ�.</td>
    </tr>
	<% end if %>
<% end if %>
</table>

<!-- ǥ �ϴܹ� ����-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
    <tr valign="top" height="25">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td valign="bottom" align="center">
          <input type="button" value="������û" onClick="SubmitForm();">
          <input type="button" value=" â �� �� " onClick="self.close()">
        </td>
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
    </tr>
    <tr valign="bottom" height="10">
        <td width="10" align="right"><img src="/images/tbl_blue_round_07.gif" width="10" height="10"></td>
        <td background="/images/tbl_blue_round_08.gif"></td>
        <td width="10" align="left"><img src="/images/tbl_blue_round_09.gif" width="10" height="10"></td>
    </tr>
    </form>
</table>
<!-- ǥ �ϴܹ� ��-->

<p>
<%
'set OUserInfo = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
