<%@ language=vbscript %>
<%
option explicit
Response.Expires = -1
%>
<%
'###########################################################
' Description : �������� ���
' Hieditor : 2011.02.24 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/md5.asp"-->
<!-- #include virtual="/common/checkPoslogin.asp"-->
<!-- #include virtual="/common/incSessionAdminorShop.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/common/lib/commonbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<% '<!-- #include virtual="/lib/checkAllowIPWithLog.asp" --> %>
<!-- #include virtual="/lib/classes/offshop/upche/upchebeasong_cls.asp" -->

<%
dim i , orderno, ojumun, shopid, showshopselect
dim yyyy1,mm1,dd1,yyyy2,mm2,dd2, yyyymmdd1,yyymmdd2, fromDate,toDate
	orderno = requestcheckvar(request("orderno"),16)
	shopid = requestcheckvar(request("shopid"),32)
	yyyy1 = requestCheckVar(request("yyyy1"),4)
	mm1 = requestCheckVar(request("mm1"),2)
	dd1 = requestCheckVar(request("dd1"),2)
	yyyy2 = requestCheckVar(request("yyyy2"),4)
	mm2 = requestCheckVar(request("mm2"),2)
	dd2 = requestCheckVar(request("dd2"),2)

showshopselect = false

if (yyyy1="") then
	fromDate = DateSerial(Cstr(Year(now())), Cstr(Month(now())), Cstr(day(now()))-30)
else
	fromDate = DateSerial(yyyy1, mm1, dd1)
end if

if (yyyy2="") then yyyy2 = Cstr(Year(now()))
if (mm2="") then mm2 = Cstr(Month(now()))
if (dd2="") then dd2 = Cstr(day(now()))

toDate = DateSerial(yyyy2, mm2, dd2+1)
yyyy1 = left(fromDate,4)
mm1 = Mid(fromDate,6,2)
dd1 = Mid(fromDate,9,2)

if C_ADMIN_USER or C_IS_OWN_SHOP then
	showshopselect = true
	shopid = request("shopid")
elseif (C_IS_SHOP) then
	'����/������
	shopid = C_STREETSHOPID
else
	if (C_IS_Maker_Upche) then
		shopid = session("ssBctID")
	else
		if (Not C_ADMIN_USER) then
			loginidshopormaker = "--"		'ǥ�þ��Ѵ�. ����.
		else
			showshopselect = true
			shopid = request("shopid")
		end if
	end if
end if

set ojumun = new cupchebeasong_list
	ojumun.frectorderno = orderno
	ojumun.frectshopid = shopid
	'ojumun.FRectStartDay = fromDate
	'ojumun.FRectEndDay = toDate
	ojumun.FPageSize = 500
	ojumun.FCurrPage = 1
	ojumun.fshopbeasong_cancel()

%>

<script type="text/javascript">

	//������
	function gosubmit(){
		frm.submit();
	}

</script>

<!-- �˻� ���� -->
<form name="frm" method="post" action="">
<input type="hidden" name="itemgubunarr">
<input type="hidden" name="itemidarr">
<input type="hidden" name="itemoptionarr">
<input type="hidden" name="shopidarr">
<input type="hidden" name="masteridxarr">
<input type="hidden" name="mode">
<input type="hidden" name="masteridx">
<input type="hidden" name="odlvTypearr">
<input type="hidden" name="detailidxarr">
<input type="hidden" name="menupos" value="<%= menupos %>">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
	<td align="left">
		<!--* ����Է��� : --><% 'DrawDateBox yyyy1,yyyy2,mm1,mm2,dd1,dd2 %>
		* ���� :
		<% if (showshopselect = true) then %>
			<% 'drawSelectBoxOffShop "shopid",shopid %>
			<% Call NewDrawSelectBoxDesignerwithNameAndUserDIV("shopid",shopid, "21") %>
		<% else %>
			<%= shopid %>
		<% end if %>
		&nbsp;&nbsp;
		* �ֹ���ȣ : <input type="text" name="orderno" value="<%= orderno %>" size="16" onKeyPress="if(window.event.keyCode==13) gosubmit('');">
	</td>
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="�˻�" onClick="gosubmit('');">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
		
	</td>
</tr>
</table>
<!-- �˻� �� -->
<br>
<!-- �׼� ���� -->
<form name="frminfo" method="post" style="margin:0px;">
<input type="hidden" name="menupos" value="<%= menupos %>">
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a">
<tr>
	<td align="left">
		<input type="button" value="���������ΰ�ħ" class="button" onclick="location.reload(); return false;">
	</td>
	<td align="right"></td>
</tr>
</table>
</form>
<!-- �׼� �� -->

<!-- ����Ʈ ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="25">
		�˻���� : <b><%= ojumun.FTotalCount %></b> &nbsp; �� 500 �� ���� �˻� �˴ϴ�.
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td>�����</td>
	<td>���ֹ���ȣ</td>
	<td>��ǰ�ڵ�</td>
	<td>�귣��ID</td>
	<td>��ǰ��[�ɼǸ�]</td>
	<td>����</td>
	<td>���������</td>
	<td>��ۿ�û��</td>
	<td>�����</td>
	<td>��ۻ���</td>
	<td>��������</td>
	<td>���</td>
</tr>
<% if ojumun.FTotalCount>0 then %>
<%
for i=0 to ojumun.FTotalCount-1
%>
<tr align="center" bgcolor="#FFFFFF">
	<td>
		<%= ojumun.FItemList(i).fshopname %>
	</td>
	<td>
		<%= ojumun.FItemList(i).forderno %>
	</td>
	<td>
		<%=ojumun.FItemList(i).fitemgubun%>-<%=CHKIIF(ojumun.FItemList(i).fitemid>=1000000,Format00(8,ojumun.FItemList(i).fitemid),Format00(6,ojumun.FItemList(i).fitemid))%>-<%=ojumun.FItemList(i).fitemoption%>
	</td>
	<td>
		<%=ojumun.FItemList(i).fmakerid%>
	</td>
	<td align="left">
		<%= ojumun.FItemList(i).fitemname %>

		<% if ojumun.FItemList(i).fitemoptionname<>"" then %>
			[<%= ojumun.FItemList(i).fitemoptionname %>]
		<% end if %>
	</td>
	<td><%= ojumun.FItemList(i).fitemno %></td>
	<td>
		<%= ojumun.FItemList(i).getbeasonggubun %>
	</td>
	<td>
		<%= ojumun.FItemList(i).fregdate %>
	</td>
	<td>
		<%= ojumun.FItemList(i).fbeasongdate %>
	</td>
	<td>
		<font color="<%= ojumun.FItemList(i).shopNormalUpcheDeliverStateColor %>">
			<%= ojumun.FItemList(i).shopNormalUpcheDeliverState %>
		</font>
	</td>
	<td>
		<% if (ojumun.FItemList(i).Fsongjangno <> "") then %>
			<a href="<%= fnGetSongjangURL(ojumun.FItemList(i).Fsongjangdiv) %><%= ojumun.FItemList(i).Fsongjangno %>" onfocus="this.blur()" target="_blink">
			<br>[<%= DeliverDivCd2Nm(ojumun.FItemList(i).Fsongjangdiv) %>]<%= ojumun.FItemList(i).Fsongjangno %></a>
		<% end if %>
	</td>
	<td>
	</td>
</tr>
</form>
<%
next
%>
<% else %>
	<tr bgcolor="#FFFFFF">
		<td colspan="25" align="center" class="page_link">[�˻������ �����ϴ�.]</td>
	</tr>
<% end if %>
</table>

<%
set ojumun = nothing
%>
<!-- #include virtual="/common/lib/commonbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->