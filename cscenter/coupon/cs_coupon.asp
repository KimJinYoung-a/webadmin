<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : cs���� ��������
' History : �̻󱸻���
'			2023.05.23 �ѿ�� ����(���� üũ �߰�)
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbHelper.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/cscenter/cs_couponcls.asp" -->
<!-- #include virtual="/lib/classes/cscenter/sp_itemcouponcls.asp" -->
<!-- #include virtual="/lib/classes/cscenter/cs_aslistcls.asp" -->
<%
dim i, userid, onlyavaiable, excludedelete, searchtype, reguserid, oitemcoupon, ocscoupon, totay, expireday, baseday, daybeforeonemonth
	userid = requestcheckvar(request("userid"),32)
	reguserid = requestcheckvar(request("reguserid"),32)
	onlyavaiable = requestcheckvar(request("onlyavaiable"),1)
	excludedelete = requestcheckvar(request("excludedelete"),1)
	searchtype = requestcheckvar(request("searchtype"),16)

if (userid = "" and reguserid = "") then
	onlyavaiable = "Y"
	excludedelete = "Y"
end if

if searchtype = "" then searchtype = "all"

'��ǰ����
set oitemcoupon = new CUserItemCoupon
	oitemcoupon.FRectUserID = userid
	oitemcoupon.FRectAvailableYN = onlyavaiable
	oitemcoupon.FRectDeleteYN = excludedelete
	oitemcoupon.FPageSize = 50
	oitemcoupon.FCurrPage = 1

	if userid<>"" and (searchtype = "all" or searchtype = "item") then
		oitemcoupon.GetCouponList
	end if

'���ʽ�����
set ocscoupon = New CCSCenterCoupon
	ocscoupon.FRectExcludeUnavailable = onlyavaiable
	ocscoupon.FRectExcludeDelete = excludedelete
	ocscoupon.FRectUserID = userid
	ocscoupon.FRectRegUserID = reguserid

	if (userid<>"" or reguserid<>"") and (searchtype = "all" or searchtype = "bonus") then
		ocscoupon.GetCSCenterCouponList
	end if

totay = Left(now(), 10)
daybeforeonemonth = Left(DateAdd("d", -30, totay), 10)

%>
<script type='text/javascript'>

function openWindowModifyCoupon(coupontype, couponidx){
	var w = window.open("/cscenter/coupon/pop_coupon_modify.asp?coupontype=" + coupontype + "&couponidx=" + couponidx,"openWindowModifyCoupon","width=1400 height=600 scrollbars=yes resizable=yes");
	w.focus();
}

function openWindowCopyCoupon(coupontype, couponidx){
	var w = window.open("/cscenter/coupon/pop_coupon_copy.asp?coupontype=" + coupontype + "&couponidx=" + couponidx,"openWindowCopyCoupon","width=1400 height=600 scrollbars=yes resizable=yes");
	w.focus();
}

function popCouponIssue(userid) {
	if (userid == "") {
		alert("���̵� �����ϴ�.");
		return;
	}

    var popwin = window.open('/cscenter/coupon/pop_coupon_issue.asp?userid=' + userid,'popCouponIssue','width=1400,height=800,scrollbars=yes,resizable=yes');
    popwin.focus();
}

function PopItemCouponAssginList(itemcouponidx){
	var popwin = window.open('/admin/shopmaster/itemcouponitemlisteidt.asp?itemcouponidx=' + itemcouponidx,'EditCouponItemList','width=1400,height=800,scrollbars=yes,resizable=yes');
	popwin.focus();
}

</script>

<!-- �˻� ���� -->
<form name="frm" method="get" action="" style="margin:0px;">
<input type="hidden" name="menupos" value="<%= menupos %>">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
	<td align="left">
		* �� ���̵� : <input type="text" class="text" name="userid" value="<%= userid %>" onKeyPress="if (event.keyCode == 13) document.frm.submit();">
		&nbsp;
		* ����
		<select class="select" name="searchtype">
		<option value='all' <% if searchtype = "all" then %>selected<% end if %>>��ü</option>
		<option value='item' <% if searchtype = "item" then %>selected<% end if %>>��ǰ����</option>
		<option value='bonus' <% if searchtype = "bonus" then %>selected<% end if %>>���ʽ�����</option>
		</select>

		&nbsp;
		<input type="checkbox" name="onlyavaiable" value="Y" <% if (onlyavaiable = "Y") then %>checked<% end if %>>��ȿ�Ⱓ�� ������ ǥ��
		&nbsp;
		<input type="checkbox" name="excludedelete" value="Y" <% if (excludedelete = "Y") then %>checked<% end if %>>�������� ����
	</td>

	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="�˻�" onclick="document.frm.submit()">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
		* ����� ���̵� : <input type="text" class="text" name="reguserid" value="<%= reguserid %>" onKeyPress="if (event.keyCode == 13) document.frm.submit();">
	</td>
</tr>
</table>
</form>

<br>
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
	<tr>
		<td>
			<% if userid<>"" then %>
				<input type="button" class="button" value=" �� �� �� �� " onclick="popCouponIssue('<%= userid %>');">
				���������� �ϸ� CS�޸� ��ϵǸ�, ��� ����˴ϴ�.
			<% else %>
				���������� �� ���̵� �˻��� �����մϴ�.
			<% end if %>
		</td>
	</tr>
</table>

<% if userid<>"" and (searchtype = "all" or searchtype = "item") then %>
	<br>
	<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="11" align="left">
			��ǰ����(<%= oitemcoupon.FTotalCount %>)
		</td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td width="70">����������<br>�ڵ�</td>
		<td width="80">���̵�</td>
		<td>������</td>
		<td width="100">�������</td>
		<td width="80">�����</td>
		<td width="170">��ȿ�Ⱓ</td>
		<td width="80">�ش�귣��<br> �� ��ǰ</td>
		<td width="30">���</td>
		<td width="90">�ֹ���ȣ</td>
		<td width="30">��ȿ</td>
		<td width="30">����</td>
	</tr>
	<% if (oitemcoupon.FResultCount > 0) then %>
		<% for i = 0 to (oitemcoupon.FResultCount - 1) %>
		<tr align="center" bgcolor="#FFFFFF">
			<td><%= oitemcoupon.FItemList(i).Fitemcouponidx %></td>
			<td><%= oitemcoupon.FItemList(i).Fuserid %></td>
			<td align="left"><%= oitemcoupon.FItemList(i).Fitemcouponname %></td>

			<td align="left"><%= oitemcoupon.FItemList(i).GetDiscountStr %></td>
			<td>
				<% if oitemcoupon.FItemList(i).Fregdate<>"" and not(isnull(oitemcoupon.FItemList(i).Fregdate)) then %>
					<%= left(oitemcoupon.FItemList(i).Fregdate,10) %>
					<Br><%= mid(oitemcoupon.FItemList(i).Fregdate,12,20) %>
				<% end if %>
			</td>
			<td>
				<%'= oitemcoupon.FItemList(i).getAvailDateStr %>
				<acronym title="<%= oitemcoupon.FItemList(i).Fitemcouponstartdate %>"><%= Left(oitemcoupon.FItemList(i).Fitemcouponstartdate,10) %></acronym> ~ <acronym title="<%= oitemcoupon.FItemList(i).Fitemcouponexpiredate %>"><%= Left(oitemcoupon.FItemList(i).Fitemcouponexpiredate,10) %></acronym>
			</td>
			<td align="left"><a href="javascript:PopItemCouponAssginList('<%= oitemcoupon.FItemList(i).FitemcouponIdx %>');" class="link_ctleftred">�����ǰ����</a></td>
			<td><%= oitemcoupon.FItemList(i).Fusedyn %></td>
			<td><%= oitemcoupon.FItemList(i).Forderserial %></td>
			<td><%= oitemcoupon.FItemList(i).Fisavailable %></td>
			<td><%= oitemcoupon.FItemList(i).Fdeleteyn %></td>
		</tr>
		<% next %>
	<% else %>
		<tr bgcolor="#FFFFFF" align="center">
			<td height="25" colspan="11">�˻������ �����ϴ�.</td>
		</tr>
	<% end if %>
	</table>
<% end if %>

<% if (userid<>"" or reguserid<>"") and (searchtype = "all" or searchtype = "bonus") then %>
	<br>
	<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="20" align="left">
			���ʽ�����(<%= ocscoupon.FResultCount %>)
		</td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td width="70">����������<br>�ڵ�</td>
		<td width="80">���̵�</td>
		<td>������</td>
		<td width="70">���ΰ�</td>
		<td width="70">�ּұ���<br>�ݾ�</td>
		<td width="70">�ִ�����<br>�ݾ�</td>
		<td width="80">�����</td>
		<td width="170">��ȿ�Ⱓ</td>
		<td width="90">����ֹ���ȣ</td>
		<td width="90">CS�ֹ���ȣ</td>
		<td width="80">��ǰCODE</td>
		<td width="70">����</td>
		<td width="30">���</td>
		<td width="30">����</td>
		<td width="80">�����</td>
	</tr>
	<% if (ocscoupon.FResultCount > 0) then %>
		<% for i = 0 to (ocscoupon.FResultCount - 1) %>
		<tr align="center" <% if (ocscoupon.FItemList(i).Fdeleteyn = "Y") then %>bgcolor="#EEEEEE" class="gray"<% else %>bgcolor="#FFFFFF"<% end if %>>
			<td><%= ocscoupon.FItemList(i).Fmasteridx %></td>
			<td><%= ocscoupon.FItemList(i).Fuserid %></td>
			<td align="left"><%= ocscoupon.FItemList(i).Fcouponname %></td>
			<% if ocscoupon.FItemList(i).Fcoupontype="3" then %>
			<td align="right">������</td>
			<% else %>
			<td align="right"><%= FormatNumber(ocscoupon.FItemList(i).Fcouponvalue,0) %><%= ocscoupon.FItemList(i).GetCouponTypeUnit %></td>
			<% end if %>
			<td align="right"><%= FormatNumber(ocscoupon.FItemList(i).Fminbuyprice,0) %></td>
			<% if ocscoupon.FItemList(i).Fcoupontype="1" then %>
			<td align="right"><%= chkIIF(ocscoupon.FItemList(i).FmxCpnDiscount=0,"-",FormatNumber(ocscoupon.FItemList(i).FmxCpnDiscount,0)&"��") %></td>
			<% else %>
			<td align="right">-</td>
			<% end if %>
			<td>
				<% if ocscoupon.FItemList(i).Fregdate<>"" and not(isnull(ocscoupon.FItemList(i).Fregdate)) then %>
					<%= left(ocscoupon.FItemList(i).Fregdate,10) %>
					<Br><%= mid(ocscoupon.FItemList(i).Fregdate,12,20) %>
				<% end if %>
				<!--<acronym title="<%'= ocscoupon.FItemList(i).Fregdate %>"><%'= Left(ocscoupon.FItemList(i).Fregdate,10) %></acronym>-->
			</td>
			<td><acronym title="<%= ocscoupon.FItemList(i).Fstartdate %>"><%= Left(ocscoupon.FItemList(i).Fstartdate,10) %></acronym> ~ <acronym title="<%= ocscoupon.FItemList(i).Fexpiredate %>"><%= Left(ocscoupon.FItemList(i).Fexpiredate,10) %></acronym></td>
			<td><%= ocscoupon.FItemList(i).Forderserial %></td>
			<td><%= ocscoupon.FItemList(i).Fcsorderserial %></td>
			<td><%= ocscoupon.FItemList(i).Fexitemid %></td>
			<td>
			<% if isNull(ocscoupon.FItemList(i).FuseLevel) or ocscoupon.FItemList(i).FuseLevel<>"7" then %>
				<% if (ocscoupon.FItemList(i).Fisusing <> "Y") and (ocscoupon.FItemList(i).Fdeleteyn <> "Y") and (daybeforeonemonth <= Left(ocscoupon.FItemList(i).Fexpiredate,10)) then %>
				<input type=button class="button" value="�Ⱓ����" onclick="openWindowModifyCoupon('bonus', <%= ocscoupon.FItemList(i).Fidx %>)">
				<% end if %>
				<% if (ocscoupon.FItemList(i).Fisusing = "Y") and (ocscoupon.FItemList(i).Fdeleteyn <> "Y") and (daybeforeonemonth <= Left(ocscoupon.FItemList(i).Fexpiredate,10)) then %>
				<input type=button class="button" value="�������" onclick="openWindowCopyCoupon('bonus', <%= ocscoupon.FItemList(i).Fidx %>)">
				<% end if %>
			<% end if %>
			</td>
			<td><% if (ocscoupon.FItemList(i).Fisusing = "Y") then %>���<% end if %></td>
			<td><% if (ocscoupon.FItemList(i).Fdeleteyn = "Y") then %>����<% end if %></td>
			<td><%= ocscoupon.FItemList(i).Freguserid %></td>
		</tr>
		<% next %>
	<% else %>
		<tr bgcolor="#FFFFFF" align="center">
			<td height="25" colspan="20">�˻������ �����ϴ�.</td>
		</tr>
	<% end if %>
	</table>
<% end if %>

<%
set ocscoupon = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
