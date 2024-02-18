<%@ language=vbscript %>
<% option explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
%>
<%
'####################################################
' Description :  ���ʽ� ����
' History : 2011.05.12 �ѿ�� ����
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/offshop/bonuscoupon/bonuscoupondetail_cls.asp" -->
<%
dim ocoupon, page, lp
dim cusUserid, regUserid, couponname, coupontype, usingyn, orderserial, chkOld
	cusUserid = requestCheckVar(request("cusUserid"),32)
	regUserid = requestCheckVar(request("regUserid"),32)
	couponname = requestCheckVar(request("couponname"),128)
	coupontype = requestCheckVar(request("coupontype"),10)
	usingyn = requestCheckVar(request("usingyn"),1)
	orderserial = requestCheckVar(request("orderserial"),16)
	chkOld = requestCheckVar(request("chkOld"),1)
	page = requestCheckVar(request("page"),10)

if page="" then page=1

set ocoupon = new CCouponMaster
ocoupon.FPageSize=60
ocoupon.FCurrPage = page
ocoupon.FrectCusUserid = cusUserid
ocoupon.FrectRegUserid = regUserid
ocoupon.FrectCouponname = couponname
ocoupon.FrectCoupontype = coupontype
ocoupon.FrectUsingyn = usingyn
ocoupon.FrectOrderserial = orderserial
ocoupon.FrectChkOld = chkOld
ocoupon.GetEventCouponList


dim i
%>
<script language="javascript">
<!--
	function goPage(pg) {
		frm.page.value=pg;
		frm.submit();
	}

	function newCoupon() {
		location.href="event_coupon_edit_off.asp";
	}

	function msgOldDB(chk) {
		if(chk.checked) {
			alert("3���� ���� ���� �˻��� DB�� ���� ���ϸ� �� �� �ְ� �˻��ð��� �����ɸ��ϴ�.\n�� �ʿ��� ��쿡�� üũ���ֽʽÿ�.");
		}
	}

	function chgUsing(fm) {
		if(fm.value=='N') {
			frm.orderserial.disabled=true;
			frm.orderserial.className="text_ro";
		} else {
			frm.orderserial.disabled=false;
			frm.orderserial.className="text";
		}
	}
//-->
</script>
<!-- �˻� ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="#999999">
	<form name="frm" method="get" action="">
	<input type="hidden" name="menupos" value="<%=menupos%>">
	<input type="hidden" name="page" value="1">
	<tr align="center" bgcolor="#FFFFFF" >
		<td rowspan=2 width="50" bgcolor="#EEEEEE">�˻�<br>����</td>
		<td align="left">
    		��ID : <input type="text" class="text" name="cusUserid" value="<%=cusUserid%>" size="12" maxlength="32"> &nbsp;
    		�߱���ID : <input type="text" class="text" name="regUserid" value="<%=regUserid%>" size="12" maxlength="32"> &nbsp;
			������ : <input type="text" class="text" name="couponname" value="<%=couponname%>" size="20" maxlength="20"> &nbsp;
			/ <label><input type="checkbox" name="chkOld" value="Y" <%=chkIIF(chkOld="Y","checked","")%> onclick="msgOldDB(this)"> 3���� ���� �˻�</label>
		</td>
		<td rowspan=2 width="50" bgcolor="#EEEEEE">
			<input type="button" class="button_s" value="�˻�" onClick="javascript:document.frm.submit();">
		</td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF" >
		<td align="left">
	     	�������� :&nbsp;
			<select class="select" name="coupontype">
			<option value="">��ü</option>
			<option value="1">%����</option>
			<option value="2">������</option>
			</select> &nbsp; &nbsp; &nbsp;
	     	������뿩�� :
			<select class="select" name="usingyn" onchange="chgUsing(this)">
			<option value="">��ü</option>
			<option value="Y">�����</option>
			<option value="N">������</option>
			</select>&nbsp;
			�ֹ���ȣ : <input type="text" class="<%=chkIIF(usingyn="N","text_ro","text")%>" name="orderserial" value="<%=orderserial%>" size="18" maxlength="16"> &nbsp;
			<script language="javascript">
			document.frm.coupontype.value="<%=coupontype%>";
			document.frm.usingyn.value="<%=usingyn%>";
			</script>
		</td>
	</tr>
	</form>
</table>
<!-- �˻� �� -->

<!-- �׼� ���� -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding:10px 0 5px 0;">
<tr>
	<td align="center"><font color=red>�׽�Ʈ���Դϴ�.</font></td>
	<td align="right"><input type="button" class="button" value="�űԵ��" onClick="newCoupon()"></td>
</tr>
</table>
<!-- �׼� �� -->

<!-- ����Ʈ ���� -->
<table width="100%" border="0" cellpadding="3" cellspacing="1" bgcolor="#B2B2B2" class="a">
<% if ocoupon.FResultCount>0 then %>
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="12">
			�˻���� : <b><%= formatNumber(ocoupon.FTotalCount,0) %></b>
			&nbsp;
			������ : <b><%= formatNumber(page,0) %>/ <%= formatNumber(ocoupon.FTotalPage,0) %></b>
		</td>
	</tr>
	<tr bgcolor="#E6E6E6" height=30>
		<td width="50" align="center">IDX</td>
		<td width="50" align="center">Master<br>IDX</td>
		<td align="center">���̵�</td>
		<td align="center">���ʽ�����</td>
		<td align="center">��밡��<br>��ǰ</td>
		<td align="center">��밡��<br>�귣��</td>
		<td width="150" align="center">��� ����</td>
		<td width="50" align="center">�ּұ��� �ݾ�</td>
		<td width="150" align="center">��ȿ�Ⱓ</td>
		<td width="80" align="center">�����</td>
		<td width="30" align="center">��� ����</td>
		<td width="100" align="center">�߱���</td>
	</tr>
	<% for i=0 to ocoupon.FResultCount - 1 %>
	<tr bgcolor="#FFFFFF" height=30>
		<td align="center"><%= ocoupon.FItemList(i).FIdx %></td>
		<td align="center"><%= ocoupon.FItemList(i).Fmasteridx %></td>
		<td align="center"><%= ocoupon.FItemList(i).Fuserid %></td>
		<td><%= ocoupon.FItemList(i).Fcouponname %></td>
		<td align="center"><%= ocoupon.FItemList(i).Ftargetitemlist %></td>
		<td align="center"><%= ocoupon.FItemList(i).Ftargetbrandlist %></td>
		<td align="center"><%= ocoupon.FItemList(i).getCouponTypeStr %></td>
		<td align="center"><%= FormatNumber(ocoupon.FItemList(i).Fminbuyprice, 0) %></td>
		<td align="center"><%= ocoupon.FItemList(i).getAvailDateStr %></td>
		<td align="center"><%= Formatdatetime(ocoupon.FItemList(i).FRegDate,2) %></td>
		<td align="center"><%= ocoupon.FItemList(i).FIsUsing %></td>
		<td align="center"><%= ocoupon.FItemList(i).Freguserid %></td>
	</tr>
	<% next %>
<% else %>
	<tr bgcolor="#FFFFFF">
		<td colspan="12" align="center" class="page_link">[�˻������ �����ϴ�.]</td>
	</tr>
<% end if %>
<tr height="25" bgcolor="FFFFFF">
	<td colspan="12" align="center">
		<% if ocoupon.HasPreScroll then %>
			<a href="javascript:goPage(<%= ocoupon.StartScrollPage-1 %>)">[pre]</a>
		<% else %>
			[pre]
		<% end if %>

		<% for lp=0 + ocoupon.StartScrollPage to ocoupon.FScrollCount + ocoupon.StartScrollPage - 1 %>
			<% if lp>ocoupon.FTotalpage then Exit for %>
			<% if CStr(page)=CStr(lp) then %>
			<font color="red">[<%= lp %>]</font>
			<% else %>
			<a href="javascript:goPage(<%= lp %>)">[<%= lp %>]</a>
			<% end if %>
		<% next %>

		<% if ocoupon.HasNextScroll then %>
			<a href="javascript:goPage(<%= lp %>)">[next]</a>
		<% else %>
			[next]
		<% end if %>
	</td>
</tr>
</table>
<!-- ����Ʈ �� -->
<%
set ocoupon = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->