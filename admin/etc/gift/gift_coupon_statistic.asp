<%@ language=vbscript %>
<% option explicit %>
<%
'#######################################################
' Description : ����Ƽ��/������ 10x10��������
' History	:  ���ʻ����� ��
'              2017.07.07 �ѿ�� ����
'#######################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/etc/giftCls.asp"-->

<%
	Dim iCurrentpage, intLoop, arrList, GiftStatisticlist, GiftStatisticshortlist, i, iTotCnt1, iTotCnt, vSDate, vEDate, page, vGubun, vOrderSerial, vUserID, vUserName, vReqHP, vReqHP1, vReqHP2, vReqHP3, vTotalSum, vParam
	Dim vNoCouponNo
	vTotalSum = "x"
	iCurrentpage 	= NullFillWith(requestCheckVar(Request("iC"),10),1)
	page 			= NullFillWith(requestCheckVar(request("page"),5),1)
	vGubun			= NullFillWith(requestCheckVar(request("gubun"),10),"")
	vOrderSerial	= NullFillWith(requestCheckVar(request("orderserial"),30),"")
	vUserID			= NullFillWith(requestCheckVar(request("userid"),50),"")
	vUserName		= NullFillWith(requestCheckVar(request("username"),100),"")
	vReqHP1			= NullFillWith(requestCheckVar(request("reqhp1"),3),"")
	vReqHP2			= NullFillWith(requestCheckVar(request("reqhp2"),4),"")
	vReqHP3			= NullFillWith(requestCheckVar(request("reqhp3"),4),"")
	If vReqHP1 <> "" AND vReqHP2 <> "" AND vReqHP3 <> "" Then
		vReqHP = vReqHP1 & "-" & vReqHP2 & "-" & vReqHP3
	End If
	vSDate			= NullFillWith(requestCheckVar(request("sdate"),10),"")
	vEDate			= NullFillWith(requestCheckVar(request("edate"),10),"")
	vNoCouponNo		= NullFillWith(requestCheckVar(request("nocouponno"),1),"")
	
	vParam = "&menupos="&Request("menupos")&"&gubun="&vGubun&"&orderserial="&vOrderSerial&"&userid="&vUserID&"&username="&vUserName&"&reqhp1="&vReqHP1&"&reqhp2="&vReqHP2&"&reqhp3="&vReqHP3&"&sdate="&vSDate&"&edate="&vEDate&"&nocouponno="&vNoCouponNo&""
	
	'<!-- //-->
'	Set GiftStatisticshortlist = new ClsGift
'	GiftStatisticshortlist.FGubun = vGubun
'	GiftStatisticshortlist.FTCouponNo = vOrderSerial
'	GiftStatisticshortlist.FUserID = vUserID
'	GiftStatisticshortlist.FUSerName = vUserName
'	GiftStatisticshortlist.FReqHP = vReqHP
'	GiftStatisticshortlist.FSDate = vSDate
'	GiftStatisticshortlist.FEDate = vEDate
'	arrList = GiftStatisticshortlist.FGiftStatisticShortList
'	iTotCnt1 = GiftStatisticshortlist.ftotalcount
'	Set GiftStatisticshortlist = Nothing
	
	
	Set GiftStatisticlist = new ClsGift
	If vSDate <> "" OR vEDate <> "" Then
		vTotalSum = "o"
		GiftStatisticlist.FPageSize = "1000"
	End IF
	GiftStatisticlist.FCurrPage = page
	GiftStatisticlist.FGubun = vGubun
	GiftStatisticlist.FTCouponNo = vOrderSerial
	GiftStatisticlist.FUserID = vUserID
	GiftStatisticlist.FUSerName = vUserName
	GiftStatisticlist.FReqHP = vReqHP
	GiftStatisticlist.FSDate = vSDate
	GiftStatisticlist.FEDate = vEDate
	GiftStatisticlist.FNoCouponno = vNoCouponNo
	GiftStatisticlist.FCouponStatisticList
	
	iTotCnt = GiftStatisticlist.ftotalcount
%>

<script language="javascript">
function chkfrm()
{

	return true;
}

function popDepositLog(id)
{
	var DepositLog = window.open("/cscenter/deposit/cs_deposit.asp?userid="+id+"","DepositLog","width=950,height=600,scrollbars=yes,resizable=yes");
	DepositLog.focus();
}
</script>

<!-- ����Ʈ ���� -->
<form name="frm" method="get" action="gift_coupon_statistic.asp" onSubmit="return chkfrm()">
<input type="hidden" name="menupos" value="<%=Request("menupos")%>">
<table cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="40" bgcolor="FFFFFF">
	<td colspan="10" style="padding-right:30px;">
		���� : 
		<select name="gubun">
			<option value="">-����-</option>
			<option value="550" <%=CHKIIF(vGubun="550","selected","")%>>������</option>
			<option value="560" <%=CHKIIF(vGubun="560","selected","")%>>����Ƽ��</option>
		</select>
		&nbsp;
		Ƽ��/�� ������ȣ : 
		<input type="text" name="orderserial" value="<%=vOrderSerial%>" maxlength="30" size="15">
		&nbsp;
		���̵� : 
		<input type="text" name="userid" value="<%=vUserID%>" maxlength="50" size="15">
		&nbsp;
		ȸ���� : 
		<input type="text" name="username" value="<%=vUserName%>" maxlength="30" size="10">
		<!--
		&nbsp;
		�������ڵ��� : 
		<input type="text" name="reqhp1" value="<%=vReqHP1%>" maxlength="3" size="3">-
		<input type="text" name="reqhp2" value="<%=vReqHP2%>" maxlength="4" size="4">-
		<input type="text" name="reqhp3" value="<%=vReqHP3%>" maxlength="4" size="4">
		//-->
		<br><br>
		����� : 
		<input type="text" name="sdate" size="10" maxlength=10 readonly value="<%= vSDate %>">
		<a href="javascript:calendarOpen(frm.sdate);"><img src="/images/calicon.gif" border="0" align="absmiddle" height=21></a>
		&nbsp;~&nbsp;
		<input type="text" name="edate" size="10" maxlength=10 readonly value="<%= vEDate %>">
		<a href="javascript:calendarOpen(frm.edate);"><img src="/images/calicon.gif" border="0" align="absmiddle" height=21></a>
		&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
		<input type="checkbox" name="nocouponno" value="o" <%=CHKIIF(vNoCouponNo="o","checked","")%>>��������ȣ ���°ͱ��� �����ȸ&nbsp;&nbsp;&nbsp;
		<input type="submit" class="button" value="�� ��">
		&nbsp;
		<br><br>�� <font color="blue">�����ϸ� Ƽ��/�� ������ȣ�� ���̵� �� �� �ϳ� ������ ���� �˻��Ͻñ� �ٶ��ϴ�. �����ͷ��� ���Ƽ� �������ϴ�.</font>
	</td>
</tr>
</table>
</form>

<% If iTotCnt1 > 0 Then %>
<br>
<table cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr bgcolor="#E6E6E6">
	<td align="center">�ݾױ�</td>
	<td align="center">����</td>
	<td align="center">�Ѿ�</td>
</tr>
<%
'	IF isArray(arrList) THEN
'		For intLoop =0 To UBound(arrList,2)
%>
		<tr bgcolor="#FFFFFF">
			<td><%'arrList(0,intLoop)%></td>
			<td align="right"><%'arrList(1,intLoop)%></td>
			<td align="right"><%'FormatNumber(arrList(2,intLoop),0)%></td>
		</tr>
<%
'		Next
'	End If
%>
</tr>
</table>
<br>
<% End If %>

<table cellpadding="0" cellspacing="0" border="0" class="a">
<tr height="30">
	<td width="120">
		Total Count : <b><%= iTotCnt %></b>
	</td>
	<td width="120"></td>
	<td align="right" width="464"><input type="button" value="����������ٿ�" class="button" onClick="location.href='gift_coupon_statistic_xls.asp?1=1<%=vParam%>';"></td>
</tr>
</table>

<table cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr bgcolor="#E6E6E6">
	<td align="center">�������</td>
	<td align="center">Ƽ��/�� ������ȣ</td>
	<td align="center">�ٹ�����������ȣ</td>
	<td align="center">�ٹ�����������</td>
	<td align="center">�����ݾ�</td>
	<td align="center">UserID</td>
	<td align="center">ȸ����</td>
	<td align="center">�ֹ���ȣ</td>
	<td align="center">�����</td>
	<td align="center"></td>
</tr>
<%
	If GiftStatisticlist.FResultCount <> 0 Then
		For i = 0 To GiftStatisticlist.FResultCount -1
%>
		<tr bgcolor="FFFFFF" onmouseout="this.style.backgroundColor='#FFFFFF'" onmouseover="this.style.backgroundColor='#F1F1F1'" style="cursor:pointer">
			<td width="70" align="center"><%=GiftStatisticlist.FItemList(i).fgubun%></td>
			<td width="110" align="center"><%=GiftStatisticlist.FItemList(i).fcouponno%></td>
			<td width="110" align="center"><%=GiftStatisticlist.FItemList(i).fcouponidx%></td>
			<td width="150" align="center"><%=GiftStatisticlist.FItemList(i).fcouponname%></td>
			<td width="80" align="center"><%=GiftStatisticlist.FItemList(i).fcouponvalue%></td>
			<td width="100" align="center"><%= printUserId(GiftStatisticlist.FItemList(i).fuserid, 2, "*") %></td>
			<td width="80" align="center"><%=GiftStatisticlist.FItemList(i).fusername%></td>
			<td width="100" align="center"><%=GiftStatisticlist.FItemList(i).forderserial%></td>
			<td width="150"> <%=GiftStatisticlist.FItemList(i).fregdate %></td>
			<td><input type="button" value="��ġ�ݷα׺���" class="button" onClick="popDepositLog('<%=GiftStatisticlist.FItemList(i).fuserid%>');"></td>
		</tr>
<%
		Next
	Else
%>
		<tr bgcolor="#FFFFFF" height="30">
			<td colspan="20" align="center" class="page_link">[�����Ͱ� �����ϴ�.]</td>
		</tr>
<%
	End If
%>
<tr bgcolor="#FFFFFF">
	<td align="center" style="padding:10 0 10 0" colspan="10">
		<a href="?page=1<%=vParam%>"><img src="http://fiximage.10x10.co.kr/web2009/momo/images/btn_pageprev02.gif" width="9" height="9" border="0" /></a>
		<% if GiftStatisticlist.HasPreScroll then %>
			&nbsp;&nbsp;<a href="?page=<%= GiftStatisticlist.StartScrollPage-1 %><%=vParam%>"><img src="http://fiximage.10x10.co.kr/web2009/momo/images/btn_pageprev01.gif" width="9" height="9" border="0" /></a>
		<% else %>
			&nbsp;&nbsp;<img src="http://fiximage.10x10.co.kr/web2009/momo/images/btn_pageprev01.gif" width="9" height="9" border="0" />
		<% end if %>																												
		<% 
		for i = 0 + GiftStatisticlist.StartScrollPage to GiftStatisticlist.StartScrollPage + GiftStatisticlist.FScrollCount - 1 
		if (i > GiftStatisticlist.FTotalpage) then Exit for 
		if CStr(i) = CStr(GiftStatisticlist.FCurrPage) then 
		%>
			&nbsp;&nbsp;&nbsp;&nbsp;<span class="eng11pxblack"><b><%= i %></b></span>
		<% else %>
			&nbsp;&nbsp;&nbsp;&nbsp;<a href="?page=<%= i %><%=vParam%>" style="cursor:pointer"><%= i %></a>
		<% 
		end if 
		next 
		%>													
		<% if GiftStatisticlist.HasNextScroll then %>
			&nbsp;&nbsp;<span class="list_link"><a href="?page=<%= i %><%=vParam%>"><img src="http://fiximage.10x10.co.kr/web2009/momo/images/btn_pagenext01.gif" width="9" height="9" border="0" /></a>
		<% else %>
			&nbsp;&nbsp;<img src="http://fiximage.10x10.co.kr/web2009/momo/images/btn_pagenext01.gif" width="9" height="9" border="0" />
		<% end if %>																												
		&nbsp;&nbsp;&nbsp;<a href="?page=<%= GiftStatisticlist.FTotalpage %><%=vParam%>" onfocus="this.blur();"><img src="http://fiximage.10x10.co.kr/web2009/momo/images/btn_pagenext02.gif" width="9" height="9" border="0" /></a>
	</td>
</tr>
</table>

<%
	set GiftStatisticlist = nothing
%>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->