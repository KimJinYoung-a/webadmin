<%@ language=vbscript %>
<% option explicit %>
<%
'#######################################################
' Description : ����Ƽ��/������ �ݾױǳ���
' History	:  ���ر� ����
'              2023.05.23 �ѿ�� ����(�����ٿ�ε� ����¡��ĺ����ؼ� ��ü �ٿ�ε� �����ϰ� ������. �ҽ� ǥ�ؼҽ��� ����.)
'#######################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/etc/giftCls.asp"-->

<%
Dim iCurrentpage, intLoop, arrList, GiftStatisticlist, GiftStatisticshortlist, i, iTotCnt1, iTotCnt, vSDate, vEDate, page
dim vGubun, vOrderSerial, vUserID, vUserName, vReqHP, vReqHP1, vReqHP2, vReqHP3, vTotalSum
	vTotalSum = "x"
	iCurrentpage 	= NullFillWith(requestCheckVar(Request("iC"),10),1)
	page = requestCheckVar(getNumeric(request("page")),10)
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

if page = "" then page = 1

	Set GiftStatisticshortlist = new ClsGift
	GiftStatisticshortlist.FGubun = vGubun
	GiftStatisticshortlist.FOrderSerial = vOrderSerial
	GiftStatisticshortlist.FUserID = vUserID
	GiftStatisticshortlist.FUSerName = vUserName
	GiftStatisticshortlist.FReqHP = vReqHP
	GiftStatisticshortlist.FSDate = vSDate
	GiftStatisticshortlist.FEDate = vEDate
	arrList = GiftStatisticshortlist.FGiftStatisticShortList
	iTotCnt1 = GiftStatisticshortlist.ftotalcount
	Set GiftStatisticshortlist = Nothing
	
	
	Set GiftStatisticlist = new ClsGift
	If vSDate <> "" OR vEDate <> "" Then
		vTotalSum = "o"
		GiftStatisticlist.FPageSize = "1000"
	End IF
	GiftStatisticlist.FCurrPage = page
	GiftStatisticlist.FGubun = vGubun
	GiftStatisticlist.FOrderSerial = vOrderSerial
	GiftStatisticlist.FUserID = vUserID
	GiftStatisticlist.FUSerName = vUserName
	GiftStatisticlist.FReqHP = vReqHP
	GiftStatisticlist.FSDate = vSDate
	GiftStatisticlist.FEDate = vEDate
	GiftStatisticlist.FGiftStatisticList
	
	iTotCnt = GiftStatisticlist.ftotalcount
%>
<script type="text/javascript">

function chkfrm()
{
	if(frm.reqhp1.value != "")
	{
		if(frm.reqhp2.value == "" || frm.reqhp3.value == "")
		{
			alert("�ڵ��� ��ȣ�� ��� �Է��ϼ���.");
			return false;
		}
	}
	return true;
}

function frmsubmit(page){
	frm.page.value=page;
	frm.submit();
}

function downloadexcel(){
	alert('Ŭ���� ��ٷ� �ּ���. �˻����� ������� �����ϴ�.');
	document.frm.target = "view";
	document.frm.action = "/admin/etc/gift/gift_giftcard_statistic_xls.asp";
	document.frm.submit();
	document.frm.target = "";
	document.frm.action = "";
}

</script>

<!-- �˻� ���� -->
<form name="frm" method="get" action="/admin/etc/gift/gift_giftcard_statistic.asp" onSubmit="return chkfrm()" style="margin:0px;">
<input type="hidden" name="menupos" value="<%=Request("menupos")%>">
<input type="hidden" name="page" value="<%= page %>">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr align="center" bgcolor="#FFFFFF" >
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
		<td align="left">
			* ���� : 
			<select name="gubun">
				<option value="">-����-</option>
				<option value="550" <%=CHKIIF(vGubun="550","selected","")%>>������</option>
				<option value="560" <%=CHKIIF(vGubun="560","selected","")%>>����Ƽ��</option>
			</select>
			&nbsp;
			* ����� : 
			<input type="text" name="sdate" size="10" maxlength=10 value="<%= vSDate %>">
			<a href="javascript:calendarOpen(frm.sdate);"><img src="/images/calicon.gif" border="0" align="absmiddle" height=21></a>
			&nbsp;~&nbsp;
			<input type="text" name="edate" size="10" maxlength=10 value="<%= vEDate %>">
			<a href="javascript:calendarOpen(frm.edate);"><img src="/images/calicon.gif" border="0" align="absmiddle" height=21></a>		
		</td>	
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button_s" value="�˻�" onclick="frmsubmit('1');" >
		</td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF" >
		<td align="left">
			* �ֹ���ȣ : 
			<input type="text" name="orderserial" value="<%=vOrderSerial%>" maxlength="30" size="15">
			&nbsp;
			* ���̵� : 
			<input type="text" name="userid" value="<%=vUserID%>" maxlength="50" size="15">
			&nbsp;
			* �����θ� : 
			<input type="text" name="username" value="<%=vUserName%>" maxlength="30" size="10">
			&nbsp;
			* �������ڵ��� : 
			<input type="text" name="reqhp1" value="<%=vReqHP1%>" maxlength="3" size="3">-
			<input type="text" name="reqhp2" value="<%=vReqHP2%>" maxlength="4" size="4">-
			<input type="text" name="reqhp3" value="<%=vReqHP3%>" maxlength="4" size="4">
		</td>
	</tr>
</table>
</form>
<!-- �˻� �� -->
		
<% If iTotCnt1 > 0 Then %>
<br>
<table cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr bgcolor="#E6E6E6">
	<td align="center">�ݾױ�</td>
	<td align="center">����</td>
	<td align="center">�Ѿ�</td>
</tr>
<%
	IF isArray(arrList) THEN
		For intLoop =0 To UBound(arrList,2)
%>
		<tr bgcolor="#FFFFFF">
			<td><%=arrList(0,intLoop)%></td>
			<td align="right"><%=arrList(1,intLoop)%></td>
			<td align="right"><%=FormatNumber(arrList(2,intLoop),0)%></td>
		</tr>
<%
		Next
	End If
%>
</tr>
</table>
<% End If %>

<!-- �׼� ���� -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
<tr>
	<td align="left">
	</td>
	<td align="right">
		<input type="button" onclick="downloadexcel();" value="�����ٿ�ε�" class="button">
	</td>
</tr>
</table>
<!-- �׼� �� -->

<table cellpadding="3" cellspacing="1" width="100%" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="9">
		�˻���� : <b><%= GiftStatisticlist.ftotalcount %></b>
		&nbsp;
		������ : <b><%= page %>/ <%= GiftStatisticlist.FTotalPage %></b>
		&nbsp;
		<% If vTotalSum = "o" Then %>�ѱݾ� : <b><%=FormatNumber(GiftStatisticlist.FTotalSum,0)%></b><% End If %>
	</td>
</tr>
<tr bgcolor="#E6E6E6">
	<td align="center">�������</td>
	<td align="center">Ƽ��/�� ������ȣ</td>
	<td align="center">UserID</td>
	<td align="center">������</td>
	<td align="center">ī���</td>
	<td align="center">�ǸŰ�</td>
	<td align="center">�ǰ�����</td>
	<td align="center">ī�����</td>
	<td align="center">�����</td>
</tr>
<%
	If GiftStatisticlist.FResultCount <> 0 Then
		For i = 0 To GiftStatisticlist.FResultCount -1
%>
		<tr bgcolor="FFFFFF" onmouseout="this.style.backgroundColor='#FFFFFF'" onmouseover="this.style.backgroundColor='#F1F1F1'" style="cursor:pointer">
			<td width="70" align="center"><%=GiftStatisticlist.FItemList(i).faccountname%></td>
			<td width="110" align="center"><%=GiftStatisticlist.FItemList(i).fcouponno%></td>
			<td width="100" align="center">
				<%= printUserId(GiftStatisticlist.FItemList(i).fuserid, 2, "*") %>
			</td>
			<td width="80" align="center"><%=GiftStatisticlist.FItemList(i).fusername%></td>
			<td width="80" align="center"><%= GetCardName(GiftStatisticlist.FItemList(i).ftotalsum) %></td>
			<td width="70" align="center"><%=FormatNumber(GiftStatisticlist.FItemList(i).ftotalsum,0) %></td>
			<td width="70" align="center"><%=FormatNumber(GiftStatisticlist.FItemList(i).fsubtotalprice,0) %></td>
			<td width="70" align="center"><font color="<%= GetCardStatusColor(GiftStatisticlist.FItemList(i).fcardStatus) %>"><%= GetCardStatusName(GiftStatisticlist.FItemList(i).fcardStatus) %></font></td>
			<td width="150"> <%=GiftStatisticlist.FItemList(i).fregdate %></td>
		</tr>
	<% Next %>

    <tr height="25" bgcolor="FFFFFF">
		<td colspan="15" align="center">
	       	<% if GiftStatisticlist.HasPreScroll then %>
				<span class="list_link"><a href="#" onclick="frmsubmit('<%= GiftStatisticlist.StartScrollPage-1 %>'); return false;">[pre]</a></span>
			<% else %>
			[pre]
			<% end if %>
			<% for i = 0 + GiftStatisticlist.StartScrollPage to GiftStatisticlist.StartScrollPage + GiftStatisticlist.FScrollCount - 1 %>
				<% if (i > GiftStatisticlist.FTotalpage) then Exit for %>
				<% if CStr(i) = CStr(GiftStatisticlist.FCurrPage) then %>
				<span class="page_link"><font color="red"><b><%= i %></b></font></span>
				<% else %>
				<a href="#" onclick="frmsubmit('<%= i %>'); return false;" class="list_link"><font color="#000000"><%= i %></font></a>
				<% end if %>
			<% next %>
			<% if GiftStatisticlist.HasNextScroll then %>
				<span class="list_link"><a href="#" onclick="frmsubmit('<%= i %>'); return false;">[next]</a></span>
			<% else %>
			[next]
			<% end if %>
		</td>
	</tr>
	<% Else %>
		<tr bgcolor="#FFFFFF" height="30">
			<td colspan="9" align="center" class="page_link">[�����Ͱ� �����ϴ�.]</td>
		</tr>
	<% End If %>

</table>

<% IF application("Svr_Info")="Dev" THEN %>
	<iframe id="view" name="view" src="" width="100%" height=300 frameborder="0" scrolling="no"></iframe>
<% else %>
	<iframe id="view" name="view" src="" width=0 height=0 frameborder="0" scrolling="no"></iframe>
<% end if %>

<%
set GiftStatisticlist = nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->