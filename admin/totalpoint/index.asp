<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  ���� ȸ�� ī��
' History : 2009.07.08 ���ر� ����
'			2011.01.18 �ѿ�� ����(����¡ Ŭ���� ������� ����. ������ǳ� ����¡ �߸��Ǿ� ����)
'###########################################################
%>
<!-- #include virtual="/common/incSessionAdminOrShop.asp" -->
<!-- #include virtual = "/lib/util/htmllib.asp" -->
<!-- #include virtual = "/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/checkAllowIPWithLog.asp" -->
<!-- #include virtual="/common/lib/commonbodyhead.asp"-->
<!-- #include virtual = "/lib/function.asp" -->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #Include virtual = "/lib/classes/totalpoint/totalpointCls.asp" -->

<%
Dim vUserName, vUserID, vJumin1, vCardNo, vUseYN, vCardGubun , ix,iPerCnt, vParam
Dim opoint ,i ,shopid, fromDate,toDate ,yyyy1,mm1,dd1,yyyy2,mm2,dd2 , page, memberYn, userhp
	shopid 	= requestCheckVar(Request("shopid"),32)
	vUserName		= NullFillWith(requestCheckVar(Request("username"),20),"")
	vUserID			= NullFillWith(requestCheckVar(Request("userid"),32),"")
	vCardGubun		= NullFillWith(requestCheckVar(Request("cardgubun"),4),"")
	vCardNo			= NullFillWith(requestCheckVar(Request("cardno"),20),"")
	vUseYN			= NullFillWith(requestCheckVar(Request("useyn"),20),"")
	memberYn		= requestCheckVar(Request("memberYn"),1)
	yyyy1 = requestCheckVar(request("yyyy1"),4)
	mm1 = requestCheckVar(request("mm1"),2)
	dd1 = requestCheckVar(request("dd1"),2)
	yyyy2 = requestCheckVar(request("yyyy2"),4)
	mm2 = requestCheckVar(request("mm2"),2)
	dd2 = requestCheckVar(request("dd2"),2)
	page = requestCheckVar(request("page"),10)
	userhp		= requestCheckVar(Request("userhp"),16)

if page="" then page=1

if (yyyy1="") then
	fromDate = DateSerial(Cstr(Year(now())), Cstr(Month(now())), Cstr(day(now()))-7)
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

if C_ADMIN_USER then
'/����
elseif (C_IS_SHOP) then

	'//�������϶�
	if C_IS_OWN_SHOP then

		'/���α��� ���� �̸�
		'if getlevel_sn("",session("ssBctId")) > 6 then
			'shopid = C_STREETSHOPID		'"streetshop011"
		'end if
	else
		shopid = C_STREETSHOPID
	end if
else
	'/��ü
	if (C_IS_Maker_Upche) then

	else
		if (Not C_ADMIN_USER) then
		    shopid = "X"                ''�ٸ�������ȸ ����.
		else
		end if
	end if
end if

vParam = "&username="&vUserName&"&cardno="&vCardNo&"&userid="&vUserID&"&useyn="&vUseYN&"&shopid="&shopid&"&yyyy1="&yyyy1&"&mm1="&mm1&"&dd1="&dd1&"&yyyy2="&yyyy2&"&mm2="&mm2&"&dd2="&dd2

set opoint = new TotalPoint
	opoint.FPageSize=20
	opoint.FCurrPage=page
 	opoint.FUserName = vUserName
 	opoint.FUserID = vUserID
	opoint.FUseYN = vUseYN
 	opoint.FCardNo = vCardNo
 	opoint.FCardGubun = vCardGubun
 	opoint.frectshopid = shopid
	opoint.FRectStartDay = fromDate
	opoint.FRectEndDay = toDate
	opoint.frectmemberYn = memberYn
	opoint.frectuserhp = userhp
	opoint.GetTotalPointList
%>

<script language="javascript">

function goRead(userseq){
	if (userseq=="0" || userseq==""){
		alert('��ȸ���̰ų� ������ ���� ���� ��ȸ �ϽǼ� �����ϴ�.');
		return;
	}
	
	var popwin = window.open('point_detail.asp?userseq='+userseq+'<%=vParam%>','point_detail','width=650,height=527,scrollbars=yes,resizable=yes');
	popwin.focus();
}

function jsGoPage(page){
	document.frm.page.value = page;
	document.frm.submit();
}

</script>

<form name="frm" method="get" action="" style="margin:0px;">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="page" value="1">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="3" width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
	<td align="left">
		* ��¥ : <% DrawDateBox yyyy1,yyyy2,mm1,mm2,dd1,dd2 %>
		&nbsp;&nbsp;
		<%
		'����/������
		if (C_IS_SHOP) then
		%>
			<% if getoffshopdiv(shopid) <> "1" and shopid <> "" then %>
				* ��ϸ��� : <%=shopid%><input type="hidden" name="shopid" value="<%= shopid %>">
			<% else %>
				* ��ϸ��� : <% drawSelectBoxOffShopdiv_off "shopid",shopid,"1,3,7,11","","" %>
			<% end if %>
		<% else %>
			* ��ϸ��� : <% drawSelectBoxOffShopdiv_off "shopid",shopid,"1,3,7,11","","" %>
		<% end if %>
	</td>
	<td rowspan="3" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="submit" class="button_s" value="�˻�" onClick="jsGoPage('');">
		<!--<input type="button" class="button_s" value="�ʱ⸮��Ʈ" onClick="location.href='/admin/totalpoint/?menupos=<%=g_MenuPos%>'">-->
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
		* ȸ������
		<% drawSelectBoxisusingYN "memberYn",memberYn,"" %>
		* ī����YN
		<% drawSelectBoxisusingYN "useyn",vUseYN,"" %>
		&nbsp;&nbsp;
		* ī�屸��
		<select name="cardgubun" class="select">
			<option value="">��ü</option>
			<option value="1010" <% If vCardGubun = "1010" Then %>selected<% End If %>>POINT1010</option>
			<option value="T" <% If vCardGubun <> "" AND vCardGubun <> "1010" AND vCardGubun <> "3253" Then %>selected<% End If %>>(��)��������</option>
			<option value="3253" <% If vCardGubun = "3253" Then %>selected<% End If %>>(��)���̶��</option>
		</select>
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
		* ����: <input type="text" class="text" name="username" value="<%=vUserName%>" size="8">
		&nbsp;&nbsp;
		* ���̵�: <input type="text" class="text" name="userid" value="<%=vUserID%>" size="12">
		&nbsp;&nbsp;
		* �޴�����ȣ: <input type="text" class="text" name="userhp" value="<%= userhp %>" size="16" maxlength=16>
		&nbsp;&nbsp;
		* ī���ȣ: <input type="text" class="text" name="cardno" value="<%=vCardNo%>">
	</td>
</tr>
</table>
</form>

<Br>

<!-- �׼� ���� -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
<tr>
	<td align="left">
	</td>
	<td align="right">
	</td>
</tr>
</table>
<!-- �׼� �� -->

<!-- ����Ʈ ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="20">
		�˻���� : <b><%= opoint.FTotalCount %></b>
		&nbsp;
		������ : <b><%= page %>/ <%= opoint.FTotalPage %></b>
	</td>
</tr>
<tr align="center" height="25" bgcolor="<%= adminColor("tabletop") %>">
	<td>ȸ����ȣ</td>
	<td>����</td>
	<td>���̵�</td>
	<td>ī�屸��</td>
	<td>ī���ȣ</td>
	<td>ī����<br>YN</td>
	<td>��������Ʈ</td>
	<td>���԰�����</td>
	<td>�����</td>
	<td>���</td>
</tr>
<%
if opoint.FresultCount > 0 then

for i=0 to opoint.FresultCount-1
%>
<tr align="center" height="25" bgcolor="#FFFFFF" onmouseout="this.style.backgroundColor='#FFFFFF'" onmouseover="this.style.backgroundColor='#F1F1F1'" style="cursor:pointer">
	<td>
		<% if opoint.FItemList(i).fUserSeq<>"0" then %>
			<%= opoint.FItemList(i).fUserSeq %>
		<% else %>
			��ȸ��
		<% end if %>
	</td>
	<td>
		<%
			If opoint.FItemList(i).fUserName <> "" Then
				If opoint.FItemList(i).fGrade <> "0" Then
					Response.Write "[Ư��]" & opoint.FItemList(i).fUserName
				Else
					Response.Write opoint.FItemList(i).fUserName
				End If
			Else
				Response.Write "&nbsp;"
			End If
		%>
	</td>
	<td><%= printUserId(opoint.FItemList(i).fOnlineUserID, 2, "*") %></td>
	<td>
		<% If Left(opoint.FItemList(i).fCardNo,4) = "1010" Then %>
			POINT1010
		<% ElseIf Left(opoint.FItemList(i).fCardNo,4) = "3253" Then %>
			���̶��
		<% Else %>
			��������
		<% End If %>
	</td>
	<td><%= opoint.FItemList(i).fCardNo %></td>
	<td><%= opoint.FItemList(i).fUseYN %></td>
	<td align="right"><%=FormatNumber(opoint.FItemList(i).fPoint,0)%></td>
	<td>
		<% If opoint.FItemList(i).fshopname = "" Then %>
			�¶��ΰ���
		<% Else %>
			�������ΰ���
			<br><%= opoint.FItemList(i).fshopname %>
		<% End If %>
	</td>
	<td><%=opoint.FItemList(i).fRegdate%></td>
	<td><input type="button" class="button" value="�����󼼺���" onClick="goRead('<%=opoint.FItemList(i).fUserSeq%>')"></td>
</tr>
<% Next %>

<tr height="25" bgcolor="FFFFFF">
	<td colspan="15" align="center">
		<% if opoint.HasPreScroll then %>
			<span class="list_link"><a href="javascript:jsGoPage(<%= opoint.StartScrollPage-1 %>);">[pre]</a></span>
		<% else %>
		[pre]
		<% end if %>
		<% for i = 0 + opoint.StartScrollPage to opoint.StartScrollPage + opoint.FScrollCount - 1 %>
			<% if (i > opoint.FTotalpage) then Exit for %>
			<% if CStr(i) = CStr(opoint.FCurrPage) then %>
			<span class="page_link"><font color="red"><b><%= i %></b></font></span>
			<% else %>
			<a href="javascript:jsGoPage(<%= i %>);" class="list_link"><font color="#000000"><%= i %></font></a>
			<% end if %>
		<% next %>
		<% if opoint.HasNextScroll then %>
			<span class="list_link"><a href="javascript:jsGoPage(<%= i %>);">[next]</a></span>
		<% else %>
		[next]
		<% end if %>
	</td>
</tr>

<% Else %>
<tr bgcolor="#FFFFFF">
	<td colspan="20" align="center" class="page_link">[�˻������ �����ϴ�.]</td>
</tr>
<% End If %>

</table>

<%
set opoint = Nothing
%>
<!-- #include virtual="/common/lib/commonbodytail.asp"-->
<!--#Include Virtual = "/lib/db/dbclose.asp" -->
