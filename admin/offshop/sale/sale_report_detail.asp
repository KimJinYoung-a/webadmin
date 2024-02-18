<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  ���� ���
' History : 2012.10.25 �ѿ�� ����
'####################################################
%>
<!-- #include virtual="/lib/util/datelib.asp"-->
<!-- #include virtual="/common/incSessionAdminOrShop.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/common/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/offshop/sale/salereport_cls.asp"-->
<!-- #include virtual="/lib/classes/offshop/event_off/event_function.asp"-->

<%
dim SType , sale_code,i ,page ,shopid, yyyy1, mm1, dd1, yyyy2, mm2, dd2
dim fromDate , toDate, menupos, inc3pl
	SType = requestCheckVar(request("SType"),1)
	sale_code = requestCheckVar(request("sale_code"),10)
	shopid = requestCheckVar(request("shopid"),32)
	yyyy1 = requestCheckVar(request("yyyy1"),4)
	mm1 = requestCheckVar(request("mm1"),2)
	dd1 = requestCheckVar(request("dd1"),2)
	yyyy2 = requestCheckVar(request("yyyy2"),4)
	mm2 = requestCheckVar(request("mm2"),2)
	dd2 = requestCheckVar(request("dd2"),2)
	menupos = requestCheckVar(request("menupos"),10)
	page = requestCheckVar(request("page"),10)
    inc3pl = requestCheckVar(request("inc3pl"),32)

if SType = "" then SType = "D"
if page = "" then page = 1

if (yyyy1="") then
	fromDate = DateSerial(Cstr(Year(now())), Cstr(Month(now())), Cstr(day(now()))-90)
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

fromDate = DateSerial(yyyy1, mm1, dd1)
toDate = DateSerial(yyyy2, mm2, dd2+1)

'/����
if (C_IS_SHOP) then

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

dim oReport
set oReport = new Csalereport_list
	oReport.FRectsale_code = sale_code
	oReport.FRectshopid = shopid
	oReport.frectevt_startdate = fromDate
	oReport.frectevt_enddate = toDate
	oReport.FPageSize = 1000
	oReport.FCurrPage = page
	oReport.FRectInc3pl = inc3pl
%>

<script language="javascript">

	function regsubmit(){
		frm.submit();
	}

	//��ǰ����
	function item_detail(SType,shopid,sale_code,yyyy1,mm1,dd1,yyyy2,mm2,dd2){
		location.href='?SType='+SType+'&sale_code='+sale_code+'&shopid='+shopid+'&yyyy1='+yyyy1+'&mm1='+mm1+'&dd1='+dd1+'&yyyy2='+yyyy2+'&mm2='+mm2+'&dd2='+dd2+'&menupos=<%=menupos%>';
	}

</script>

<!-- ǥ ��ܹ� ����-->
<table width="100%" align="center" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>" class="a">
<form name="frm" method="get" action="">
<input type="hidden" name="page" value="1">
<input type="hidden" name="menupos" value="<%= request("menupos") %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
	<td align="left">
		<table border="0" width="100%" cellpadding="3" cellspacing="0" class="a">
		<tr>
			<td>
				* �Ⱓ :
				<% DrawDateBox yyyy1,yyyy2,mm1,mm2,dd1,dd2 %>
				&nbsp;&nbsp;
				<%
				'����/������
				if (C_IS_SHOP) then
				%>
					<% if (not C_IS_OWN_SHOP and shopid <> "") then %>
						* ���� : <%=shopid%><input type="hidden" name="shopid" value="<%= shopid %>">
					<% else %>
						* ���� : <% drawSelectBoxOffShop "shopid",shopid %>
					<% end if %>
				<% else %>
					<% if not(C_IS_Maker_Upche) then %>
						* ���� : <% drawSelectBoxOffShop "shopid",shopid %>
					<% else %>
						* ���� : <% drawSelectBoxOffShop "shopid",shopid %>
					<% end if %>
				<% end if %>

				<p>
				* ���ι�ȣ : <input type="text" name="sale_code" size="10" value="<%= sale_code %>">
	            &nbsp;&nbsp;
	            <b>* ����ó����</b>
	            <% Call draw3plMeachulComboBox("inc3pl",inc3pl) %>
			</td>
		</tr>
		</table>
    </td>
		<td  width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="�˻�" onClick="regsubmit();">
	</td>
</tr>
</table>
<!-- ǥ ��ܹ� ��-->
<br>
<!-- ǥ �߰��� ����-->
<table width="100%" align="center" cellpadding="1" cellspacing="1" class="a">
<tr valign="bottom">
    <td align="left">
    	��1000�� ���� �˻�����
    </td>
    <td align="right">
		�з�:
		<input type="radio" name="SType" value="D" <% If SType = "D" Then response.write "checked" %> onclick="regsubmit();">��¥��
		<input type="radio" name="SType" value="I" <% If SType = "I" Then response.write "checked" %> onclick="regsubmit();">��ǰ��
		<input type="radio" name="SType" value="B" <% If SType = "B" Then response.write "checked" %> onclick="regsubmit();">�귣�庰
    </td>
</tr>
</form>
</table>
<!-- ǥ �߰��� ��-->

<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">

<%
'// ��¥�� ���� ���
if SType = "D" then

	'//������̺��� ������
	oReport.getsaledate_sum()
%>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td>������</td>
		<td>����</td>
		<td>����<Br>�ڵ�</td>
		<td>���θ�</td>
		<!--<td>�����</td>-->
		<td>�����</td>
		<!--<td>�ֹ��Ǽ�</td>-->
		<td>�Ǹ�<br>����</td>
		<td>���</td>
	</tr>
	<%
	dim totsellprice ,totrealsellprice ,totselljumuncnt ,totsellCnt
		totsellprice = 0
		totrealsellprice = 0
		totselljumuncnt = 0
		totsellCnt = 0

	if oReport.FResultCount > 0 then

	for i=0 to oReport.FResultCount-1

	totsellprice = totsellprice + oReport.FItemList(i).ftotsellprice
	totrealsellprice = totrealsellprice + oReport.FItemList(i).ftotrealsellprice
	totselljumuncnt = totselljumuncnt + oReport.FItemList(i).ftotselljumuncnt
	totsellCnt = totsellCnt + oReport.FItemList(i).ftotsellCnt
	%>
	<tr bgcolor="#FFFFFF" align="center">
		<td width=80><%= oReport.FItemList(i).fyyyymmdd %></td>
		<td><%= oReport.FItemList(i).fshopname %></td>
		<td width=60><%= oReport.FItemList(i).fsale_code %></td>
		<td><%= oReport.FItemList(i).fsale_name %></td>
		<!--<td width=80 align="right"><%'= FormatNumber(oReport.FItemList(i).ftotsellprice,0) %></td>-->
		<td width=80 align="right" bgcolor="#E6B9B8"><%= FormatNumber(oReport.FItemList(i).ftotrealsellprice,0) %></td>
		<!--<td width=50 align="right"><%'= FormatNumber(oReport.FItemList(i).ftotselljumuncnt,0) %></td>-->
		<td width=50 align="right"><%= FormatNumber(oReport.FItemList(i).ftotsellCnt,0) %></td>
		<td width=80>
			<input type="button" class="button" value="��ǰ��" onclick="item_detail('I','<%= oReport.FItemList(i).fshopid %>','<%= oReport.FItemList(i).fsale_code %>','<%= left(oReport.FItemList(i).fyyyymmdd,4) %>','<%= mid(oReport.FItemList(i).fyyyymmdd,6,2) %>','<%= right(oReport.FItemList(i).fyyyymmdd,2) %>','<%= left(oReport.FItemList(i).fyyyymmdd,4) %>','<%= mid(oReport.FItemList(i).fyyyymmdd,6,2) %>','<%= right(oReport.FItemList(i).fyyyymmdd,2) %>');">
		</td>
	</tr>
	<%
	next
	%>
	<tr bgcolor="#FFFFFF" align="center">
		<td colspan=4>����</td>
		<!--<td align="right"><%'= FormatNumber(totsellprice,0) %></td>-->
		<td align="right"><%= FormatNumber(totrealsellprice,0) %></td>
		<!--<td align="right"><%'= FormatNumber(totselljumuncnt,0) %></td>-->
		<td align="right"><%= FormatNumber(totsellCnt,0) %></td>
		<td></td>
	</tr>

	<% ELSE %>
	<tr  align="center" bgcolor="#FFFFFF">
		<td colspan="20">��ϵ� ������ �����ϴ�.</td>
	</tr>
	<%
	end if
	%>
<%
'// ��ǰ�� ���� ���
elseif SType = "I" then

	oReport.getsaleitem_sum()
%>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td>����</td>
		<td>����<Br>�ڵ�</td>
		<td>���θ�</td>
		<td>��ǰ�ڵ�</td>
		<td>�귣��</td>
		<td>��ǰ��<font color='blue'>(�ɼǸ�)<font></td>
		<!--<td>�����</td>-->
		<td>�����</td>
		<td>�Ǹ�<br>����</td>
	</tr>
	<%
	dim totsuplyprice, totbuyprice, totitemno
		totsellprice = 0
		totrealsellprice = 0
		totsuplyprice = 0
		totbuyprice = 0
		totitemno = 0

	if oReport.FResultCount > 0 then

	for i=0 to oReport.FResultCount-1

	totsellprice = totsellprice + oReport.FItemList(i).ftotsellprice
	totrealsellprice = totrealsellprice + oReport.FItemList(i).ftotrealsellprice
	totsuplyprice = totsuplyprice + oReport.FItemList(i).ftotsuplyprice
	totbuyprice = totbuyprice + oReport.FItemList(i).ftotbuyprice
	totitemno = totitemno + oReport.FItemList(i).ftotitemno
	%>
	<tr bgcolor="#FFFFFF" align="center">
		<td>
			<%= oReport.FItemList(i).fshopname %>
		</td>
		<td width=60>
			<%= oReport.FItemList(i).fsale_code %>
		</td>
		<td>
			<%= oReport.FItemList(i).fsale_name %>
		</td>
		<td width=90>
			<%= oReport.FItemList(i).fitemgubun %><%= CHKIIF(oReport.FItemList(i).FItemid>=1000000,Format00(8,oReport.FItemList(i).FItemid),Format00(6,oReport.FItemList(i).FItemid)) %><%= oReport.FItemList(i).fitemoption %>
		</td>
		<td>
			<%= oReport.FItemList(i).fmakerid %>
		</td>
		<td>
			<%= oReport.FItemList(i).fitemname %>
			<%
			if oReport.FItemList(i).fitemoption <> "0000" then
				response.write "<font color='blue'>("&oReport.FItemList(i).fitemoptionname&")<font>"
			end if
			%>
		</td>
		<!--<td width=80 align="right">
			<%'= FormatNumber(oReport.FItemList(i).ftotsellprice,0) %>
		</td>-->
		<td width=80 align="right" bgcolor="#E6B9B8">
			<%= FormatNumber(oReport.FItemList(i).ftotrealsellprice,0) %>
		</td>
		<td width=50 align="right">
			<%= oReport.FItemList(i).ftotitemno %>
		</td>
	</tr>
	<% next %>
	<tr bgcolor="#FFFFFF">
		<td colspan=6 align="center">����</td>
		<!--<td align="right">
			<%'= FormatNumber(totsellprice,0) %>
		</td>-->
		<td align="right">
			<%= FormatNumber(totrealsellprice,0) %>
		</td>
		<td align="right">
			<%= FormatNumber(totitemno,0) %>
		</td>
	</tr>
	<% ELSE %>
	<tr  align="center" bgcolor="#FFFFFF">
		<td colspan="20">��ϵ� ������ �����ϴ�.</td>
	</tr>
	<% end if %>
<%
'// �귣�庰 ���� ���
elseif SType = "B" then

	oReport.getsalebrand_sum()
%>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td>����</td>
		<td>����<Br>�ڵ�</td>
		<td>���θ�</td>
		<td>�귣��</td>
		<!--<td>�����</td>-->
		<td>�����</td>
		<td>�Ǹ�<br>����</td>
	</tr>
	<%
		totsellprice = 0
		totrealsellprice = 0
		totsuplyprice = 0
		totbuyprice = 0
		totitemno = 0

	if oReport.FResultCount > 0 then

	for i=0 to oReport.FResultCount-1

	totsellprice = totsellprice + oReport.FItemList(i).ftotsellprice
	totrealsellprice = totrealsellprice + oReport.FItemList(i).ftotrealsellprice
	totsuplyprice = totsuplyprice + oReport.FItemList(i).ftotsuplyprice
	totbuyprice = totbuyprice + oReport.FItemList(i).ftotbuyprice
	totitemno = totitemno + oReport.FItemList(i).ftotitemno
	%>
	<tr bgcolor="#FFFFFF" align="center">
		<td>
			<%= oReport.FItemList(i).fshopname %>
		</td>
		<td width=60>
			<%= oReport.FItemList(i).fsale_code %>
		</td>
		<td>
			<%= oReport.FItemList(i).fsale_name %>
		</td>
		<td>
			<%= oReport.FItemList(i).fmakerid %>
		</td>
		<!--<td width=80 align="right">
			<%'= FormatNumber(oReport.FItemList(i).ftotsellprice,0) %>
		</td>-->
		<td width=80 align="right" bgcolor="#E6B9B8">
			<%= FormatNumber(oReport.FItemList(i).ftotrealsellprice,0) %>
		</td>
		<td width=50 align="right">
			<%= oReport.FItemList(i).ftotitemno %>
		</td>
	</tr>
	<% next %>
	<tr bgcolor="#FFFFFF">
		<td colspan=4 align="center">����</td>
		<!--<td align="right">
			<%'= FormatNumber(totsellprice,0) %>
		</td>-->
		<td align="right">
			<%= FormatNumber(totrealsellprice,0) %>
		</td>
		<td align="right">
			<%= FormatNumber(totitemno,0) %>
		</td>
	</tr>
	<% ELSE %>
	<tr  align="center" bgcolor="#FFFFFF">
		<td colspan="20">��ϵ� ������ �����ϴ�.</td>
	</tr>
	<% end if %>	
<%
end if
%>
</table>

<%
set oReport = Nothing
%>
<!-- #include virtual="/common/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->