<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  ���������̺�Ʈ ���
' History : 2010.03.25 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/common/incSessionAdminOrShop.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/common/lib/commonbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/offshop/event_off/event_function.asp"-->
<!-- #include virtual="/lib/classes/offshop/event_off/eventreport_Cls.asp"-->
<%
dim SType , evt_code,i ,page, grpWidth ,shopid ,t_TotalCost, t_FTotalNo,yyyy1,mm1,dd1,yyyy2,mm2,dd2
dim fromDate , toDate ,ftotselljumuncnt ,datefg, inc3pl
	datefg = requestCheckVar(request("datefg"),10)
	SType = requestCheckVar(request("SType"),10)
	evt_code = requestCheckVar(request("evt_code"),10)
	shopid = requestCheckVar(request("shopid"),32)
	yyyy1 = requestCheckVar(request("yyyy1"),4)
	mm1 = requestCheckVar(request("mm1"),2)
	dd1 = requestCheckVar(request("dd1"),2)
	yyyy2 = requestCheckVar(request("yyyy2"),4)
	mm2 = requestCheckVar(request("mm2"),2)
	dd2 = requestCheckVar(request("dd2"),2)
    inc3pl = requestCheckVar(request("inc3pl"),32)

if datefg = "" then datefg = "event"
if (yyyy1="") then yyyy1 = Cstr(Year(now()))
if (mm1="") then mm1 = Cstr(Month(now()))
if (dd1="") then dd1 = Cstr(day(now()))
if (yyyy2="") then yyyy2 = Cstr(Year(now()))
if (mm2="") then mm2 = Cstr(Month(now()))
if (dd2="") then dd2 = Cstr(day(now()))
			
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
set oReport = new Cevtreport_list
	oReport.FRectevt_code = evt_code
	oReport.FRectshopid = shopid
	oReport.FRectStartDay = fromDate
	oReport.FRectEndDay = toDate
	oReport.frectdatefg = datefg
	oReport.FRectInc3pl = inc3pl
	
t_TotalCost = 0
t_FTotalNo  = 0
ftotselljumuncnt = 0
%>

<script language="javascript">
	
	function jsPopCal(sName){
		var winCal;
		winCal = window.open('/lib/common_cal.asp?DN='+sName,'pCal','width=250, height=200');
		winCal.focus();
	}
	
	function regsubmit(){
		frm.submit();
	}

	//��ǰ����
	function item_detail(shopid,evt_code,yyyy1,mm1,dd1,yyyy2,mm2,dd2){
		location.href='?evt_code='+evt_code+'&shopid='+shopid+'&SType=T&yyyy1='+yyyy1+'&mm1='+mm1+'&dd1='+dd1+'&yyyy2='+yyyy2+'&mm2='+mm2+'&dd2='+dd2+'&menupos=<%=menupos%>&datefg=jumun';
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
				<% draweventmaechul_datefg "datefg" ,datefg ," onchange='regsubmit()'"%>				
				<% DrawDateBox yyyy1,yyyy2,mm1,mm2,dd1,dd2 %>
				&nbsp;&nbsp;
				<%
				'����/������
				if (C_IS_SHOP) then
				%>	
				
					<% if not C_IS_OWN_SHOP and shopid <> "" then %>
						* ���� : <%=shopid%><input type="hidden" name="shopid" value="<%= shopid %>">
					<% else %>
						* ���� : <% drawSelectBoxOffShop "shopid",shopid %>
					<% end if %>
				<% else %>
					* ���� : <% drawSelectBoxOffShop "shopid",shopid %>
				<% end if %>
				<p>
				* �̺�Ʈ��ȣ : <input type="text" name="evt_code" size="10" value="<%= evt_code %>">
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
    </td>
    <td align="right">
		�з�: 
		<input type="radio" name="SType" value="D" <% If SType = "D" Then response.write "checked" %> onclick="regsubmit();">��¥��
		<input type="radio" name="SType" value="T" <% If SType = "T" Then response.write "checked" %> onclick="regsubmit();">��ǰ��    	
    </td>        
</tr>
</form>
</table>
<!-- ǥ �߰��� ��-->

<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<%
'// ��¥�� �̺�Ʈ ��� 
if SType = "D" then

	'//��迡�� ������
	oReport.geteventdate_sum()
%>
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="20">
			�˻���� : <b><%= oReport.FResultCount %></b> �� ��1000�� ������ �˻� �˴ϴ�.
		</td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td>������</td>
		<td>����</td>
		<td>�̺�Ʈ<Br>�ڵ�</td>
		<td>�̺�Ʈ��</td>
		<td>�����</td>
		<td>����<br>�Ǽ�</td>
		<td>�Ǹ�<br>����</td>
		<td>�׷���</td>
		<td>���</td>
	</tr>
	<% 
	if oReport.FResultCount > 0 then 
	for i=0 to oReport.FResultCount-1
	
	t_TotalCost = t_TotalCost + oReport.FItemList(i).fsellsum
	t_FTotalNo  = t_FTotalNo + oReport.FItemList(i).fsum_cnt
	ftotselljumuncnt = ftotselljumuncnt + oReport.FItemList(i).ftotselljumuncnt
	%>
	<tr bgcolor="#FFFFFF" align="center">
		<td><%= oReport.FItemList(i).fshopregdate %></td>
		<td><%= oReport.FItemList(i).fshopid %></td>
		<td><%= oReport.FItemList(i).fevt_code %></td>
		<td><%= oReport.FItemList(i).fevt_name %></td>
		<td align="right" bgcolor="#E6B9B8"><%= FormatNumber(oReport.FItemList(i).fsellsum,0) %></td>
		<td align="right"><%= FormatNumber(oReport.FItemList(i).ftotselljumuncnt,0) %></td>
		<td align="right"><%= oReport.FItemList(i).fsum_cnt %></td>
		<td width="200" align="left">
			<%				
				if oReport.maxc>0 then
					grpWidth = Clng(oReport.FItemList(i).fsellsum/oReport.maxc*200)
				else
					grpWidth = 0
				end if
			%>
			<img src="http://partner.10x10.co.kr/images/dot1.gif" height="4" width="<%=grpWidth%>">
		</td>
		<td width=100>
			<input type="button" class="button" value="��ǰ��" onclick="item_detail('<%= oReport.FItemList(i).fshopid %>','<%= oReport.FItemList(i).fevt_code %>','<%= left(oReport.FItemList(i).fshopregdate,4) %>','<%= mid(oReport.FItemList(i).fshopregdate,6,2) %>','<%= right(oReport.FItemList(i).fshopregdate,2) %>','<%= left(oReport.FItemList(i).fshopregdate,4) %>','<%= mid(oReport.FItemList(i).fshopregdate,6,2) %>','<%= right(oReport.FItemList(i).fshopregdate,2) %>');">		
		</td>		
	</tr>
	<% 
	next
	%>
	<tr bgcolor="#FFFFFF" align="center">
		<td colspan=4>����</td>		
		<td align="right"><%= FormatNumber(t_TotalCost,0) %></td>
		<td align="right"><%= FormatNumber(ftotselljumuncnt,0) %></td>
		<td align="right"><%= FormatNumber(t_FTotalNo,0) %></td>
		<td colspan=3></td>
	</tr>
		
	<% ELSE %>
	<tr  align="center" bgcolor="#FFFFFF">
		<td colspan="20">��ϵ� ������ �����ϴ�.</td>
	</tr>
	<% 
	end if 
	%>
<% 
'// ��ǰ�� �̺�Ʈ ��� 
elseif SType = "T" then 

	oReport.geteventitem_sum()
%>
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="20">
			�˻���� : <b><%= oReport.FResultCount %></b> �� ��1000�� ������ �˻� �˴ϴ�.
		</td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td>����</td>
		<td>�̺�Ʈ<Br>�ڵ�</td>	
		<td>�̺�Ʈ��</td>
		<td>��ǰ�ڵ�<br>�귣��</td>
		<td>��ǰ��<font color='blue'>(�ɼǸ�)<font></td>		
		<td>�����</td>
		<td>�Ǹ�<br>����</td>
		<td>�׷���</td>
	</tr>
	<% if oReport.FResultCount > 0 then %>
	<% for i=0 to oReport.FResultCount-1 %>
	<%
	t_TotalCost = t_TotalCost + oReport.FItemList(i).fsellsum
	t_FTotalNo  = t_FTotalNo + oReport.FItemList(i).fsum_cnt
	%>
	<tr bgcolor="#FFFFFF" align="center">
		<td>
			<%= oReport.FItemList(i).fshopid %>
		</td>		
		<td>
			<%= oReport.FItemList(i).fevt_code %>
		</td>
		<td>
			<%= oReport.FItemList(i).fevt_name %>
		</td>				
		<td>
			<%= oReport.FItemList(i).fitemgubun %>-<%= oReport.FItemList(i).FItemid %>-<%= oReport.FItemList(i).fitemoption %><br><%= oReport.FItemList(i).fmakerid %>
		</td>
		<td>			
			<%= oReport.FItemList(i).fitemname %>
			<% 
			if oReport.FItemList(i).fitemoption <> "0000" then
				response.write "<font color='blue'>("&oReport.FItemList(i).fitemoptionname&")<font>"
			end if
			%>
		</td>
		<td align="right" bgcolor="#E6B9B8">
			<%= FormatNumber(oReport.FItemList(i).fsellsum,0) %>
		</td>	
		<td align="right">
			<%= oReport.FItemList(i).fsum_cnt %>
		</td>
		<td align="left" width=200>
			<%			
			if oReport.maxc>0 then
				grpWidth = Clng(oReport.FItemList(i).fsellsum/oReport.maxc*200)
			else
				grpWidth = 0
			end if
			%>
			<img src="http://partner.10x10.co.kr/images/dot1.gif" height="4" width="<%=grpWidth%>">
		</td>
	</tr>
	<% next %>
	<tr bgcolor="#FFFFFF">
		<td colspan=5 align="center">����</td>
		<td align="right">
			<%= FormatNumber(t_TotalCost,0) %>
		</td>
		<td align="right">
			<%= FormatNumber(t_FTotalNo,0) %>
		</td>		
		<td></td>
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
<!-- #include virtual="/common/lib/commonbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->