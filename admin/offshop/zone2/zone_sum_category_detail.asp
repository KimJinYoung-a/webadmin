<%@ language=vbscript %>
<% option explicit %>
<%
	Response.AddHeader "Cache-Control","no-cache"
	Response.AddHeader "Expires","0"
	Response.AddHeader "Pragma","no-cache"
%>
<%
'###########################################################
' Description : ���׺�����
' Hieditor : 2010.12.29 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/common/incSessionAdminOrShop.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/common/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/offshop/zone2/zone_cls.asp"-->

<%
Dim ozone,i,page , parameter , shopid ,yyyy1,mm1,dd1,yyyy2,mm2,dd2 ,fromDate ,toDate
dim designer , sellgubun ,cdl ,cdm ,cds ,datefg ,menupos , zoneidx
dim dategubun ,totsumshopsuplycash, inc3pl
dim totsellsum ,totprofit , totitemno ,totrealsellprice ,totshopsuplycash
	designer = RequestCheckVar(request("designer"),32)
	page = requestCheckVar(request("page"),10)
	shopid = requestCheckVar(request("shopid"),32)
	yyyy1 = requestCheckVar(request("yyyy1"),4)
	mm1 = requestCheckVar(request("mm1"),2)
	dd1 = requestCheckVar(request("dd1"),2)
	yyyy2 = requestCheckVar(request("yyyy2"),4)
	mm2 = requestCheckVar(request("mm2"),2)
	dd2 = requestCheckVar(request("dd2"),2)
	sellgubun = requestCheckVar(request("sellgubun"),10)
	cdl     = requestCheckVar(request("cdl"),3)
	cdm     = requestCheckVar(request("cdm"),3)
	cds     = requestCheckVar(request("cds"),3)
	datefg = requestCheckVar(request("datefg"),10)
	zoneidx = requestCheckVar(request("zoneidx"),10)
	menupos = requestCheckVar(request("menupos"),10)
	dategubun = requestCheckVar(request("dategubun"),10)
    inc3pl = requestCheckVar(request("inc3pl"),32)
    
if dategubun = "" then dategubun = "G"
if datefg = "" then datefg = "maechul"			
if sellgubun = "" then sellgubun = "S"
if (yyyy1="") then yyyy1 = Cstr(Year(now()))
if (mm1="") then mm1 = Cstr(Month(now()))
if (dd1="") then dd1 = Cstr(day(now()))
if (yyyy2="") then yyyy2 = Cstr(Year(now()))
if (mm2="") then mm2 = Cstr(Month(now()))
if (dd2="") then dd2 = Cstr(day(now()))
			
fromDate = DateSerial(yyyy1, mm1, dd1)
toDate = DateSerial(yyyy2, mm2, dd2+1)


if page = "" then page = 1
if cdl<>"" and cdm<>"" then cds=""
		
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
		designer = session("ssBctID")	'"7321"

	else
		if (Not C_ADMIN_USER) then
		    shopid = "X"                ''�ٸ�������ȸ ����.
		else
		end if
	end if
end if	
			
set ozone = new czone_list
	ozone.FPageSize = 100
	ozone.FCurrPage = page
	ozone.frectdategubun = dategubun
	ozone.frectshopid = shopid
	ozone.FRectStartDay = fromDate
	ozone.FRectEndDay = toDate
	ozone.FRectmakerid = designer
	ozone.FRectCDL = cdl
	ozone.FRectCDM = cdm
	ozone.FRectCDN = cds	
	ozone.frectdatefg = datefg
	ozone.frectzoneidx = zoneidx
	ozone.frectsellgubun = sellgubun
	ozone.frectidx = zoneidx
	ozone.FRectInc3pl = inc3pl
	
	if shopid <> "" then
		ozone.Getoffshopzone_detailCategory
	end if
	
	if shopid = "" then response.write "<script>alert('������ �������ּ���');</script>"
		
	totsellsum = 0
	totprofit =0
	totitemno = 0
	totrealsellprice = 0
	totshopsuplycash = 0
	totsumshopsuplycash = 0	
	
parameter = "designer="&designer&"&shopid="&shopid&"&yyyy1="&yyyy1&"&mm1="&mm1&"&dd1="&dd1&"&yyyy2="&yyyy2&"&mm2="&mm2&"&dd2="&dd2&"&sellgubun="&sellgubun
parameter = parameter & "&datefg="&datefg&"&zoneidx="&zoneidx&"&dategubun="&dategubun&"&inc3pl="&inc3pl&"&menupos="&menupos
%>

<script language="javascript">
	
	function gopage(page){
		frm.page.value=page;
		frm.submit();
	}

	//��ǰ����
	function item_detail(cdl,cdm,cds){
		var item_detail = window.open('zone_sum_item_detail.asp?cdl='+cdl+'&cdm='+cdm+'&cds='+cds+'&<%=parameter%>','item_detail','width=1024,height=768,scrollbars=yes,resizable=yes');
		item_detail.focus();
	}

</script>
	
<!-- �˻� ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get" action="">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="page" value="<%= page %>">
<input type="hidden" name="dategubun" value="<%= dategubun %>">
<input type="hidden" name="cdl" value="<%= cdl %>">
<input type="hidden" name="cdm" value="<%= cdm %>">
<input type="hidden" name="cds" value="<%= cds %>">	
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
	<td align="left">
		* �Ⱓ : <% drawmaechul_datefg "datefg" ,datefg ,""%> 
		<% DrawDateBox yyyy1,yyyy2,mm1,mm2,dd1,dd2 %>	
		&nbsp;&nbsp;
		<%
		'����/������
		if (C_IS_SHOP) then
		%>	
			<% if (not C_IS_OWN_SHOP and shopid <> "") then %>
				* ���� : <%=shopid%><input type="hidden" name="shopid" value="<%= shopid %>">
			<% else %>
				* ����:<% drawSelectBoxOffShop "shopid",shopid %>
			<% end if %>
		<% else %>
			<% if not(C_IS_Maker_Upche) then %> 
				* ����:<% drawSelectBoxOffShop "shopid",shopid %>
			<% else %>
				* ����:<% drawSelectBoxOffShop "shopid",shopid %>
			<% end if %>
		<% end if %>
	</td>	
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="�˻�" onClick="gopage('');">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
		* �귣�� : <% drawSelectBoxDesignerwithName "designer",designer %>
		&nbsp;&nbsp;
		<% Call zoneselectbox(shopid,"zoneidx",zoneidx,"") %>
		&nbsp;&nbsp;		
		<input type="radio" name="sellgubun" value="S" <% if sellgubun="S" then response.write " checked" %>>������������
		<input type="radio" name="sellgubun" value="N" <% if sellgubun="N" then response.write " checked" %>>�����ϳ�������
        <p>
        <b>* ����ó����</b>
        <% Call draw3plMeachulComboBox("inc3pl",inc3pl) %>			
	</td>
</tr>
</form>
</table>
<!-- �˻� �� -->
<br>
		
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
	<td colspan="15">
		�˻���� : <b><%= ozone.FTotalCount %></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td>����</td>
	<td>���׸�</td>
	<td>ī������</font></td>
	<td width="400" align="left">
		<img src="/images/dot1.gif" height="4" width=10>�Ǹż���
		<br><img src="/images/dot2.gif" height="4" width=10>�����
	</td>
	<td>�����</td>
	<td>�����<br>������</td>
	
	<% if NOT(C_IS_SHOP) then %>
		<td>�����<br>����</td>
		<td>�����<br>����������</td>
	<% end if %>
	
	<td>���</td>
</tr>
<% if ozone.FtotalCount>0 then %>
<%
for i=0 to ozone.FtotalCount-1

totsellsum = totsellsum + ozone.FItemList(i).fsellsum
totprofit = totprofit + (ozone.FItemList(i).fsellsum-ozone.FItemList(i).fsuplypricesum)
%>
<tr align="center" bgcolor="#FFFFFF" onmouseover=this.style.background="#f1f1f1"; onmouseout=this.style.background='#ffffff';>
	<td>
		<%= ozone.FItemList(i).fshopid %>
	</td>
	<td>
		<%= ozone.FItemList(i).fzonename %>
	</td>
	<td>
		<% if ozone.FItemList(i).FCateCDm="" then %>
			<a href="zone_sum_category_detail.asp?cdl=<%= ozone.FItemList(i).FCateCDL %>&cdm=<%= ozone.FItemList(i).FCateCDM %>&cds=<%= ozone.FItemList(i).FCateCDs %>&<%=parameter%>"><%= ozone.FItemList(i).FCateName %></a>
		<% else %>
			<%= ozone.FItemList(i).FCateName %>
		<% end if %>
	</td>
	<td height="10" width="400">
		<% if  (ozone.FItemList(i).fsellsum<>0) then %>
			<div align="left">
				<img src="/images/dot1.gif" height="4" width="<%= CLng((ozone.FItemList(i).fsellsum/ozone.maxt)*400) %>">			
			</div>
			<br><div align="left">
				<img src="/images/dot2.gif" height="4" width="<%= CLng((ozone.FItemList(i).fitemcnt/ozone.maxc)*400) %>">
			</div>
		<% end if %>
	</td>
	<td bgcolor="#E6B9B8">
		<%= FormatNumber(ozone.FItemList(i).fsellsum,0) %>
	</td>	
	<td>
		<% if ozone.FItemList(i).fsellsum<>0 then %>
			<%= Clng( ((ozone.FItemList(i).fsellsum / ozone.FSumTotal) * 10000)) / 100 %> %
		<% else %>
			0 %
		<% end if %>
	</td>
	<% if NOT(C_IS_SHOP) then %>		
		<td>
			<%= FormatNumber(ozone.FItemList(i).fsellsum-ozone.FItemList(i).fsuplypricesum,0) %>
		</td>
		<td>
			<% if ozone.FItemList(i).fsellsum-ozone.FItemList(i).fsuplypricesum<>0 then %>
				<%= Clng( (((ozone.FItemList(i).fsellsum-ozone.FItemList(i).fsuplypricesum) / ozone.fprofitTotal) * 10000)) / 100 %> %
			<% else %>
				0 %
			<% end if %>
		</td>
	<% end if %>	
	<td>
		<input type="button" onclick="javascript:item_detail('<%= ozone.FItemList(i).FCateCDL %>','<%= ozone.FItemList(i).FCateCDM %>','<%= ozone.FItemList(i).FCateCDs %>');" value="��ǰ��" class="button">
	</td>	
</tr>
<% next %>
<tr bgcolor="#FFFFFF" align="center">
	<td colspan=4>�հ�</td>
	<td><%= FormatNumber(totsellsum,0) %></td>
	<td></td>
	<% if NOT(C_IS_SHOP) then %>
		<td>
			<%= FormatNumber(totprofit,0) %>
		</td>	
		<td></td>
	<% end if %>
	<td></td>
</tr>	
<% else %>
	<tr bgcolor="#FFFFFF">
		<td colspan="15" align="center" class="page_link">[�˻������ �����ϴ�.]</td>
	</tr>
<% end if %>
</table>

<%
set ozone = nothing
%>
<!-- #include virtual="/common/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->