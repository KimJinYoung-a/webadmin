<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  �������� ���� (����Ÿ��Ʈ ��輭������ ������)
' History : 2010.05.10 ������ ����
'			2012.02.07 �ѿ�� ����
'####################################################
%>
<!-- #include virtual="/common/incSessionAdminOrShop.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db3open.asp" -->
<!-- #include virtual="/common/lib/commonbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/offshop/newoffshopsellcls.asp"-->
<%
dim yyyy1,mm1,dd1,yyyy2,mm2,dd2 , fromDate,toDate , searchtype , shopid ,i , datelen, datelen2 ,makerid
dim datefg , tmpdate ,myyyy1,mmm1,mdd1,myyyy2,mmm2,mdd2 ,mfromDate,mtoDate, offgubun, reload, inc3pl
dim cdl, cdm, offmduserid
	makerid = requestCheckVar(request("makerid"),32)
	shopid = requestCheckVar(request("shopid"),32)
	searchtype = requestCheckVar(request("searchtype"),32)
	yyyy1 = requestCheckVar(request("yyyy1"),4)
	mm1 = requestCheckVar(request("mm1"),2)
	dd1 = requestCheckVar(request("dd1"),2)
	yyyy2 = requestCheckVar(request("yyyy2"),4)
	mm2 = requestCheckVar(request("mm2"),2)
	dd2 = requestCheckVar(request("dd2"),2)
	myyyy1 = requestCheckVar(request("myyyy1"),4)
	mmm1 = requestCheckVar(request("mmm1"),2)
	mdd1 = requestCheckVar(request("mdd1"),2)
	myyyy2 = requestCheckVar(request("myyyy2"),4)
	mmm2 = requestCheckVar(request("mmm2"),2)
	mdd2 = requestCheckVar(request("mdd2"),2)
	offgubun = requestCheckVar(request("offgubun"),10)
	datefg = requestCheckVar(request("datefg"),32)
	reload = requestCheckVar(request("reload"),2)
	cdl = requestCheckVar(request("selC"),3)
	cdm = requestCheckVar(request("selCM"),3)
	offmduserid = requestCheckVar(request("offmduserid"),32)
    inc3pl = requestCheckVar(request("inc3pl"),32)

	if reload <> "on" and offgubun = "" then offgubun = "95"	
	if datefg = "" then datefg = "maechul"
		
	if searchtype="" then searchtype="ipgo"	

	tmpdate = dateadd("m",-1,date)
	
	if (yyyy1="") then yyyy1 = Cstr(Year(tmpdate))
	if (mm1="") then mm1 = Cstr(Month(tmpdate))
	if (dd1="") then dd1 = Cstr(day(tmpdate))
	if (yyyy2="") then yyyy2 = Cstr(Year(now()))
	if (mm2="") then mm2 = Cstr(Month(now()))
	if (dd2="") then dd2 = Cstr(day(now()))
	fromDate = DateSerial(yyyy1, mm1, dd1)
	toDate = DateSerial(yyyy2, mm2, dd2+1)

	if (myyyy1="") then myyyy1 = Cstr(Year(tmpdate))
	if (mmm1="") then mmm1 = Cstr(Month(tmpdate))
	if (mdd1="") then mdd1 = Cstr(day(tmpdate))
	if (myyyy2="") then myyyy2 = Cstr(Year(now()))
	if (mmm2="") then mmm2 = Cstr(Month(now()))
	if (mdd2="") then mdd2 = Cstr(day(now()))
	mfromDate = DateSerial(myyyy1, mmm1, mdd1)
	mtoDate = DateSerial(myyyy2, mmm2, mdd2+1)

'C_IS_SHOP = TRUE
'C_IS_Maker_Upche = TRUE

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
		makerid = session("ssBctID")	'"7321"
	else
		if (Not C_ADMIN_USER) then
		    shopid = "X"                ''�ٸ�������ȸ ����.
		else
		end if
	end if
end if	

if shopid<>"" then offgubun=""
	
dim oreport
set oreport = new COffShopSell
	oreport.FPageSize = 2500
	oreport.FRectSearchType = searchtype
	oreport.FRectFromDate = fromDate
	oreport.FRectToDate = toDate
	oreport.frectdatefg = datefg	
	oreport.FRectmFromDate = mfromDate
	oreport.FRectmToDate = mtoDate	
	oreport.FRectShopID = shopid
	oreport.frectmakerid = makerid
	oreport.FRectOffgubun = offgubun
	oreport.frectoffmduserid = OffMDUserID	
	oreport.FRectcdl = cdl
	oreport.FRectcdm = cdm
	oreport.FRectInc3pl = inc3pl
	
	'/����Ÿ��Ʈ
	oreport.GetNewBrandSell_datamart
	
	'/���ε�� �ǽð�
	'oreport.GetNewBrandSell
%>

<script language='javascript'>

function detailitem(makerid,shopid,myyyy1,mmm1,mdd1,myyyy2,mmm2,mdd2,datefg){
	var detailitem = window.open('newbrandsum_detailitem.asp?makerid='+makerid+'&shopid='+shopid+'&yyyy1='+myyyy1+'&mm1='+mmm1+'&dd1='+mdd1+'&yyyy2='+myyyy2+'&mm2='+mmm2+'&dd2='+mdd2+'&datefg='+datefg,'detailitem','width=1024,height=768,scrollbars=yes,resizable=yes');
	detailitem.focus();
}

function reg(){
	frm.submit();
}

</script>

<!-- ǥ ��ܹ� ����-->
<table width="100%" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>" class="a">
<form name="frm" method="get" action="">
<input type="hidden" name="showtype" value="showtype">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="reload" value="on">
<tr align="center" bgcolor="#FFFFFF" >
	<td width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
	<td align="left">  
		<table border="0" width="100%" cellpadding="1" cellspacing="1" class="a">
		<tr>
			<td>
				* �űԾ�ü���� : <% drawnewupche_datefg "searchtype",searchtype , " onchange='reg();'" %>
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
				* �귣�� : <% drawSelectBoxDesignerwithName "makerid",makerid %>
				&nbsp;&nbsp;
				* ���� ���� : <% drawoffshop_commoncode "offgubun", offgubun, "shopdivithinkso", "", "", " onchange='reg();'" %>
				&nbsp;&nbsp;				
				* �������δ��MD : <% drawSelectBoxCoWorker_OnOff "offmduserid", offmduserid, "off" %>				
				<p>
	            <b>* ����ó����</b>
	            <% Call draw3plMeachulComboBox("inc3pl",inc3pl) %>
	            &nbsp;&nbsp;	            
				<!-- #include virtual="/common/module/categoryselectbox_cdl.asp"-->						
			</td>
		</tr>
		</table> 
    </td>
	<td  width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="�˻�" onClick="frm.submit();">
	</td>
</tr>	
</table>
<!-- ǥ ��ܹ� ��-->
<br>
<!-- ǥ �߰��� ����-->
<table width="100%" cellpadding="1" cellspacing="1" class="a">	
<tr valign="bottom">       
    <td align="left">
		�� �Ϸ� �������� �Ǹŵ� ���� ����̸�, �Ϸ翡 �ѹ� ������ ������Ʈ �˴ϴ�.
		<br>&nbsp; &nbsp; ù����� : ������ü ��������� ó�� ��ϵ� ��¥
    </td>
    <td align="right">
		����Ⱓ :
		<%' drawmaechuldatefg "datefg" ,datefg ,""%>
		<input type="hidden" name="datefg" value="<%=datefg%>">
		<% DrawDateBoxdynamic myyyy1,"myyyy1",myyyy2,"myyyy2",mmm1,"mmm1",mmm2,"mmm2",mdd1,"mdd1",mdd2,"mdd2" %>
    </td>        
</tr>
</form>
</table>
<!-- ǥ �߰��� ��-->

<table width="100%" cellspacing="1" cellpadding="3" class="a" bgcolor=#3d3d3d>
<tr height="25" bgcolor="FFFFFF">
    <td colspan="25">
        �˻���� : <b><%= oreport.FTotalcount %></b> �� �˻� ������ �ִ� 2500�Ǳ��� �������ϴ�.
    </td>
</tr>
<tr bgcolor="#EEEEEE" align="center">
	<td>�귣��</td>
	<td>��������<br>ī�װ�</td>
	<td>����<br>����</td>
	
	<% if C_ADMIN_USER or C_IS_OWN_SHOP then %>
		<td>����</td>
	<% end if %>
	
	<td>����</td>
	
	<% if searchtype = "ipgo" then %>
		<td>ù�԰���</td>
	<% else %>
		<td>ù�����</td>
	<% end if %>

	<td>�Ǹ�<Br>�ϼ�</td>
	<td>
		�����
	</td>
	
	<% if C_ADMIN_USER or C_IS_OWN_SHOP then %>
		<td>
			���԰�
		</td>
	<% end if %>
	
	<td>�Ǹ�<br>����</td>
	<td>�Ǹ�<br>�Ǽ�</td>
	
	<% if C_ADMIN_USER or C_IS_OWN_SHOP then %>
		<td>
			�������
		</td>
	<% end if %>
	
	<td>
		���<br>���ܰ�
	</td>	
	<td>
		�����<br>����
	</td>	
	<td>��������<br>���MD</td>
	<td>���<br>����</td>
	<td>�귣��<br>��ǰ��</td>
	<td>���</td>
</tr>
<%
if oreport.FResultCount > 0 then
	
for i=0 to oreport.FResultCount - 1

'/�԰��� ����
if searchtype = "ipgo" then
	datelen = datediff("d",Left(oreport.FItemList(i).ffirstipgodate,10),mtoDate)
'/����� ����
else
	datelen = datediff("d",Left(oreport.FItemList(i).fregdate,10),mtoDate)
end if

datelen2 = datediff("d",mfromDate,mtoDate)
if datelen2<datelen then datelen=datelen2
%>
<tr bgcolor="#FFFFFF" height=24 align="center">
	<td>
		<%= oreport.FItemList(i).Fsocname_kor %>(<%= oreport.FItemList(i).FUserId %>)
	</td>
	<td><%= oreport.FItemList(i).fcate_nm1 %></td>	
	<td><%= oreport.FItemList(i).GetMaeipDivName %></td>
	
	<% if C_ADMIN_USER or C_IS_OWN_SHOP then %>
		<td><%= oreport.FItemList(i).Fdefaultmargine %></td>
	<% end if %>
	
	<td><%= oreport.FItemList(i).fshopid %></td>
	
	<% if searchtype = "ipgo" then %>
		<td><%= Left(oreport.FItemList(i).ffirstipgodate,10) %></td>
	<% else %>
		<td><%= Left(oreport.FItemList(i).fregdate,10) %></td>
	<% end if %>

	<td align="center">
		<%= datelen %>
	</td>
	<td align="right" bgcolor="#E6B9B8"><%= FormatNumber(oreport.FItemList(i).Fsellttl,0) %></td>
	
	<% if C_ADMIN_USER or C_IS_OWN_SHOP then %>
		<td align="right"><%= FormatNumber(oreport.FItemList(i).Fbuyttl,0) %></td>
	<% end if %>
	
	<td align="right"><%= FormatNumber(oreport.FItemList(i).fsellcntsum,0) %></td>
	<td align="right"><%= FormatNumber(oreport.FItemList(i).fsellcnt,0) %></td>
	
	<% if C_ADMIN_USER or C_IS_OWN_SHOP then %>
		<td align="right"><%= FormatNumber(oreport.FItemList(i).Fsellttl-oreport.FItemList(i).Fbuyttl,0) %></td>
	<% end if %>
	
	<td align="right">
		<% if oreport.FItemList(i).Fsellttl <> 0 and oreport.FItemList(i).fsellcnt <> 0 then %>
			<%= FormatNumber(oreport.FItemList(i).Fsellttl/oreport.FItemList(i).fsellcnt,0) %>
		<% else %>
			0
		<% end if %>
	</td>
	<td align="right">
		<% if datelen<>0 then %>
			<%= FormatNumber(oreport.FItemList(i).Fsellttl/datelen,0) %>
		<% end if %>
	</td>	
	<td><%= oreport.FItemList(i).Fmduserid %></td>
	<td><%= oreport.FItemList(i).Fisusing %></td>
	<td><%= oreport.FItemList(i).Fitemcount %></td>
	<td align="center">
		<input type="button" onclick="detailitem('<%= oreport.FItemList(i).FUserId %>','<%= oreport.FItemList(i).fshopid %>','<%=myyyy1%>','<%=mmm1%>','<%=mdd1%>','<%=myyyy2%>','<%=mmm2%>','<%=mdd2%>','<%=datefg%>');" value="��" class="button">
	</td>
</tr>
<% next %>
<% else %>
<tr bgcolor="#FFFFFF" height=24>
	<td align="center" colspan=25>�˻� ����� �����ϴ�.</td>
</tr>
<% end if %>
</table>

<%
set oreport = Nothing
%>
<!-- #include virtual="/common/lib/commonbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/db3close.asp" -->