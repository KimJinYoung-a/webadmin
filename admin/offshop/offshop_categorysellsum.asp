<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  �������� ī�װ��� ���
'		[OFF]����_������>>�귣�庰������� /admin/offshop/sellreportbrand.asp ���� ���
'		[OFF]����_������>>���Ϻ�����м� /admin/offshop/weeklysellreport.asp ���� ���
' History : 2009.04.07 ������ ����
'			2022.02.09 �ѿ�� ����(�������� ��񿡼� �������� �����۾�)
'####################################################
%>
<!-- #include virtual="/common/incSessionAdminOrShop.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/common/lib/commonbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/offshopclass/offshop_reportcls.asp"-->
<%
const Maxlines = 10

dim i , yyyy1,mm1,dd1,yyyy2,mm2,dd2 , yyyymmdd1,yyymmdd2 , ojumun , fromDate,toDate , weekdate, vTotalSum, vTotalPercent
dim shopid , oldlist , cdl,cdm,cds, searchtype, page ,totalprice,totalea ,totIorgsellprice, catecdnull ,olddatay
dim totsuplyprice, totprofit , totprofit2 , totrealsellprice ,datefg ,offgubun ,makerid, inc3pl
dim fromDateolddatay ,toDateolddatay ,ojumun2, vPurchaseType ,reload, BanPum, buyergubun
	olddatay = RequestCheckVar(request("olddatay"),10)
	shopid  = requestCheckVar(request("shopid"),32)
	yyyy1   = requestCheckVar(request("yyyy1"),4)
	mm1     = requestCheckVar(request("mm1"),2)
	dd1     = requestCheckVar(request("dd1"),2)
	yyyy2   = requestCheckVar(request("yyyy2"),4)
	mm2     = requestCheckVar(request("mm2"),2)
	dd2     = requestCheckVar(request("dd2"),2)
	oldlist = requestCheckVar(request("oldlist"),10)
	cdl     = requestCheckVar(request("cdl"),3)
	cdm     = requestCheckVar(request("cdm"),3)
	cds     = requestCheckVar(request("cds"),3)
	page    = requestCheckVar(request("page"),10)
	searchtype = requestCheckVar(request("searchtype"),1)
	datefg = requestCheckVar(request("datefg"),10)
	offgubun = requestCheckVar(request("offgubun"),32)
	makerid = requestCheckVar(request("makerid"),32)
	catecdnull    = requestCheckVar(request("catecdnull"),32)
	weekdate = requestCheckVar(request("weekdate"),10)
	vPurchaseType = requestCheckVar(request("purchasetype"),2)
	reload = requestCheckVar(request("reload"),2)
	BanPum = requestCheckVar(request("BanPum"),1)
	buyergubun = requestCheckVar(request("buyergubun"),10)
    inc3pl = requestCheckVar(request("inc3pl"),32)

if datefg = "" then datefg = "maechul"
if searchtype="" then searchtype="c"
if page="" then page="1"
if (yyyy1="") then yyyy1 = Cstr(Year(now()))
if (mm1="") then mm1 = Cstr(Month(now()))
if (dd1="") then dd1 = Cstr(day(now()))
if (yyyy2="") then yyyy2 = Cstr(Year(now()))
if (mm2="") then mm2 = Cstr(Month(now()))
if (dd2="") then dd2 = Cstr(day(now()))
if reload <> "on" and offgubun = "" then offgubun = "95"

fromDate = DateSerial(yyyy1, mm1, dd1)
toDate = DateSerial(yyyy2, mm2, dd2+1)

Dim vYYYYold1, vYYYYold2, vMMold1, vMMold2, vDDold1, vDDold2
vYYYYold1 = NullFillWith(request("yyyyold1"), yyyy1-1)
vYYYYold2 = NullFillWith(request("yyyyold2"), yyyy2-1)
vMMold1 = NullFillWith(request("mmold1"), mm1)
vMMold2 = NullFillWith(request("mmold2"), mm2)
vDDold1 = NullFillWith(request("ddold1"), dd1)
vDDold2 = NullFillWith(request("ddold2"), dd2)
fromDateolddatay = DateSerial(vYYYYold1, vMMold1, vDDold1)
toDateolddatay = DateSerial(vYYYYold2, vMMold2, vDDold2)

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
		makerid = session("ssBctID")
	else
		if (Not C_ADMIN_USER) then
		else
		end if
	end if
end if

if shopid<>"" then offgubun=""
if (searchtype="c") and ((cdl<>"") and (cdm<>"") and (cds<>"")) then cds=""

set ojumun = new COffshopReport
	ojumun.FPageSize = 5000
	ojumun.FCurrPage = page
	ojumun.FRectShopID = shopid
	ojumun.FRectStartDay = fromDate
	ojumun.FRectEndDay = toDate
	ojumun.FRectOldData = oldlist
	ojumun.FRectCDL = cdl
	ojumun.FRectCDM = cdm
	ojumun.FRectCDN = cds
	ojumun.frectdatefg = datefg
	ojumun.FRectOffgubun = offgubun
	ojumun.frectmakerid = makerid
	ojumun.frectcatecdnull = catecdnull
	ojumun.frectweekdate = weekdate
	ojumun.FRectBrandPurchaseType = vPurchaseType
	ojumun.frectBanPum = BanPum
	ojumun.frectbuyergubun = buyergubun
	ojumun.FRectInc3pl = inc3pl

if searchtype="i" then
	ojumun.SearchCategorySellItems
elseif cdl="" then
	ojumun.SearchCategorySellrePort
elseif cdm="" then
	ojumun.SearchCategorySellrePort2
elseif cds="" then
	ojumun.SearchCategorySellrePort3
else
	ojumun.SearchCategorySellItems
end if


totprofit2 = 0
totprofit = 0
totsuplyprice = 0
totalprice = 0
totalea = 0
totrealsellprice = 0
vTotalSum = 0
vTotalPercent = 0
totIorgsellprice = 0
%>

<script language='javascript'>

function popOffItemEdit(ibarcode){
	var popwin = window.open('popoffitemedit.asp?barcode=' + ibarcode,'offitemedit','width=500,height=800,resizable=yes,scrollbars=yes');
	popwin.focus();
}

function ReSearch(page,cholddatay){

	//��¥ ��
	var startdate = frm.yyyy1.value + "-" + frm.mm1.value + "-" + frm.dd1.value;
	var enddate = frm.yyyy2.value + "-" + frm.mm2.value + "-" + frm.dd2.value;
    var diffDay = 0;
    var start_yyyy = startdate.substring(0,4);
    var start_mm = startdate.substring(5,7);
    var start_dd = startdate.substring(8,startdate.length);
    var sDate = new Date(start_yyyy, start_mm-1, start_dd);
    var end_yyyy = enddate.substring(0,4);
    var end_mm = enddate.substring(5,7);
    var end_dd = enddate.substring(8,enddate.length);
    var eDate = new Date(end_yyyy, end_mm-1, end_dd);

    diffDay = Math.ceil((eDate.getTime() - sDate.getTime())/(1000*60*60*24));

	if (diffDay > 1095){
		alert('3���������� �ǽð��˻��� �����մϴ�.');
		return;
	}

	if(cholddatay=='RESETOLDDATAY'){
		frm.olddatay.value = '';
	}

	frm.page.value=page;
	frm.submit();
}

function cholddatay(){
	//cholddatay = document.getElementsByName("cholddatay")

	if(frm.olddatay.value==''){
		frm.olddatay.value = 'ON';
	} else {
		frm.olddatay.value = '';
	}

	ReSearch('','');
}

function cholddatayButton()
{
	document.getElementById("warningtext").style.display = "block";
	document.frm.yyyyold1.value = document.getElementsByName("yyyyold11")[0].value;
	document.frm.yyyyold2.value = document.getElementsByName("yyyyold22")[0].value;
	document.frm.mmold1.value = document.getElementsByName("mmold11")[0].value;
	document.frm.mmold2.value = document.getElementsByName("mmold22")[0].value;
	document.frm.ddold1.value = document.getElementsByName("ddold11")[0].value;
	document.frm.ddold2.value = document.getElementsByName("ddold22")[0].value;

	frm.olddatay.value = 'ON';
	ReSearch('','');
}
</script>

<!-- �˻� ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get" action="">
<input type="hidden" name="yyyyold1" value="<%=vYYYYold1%>">
<input type="hidden" name="yyyyold2" value="<%=vYYYYold2%>">
<input type="hidden" name="mmold1" value="<%=vMMold1%>">
<input type="hidden" name="mmold2" value="<%=vMMold2%>">
<input type="hidden" name="ddold1" value="<%=vDDold1%>">
<input type="hidden" name="ddold2" value="<%=vDDold2%>">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="page" value=1>
<input type="hidden" name="olddatay" value="<%= olddatay %>">
<input type="hidden" name="reload" value="on">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
	<td align="left">
		* �Ⱓ :
		<% drawmaechul_datefg "datefg" ,datefg ,""%>
		<% DrawDateBox yyyy1,yyyy2,mm1,mm2,dd1,dd2 %>
		<!--<input type="checkbox" name="oldlist" <%' if oldlist="on" then response.write "checked" %> >2009������������ȸ-->
		&nbsp;&nbsp;&nbsp;
		<%
		'����/������
		if (C_IS_SHOP) then
		%>
			<% if not C_IS_OWN_SHOP and shopid <> "" then %>
				* ���� : <%=shopid%><input type="hidden" name="shopid" value="<%= shopid %>">
			<% else %>
				* ���� : <% Call NewDrawSelectBoxDesignerwithNameAndUserDIV("shopid",shopid, "21") %>
			<% end if %>
		<% else %>
			* ���� : <% Call NewDrawSelectBoxDesignerwithNameAndUserDIV("shopid",shopid, "21") %>
		<% end if %>
		&nbsp;&nbsp;
		* ��ǰ���� :
		<% drawSelectBoxisusingYN "BanPum" , BanPum ," onchange='ReSearch("""",""RESETOLDDATAY"");'" %>
		<p>
		<% if (C_IS_Maker_Upche) then %>
			* �귣�� : <%= makerid %><br>
			<input type="hidden" name="makerid" value="<%= makerid %>">
		<% else %>
			* �귣�� : <% drawSelectBoxDesignerwithName "makerid",makerid %>
		<% end if %>
		&nbsp;&nbsp;
		* �������� : 
		<% drawPartnerCommCodeBox true,"purchasetype","purchasetype",vPurchaseType,"" %>
		&nbsp;&nbsp;
		* ���� ���� : <% drawoffshop_commoncode "offgubun", offgubun, "shopdivithinkso", "", "", " onchange='ReSearch("""",""RESETOLDDATAY"");'" %>
		&nbsp;&nbsp;
		* ����:<% drawweekday_select "weekdate" , weekdate ," onchange='ReSearch("""",""RESETOLDDATAY"");'" %>
		&nbsp;&nbsp;
		* ��������: <% drawoffshop_commoncode "buyergubun", buyergubun, "buyergubun", "MAIN", "", " onchange='ReSearch("""",""RESETOLDDATAY"");'" %>
	</td>
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="�˻�" onClick="ReSearch('','RESETOLDDATAY');">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
        <b>* ����ó����</b>
        <% Call draw3plMeachulComboBox("inc3pl",inc3pl) %>
        &nbsp;&nbsp;
		<!-- #include virtual="/common/module/categoryselectbox.asp"-->
	</td>
</tr>
</table>
<!-- �˻� �� -->

<br>
<!-- ǥ �߰��� ����-->
<table width="100%" align="center" cellpadding="1" cellspacing="1" class="a">
<tr valign="bottom">
    <td align="left">
    	�� ������ �ֹ��� �������� ���� �Ǹ�, �ǽð� �����ʹ� �ֱ� 3�� �����͸� �˻� �����մϴ�.
    </td>
    <td align="right">
		<input type="radio" name="searchtype" value="c" <% if searchtype="c" then response.write "checked" %> >ī�װ��հ�
		<input type="radio" name="searchtype" value="i" <% if searchtype="i" then response.write "checked" %> >�ǸŻ�ǰ���
    </td>
</tr>
</form>
</table>
<!-- ǥ �߰��� ��-->

<%
'/�ǸŻ�ǰ���
if (searchtype="i") then
%>
	<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="15">
			�˻���� : <b><%= ojumun.FResultCount %></b> (�ִ� <%= ojumun.FPageSize %>�� �˻�)
		</td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td>�귣��</td>
		<td>���ڵ�</td>
		<td>��ǰ��</td>
		<td>�ɼ�</td>
		<td>��ǰ����</td>
		<% if (NOT C_InspectorUser) then %>
		<td>�Һ��ڰ�[��ǰ]</td>
		<% end if %>
		<td>�����</td>

		<% if C_ADMIN_USER or C_IS_OWN_SHOP then %>
			<td>�����Ѿ�[��ǰ]</td>
			<td><b>�������</b></td>
			<td>������</td>
		<% end if %>

		<td>%</td>
	</tr>
	<% if ojumun.FResultCount<1 then %>
	<tr bgcolor="#FFFFFF">
		<td colspan="15" align="center">[�˻������ �����ϴ�.]</td>
	</tr>
	<% else %>
	<%
	vTotalSum = ojumun.FTotalSum
	for i=0 to ojumun.FResultCount - 1

	totalea = totalea + ojumun.FItemList(i).Fsellcnt
	totsuplyprice = totsuplyprice + ojumun.FItemList(i).fsuplyprice
	totprofit = totprofit + (ojumun.FItemList(i).FsellSum - ojumun.FItemList(i).fsuplyprice	)
	totrealsellprice = totrealsellprice + ojumun.FItemList(i).Fsellsum
	totIorgsellprice = totIorgsellprice + ojumun.FItemList(i).fIorgsellprice

	if ojumun.FItemList(i).fsuplyprice <> 0 and ojumun.FItemList(i).FsellSum <> 0 then
	totprofit2 = totprofit2 + (100-((ojumun.FItemList(i).fsuplyprice)/(ojumun.FItemList(i).FsellSum)*100*100)/100)
	end if

	if ojumun.FItemList(i).Fsellsum <> 0 and ojumun.FItemList(i).Fsellsum <> "" and vTotalSum <> 0 and vTotalSum <> "" then
		vTotalPercent = vTotalPercent + (ojumun.FItemList(i).Fsellsum/vTotalSum)*100
	else
		vTotalPercent = vTotalPercent + 0
	end if
	%>
	<tr bgcolor="#FFFFFF" onmouseover=this.style.background="#f1f1f1"; onmouseout=this.style.background="#FFFFFF"; align="center">
		<td><%= ojumun.FItemList(i).FMakerid %></td>
		<td><a href="javascript:popOffItemEdit('<%= ojumun.FItemList(i).GetBarCode %>');"><%= ojumun.FItemList(i).GetBarCode %></a></td>
		<td><%= ojumun.FItemList(i).FItemName %></td>
		<td><%= ojumun.FItemList(i).FItemOptionName %></td>
		<td><%= ojumun.FItemList(i).FSellCnt %></td>
		<% if (NOT C_InspectorUser) then %>
		<td align="right"><%= FormatNumber(ojumun.FItemList(i).fIorgsellprice,0) %></td>
	    <% end if %>
		<td align="right" bgcolor="#E6B9B8"><%= FormatNumber(ojumun.FItemList(i).Fsellsum,0) %></td>

		<% if C_ADMIN_USER or C_IS_OWN_SHOP then %>
			<td align="right"><%= FormatNumber(ojumun.FItemList(i).fsuplyprice,0) %></td>
			<td align="right"><b><%= FormatNumber(ojumun.FItemList(i).fsellsum-ojumun.FItemList(i).fsuplyprice,0) %></b></td>
			<td>
				<%
				if ojumun.FItemList(i).fsuplyprice <> 0 and ojumun.FItemList(i).fsellsum <> 0 then
					response.write round(100-((ojumun.FItemList(i).fsuplyprice)/(ojumun.FItemList(i).fsellsum)*100*100)/100,1)&"%"
				else
					response.write "0%"
				end if
				%>
			</td>
		<% end if %>

		<td>
			<% if ojumun.FItemList(i).Fsellsum <> 0 and ojumun.FItemList(i).Fsellsum <> "" and vTotalSum <> 0 and vTotalSum <> "" then %>
				<%=round((ojumun.FItemList(i).Fsellsum/vTotalSum)*100,1)%>%
			<% else %>
				0 %
			<% end if %>
		</td>
	</tr>
	<% next %>
	<tr bgcolor="<%= adminColor("tabletop") %>" align="center">
		<td colspan=4>�Ѱ�</td>
		<td><% = FormatNumber(totalea,0) %></td>
		<% if (NOT C_InspectorUser) then %>
		<td align="right"><% = FormatNumber(totIorgsellprice,0) %></td>
	    <% end if %>
		<td align="right"><% = FormatNumber(totrealsellprice,0) %></td>

		<% if C_ADMIN_USER or C_IS_OWN_SHOP then %>
			<td align="right"><% = FormatNumber(totsuplyprice,0) %></td>
			<td align="right"><b><% = FormatNumber(totprofit,0) %></b></td>
			<td><%if   totrealsellprice <> 0 then%><% = round(100-(totsuplyprice/totrealsellprice*100*100)/100,1)  %><%else%>0<%end if%>%</td>
		<% end if %>

		<td><%= round(vTotalPercent,1) %>%</td>
	</tr>
	<% end if %>
	</table>

<%
'/ī�װ��հ�
else
%>
	<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="15">
			�˻���� : <b><%= ojumun.FResultCount %></b> (�ִ� <%= ojumun.FPageSize %>�� �˻�)
		</td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td class="a">ī����</font></td>
		<!--<td width="600"></td>//-->
		<td>��ǰ����</td>
		<% if (NOT C_InspectorUser) then %>
		<td>�Һ��ڰ�[��ǰ]</td>
		<% end if %>
		<td>�����</td>

		<% if C_ADMIN_USER or C_IS_OWN_SHOP then %>
			<td>�����Ѿ�[��ǰ]</td>
			<td><b>�������</b></td>
			<td>������</td>
		<% end if %>

		<td>%</td>
	</tr>
	<% if ojumun.FresultCount<1 then %>
	<tr bgcolor="#FFFFFF">
		<td colspan="15" align="center">[�˻������ �����ϴ�.]</td>
	</tr>
	<% else %>
	<%
	vTotalSum = ojumun.FTotalSum
	for i=0 to ojumun.FResultCount - 1

	totalprice = totalprice + ojumun.FItemList(i).Fsellsum
	totalea = totalea + ojumun.FItemList(i).Fsellcnt
	totsuplyprice = totsuplyprice + ojumun.FItemList(i).fsuplyprice
	totprofit = totprofit + (ojumun.FItemList(i).FsellSum - ojumun.FItemList(i).fsuplyprice)
	totIorgsellprice = totIorgsellprice + ojumun.FItemList(i).fIorgsellprice

	if ojumun.FItemList(i).fsuplyprice <> 0 and ojumun.FItemList(i).FsellSum <> 0 then
		totprofit2 = totprofit2 + (100-((ojumun.FItemList(i).fsuplyprice)/(ojumun.FItemList(i).FsellSum)*100*100)/100)
	end if

	if ojumun.FItemList(i).Fsellsum <> 0 and ojumun.FItemList(i).Fsellsum <> "" and vTotalSum <> 0 and vTotalSum <> "" then
		vTotalPercent = vTotalPercent + (ojumun.FItemList(i).Fsellsum/vTotalSum)*100
	else
		vTotalPercent = vTotalPercent + 0
	end if
	%>
	<tr bgcolor="#FFFFFF" onmouseover=this.style.background="#f1f1f1"; onmouseout=this.style.background="#FFFFFF"; align="center">
		<td>
			<% if (IsNULL(ojumun.FItemList(i).FCateCDL)) or ((ojumun.FItemList(i).FCateCDL="") and (ojumun.FItemList(i).FCateCDM="") and (ojumun.FItemList(i).FCateCDN="")) then %>
				<a href="?searchtype=i&datefg=<%=datefg%>&offgubun=<%=offgubun%>&makerid=<%=makerid%>&oldlist=<%= oldlist %>&shopid=<%= shopid %>&yyyy1=<%= yyyy1 %>&yyyy2=<%= yyyy2 %>&mm1=<%= mm1 %>&mm2=<%= mm2 %>&dd1=<%= dd1 %>&dd2=<%= dd2 %>&cdl=<%= ojumun.FItemList(i).FCateCDL %>&cdm=<%= ojumun.FItemList(i).FCateCDM %>&cds=<%= ojumun.FItemList(i).FCateCDN %>&catecdnull=ON&weekdate=<%=weekdate%>&purchasetype=<%=vpurchasetype%>&inc3pl=<%=inc3pl%>&menupos=<%= menupos %>">
				<%= ojumun.FItemList(i).FCateName %>...</a>
			<% else %>
				<a href="?searchtype=<%= chkIIF(cdl<>"" and cdm<>"" and ojumun.FItemList(i).FCateCDN<>"","i",searchtype) %>&datefg=<%=datefg%>&offgubun=<%=offgubun%>&makerid=<%=makerid%>&oldlist=<%= oldlist %>&shopid=<%= shopid %>&yyyy1=<%= yyyy1 %>&yyyy2=<%= yyyy2 %>&mm1=<%= mm1 %>&mm2=<%= mm2 %>&dd1=<%= dd1 %>&dd2=<%= dd2 %>&cdl=<%= ojumun.FItemList(i).FCateCDL %>&cdm=<%= ojumun.FItemList(i).FCateCDM %>&cds=<%= ojumun.FItemList(i).FCateCDN %>&weekdate=<%=weekdate%>&purchasetype=<%=vpurchasetype%>&inc3pl=<%=inc3pl%>&menupos=<%= menupos %>">
				<%= ojumun.FItemList(i).FCateName %>
				<%= ChkIIF(IsNULL(ojumun.FItemList(i).FCateName) or (ojumun.FItemList(i).FCateName=""),ojumun.FItemList(i).FCateCDL & "-" & ojumun.FItemList(i).FCateCDM & "-" & ojumun.FItemList(i).FCateCDN,"") %></a>
			<% end if %>
		</td>
		<!--
		<td height="10" width="600">
			<%' if  ojumun.FItemList(i).Fsellsum<>0 and ojumun.FItemList(i).Fsellsum <> "" and ojumun.maxt <> 0 and ojumun.maxt <> "" then %>
				<div align="left"> <img src="/images/dot1.gif" height="4" width="<%' CLng((ojumun.FItemList(i).Fsellsum/ojumun.maxt)*600) %>"></div><br>
				<div align="left"> <img src="/images/dot2.gif" height="4" width="<%' CLng((ojumun.FItemList(i).Fsellcnt/ojumun.maxc)*600) %>"></div>
			<%' end if %>
		</td>
		//-->
		<td><%= ojumun.FItemList(i).Fsellcnt %></td>
		<% if (NOT C_InspectorUser) then %>
		<td align="right">
			<%= FormatNumber(FormatCurrency(ojumun.FItemList(i).fIorgsellprice),0) %>
		</td>
	    <% end if %>
		<td align="right" bgcolor="#E6B9B8">
			<%= FormatNumber(FormatCurrency(ojumun.FItemList(i).Fsellsum),0) %>
		</td>

		<% if C_ADMIN_USER or C_IS_OWN_SHOP then %>
			<td align="right"><%= FormatNumber(ojumun.FItemList(i).fsuplyprice,0) %></td>
			<td align="right"><b><%= FormatNumber(ojumun.FItemList(i).FsellSum - ojumun.FItemList(i).fsuplyprice,0) %></b></td>
			<td>
				<%
				if ojumun.FItemList(i).fsuplyprice <> 0 and ojumun.FItemList(i).fsellsum <> 0 then
					response.write round(100-((ojumun.FItemList(i).fsuplyprice)/(ojumun.FItemList(i).fsellsum)*100*100)/100,1)&"%"
				else
					response.write "0"
				end if
				%>
			</td>
		<% end if %>

		<td>
			<% if ojumun.FItemList(i).Fsellsum <> 0 and ojumun.FItemList(i).Fsellsum <> "" and vTotalSum <> 0 and vTotalSum <> "" then %>
				<%=round((ojumun.FItemList(i).Fsellsum/vTotalSum)*100,1)%>%
			<% else %>
				0 %
			<% end if %>
		</td>
	</tr>
	<% next %>
	<tr bgcolor="<%= adminColor("tabletop") %>" align="center">
		<td>�Ѱ�</td>
		<td><%=FormatNumber(totalea,0)%></td>
		<% if (NOT C_InspectorUser) then %>
		<td align="right"><%=FormatNumber(totIorgsellprice,0)%></td>
	    <% end if %>
		<td align="right"><% = FormatNumber(totalprice,0) %></td>

		<% if C_ADMIN_USER or C_IS_OWN_SHOP then %>
			<td align="right"><% = FormatNumber(totsuplyprice,0) %></td>
			<td align="right"><b><% = FormatNumber(totprofit,0) %></b></td>
			<td> <%if  totalprice <> 0 then%><% = round(100-(totsuplyprice/totalprice*100*100)/100,1)  %><%else%>0<%end if%>%</td>
		<% end if %>

		<td><%=vTotalPercent%>%</td>
	</tr>
	<% end if %>
	</table>

	<Br>
	<!-- �׼� ���� -->

	<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a">
	<tr>
		<td align="left">
			<input type="checkbox" name="cholddatay" <% if olddatay="ON" then response.write " checked" %> onclick='cholddatay();'>
			<% DrawDateBoxdynamic vYYYYold1,"yyyyold11",vYYYYold2,"yyyyold22",vMMold1,"mmold11",vMMold2,"mmold22",vDDold1,"ddold11",vDDold2,"ddold22" %>
			<input type="button" value="��˻�" class="button" onClick="cholddatayButton();">
			<center><div id="warningtext" style="display:none"><br><b>�� <font color="blue">�˻� ��ư�� Ŭ���� �� DB�� �д� ���̴� �������� ���ٰ� </font><font color="red">���� Ŭ������ ������. DB�����ϴ�.</font></b></div></center>
		</td>
		<td align="right">
		</td>
	</tr>
	</table>
	<!-- �׼� �� -->

	<%
	if olddatay = "ON" then

		set ojumun2 = new COffshopReport
			ojumun2.FPageSize = 5000
			ojumun2.FCurrPage = page
			ojumun2.FRectShopID = shopid
			ojumun2.FRectStartDay = fromDateolddatay
			ojumun2.FRectEndDay = dateadd("d",+1,toDateolddatay)
			ojumun2.FRectOldData = oldlist
			ojumun2.FRectCDL = cdl
			ojumun2.FRectCDM = cdm
			ojumun2.FRectCDN = cds
			ojumun2.frectdatefg = datefg
			ojumun2.FRectOffgubun = offgubun
			ojumun2.frectmakerid = makerid
			ojumun2.frectcatecdnull = catecdnull
			ojumun2.frectweekdate = weekdate
			ojumun2.FRectBrandPurchaseType = vPurchaseType
			ojumun2.frectBanPum = BanPum

			if searchtype="i" then
				'ojumun2.SearchCategorySellItems
			elseif cdl="" then
				ojumun2.SearchCategorySellrePort
			elseif cdm="" then
				ojumun2.SearchCategorySellrePort2
			elseif cds="" then
				ojumun2.SearchCategorySellrePort3
			else
				'ojumun2.SearchCategorySellItems
			end if

		totprofit2 = 0
		totprofit = 0
		totsuplyprice = 0
		totalprice = 0
		totalea = 0
		totrealsellprice = 0
		vTotalSum = 0
		vTotalPercent = 0
		totIorgsellprice = 0
	%>
		<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
		<tr height="25" bgcolor="FFFFFF">
			<td colspan="15">
				�˻���� : <b><%= ojumun2.FResultCount %></b> (�ִ� <%= ojumun2.FPageSize %>�� �˻�)
			</td>
		</tr>
		<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
			<td class="a">ī����</font></td>
			<!--<td width="600"></td>//-->
			<td>��ǰ����</td>
			<% if (NOT C_InspectorUser) then %>
			<td>�Һ��ڰ�[��ǰ]</td>
		    <% end if %>
			<td>�����</td>

			<% if C_ADMIN_USER or C_IS_OWN_SHOP then %>
				<td>�����Ѿ�[��ǰ]</td>
				<td><b>�������</b></td>
				<td>������</td>
			<% end if %>

			<td>%</td>
		</tr>
		<% if ojumun2.FresultCount<1 then %>
		<tr bgcolor="#FFFFFF">
			<td colspan="15" align="center">[�˻������ �����ϴ�.]</td>
		</tr>
		<% else %>
		<%
		vTotalSum = ojumun2.FTotalSum
		vTotalPercent = 0

		'''2013/01/15 ������ �߰�
		IF cdl<>"" and cdm<>""  then
		    yyyy1 = vYYYYold1
		    yyyy2 = vYYYYold2
		    mm1  = vMMold1
		    mm2  = vMMold2
		    dd1  = vDDold1
		    dd2  = vDDold2
		end if

		for i=0 to ojumun2.FResultCount - 1

		totalprice = totalprice + ojumun2.FItemList(i).Fsellsum
		totalea = totalea + ojumun2.FItemList(i).Fsellcnt
		totsuplyprice = totsuplyprice + ojumun2.FItemList(i).fsuplyprice
		totprofit = totprofit + (ojumun2.FItemList(i).FsellSum - ojumun2.FItemList(i).fsuplyprice)
		totIorgsellprice = totIorgsellprice + ojumun2.FItemList(i).fIorgsellprice

		if ojumun2.FItemList(i).fsuplyprice <> 0 and ojumun2.FItemList(i).FsellSum <> 0 then
		totprofit2 = totprofit2 + (100-((ojumun2.FItemList(i).fsuplyprice)/(ojumun2.FItemList(i).FsellSum)*100*100)/100)
		end if

		if ojumun2.FItemList(i).Fsellsum <> 0 and ojumun2.FItemList(i).Fsellsum <> "" and vTotalSum <> 0 and vTotalSum <> "" then
			vTotalPercent = vTotalPercent + (ojumun2.FItemList(i).Fsellsum/vTotalSum)*100
		else
			vTotalPercent = vTotalPercent + 0
		end if
		%>
		<tr bgcolor="#FFFFFF" onmouseover=this.style.background="#f1f1f1"; onmouseout=this.style.background="#FFFFFF"; align="center">
			<td>
				<% if (IsNULL(ojumun2.FItemList(i).FCateCDL)) or ((ojumun2.FItemList(i).FCateCDL="") and (ojumun2.FItemList(i).FCateCDM="") and (ojumun2.FItemList(i).FCateCDN="")) then %>
					<a href="?searchtype=i&menupos=<%= menupos %>&datefg=<%=datefg%>&offgubun=<%=offgubun%>&makerid=<%=makerid%>&oldlist=<%= oldlist %>&shopid=<%= shopid %>&yyyy1=<%= vYYYYold1 %>&yyyy2=<%= vYYYYold2 %>&mm1=<%= vMMold1 %>&mm2=<%= vMMold2 %>&dd1=<%= vDDold1 %>&dd2=<%= vDDold2 %>&cdl=<%= ojumun2.FItemList(i).FCateCDL %>&cdm=<%= ojumun2.FItemList(i).FCateCDM %>&cds=<%= ojumun2.FItemList(i).FCateCDN %>&catecdnull=ON&weekdate=<%=weekdate%>"><%= ojumun2.FItemList(i).FCateName %>...</a>
				<% else %>
					<a href="?searchtype=<%= chkIIF(cdl<>"" and cdm<>"" and ojumun2.FItemList(i).FCateCDN<>"","i",searchtype) %>&menupos=<%= menupos %>&datefg=<%=datefg%>&offgubun=<%=offgubun%>&makerid=<%=makerid%>&oldlist=<%= oldlist %>&shopid=<%= shopid %>&yyyy1=<%= yyyy1 %>&yyyy2=<%= yyyy2 %>&mm1=<%= mm1 %>&mm2=<%= mm2 %>&dd1=<%= dd1 %>&dd2=<%= dd2 %>&cdl=<%= ojumun2.FItemList(i).FCateCDL %>&cdm=<%= ojumun2.FItemList(i).FCateCDM %>&cds=<%= ojumun2.FItemList(i).FCateCDN %>&weekdate=<%=weekdate%>"><%= ojumun2.FItemList(i).FCateName %> <%= ChkIIF(IsNULL(ojumun2.FItemList(i).FCateName) or (ojumun2.FItemList(i).FCateName=""),ojumun2.FItemList(i).FCateCDL & "-" & ojumun2.FItemList(i).FCateCDM & "-" & ojumun2.FItemList(i).FCateCDN,"") %></a>
				<% end if %>
			</td>
			<!--
			<td height="10" width="600">
				<%' if  ojumun2.FItemList(i).Fsellsum<>0 and ojumun2.FItemList(i).Fsellsum <> "" and ojumun.maxt <> 0 and ojumun.maxt <> "" then %>
					<div align="left"> <img src="/images/dot1.gif" height="4" width="<%' CLng((ojumun2.FItemList(i).Fsellsum/ojumun2.maxt)*600) %>"></div><br>
					<div align="left"> <img src="/images/dot2.gif" height="4" width="<%' CLng((ojumun2.FItemList(i).Fsellcnt/ojumun2.maxc)*600) %>"></div>
				<%' end if %>
			</td>
			//-->
			<td><%= ojumun2.FItemList(i).Fsellcnt %></td>
			<% if (NOT C_InspectorUser) then %>
			<td align="right">
				<%= FormatNumber(FormatCurrency(ojumun2.FItemList(i).fIorgsellprice),0) %>
			</td>
		    <% end if %>
			<td bgcolor="#E6B9B8" align="right">
				<%= FormatNumber(FormatCurrency(ojumun2.FItemList(i).Fsellsum),0) %>
			</td>

			<% if C_ADMIN_USER or C_IS_OWN_SHOP then %>
				<td align="right"><%= FormatNumber(ojumun2.FItemList(i).fsuplyprice,0) %></td>
				<td align="right"><b><%= FormatNumber(ojumun2.FItemList(i).FsellSum - ojumun2.FItemList(i).fsuplyprice,0) %></b></td>
				<td>
					<%
					if ojumun2.FItemList(i).fsuplyprice <> 0 and ojumun2.FItemList(i).fsellsum <> 0 then
						response.write round(100-((ojumun2.FItemList(i).fsuplyprice)/(ojumun2.FItemList(i).fsellsum)*100*100)/100,1)&"%"
					else
						response.write "0"
					end if
					%>
				</td>
			<% end if %>

			<td>
				<% if ojumun2.FItemList(i).Fsellsum <> 0 and ojumun2.FItemList(i).Fsellsum <> "" and vTotalSum <> 0 and vTotalSum <> "" then %>
					<%=round((ojumun2.FItemList(i).Fsellsum/vTotalSum)*100,1)%>%
				<% else %>
					0 %
				<% end if %>
			</td>
		</tr>
		<% next %>
		<tr bgcolor="<%= adminColor("tabletop") %>" align="center">
			<td>�Ѱ�</td>
			<td><%=FormatNumber(totalea,0)%></td>
			<% if (NOT C_InspectorUser) then %>
			<td align="right"><%=FormatNumber(totIorgsellprice,0)%></td>
		    <% end if %>
			<td align="right"><% = FormatNumber(totalprice,0) %></td>

			<% if C_ADMIN_USER or C_IS_OWN_SHOP then %>
				<td align="right"><% = FormatNumber(totsuplyprice,0) %></td>
				<td align="right"><b><% = FormatNumber(totprofit,0) %></b></td>
				<td><% = round(100-(totsuplyprice/totalprice*100*100)/100,1)  %>%</td>
			<% end if %>

			<td><%=vTotalPercent%>%</td>
		</tr>
		<% end if %>
		</table>
	<% end if %>
<% end if %>

<%
set ojumun = Nothing
set ojumun2 = Nothing
%>
<!-- #include virtual="/common/lib/commonbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
