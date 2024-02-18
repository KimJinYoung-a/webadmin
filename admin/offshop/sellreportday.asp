<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  �������� ���� (��Ÿ���� ��� ĳ�ű��ѵ� ������ ó�� ���Ұ�)
' History : 2009.04.07 ������ ����
'			2010.03.26 �ѿ�� ����
'####################################################
%>
<!-- #include virtual="/common/incSessionAdminOrShop.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/common/lib/commonbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/offshop/offshopsellcls.asp"-->
<%
dim shopid , yyyy1,mm1,dd1,yyyy2,mm2,dd2, yyyymmdd1,yyymmdd2 ,oldlist ,i , datefg ,ooffsell2 ,tmpcnt
dim fromDate,toDate , totalcount, totalitemcnt, totalsellsum, page
dim totsuplyprice , totprofit , totprofit2 , custa ,makerid ,olddatay ,fromDateolddatay ,toDateolddatay
dim tmpselldate , tmptargetmaechul, buyergubun, inc3pl
	olddatay = RequestCheckVar(request("olddatay"),10)
	shopid = requestCheckVar(request("shopid"),32)
	yyyy1 = requestCheckVar(request("yyyy1"),4)
	mm1 = requestCheckVar(request("mm1"),2)
	dd1 = requestCheckVar(request("dd1"),2)
	yyyy2 = requestCheckVar(request("yyyy2"),4)
	mm2 = requestCheckVar(request("mm2"),2)
	dd2 = requestCheckVar(request("dd2"),2)
	oldlist = requestCheckVar(request("oldlist"),2)
	makerid = requestCheckVar(request("makerid"),32)
	page = requestCheckVar(request("page"),10)
	datefg = requestCheckVar(request("datefg"),16)
	buyergubun = requestCheckVar(request("buyergubun"),10)
    inc3pl = requestCheckVar(request("inc3pl"),32)

if datefg = "" then datefg = "maechul"
if page = "" then page = 1

if (yyyy1="") then
	fromDate = DateSerial(Cstr(Year(now())), Cstr(Month(now())), Cstr(day(now()))-7)
else
	fromDate = DateSerial(yyyy1, mm1, dd1)
end if

if (yyyy2="") then yyyy2 = Cstr(Year(now()))
if (mm2="") then mm2 = Cstr(Month(now()))
if (dd2="") then dd2 = Cstr(day(now()))

toDate = DateSerial(yyyy2, mm2, dd2+1)
fromDateolddatay = dateadd("m",-12,fromDate)
toDateolddatay = dateadd("m",-12,dateadd("d",-1,toDate))

yyyy1 = left(fromDate,4)
mm1 = Mid(fromDate,6,2)
dd1 = Mid(fromDate,9,2)

'C_IS_Maker_Upche = TRUE
'C_IS_SHOP = TRUE

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

''��Ÿ�� ������ȸ ����
Dim isFixShopView
IF (session("ssBctID")="doota01") then
    shopid="streetshop014"
    C_IS_SHOP = TRUE
    isFixShopView = TRUE
ENd If

dim ooffsell
set ooffsell = new COffShopSellReport
	ooffsell.FRectShopID = shopid
	ooffsell.FPageSize = 500
	ooffsell.FCurrPage = page
	ooffsell.FRectNormalOnly = "on"
	ooffsell.FRectStartDay = fromDate
	ooffsell.frectdatefg = datefg
	ooffsell.FRectEndDay = toDate
	ooffsell.FRectOldData = oldlist
	ooffsell.frectmakerid = makerid
	ooffsell.frectbuyergubun = buyergubun
	ooffsell.FRectInc3pl = inc3pl

	if C_IS_Maker_Upche then
		ooffsell.GetDaylySumList
	else
		if shopid <> "" then
			ooffsell.GetDaylySumList
		else
			response.write "<script type='text/javascript'>"
			response.write "alert('������ �����Ͻ� �� �˻��ϼ���.');"
			response.write "</script>"
		end if
	end if

%>

<script type='text/javascript'>

function frmsubmit(cholddatay){

	if(cholddatay=='RESETOLDDATAY'){
		frm.olddatay.value = '';
	}

	frm.submit();
}

function cholddatay(){
	//cholddatay = document.getElementsByName("cholddatay")

	if(frm.olddatay.value==''){
		frm.olddatay.value = 'ON';
	} else {
		frm.olddatay.value = '';
	}

	frmsubmit('');
}

function popitemdetail(yyyy1,mm1,dd1,shopid, makerid){
	var popitemdetail = window.open('/admin/offshop/todayselldetail.asp?oldlist=<%=oldlist%>&datefg=<%=datefg%>&yyyy1='+yyyy1+'&mm1='+mm1+'&dd1='+dd1+'&yyyy2='+yyyy1+'&mm2='+mm1+'&dd2='+dd1+'&shopid='+shopid+'&makerid='+makerid+'&buyergubun=<%=buyergubun%>&inc3pl=<%=inc3pl%>&menupos=<%= menupos %>','popitemdetail','width=1024,height=768,scrollbars=yes,resizable=yes');
	popitemdetail.focus();
}

function popjumundetail(yyyy1,mm1,dd1,shopid){
	var popjumundetail = window.open('/admin/offshop/todaysellmaster.asp?oldlist=<%=oldlist%>&datefg=<%=datefg%>&yyyy1='+yyyy1+'&mm1='+mm1+'&dd1='+dd1+'&yyyy2='+yyyy1+'&mm2='+mm1+'&dd2='+dd1+'&shopid='+shopid+'&buyergubun=<%=buyergubun%>&inc3pl=<%=inc3pl%>&menupos=<%= menupos %>','popjumundetail','width=1024,height=768,scrollbars=yes,resizable=yes');
	popjumundetail.focus();
}

function viewcomment(dname)
{
	document.getElementById(""+dname+"").style.display = "block";
}

function notviewcomment(dname)
{
	document.getElementById(""+dname+"").style.display = "none";
}

function pop_manualmaechul(shopid){

	if (shopid==''){
		alert('������ �˻��Ͻ���, ��밡�� �մϴ�');
		return;
	}

	var pop_manualmaechul = window.open('/admin/offshop/maechul/manualmaechul.asp?shopid='+shopid,'pop_manualmaechul','width=1024,height=768,scrollbars=yes,resizable=yes');
	pop_manualmaechul.focus();
}

function pop_manualXLmaechul(shopid){
	var pop_manualmaechul = window.open('<%=stsAdmURL%>/admin/offshop/maechul/manualXLmaechul.asp?shopid='+shopid,'pop_manualXLmaechul','width=1100,height=600,scrollbars=yes,resizable=yes');
	pop_manualmaechul.focus();
}

</script>

<!-- ǥ ��ܹ� ����-->
<table width="100%" align="center" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>" class="a">
<form name="frm" method="get" action="">
<input type="hidden" name="olddatay" value="<%= olddatay %>">
<input type="hidden" name="menupos" value="<%= menupos %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
	<td align="left">
		<table border="0" width="100%" cellpadding="3" cellspacing="0" class="a">
		<tr>
			<td>
				* �Ⱓ : <% drawmaechuldatefg "datefg" ,datefg ,""%>
				<% DrawDateBox yyyy1,yyyy2,mm1,mm2,dd1,dd2 %>
				<input type="checkbox" name="oldlist" <% if oldlist="on" then response.write "checked" %> >3������
				&nbsp;&nbsp;
				<%
				'����/������
				if (C_IS_SHOP) then
				%>
					<% if (not C_IS_OWN_SHOP and shopid <> "") or (isFixShopView) then %>
						* ���� : <%=shopid%><input type="hidden" name="shopid" value="<%= shopid %>">
					<% else %>
						* ���� : <% drawSelectBoxOffShop "shopid",shopid %>
					<% end if %>
				<% else %>
					<% if not(C_IS_Maker_Upche) then %>
						* ���� : <% drawSelectBoxOffShop "shopid",shopid %>
					<% else %>
						* ���� : <% drawBoxDirectIpchulOffShopByMakerchfg "shopid",shopid,makerid," onchange='frmsubmit(""RESETOLDDATAY"");'","" %>
					<% end if %>
				<% end if %>
				<p>
				<% if (C_IS_Maker_Upche) then %>
					* �귣�� : <%= makerid %><br>
					<input type="hidden" name="makerid" value="<%= makerid %>">
				<% else %>
					* �귣�� : <% drawSelectBoxDesignerwithName "makerid",makerid %>
				<% end if %>
				&nbsp;&nbsp;
				* ��������: <% drawoffshop_commoncode "buyergubun", buyergubun, "buyergubun", "MAIN", "", " onchange='frmsubmit(""RESETOLDDATAY"");'" %>
	            &nbsp;&nbsp;
	            <b>* ����ó����</b>
	            <% Call draw3plMeachulComboBox("inc3pl",inc3pl) %>
			</td>
		</tr>
		</table>
    </td>
		<td  width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="�˻�" onClick="frmsubmit('RESETOLDDATAY');">
	</td>
</tr>
</form>
</table>
<!-- ǥ ��ܹ� ��-->
<br>
<!-- ǥ �߰��� ����-->
<table width="100%" align="center" cellpadding="1" cellspacing="1" class="a">
    <tr valign="bottom">
        <td align="left">
        	�� ������ �ֹ��� �������� ���� �˴ϴ�.
	    </td>
	    <td align="right">
	    	<% if C_ADMIN_USER then %>
				<input type="button" onclick="pop_manualmaechul('<%=shopid%>')" value="���������" class="button">
				<input type="button" onclick="pop_manualXLmaechul('<%=shopid%>')" value="���������(�ϰ�)" class="button">
	    	<% end if %>
        </td>
	</tr>
</table>
<!-- ǥ �߰��� ��-->

<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<tr bgcolor="#FFFFFF">
	<td colspan="25">
		�˻���� : <b><%=ooffsell.FresultCount%></b>&nbsp;&nbsp;<% if ooffsell.FresultCount = "400" then response.write "�ִ� 400��" %>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td>����</td>
	<td>
		<% if datefg = "maechul" then %>
			������
		<% else %>
			�ֹ���
		<% end if %>
	</td>
	<td>����</td>
	
	<% if not(C_IS_Maker_Upche) then %>
		<td>����</td>
	<% end if %>

	<td>�ֹ�<br>�Ǽ�</td>
	<td>�����</td>

	<% if C_ADMIN_USER or (C_IS_OWN_SHOP and session("ssBctId") <> "doota01") then %>
		<% if not(C_IS_Maker_Upche) then %>
			<td>��ǥ<Br>�޼���</td>
		<% end if %>
	<% end if %>

	<% if C_ADMIN_USER or C_IS_OWN_SHOP or C_IS_Maker_Upche then %>
		<td>���Ծ�</td>
	<% end if %>

	<% if C_ADMIN_USER or C_IS_OWN_SHOP then %>
		<td>������</td>
		<td>���ͱݾ�</td>
	<% end if %>

	<% if not(C_IS_Maker_Upche) then %>
		<td>���ܰ�</td>
	<% end if %>

	<td>���</td>
</tr>

<%
tmpselldate = ""
totalcount = 0
totalitemcnt = 0
totalsellsum = 0
totsuplyprice = 0
totprofit2 = 0
totprofit = 0
custa = 0
tmpcnt = 0
tmptargetmaechul = 0

if ooffsell.FresultCount>0 then

For i = 0 To ooffsell.FResultCount - 1

if tmpselldate <> left(ooffsell.FItemList(i).FTerm,7) and i <> 0 then
%>
	<tr align="center" bgcolor="#f1f1f1">
		<td colspan=3><%= tmpselldate %> ���հ�</td>
		
		<% if not(C_IS_Maker_Upche) then %>
			<td></td>
		<% end if %>

		<td>
			<%= FormatNumber(totalcount,0) %>
		</td>
		<td align="right"><% = FormatNumber(totalsellsum,0) %></td>

		<% if C_ADMIN_USER or (C_IS_OWN_SHOP and session("ssBctId") <> "doota01") then %>
			<% if not(C_IS_Maker_Upche) then %>
				<td align="right">
					<% if totalsellsum <> 0 and tmptargetmaechul <> 0 then %>
						<% response.write round(((totalsellsum/tmptargetmaechul) *100),1) %> %
					<% end if %>
				</td>
			<% end if %>
		<% end if %>

		<% if C_ADMIN_USER or C_IS_OWN_SHOP or C_IS_Maker_Upche then %>
			<td align="right"><% = FormatNumber(totsuplyprice,0) %></td>
		<% end if %>

		<% if C_ADMIN_USER or C_IS_OWN_SHOP then %>
			<td align="right"><% = round(totprofit2/tmpcnt,0) %>%</td>
			<td align="right"><% = FormatNumber(totprofit,0) %></td>
		<% end if %>

		<% if not(C_IS_Maker_Upche) then %>
			<td align="right">
				<%'= FormatNumber(custa/tmpcnt,0) %>
				<%= FormatNumber(totalsellsum/totalcount,0) %>
			</td>
		<% end if %>

		<td></td>
	</tr>
<%
	tmpselldate = ""
	totalcount = 0
	totalitemcnt = 0
	totalsellsum = 0
	totsuplyprice = 0
	totprofit2 = 0
	totprofit = 0
	custa = 0
	tmpcnt = 0
	tmptargetmaechul = 0
end if

	tmpcnt = tmpcnt + 1
	tmpselldate = left(ooffsell.FItemList(i).FTerm,7)
	tmptargetmaechul = ooffsell.FItemList(i).ftargetmaechul

	totalitemcnt = totalitemcnt + ooffsell.FItemList(i).fitemcnt
	totalcount = totalcount + ooffsell.FItemList(i).FCount
	totalsellsum = totalsellsum + ooffsell.FItemList(i).FSum
	totsuplyprice = totsuplyprice + ooffsell.FItemList(i).fsuplyprice
	totprofit = totprofit + ooffsell.FItemList(i).FSum - ooffsell.FItemList(i).fsuplyprice

	if ooffsell.FItemList(i).fsuplyprice <> 0 and ooffsell.FItemList(i).FSum <> 0 then
		totprofit2 = totprofit2 + (100-((ooffsell.FItemList(i).fsuplyprice)/(ooffsell.FItemList(i).FSum)*100*100)/100)
	end if
	if ooffsell.FItemList(i).FSum <> 0 and ooffsell.FItemList(i).FCount <> 0 then
		custa = custa + (ooffsell.FItemList(i).FSum / ooffsell.FItemList(i).FCount)
	end if
%>
<tr align="center" bgcolor="#FFFFFF">
	<td>
		<%= ooffsell.FItemList(i).FShopid %>
	</td>

	<td>
		<%= getweekendcolor(ooffsell.FItemList(i).FTerm) %>
	</td>
	<td>
		<%= getweekend(ooffsell.FItemList(i).FTerm) %>
	</td>
	
	<% if not(C_IS_Maker_Upche) then %>
		<td>
			<%
				If ooffsell.FItemList(i).FWeather <> "" Then
					Response.Write WeatherImage(Split(ooffsell.FItemList(i).FWeather,"||")(0),"22","")
					If Split(ooffsell.FItemList(i).FWeather,"||")(1) <> "" Then
					%>
						&nbsp;<span style="cursor:pointer;" onMouseOver="viewcomment('div<%=i%>');" onMouseOut="notviewcomment('div<%=i%>');">[��]</span>
						<div id="div<%=i%>" style="display:none;border-width:1px; width:100px; border-style:solid;position:absolute;z-index:1;background-color:white;padding:2 2 2 2;"><%=Split(ooffsell.FItemList(i).FWeather,"||")(1)%></div>
					<%
					End IF
				End IF
			%>
		</td>
	<% end if %>

	<td>
		<%= ooffsell.FItemList(i).FCount %>
	</td>
	<td align="right" bgcolor="#E6B9B8"><%= FormatNumber(ooffsell.FItemList(i).FSum,0) %></td>

	<% if C_ADMIN_USER or (C_IS_OWN_SHOP and session("ssBctId") <> "doota01") then %>
		<% if not(C_IS_Maker_Upche) then %>
			<td align="right">
				<% if ooffsell.FItemList(i).FSum <> 0 and ooffsell.FItemList(i).ftargetmaechul <> 0 then %>
					<%'= FormatNumber(ooffsell.FItemList(i).ftargetmaechul,0) %>
					<% response.write round(((ooffsell.FItemList(i).FSum/ooffsell.FItemList(i).ftargetmaechul) *100),1) %> %
				<% end if %>
			</td>
		<% end if %>
	<% end if %>

	<% if C_ADMIN_USER or C_IS_OWN_SHOP or C_IS_Maker_Upche then %>
		<td align="right"><%= FormatNumber(ooffsell.FItemList(i).fsuplyprice,0) %></td>
	<% end if %>

	<% if C_ADMIN_USER or C_IS_OWN_SHOP then %>
		<td align="right">
			<%
			if ooffsell.FItemList(i).fsuplyprice <> 0 and ooffsell.FItemList(i).FSum <> 0 then
				response.write round(100-((ooffsell.FItemList(i).fsuplyprice)/(ooffsell.FItemList(i).FSum)*100*100)/100,1)&"%"
			else
				response.write "0"
			end if
			%>
		</td>
		<td align="right"><%= FormatNumber(ooffsell.FItemList(i).FSum - ooffsell.FItemList(i).fsuplyprice,0) %></td>
	<% end if %>

	<% if not(C_IS_Maker_Upche) then %>
		<td align="right">
			<%
			if ooffsell.FItemList(i).FSum <> 0 and ooffsell.FItemList(i).FCount <> 0 then
				response.write  FormatNumber(ooffsell.FItemList(i).FSum / ooffsell.FItemList(i).FCount,0)
			else
				response.write "0"
			end if
			%>
		</td>
	<% end if %>

	<td width=200>
		<input type="button" onclick="popitemdetail('<%= left(ooffsell.FItemList(i).FTerm,4) %>','<%= mid(ooffsell.FItemList(i).FTerm,6,2) %>','<%= right(ooffsell.FItemList(i).FTerm,2) %>','<%= ooffsell.FItemList(i).FShopid %>','<%= makerid %>');" value="��ǰ��" class="button">

		<% if not(C_IS_Maker_Upche) then %>
			<input type="button" onclick="popjumundetail('<%= left(ooffsell.FItemList(i).FTerm,4) %>','<%= mid(ooffsell.FItemList(i).FTerm,6,2) %>','<%= right(ooffsell.FItemList(i).FTerm,2) %>','<%= ooffsell.FItemList(i).FShopid %>');" value="�ֹ���" class="button">
		<% end if %>
	</td>
</tr>

<% Next %>
<tr align="center" bgcolor="#f1f1f1">
	<td colspan=3><%= tmpselldate %> ���հ�</td>
	
	<% if not(C_IS_Maker_Upche) then %>
		<td></td>
	<% end if %>

	<td>
		<%= FormatNumber(totalcount,0) %>
	</td>
	<td align="right"><% = FormatNumber(totalsellsum,0) %></td>

	<% if C_ADMIN_USER or (C_IS_OWN_SHOP and session("ssBctId") <> "doota01") then %>
		<% if not(C_IS_Maker_Upche) then %>
			<td align="right">
				<% if totalsellsum <> 0 and tmptargetmaechul <> 0 then %>
					<% response.write round(((totalsellsum/tmptargetmaechul) *100),1) %> %
				<% end if %>
			</td>
		<% end if %>
	<% end if %>

	<% if C_ADMIN_USER or C_IS_OWN_SHOP or C_IS_Maker_Upche then %>
		<td align="right"><% = FormatNumber(totsuplyprice,0) %></td>
	<% end if %>

	<% if C_ADMIN_USER or C_IS_OWN_SHOP then %>
		<td align="right">
		<% if totalsellsum<>0 then %>
		<% response.write round((100-(totsuplyprice/totalsellsum) *100),1) %> %
		<% end if %>
		</td>
		<td align="right"><% = FormatNumber(totprofit,0) %></td>
	<% end if %>

	<% if not(C_IS_Maker_Upche) then %>
		<td align="right">
			<%'= FormatNumber(custa/tmpcnt,0) %>
			<%= FormatNumber(totalsellsum/totalcount,0) %>
		</td>
	<% end if %>

	<td></td>
</tr>
<% ELSE %>
<tr  align="center" bgcolor="#FFFFFF">
	<td colspan="25">��ϵ� ������ �����ϴ�.</td>
</tr>
<%END IF%>
</table>

<Br>
<!-- �׼� ���� -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a">
<tr>
	<td align="left">
		<input type="checkbox" name="cholddatay" <% if olddatay="ON" then response.write " checked" %> onclick='cholddatay();'>���⵵ �񱳳��� ǥ��( <%= fromDateolddatay%> - <%=toDateolddatay%> )
	</td>
	<td align="right">
	</td>
</tr>
</table>
<!-- �׼� �� -->

<%
'/���⵵ �񱳳��� ǥ��
if olddatay = "ON" then

set ooffsell2 = new COffShopSellReport
	ooffsell2.FRectShopID = shopid
	ooffsell2.FPageSize = 500
	ooffsell2.FCurrPage = page
	ooffsell2.FRectNormalOnly = "on"
	ooffsell2.FRectStartDay = fromDateolddatay
	ooffsell2.frectdatefg = datefg
	ooffsell2.FRectEndDay = dateadd("d",+1,toDateolddatay)
	ooffsell2.FRectOldData = oldlist
	ooffsell2.frectmakerid = makerid
	ooffsell2.FRectInc3pl = inc3pl

	if C_IS_Maker_Upche then
		ooffsell2.GetDaylySumList
	else
		if shopid <> "" then
			ooffsell2.GetDaylySumList
		end if
	end if

%>
<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<tr bgcolor="#FFFFFF">
	<td colspan="25">
		�˻���� : <b><%=ooffsell2.FresultCount%></b>&nbsp;&nbsp;<% if ooffsell2.FresultCount = "400" then response.write "�ִ� 400��" %>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td>����</td>
	<td>
		<% if datefg = "maechul" then %>
			������
		<% else %>
			�ֹ���
		<% end if %>
	</td>
	<td>����</td>
	
	<% if not(C_IS_Maker_Upche) then %>
		<td>����</td>
	<% end if %>

	<% if not(C_IS_Maker_Upche) then %>
		<td>�ֹ�<br>�Ǽ�</td>
	<% else %>
		<td>�Ǹ�<br>����</td>
	<% end if %>

	<td>�����</td>

	<% if C_ADMIN_USER or (C_IS_OWN_SHOP and session("ssBctId") <> "doota01") then %>
		<% if not(C_IS_Maker_Upche) then %>
			<td>��ǥ<Br>�޼���</td>
		<% end if %>
	<% end if %>

	<% if C_ADMIN_USER or C_IS_OWN_SHOP or C_IS_Maker_Upche then %>
		<td>���Ծ�</td>
	<% end if %>

	<% if C_ADMIN_USER or C_IS_OWN_SHOP then %>
		<td>������</td>
		<td>���ͱݾ�</td>
	<% end if %>

	<% if not(C_IS_Maker_Upche) then %>
		<td>���ܰ�</td>
	<% end if %>

	<td>���</td>
</tr>

<%
tmpselldate = ""
totalcount = 0
totalitemcnt = 0
totalsellsum = 0
totsuplyprice = 0
totprofit2 = 0
totprofit = 0
custa = 0
tmpcnt = 0
tmptargetmaechul = 0

if ooffsell2.FresultCount>0 then

For i = 0 To ooffsell2.FResultCount - 1

if tmpselldate <> left(ooffsell2.FItemList(i).FTerm,7) and i <> 0 then
%>
	<tr align="center" bgcolor="#f1f1f1">
		<td colspan=3><%= tmpselldate %> ���հ�</td>
		
		<% if not(C_IS_Maker_Upche) then %>
			<td></td>
		<% end if %>

		<td>
			<%= FormatNumber(totalcount,0) %>
		</td>
		<td align="right"><% = FormatNumber(totalsellsum,0) %></td>

		<% if C_ADMIN_USER or (C_IS_OWN_SHOP and session("ssBctId") <> "doota01") then %>
			<% if not(C_IS_Maker_Upche) then %>
				<td align="right">
					<% if totalsellsum <> 0 and tmptargetmaechul <> 0 then %>
						<% response.write round(((totalsellsum/tmptargetmaechul) *100),1) %> %
					<% end if %>
				</td>
			<% end if %>
		<% end if %>

		<% if C_ADMIN_USER or C_IS_OWN_SHOP or C_IS_Maker_Upche then %>
			<td align="right"><% = FormatNumber(totsuplyprice,0) %></td>
		<% end if %>

		<% if C_ADMIN_USER or C_IS_OWN_SHOP then %>
			<td align="right">
			<% if totalsellsum<>0 then %>
    		<% response.write round((100-(totsuplyprice/totalsellsum) *100),1) %> %
    		<% end if %>
			</td>
			<td align="right"><% = FormatNumber(totprofit,0) %></td>
		<% end if %>

		<% if not(C_IS_Maker_Upche) then %>
			<td align="right">
				<%'= FormatNumber(custa/tmpcnt,0) %>
				<%= FormatNumber(totalsellsum/totalcount,0) %>
			</td>
		<% end if %>

		<td></td>
	</tr>
<%
	tmpselldate = ""
	totalcount = 0
	totalitemcnt = 0
	totalsellsum = 0
	totsuplyprice = 0
	totprofit2 = 0
	totprofit = 0
	custa = 0
	tmpcnt = 0
	tmptargetmaechul = 0
end if

	tmpcnt = tmpcnt + 1
	tmpselldate = left(ooffsell2.FItemList(i).FTerm,7)
	tmptargetmaechul = ooffsell2.FItemList(i).ftargetmaechul

	totalitemcnt = totalitemcnt + ooffsell2.FItemList(i).fitemcnt
	totalcount = totalcount + ooffsell2.FItemList(i).FCount
	totalsellsum = totalsellsum + ooffsell2.FItemList(i).FSum
	totsuplyprice = totsuplyprice + ooffsell2.FItemList(i).fsuplyprice
	totprofit = totprofit + ooffsell2.FItemList(i).FSum - ooffsell2.FItemList(i).fsuplyprice

	if ooffsell2.FItemList(i).fsuplyprice <> 0 and ooffsell2.FItemList(i).FSum <> 0 then
		totprofit2 = totprofit2 + (100-((ooffsell2.FItemList(i).fsuplyprice)/(ooffsell2.FItemList(i).FSum)*100*100)/100)
	end if
	if ooffsell2.FItemList(i).FSum <> 0 and ooffsell2.FItemList(i).FCount <> 0 then
		custa = custa + (ooffsell2.FItemList(i).FSum / ooffsell2.FItemList(i).FCount)
	end if
%>
<tr align="center" bgcolor="#FFFFFF">
	<td>
		<%= ooffsell2.FItemList(i).FShopid %>
	</td>

	<td>
		<%= getweekendcolor(ooffsell2.FItemList(i).FTerm) %>
	</td>
	<td>
		<%= getweekend(ooffsell2.FItemList(i).FTerm) %>
	</td>
	
	<% if not(C_IS_Maker_Upche) then %>
		<td>
			<%
				If ooffsell2.FItemList(i).FWeather <> "" Then
					Response.Write WeatherImage(Split(ooffsell2.FItemList(i).FWeather,"||")(0),"22","")
					If Split(ooffsell2.FItemList(i).FWeather,"||")(1) <> "" Then
					%>
						&nbsp;<span style="cursor:pointer;" onMouseOver="viewcomment('div<%=i%>');" onMouseOut="notviewcomment('div<%=i%>');">[��]</span>
						<div id="div<%=i%>" style="display:none;border-width:1px; width:100px; border-style:solid;position:absolute;z-index:1;background-color:white;padding:2 2 2 2;"><%=Split(ooffsell2.FItemList(i).FWeather,"||")(1)%></div>
					<%
					End IF
				End IF
			%>
		</td>
	<% end if %>

	<td>
		<%= ooffsell2.FItemList(i).FCount %>
	</td>
	<td align="right" bgcolor="#E6B9B8"><%= FormatNumber(ooffsell2.FItemList(i).FSum,0) %></td>

	<% if C_ADMIN_USER or (C_IS_OWN_SHOP and session("ssBctId") <> "doota01") then %>
		<% if not(C_IS_Maker_Upche) then %>
			<td align="right">
				<% if ooffsell2.FItemList(i).FSum <> 0 and ooffsell2.FItemList(i).ftargetmaechul <> 0 then %>
					<%'= FormatNumber(ooffsell2.FItemList(i).ftargetmaechul,0) %>
					<% response.write round(((ooffsell2.FItemList(i).FSum/ooffsell2.FItemList(i).ftargetmaechul) *100),1) %> %
				<% end if %>
			</td>
		<% end if %>
	<% end if %>

	<% if C_ADMIN_USER or C_IS_OWN_SHOP or C_IS_Maker_Upche then %>
		<td align="right"><%= FormatNumber(ooffsell2.FItemList(i).fsuplyprice,0) %></td>
	<% end if %>

	<% if C_ADMIN_USER or C_IS_OWN_SHOP then %>
		<td align="right">
			<%
			if ooffsell2.FItemList(i).fsuplyprice <> 0 and ooffsell2.FItemList(i).FSum <> 0 then
				response.write round(100-((ooffsell2.FItemList(i).fsuplyprice)/(ooffsell2.FItemList(i).FSum)*100*100)/100,1)&"%"
			else
				response.write "0"
			end if
			%>
		</td>
		<td align="right"><%= FormatNumber(ooffsell2.FItemList(i).FSum - ooffsell2.FItemList(i).fsuplyprice,0) %></td>
	<% end if %>

	<% if not(C_IS_Maker_Upche) then %>
		<td align="right">
			<%
			if ooffsell2.FItemList(i).FSum <> 0 and ooffsell2.FItemList(i).FCount <> 0 then
				response.write  FormatNumber(ooffsell2.FItemList(i).FSum / ooffsell2.FItemList(i).FCount,0)
			else
				response.write "0"
			end if
			%>
		</td>
	<% end if %>

	<td width=200>
		<input type="button" onclick="popitemdetail('<%= left(ooffsell2.FItemList(i).FTerm,4) %>','<%= mid(ooffsell2.FItemList(i).FTerm,6,2) %>','<%= right(ooffsell2.FItemList(i).FTerm,2) %>','<%= ooffsell2.FItemList(i).FShopid %>','<%= makerid %>');" value="��ǰ��" class="button">
		<% if not(C_IS_Maker_Upche) then %>
			<input type="button" onclick="popjumundetail('<%= left(ooffsell2.FItemList(i).FTerm,4) %>','<%= mid(ooffsell2.FItemList(i).FTerm,6,2) %>','<%= right(ooffsell2.FItemList(i).FTerm,2) %>','<%= ooffsell2.FItemList(i).FShopid %>');" value="�ֹ���" class="button">
		<% end if %>
	</td>
</tr>

<% Next %>
<tr align="center" bgcolor="#f1f1f1">
	<td colspan=3><%= tmpselldate %> ���հ�</td>
	
	<% if not(C_IS_Maker_Upche) then %>
		<td></td>
	<% end if %>

	<td>
		<%= FormatNumber(totalcount,0) %>
	</td>
	<td align="right"><% = FormatNumber(totalsellsum,0) %></td>

	<% if C_ADMIN_USER or (C_IS_OWN_SHOP and session("ssBctId") <> "doota01") then %>
		<% if not(C_IS_Maker_Upche) then %>
			<td align="right">
				<% if totalsellsum <> 0 and tmptargetmaechul <> 0 then %>
					<% response.write round(((totalsellsum/tmptargetmaechul) *100),1) %> %
				<% end if %>
			</td>
		<% end if %>
	<% end if %>

	<% if C_ADMIN_USER or C_IS_OWN_SHOP or C_IS_Maker_Upche then %>
		<td align="right"><% = FormatNumber(totsuplyprice,0) %></td>
	<% end if %>

	<% if C_ADMIN_USER or C_IS_OWN_SHOP then %>
		<td align="right">
		    <% if totalsellsum<>0 then %>
    		<% response.write round((100-(totsuplyprice/totalsellsum) *100),1) %> %
    		<% end if %>
		</td>
		<td align="right"><% = FormatNumber(totprofit,0) %></td>
	<% end if %>

	<% if not(C_IS_Maker_Upche) then %>
		<td align="right">
			<%'= FormatNumber(custa/tmpcnt,0) %>
			<%= FormatNumber(totalsellsum/totalcount,0) %>
		</td>
	<% end if %>

	<td></td>
</tr>
<% ELSE %>
<tr  align="center" bgcolor="#FFFFFF">
	<td colspan="25">��ϵ� ������ �����ϴ�.</td>
</tr>
<%END IF%>
</table>

<% end if %>

<%
set ooffsell= Nothing
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/common/lib/commonbodytail.asp"-->