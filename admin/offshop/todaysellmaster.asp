<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  �������� ���� �ֹ��� �� ���������� NO ����¡ ����
' History : 2009.04.07 ������ ����
'			2010.03.26 �ѿ�� ����
'####################################################
%>
<!-- #include virtual="/common/incSessionAdminOrShop.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/common/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/offshop/offshopsellcls.asp"-->
<%
dim shopid, oldlist , datefg , prejumunno , makerid , menupos ,yyyy1,mm1,dd1,yyyy2,mm2,dd2, toDate,fromDate
dim cardMinusTotal, cashMinusTotal, cardMinusCnt, cashMinusCnt, buyergubun, inc3pl
dim etcTotal, etcCnt, etcMinusTotal, etcMinusCnt ,i,totalsum ,cardtotal, cashtotal, cardcnt, cashcnt
dim extTotal, extCnt, extMinusTotal, extMinusCnt
	yyyy1 = requestCheckVar(request("yyyy1"),4)
	mm1 = requestCheckVar(request("mm1"),2)
	dd1 = requestCheckVar(request("dd1"),2)
	yyyy2 = requestCheckVar(request("yyyy2"),4)
	mm2 = requestCheckVar(request("mm2"),2)
	dd2 = requestCheckVar(request("dd2"),2)
	menupos = requestCheckVar(request("menupos"),10)
	shopid = requestCheckVar(request("shopid"),32)
	oldlist = requestCheckVar(request("oldlist"),10)
	datefg = requestCheckVar(request("datefg"),32)
	makerid = requestCheckVar(request("makerid"),32)
	buyergubun = requestCheckVar(request("buyergubun"),32)
    inc3pl = requestCheckVar(request("inc3pl"),32)

if datefg = "" then datefg = "maechul"

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
	ooffsell.FRectShopid = shopid
	ooffsell.FRectNormalOnly = "on"
	ooffsell.FRectOldData = oldlist
    ooffsell.FRectTerms = ""
    ooffsell.FRectStartDay = fromDate
    ooffsell.frectdatefg = datefg
    ooffsell.FRectDesigner = makerid
    ooffsell.FRectEndDay = toDate
    ooffsell.FRectbuyergubun = buyergubun
	ooffsell.FRectInc3pl = inc3pl    
	ooffsell.GetDaylySellJumunList

totalsum =0
cardtotal =0
cashtotal =0
cardcnt   =0
cashcnt   =0
cardMinusTotal =0
cashMinusTotal =0
cardMinusCnt   =0
cashMinusCnt   =0
etcTotal        =0
etcCnt          =0
etcMinusTotal   =0
etcMinusCnt     =0
extTotal        =0
extCnt          =0
extMinusTotal   =0
extMinusCnt     =0
%>

<script type='text/javascript'>

function frmsubmit(){

	frm.submit();
}

</script>

<!-- ǥ ��ܹ� ����-->
<table width="100%" align="center" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>" class="a">
<form name="frm" method="get" action="">
<input type="hidden" name="menupos" value="<%= menupos %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
	<td align="left">  
		<table border="0" width="100%" cellpadding="3" cellspacing="0" class="a">
		<tr>
			<td>
				* �Ⱓ :
				<% drawmaechuldatefg "datefg" ,datefg ,""%> 
				<% DrawDateBox yyyy1,yyyy2,mm1,mm2,dd1,dd2 %>
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
					* ���� : <% drawSelectBoxOffShop "shopid",shopid %>
				<% end if %>				
				<p>
				<% if C_IS_Maker_Upche then %>
					* �귣�� : <%= makerid %><input type="hidden" name="makerid" value="<%= makerid %>">
				<% else %>
					* �귣�� : <% drawSelectBoxDesignerwithName "makerid",makerid %>
				<% end if %>
				&nbsp;&nbsp;
				* ��������: <% drawoffshop_commoncode "buyergubun", buyergubun, "buyergubun", "MAIN", "", " onchange='frmsubmit();'" %>
	            &nbsp;&nbsp;
	            <b>* ����ó����</b>
	            <% Call draw3plMeachulComboBox("inc3pl",inc3pl) %>				
			</td>
		</tr>
		</table> 
    </td>
		<td  width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="�˻�" onClick="frmsubmit();">
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
    </td>
    <td align="right">	       
    </td>        
</tr>	
</table>
<!-- ǥ �߰��� ��-->
<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<tr bgcolor="#FFFFFF" height="25">
	<td colspan="20">
		�˻���� : <b><%=ooffsell.FresultCount%></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td>
		<% if datefg = "maechul" then %>
			<%= chkIIF(shopid="cafe002","������","�ֹ���ȣ") %>
		<% else %>
			<%= chkIIF(shopid="cafe002","�ֹ���","�ֹ���ȣ") %>
		<% end if %>	
	</td>
	<td></td>
	<td>��ǰ��</td>
	<td>�ǸŰ�</td>
	<td>�����ݾ�</td>
	<td>����</td>
	
	<% if shopid<>"cafe002" then %>
		<td>���ֹ���</td>
	<% end if %>
</tr>
<%
if ooffsell.FresultCount > 0 then
	
for i=0 to ooffsell.FresultCount-1

if prejumunno<>ooffsell.FItemList(i).ForderNo then
	
	totalsum = totalsum + ooffsell.FItemList(i).Frealsum
	if (ooffsell.FItemList(i).Fcardsum>0) then
        cardtotal = cardtotal + ooffsell.FItemList(i).Fcardsum
        cardcnt   = cardcnt + 1
    elseif (ooffsell.FItemList(i).Fcardsum<0) then
        cardMinusTotal = cardMinusTotal + ooffsell.FItemList(i).Fcardsum
        cardMinusCnt   =cardMinusCnt + 1
    end if
    
    if (ooffsell.FItemList(i).Fcashsum>0) then
        cashtotal = cashtotal + ooffsell.FItemList(i).Fcashsum
        cashcnt   = cashcnt + 1
    elseif (ooffsell.FItemList(i).Fcashsum<0) then
        cashMinusTotal = cashMinusTotal + ooffsell.FItemList(i).Fcashsum
        cashMinusCnt   =cashMinusCnt + 1
    end if
    
    if (ooffsell.FItemList(i).FgiftcardPaysum>0) then
        etcTotal = etcTotal + ooffsell.FItemList(i).FgiftcardPaysum
        etcCnt   = etcCnt + 1
    elseif (ooffsell.FItemList(i).FgiftcardPaysum<0) then
        etcMinusTotal = etcMinusTotal + ooffsell.FItemList(i).FgiftcardPaysum
        etcMinusCnt   =etcMinusCnt + 1
    end if

    if (ooffsell.FItemList(i).FextPaysum>0) then
        extTotal = extTotal + ooffsell.FItemList(i).FextPaysum
        extCnt   = extCnt + 1
    elseif (ooffsell.FItemList(i).FextPaysum<0) then
        extMinusTotal = extMinusTotal + ooffsell.FItemList(i).FextPaysum
        extMinusCnt   =extMinusCnt + 1
    end if

prejumunno = ooffsell.FItemList(i).ForderNo
%>
<tr bgcolor="#EEEEEE" align="center">
	<td align="center"><%= chkIIF(shopid="cafe002",ooffsell.FItemList(i).Fshopregdate,ooffsell.FItemList(i).ForderNo) %></td>
	<td><font color="<%= ooffsell.FItemList(i).JumunMethodColor %>"><%= ooffsell.FItemList(i).JumunMethodName %></font></td>
	<td><%= ooffsell.FItemList(i).Fpointuserno %></td>
	<td align="right"><%= FormatNumber(ooffsell.FItemList(i).Ftotalsum,0) %></td>
	<td align="right"><%= FormatNumber(ooffsell.FItemList(i).Frealsum,0) %></td>
	<td align="center"></td>
	
	<% if shopid<>"cafe002" then %>
		<td><%= ooffsell.FItemList(i).Fshopregdate %></td>
	<% end if %>
</tr>
<% end if %>
<tr bgcolor="#FFFFFF" align="center">
	<td></td>
	<td><%= ooffsell.FItemList(i).FMakerID %></td>
	<td><%= ooffsell.FItemList(i).FItemName %> <%= ooffsell.FItemList(i).FItemOptionName %></td>
	
	<% if ooffsell.FItemList(i).FItemNo<0 then %>
		<td align="right"><font color=red><%= FormatNumber(ooffsell.FItemList(i).FSellPrice,0) %></font></td>
		<td align="right"><font color=red><%= FormatNumber(ooffsell.FItemList(i).FRealSellPrice,0) %></font></td>
		<td align="center"><font color=red><%= ooffsell.FItemList(i).FItemNo %></font></td>
	<% else %>
		<td align="right"><%= FormatNumber(ooffsell.FItemList(i).FSellPrice,0) %></td>
		<td align="right"><%= FormatNumber(ooffsell.FItemList(i).FRealSellPrice,0) %></td>
		<td align="center"><%= ooffsell.FItemList(i).FItemNo %></td>
	<% end if %>
	
	<% if shopid<>"cafe002" then %>
		<td align="right"></td>
	<% end if %>
</tr>
<% next %>
<tr bgcolor="#FFFFFF" align="center">
	<td colspan="2"><b>�Ѱ�</b></td>
	<td colspan="6" align="right">
		<table width=440 border=0 cellspacing=0 cellpadding=0 class="a">
		<tr>
		    <td>���� :</td>
		    <td align="right"><%= FormatNumber(cashtotal,0) %> ��</td>
		    <td align="center">(<%= FormatNumber(cashcnt,0) %> ��)</td>
		    <td width=10></td>
		    <td align="right"><font color="red"><%= FormatNumber(cashMinusTotal,0) %> ��</td>
		    <td align="center"><font color="red">(<%= FormatNumber(cashMinusCnt,0) %> ��)</font></td>
		    <td align="right"><%= FormatNumber(cashtotal + cashMinusTotal,0) %> ��</td>
		</tr>
		<tr>
		    <td>ī�� :</td>
		    <td align="right"><%= FormatNumber(cardtotal,0) %> ��</td>
		    <td align="center">(<%= FormatNumber(cardcnt,0) %> ��)</td>
		    <td></td>
		    <td align="right"><font color="red"><%= FormatNumber(cardMinusTotal,0) %> ��</font></td>
		    <td align="center"><font color="red">(<%= FormatNumber(cardMinusCnt,0) %> ��)</font></td>
		    <td align="right"><%= FormatNumber(cardtotal + cardMinusTotal,0) %> ��</td>
		</tr>
		<tr>
		    <td>��ǰ�� :</td>
		    <td align="right"><%= FormatNumber(etcTotal,0) %> ��</td>
		    <td align="center">(<%= FormatNumber(etccnt,0) %> ��)</td>
		    <td></td>
		    <td align="right"><font color="red"><%= FormatNumber(etcMinusTotal,0) %> ��</font></td>
		    <td align="center"><font color="red">(<%= FormatNumber(etcMinusCnt,0) %> ��)</font></td>
		    <td align="right"><%= FormatNumber(etcTotal + etcMinusTotal,0) %> ��</td>
		</tr>
		<tr>
		    <td>��Ÿ���� :</td>
		    <td align="right"><%= FormatNumber(extTotal,0) %> ��</td>
		    <td align="center">(<%= FormatNumber(extcnt,0) %> ��)</td>
		    <td></td>
		    <td align="right"><font color="red"><%= FormatNumber(extMinusTotal,0) %> ��</font></td>
		    <td align="center"><font color="red">(<%= FormatNumber(extMinusCnt,0) %> ��)</font></td>
		    <td align="right"><%= FormatNumber(extTotal + extMinusTotal,0) %> ��</td>
		</tr>
		<tr>
		    <td>�հ� :</td>
		    <td align="right"><%= FormatNumber(cashtotal + cardtotal + etcTotal + extTotal,0) %> ��</td>
		    <td align="center">(<%= FormatNumber(cashcnt + cardcnt + etccnt + extcnt,0) %> ��)</td>
		    <td></td>
		    <td align="right"><font color="red"><%= FormatNumber(cashMinusTotal + cardMinusTotal + etcMinusTotal + extMinusTotal,0) %> ��</td>
		    <td align="center"><font color="red">(<%= FormatNumber(cashMinusCnt + cardMinusCnt + etcMinusCnt + extMinusCnt,0) %> ��)</font></td>
		    <td align="right"><%= FormatNumber(totalsum,0) %> ��</td>
		</tr>
		</table>
	</td>
</tr>
<% else %>
<tr  align="center" bgcolor="#FFFFFF">
	<td colspan="20">��ϵ� ������ �����ϴ�.</td>
</tr>
<% end if %>
</table>

<%
set ooffsell = Nothing
%>
<!-- #include virtual="/common/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->