<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/report/reportcls.asp"-->

<%
dim DELIVERDANGA: DELIVERDANGA = chkIIF(date()>="2019-01-01",2500,2000)
dim yyyy1,mm1,dd1,yyyy2,mm2,dd2
dim yyyymmdd1,yyymmdd2
dim fromDate,toDate
dim ck_joinmall,ck_ipjummall,ck_pointmall,research, rdsite
dim oldlist

yyyy1 = request("yyyy1")
mm1 = request("mm1")
dd1 = request("dd1")
yyyy2 = request("yyyy2")
mm2 = request("mm2")
dd2 = request("dd2")

research = request("research")
ck_joinmall = request("ck_joinmall")
ck_ipjummall = request("ck_ipjummall")
ck_pointmall = request("ck_pointmall")
oldlist = request("oldlist")
rdsite = request("rdsite")

if research<>"on" then
	if ck_joinmall="" then ck_joinmall="on"
	if ck_ipjummall="" then ck_ipjummall="on"
	'if ck_pointmall="" then ck_pointmall="on"
end if

if (yyyy1="") then 
    yyyymmdd1 = dateAdd("d",-7,now())
    yyyy1 = Cstr(Year(yyyymmdd1))
    mm1 = Cstr(Month(yyyymmdd1))
    dd1 = Cstr(day(yyyymmdd1))
end if

if (yyyy2="") then yyyy2 = Cstr(Year(now()))
if (mm2="") then mm2 = Cstr(Month(now()))
if (dd2="") then dd2 = Cstr(day(now()))

fromDate = DateSerial(yyyy1, mm1, dd1)
toDate = DateSerial(yyyy2, mm2, dd2+1)

dim oreport
set oreport = new CJumunMaster
oreport.FRectFromDate = fromDate
oreport.FRectToDate = toDate
oreport.FRectJoinMallNotInclude = ck_joinmall
oreport.FRectExtMallNotInclude = ck_ipjummall
oreport.FRectPointNotInclude = ck_pointmall
oreport.FRectOldJumun = oldlist
oreport.FRectRdsite = rdsite

oreport.SearchMallSellrePort5

dim i,p1,p2,p3
dim plussum,plusbuysum,minussum,pluscount,minuscount
dim deliversum

dim jumuntotalsum, miletotalprice, spendmembership, tencardspend, allatdiscountprice
%>


<table width="100%" border="0" cellpadding="5" cellspacing="0" bgcolor="#CCCCCC">
	<form name="frm" method="get" action="">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<input type="hidden" name="research" value="on">
	<tr>
		<td class="a" >
		<input type="checkbox" name="oldlist" <% if oldlist="on" then response.write "checked" %> >6������������
		<br>
		�˻��Ⱓ :
		<% DrawDateBox yyyy1,yyyy2,mm1,mm2,dd1,dd2 %>
		<input type="checkbox" name="ck_joinmall" <% if ck_joinmall="on" then response.write "checked" %> >���޸� ����
		<input type="checkbox" name="ck_ipjummall" <% if ck_ipjummall="on" then response.write "checked" %> >������ ����
		<input type="checkbox" name="ck_pointmall" <% if ck_pointmall="on" then response.write "checked" %> >����Ʈ�� ����
		<input type="checkbox" name="rdsite" <% if rdsite="on" then response.write "checked" %> >������ǸŸ�
		<td class="a" align="right">
			<a href="javascript:document.frm.submit();"><img src="/admin/images/search2.gif" width="74" height="22" border="0"></a>
		</td>
	</tr>
	</form>
</table>

<table width="100%" border="0" cellspacing="1" cellpadding="3" bgcolor="#FFFFFF" class="a">
<tr>
	<td>�ǰ����ݾױ���. <font color="#000000">������ + �÷��� ����</font> <font color="#AAAAAA">ȸ�� - ���̳ʽ� ����</font>
	<img src="/images/dot1.gif" height="5" width="10">����(��ۺ����� �Ǹ���) 
	<img src="/images/dot2.gif" height="5" width="10">����(��ۿ�������) 
	</td>
</tr>
</table>

<table width="100%" border="0" cellspacing="1" cellpadding="3" bgcolor="#EFBE00">
        <tr align="center">
          <td width="120" class="a"><font color="#FFFFFF">�Ⱓ</font></td>
          <td class="a"><font color="#FFFFFF"></font></td>
          <td class="a" width="80"><font color="#FFFFFF">����(��)</font></td>
          <td class="a" width="50"><font color="#FFFFFF">���Կ���</font></td>
          <td class="a" width="50"><font color="#FFFFFF">��ۿ���</font></td>
          <td class="a" width="50"><font color="#FFFFFF">����</font></td>
          <td class="a" width="50"><font color="#FFFFFF">�Ǽ�</font></td>
          <td class="a" width="80"><font color="#FFFFFF">(-)����(��)</font></td>
          <td class="a" width="50"><font color="#FFFFFF">(-)�Ǽ�</font></td>
          
        </tr>

		<% for i=0 to oreport.FResultCount-1 %>
		<%
			if oreport.maxt<>0 then
				p1 = Clng(oreport.FMasterItemList(i).Fselltotal/oreport.maxt*100)
			end if
            
            if oreport.maxt<>0 then
				p3 = Clng(oreport.FMasterItemList(i).Fbuytotal/oreport.maxt*100)
			end if
			
			if oreport.maxc<>0 then
				p2 = Clng(oreport.FMasterItemList(i).Fsellcnt/oreport.maxc*100)
			end if
            
            DELIVERDANGA = chkIIF(oreport.FMasterItemList(i).Fsitename>="2019-01-01",2500,2000)	'�ñ⿡���� ��ۺ� ����
			
			plussum		=	plussum + oreport.FMasterItemList(i).Fselltotal
			plusbuysum  =	plusbuysum + oreport.FMasterItemList(i).Fbuytotal
			minussum	=	minussum + oreport.FMasterItemList(i).Fminustotal
			pluscount	=	pluscount + oreport.FMasterItemList(i).Fsellcnt
			minuscount	=	minuscount + oreport.FMasterItemList(i).Fminuscount
            deliversum  =   deliversum + oreport.FMasterItemList(i).Ftenbeasongcount*DELIVERDANGA
            
			jumuntotalsum = jumuntotalsum + oreport.FMasterItemList(i).Fjumuntotalsum

			miletotalprice = miletotalprice + oreport.FMasterItemList(i).Fmiletotalprice
			spendmembership = spendmembership + oreport.FMasterItemList(i).Fspendmembership
			tencardspend = tencardspend + oreport.FMasterItemList(i).Ftencardspend
			allatdiscountprice = allatdiscountprice + oreport.FMasterItemList(i).Fallatdiscountprice
            
            
		%>
        <tr bgcolor="#FFFFFF" height="10"  class="a">
			<td width="120" height="10" rowspan="2">
	          	<%= oreport.FMasterItemList(i).Fsitename %>(<%= oreport.FMasterItemList(i).GetDpartName %>)
	        </td>
	        <td  height="35" width="50%">
				<div align="left"> <img src="/images/dot1.gif" height="5" width="<%= p1 %>%"></div><br>
				<div align="left"> <img src="/images/dot2.gif" height="5" width="<%= p3 %>%"></div>
				<!--
				<br>
	        	<div align="left"> <img src="/images/dot4.gif" height="3" width="<%= p2 %>%"></div>
	        	-->
	        	
	        </td>
			<td class="a" align="right" width="80">
			<%= FormatNumber(oreport.FMasterItemList(i).Fselltotal,0) %><br>
    		</td>
    		<td class="a" align="right" width="80">
			<%= FormatNumber(oreport.FMasterItemList(i).Fbuytotal,0) %><br>
		    </td>
		    <td class="a" align="right" width="80">
			<%= FormatNumber(oreport.FMasterItemList(i).Ftenbeasongcount*DELIVERDANGA,0) %><br>
		    </td>
		    <td class="a" align="center" width="50">
		    <% if oreport.FMasterItemList(i).Fselltotal<>0 then %>
		        <%= CLng(10000-(oreport.FMasterItemList(i).Fbuytotal+oreport.FMasterItemList(i).Ftenbeasongcount*DELIVERDANGA)/oreport.FMasterItemList(i).Fselltotal*100*100)/100 %>%
		    <% end if %>
		    </td>
			<td class="a" width="50" align="right">
			<%= FormatNumber(oreport.FMasterItemList(i).Fsellcnt,0) %><br>
			</td>
			<td class="a" align="right" width="80">
				<font color="#AAAAAA"><%= FormatNumber(oreport.FMasterItemList(i).Fminustotal,0) %></font>
			</td>
			<td class="a" align="right" width="50">
	          	<font color="#AAAAAA"><%= FormatNumber(oreport.FMasterItemList(i).Fminuscount,0) %></font>
			</td>
        </tr>
        <tr bgcolor="#FFFFFF"  class="a">
        	<td colspan="9" align="right" >
        	<table width="560" border=0 cellspacing=0 cellpadding=0 class="gray">
        	<tr>
        		<td >���ϸ��� : </td>
        		<td><%= FormatNumber(oreport.FMasterItemList(i).Fmiletotalprice,0) %> <font color="#000000">(<% if oreport.FMasterItemList(i).Fjumuntotalsum<>0 then response.write CLng(oreport.FMasterItemList(i).Fmiletotalprice/oreport.FMasterItemList(i).Fjumuntotalsum*100*10)/10 %> %)</font></td>
        		<td>SKT : </td>
        		<td><%= FormatNumber(oreport.FMasterItemList(i).Fspendmembership,0) %> <font color="#000000">(<% if oreport.FMasterItemList(i).Fjumuntotalsum<>0 then response.write CLng(oreport.FMasterItemList(i).Fspendmembership/oreport.FMasterItemList(i).Fjumuntotalsum*100*10)/10 %> %)</font></td>
        		<td>���� : </td>
        		<td><%= FormatNumber(oreport.FMasterItemList(i).Ftencardspend,0) %> <font color="#000000">(<% if oreport.FMasterItemList(i).Fjumuntotalsum<>0 then response.write CLng(oreport.FMasterItemList(i).Ftencardspend/oreport.FMasterItemList(i).Fjumuntotalsum*100*10)/10 %> %)</font></td>
        		<td>�ÿ�ī�� : </td>
        		<td><%= FormatNumber(oreport.FMasterItemList(i).Fallatdiscountprice,0) %> <font color="#000000">(<% if oreport.FMasterItemList(i).Fjumuntotalsum<>0 then response.write CLng(oreport.FMasterItemList(i).Fallatdiscountprice/oreport.FMasterItemList(i).Fjumuntotalsum*100*10)/10 %> %)</font></td>
        	</tr>
	        </table>
        	</td>
        </tr>
        <% next %>
        <tr bgcolor="#FFFFFF" height="10"  class="a">
        	<td>Total</td>
        	<td></td>
        	<td align="right"><%= Formatnumber(plussum,0) %></td>
        	<td align="right"><%= Formatnumber(plusbuysum,0) %></td>
        	<td align="right"><%= Formatnumber(deliversum,0) %></td>
        	
        	<td align="center">
        	    <% if plussum<>0 then %>
		            <%= CLng(10000-(plusbuysum+deliversum)/plussum*100*100)/100 %>%
		        <% end if %>
        	</td>
        	<td align="right"><%= Formatnumber(pluscount,0) %></td>
        	<td align="right"><%= Formatnumber(minussum,0) %></td>
        	<td align="right"><%= Formatnumber(minuscount,0) %></td>
        </tr>
        <tr bgcolor="#FFFFFF">
        	<td colspan="9" align="right">

            	<table width="800" border=0 cellspacing=0 cellpadding=0 class="gray">
            	<tr>
            		<td >�ֹ��� : </td>
            		<td><%= FormatNumber(jumuntotalsum,0) %> </td>
            		<td >���ϸ��� : </td>
            		<td><%= FormatNumber(miletotalprice,0) %> <font color="#000000">(<% if jumuntotalsum<>0 then response.write CLng(miletotalprice/jumuntotalsum*100*10)/10 %> %)</font></td>
            		<td>SKT : </td>
            		<td><%= FormatNumber(spendmembership,0) %> <font color="#000000">(<% if jumuntotalsum<>0 then response.write CLng(spendmembership/jumuntotalsum*100*10)/10 %> %)</font></td>
            		<td>���� : </td>
            		<td><%= FormatNumber(tencardspend,0) %> <font color="#000000">(<% if jumuntotalsum<>0 then response.write CLng(tencardspend/jumuntotalsum*100*10)/10 %> %)</font></td>
            		<td>�ÿ�ī�� : </td>
            		<td><%= FormatNumber(allatdiscountprice,0) %> <font color="#000000">(<% if jumuntotalsum<>0 then response.write CLng(allatdiscountprice/jumuntotalsum*100*10)/10 %> %)</font></td>
            	</tr>
    	        </table>
        	</td>
        </tr>
</table>
<%
set oreport = Nothing
%>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->