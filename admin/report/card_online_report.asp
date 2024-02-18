<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/report/reportcls.asp"-->

<%
dim yyyy1,mm1,dd1,yyyy2,mm2,dd2, vIsOldOrder
dim yyyymmdd1,yyymmdd2
dim fromDate,toDate
dim ordertype

yyyy1 = request("yyyy1")
mm1 = request("mm1")
dd1 = request("dd1")
yyyy2 = request("yyyy2")
mm2 = request("mm2")
dd2 = request("dd2")
vIsOldOrder = request("isoldorder")
If vIsOldOrder = "" Then
	vIsOldOrder = "n"
End If

ordertype = request("ordertype")
if ordertype = "" then ordertype = "D"

if (yyyy1="") then yyyy1 = Cstr(Year(now()))
if (mm1="") then mm1 = Cstr(Month(now()))
if (dd1="") then dd1 = "1"

if (yyyy2="") then yyyy2 = Cstr(Year(now()))
if (mm2="") then mm2 = Cstr(Month(now()))
if (dd2="") then dd2 = Cstr(day(now()))

fromDate = DateSerial(yyyy1, mm1, dd1)
toDate = DateSerial(yyyy2, mm2, dd2+1)

dim oreport
set oreport = new CJumunMaster
oreport.FRectFromDate = fromDate
oreport.FRectToDate = toDate
'oreport.FRectOrdertype = ordertype

oreport.FRectOldJumun = vIsOldOrder

if ordertype="D" then
oreport.SearchCardOnline
else
oreport.SearchCardOnlineMonth
end if

dim i,p1,p2
dim prename
dim buftext, bufname, bufimage
dim sumtotal
dim ch1,ch2,ch3,ch4,ch5,ch6
%>

<table width="100%" border="0" cellpadding="5" cellspacing="0" bgcolor="#CCCCCC">
	<form name="frm" method="get" action="">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<input type="hidden" name="research" value="on">
	<tr>
		<td class="a" >
		검색기간 :
		<% DrawDateBox yyyy1,yyyy2,mm1,mm2,dd1,dd2 %>
 		<input type="radio" name="ordertype" value="D" <% if ordertype = "D" then response.write "checked" %>> 일별 <input type="radio" name="ordertype" value="M" <% if ordertype = "M" then response.write "checked" %>> 월별
 		<br>
 		<input type="radio" name="isoldorder" value="n" <% If vIsOldOrder = "n" Then Response.Write "checked" End If %>><%=Month(DateAdd("m",-6,now))%>월~현재<%=Month(now)%>월&nbsp;&nbsp;&nbsp;
 		<input type="radio" name="isoldorder" value="y" <% If vIsOldOrder = "y" Then Response.Write "checked" End If %>><%=Month(DateAdd("m",-6,now))%>월~현재<%=Month(now)%>월 이전 데이터
		<td class="a" align="right">
			<a href="javascript:document.frm.submit();"><img src="/admin/images/search2.gif" width="74" height="22" border="0"></a>
		</td>
	</tr>
	</form>
</table>
<table width="100%" border="0" cellspacing="1" cellpadding="3" bgcolor="#EFBE00">
        <tr align="center">
          <td width="120" class="a"><font color="#FFFFFF">기간</font></td>
          <td class="a" width="600"><font color="#FFFFFF"></font></td>
          <td class="a" width="200"><font color="#FFFFFF"></font></td>
        </tr>

		<% for i=0 to oreport.FResultCount-1 %>
		<%
			if oreport.maxt<>0 then
				p1 = Clng(oreport.FMasterItemList(i).Fselltotal/oreport.maxt*70)
			end if
		%>

		<% if (prename<>oreport.FMasterItemList(i).Fsitename) then %>
		<% if (prename<>"") then %>
        <tr bgcolor="#FFFFFF" height="10"  class="a">
			  <td width="120" height="10">
			  <%= bufname %>
	          </td>
	          <td >
		          <table border="0" class="a" width="500">
				  <tr>
				  <%
				  if Clng(ch1/sumtotal*100)>=10 then
				   bufimage = replace(bufimage,"[7]",CStr(Clng(ch1/sumtotal*100)) + "%")
				  else
				   bufimage = replace(bufimage,"[7]","")
				  end if
				  %>

				  <%
				  if Clng(ch2/sumtotal*100)>=10 then
				   bufimage = replace(bufimage,"[100]",CStr(Clng(ch2/sumtotal*100)) + "%")
				  else
				   bufimage = replace(bufimage,"[100]","")
				  end if
				  %>

				  <%
				  if Clng(ch3/sumtotal*100)>=10 then
				   bufimage = replace(bufimage,"[30]",CStr(Clng(ch3/sumtotal*100)) + "%")
				  else
				   bufimage = replace(bufimage,"[30]","")
				  end if
				  %>

				  <%
				  if Clng(ch4/sumtotal*100)>=10 then
				   bufimage = replace(bufimage,"[50]",CStr(Clng(ch4/sumtotal*100)) + "%")
				  else
				   bufimage = replace(bufimage,"[50]","")
				  end if
				  %>

				  <%
				  if Clng(ch5/sumtotal*100)>=10 then
				   bufimage = replace(bufimage,"[80]",CStr(Clng(ch5/sumtotal*100)) + "%")
				  else
				   bufimage = replace(bufimage,"[80]","")
				  end if
				  %>

				  <%
				  if Clng(ch6/sumtotal*100)>=10 then
				   bufimage = replace(bufimage,"[90]",CStr(Clng(ch6/sumtotal*100)) + "%")
				  else
				   bufimage = replace(bufimage,"[90]","")
				  end if

				  bufimage = replace(bufimage,"[0]","")
				  %>


				  <%= bufimage %>
				  <td><%= FormatNumber(sumtotal,0) %></td>
				  </tr>
				  </table>
	          </td>
			  <td class="a" width="200" align="right">
			    <%= buftext %>
			  </td>
		  </td>
        </tr>
        <%
        	buftext = ""
        	bufimage = ""
        	sumtotal = 0
        	ch1 = 0
        	ch2 = 0
        	ch3 = 0
        	ch4 = 0
        	ch5 = 0
        	ch6 = 0
        %>
        <% end if %>
        <% end if %>
        <%
        bufname = oreport.FMasterItemList(i).Fsitename + "(" + oreport.FMasterItemList(i).GetDpartName + ")"
		buftext = buftext +  FormatNumber(oreport.FMasterItemList(i).Fselltotal,0) + "원 (" + FormatNumber(oreport.FMasterItemList(i).Fsellcnt,0) + "건)" + "<img src='/images/dot" + Trim(oreport.FMasterItemList(i).Faccountdiv) + ".gif' height='4' width='20'>" + Trim(oreport.FMasterItemList(i).JumunMethodName) + "<br>"
		bufimage = bufimage + "<td background='/images/dot" + Trim(oreport.FMasterItemList(i).Faccountdiv) + ".gif' height='20' width='" +  CStr(p1) + "%'>[" + Trim(oreport.FMasterItemList(i).Faccountdiv) + "]</td>"
		sumtotal = sumtotal + oreport.FMasterItemList(i).Fselltotal
        prename = oreport.FMasterItemList(i).Fsitename

        if oreport.FMasterItemList(i).Faccountdiv=7 then
        	ch1 = oreport.FMasterItemList(i).Fselltotal
        elseif oreport.FMasterItemList(i).Faccountdiv=100 then
        	ch2 = oreport.FMasterItemList(i).Fselltotal
        elseif oreport.FMasterItemList(i).Faccountdiv=30 then
        	ch3 = oreport.FMasterItemList(i).Fselltotal
        elseif oreport.FMasterItemList(i).Faccountdiv=50 then
        	ch4 = oreport.FMasterItemList(i).Fselltotal
        elseif oreport.FMasterItemList(i).Faccountdiv=80 then
        	ch5 = oreport.FMasterItemList(i).Fselltotal
        elseif oreport.FMasterItemList(i).Faccountdiv=90 then
        	ch6 = oreport.FMasterItemList(i).Fselltotal
        end if
        %>
        <% next %>
        <tr bgcolor="#FFFFFF" height="10"  class="a">
			  <td width="120" height="10">
			  <%= bufname %>
	          </td>
	          <td >
		          <table border="0" class="a" width="500">
				  <tr>
				  <%
				  if Clng(ch1/sumtotal*100)>=10 then
				   bufimage = replace(bufimage,"[7]",CStr(Clng(ch1/sumtotal*100)) + "%")
				  else
				   bufimage = replace(bufimage,"[7]","")
				  end if
				  %>

				  <%
				  if Clng(ch2/sumtotal*100)>=10 then
				   bufimage = replace(bufimage,"[100]",CStr(Clng(ch2/sumtotal*100)) + "%")
				  else
				   bufimage = replace(bufimage,"[100]","")
				  end if
				  %>

				  <%
				  if Clng(ch3/sumtotal*100)>=10 then
				   bufimage = replace(bufimage,"[30]",CStr(Clng(ch3/sumtotal*100)) + "%")
				  else
				   bufimage = replace(bufimage,"[30]","")
				  end if
				  %>

				  <%
				  if Clng(ch4/sumtotal*100)>=10 then
				   bufimage = replace(bufimage,"[50]",CStr(Clng(ch4/sumtotal*100)) + "%")
				  else
				   bufimage = replace(bufimage,"[50]","")
				  end if
				  %>

				  <%
				  if Clng(ch5/sumtotal*100)>=10 then
				   bufimage = replace(bufimage,"[80]",CStr(Clng(ch5/sumtotal*100)) + "%")
				  else
				   bufimage = replace(bufimage,"[80]","")
				  end if
				  %>

				  <%
				  if Clng(ch6/sumtotal*100)>=10 then
				   bufimage = replace(bufimage,"[90]",CStr(Clng(ch6/sumtotal*100)) + "%")
				  else
				   bufimage = replace(bufimage,"[90]","")
				  end if

				  bufimage = replace(bufimage,"[0]","")
				  %>
				  <%= bufimage %>
				  <td><%= FormatNumber(sumtotal,0) %></td>
				  </tr>
				  </table>
	          </td>
			  <td class="a" width="200" align="right">
			    <%= buftext %>
			  </td>
		  </td>
        </tr>
</table>
<%
set oreport = Nothing
%>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->