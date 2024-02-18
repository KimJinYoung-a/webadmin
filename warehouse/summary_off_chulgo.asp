<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dblogicsopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/order/baljucls_off.asp" -->


<%
dim yyyy1,yyyy2,mm1,mm2,dd1,dd2
yyyy1 = request("yyyy1")
mm1 = request("mm1")
dd1 = request("dd1")

yyyy2 = request("yyyy2")
mm2 = request("mm2")
dd2 = request("dd2")

dim nowdate,date1,date2,Edate
nowdate = now

if (yyyy1="") then
	date1 = dateAdd("d",0,nowdate)
	yyyy1 = Left(CStr(date1),4)
	mm1   = Mid(CStr(date1),6,2)
	dd1   = Mid(CStr(date1),9,2)

	yyyy2 = Left(CStr(nowdate),4)
	mm2   = Mid(CStr(nowdate),6,2)
	dd2   = Mid(CStr(nowdate),9,2)

	Edate = Left(CStr(nowdate+1),10)
else
	Edate = Left(CStr(dateserial(yyyy2, mm2 , dd2)+1),10)
end if

%>
<%

dim baljunum, baljuid, baljudate
baljunum = request("baljunum")
baljuid = request("baljuid")
baljudate = request("baljudate")

dim baljuoff

set baljuoff = new COfflineBalju
baljuoff.FRectStartDate = yyyy1 + "-" + mm1 + "-" + dd1
baljuoff.FRectEndDate = Edate
baljuoff.GetOfflineBaljuList


dim i, isdayfinal,predate

dim SubTotalBaljucount, SubTotalUpchecount, SubTotalTenBaljucount, SubTotalOffBaljucount
dim SubTotalNoPackCount, SubTotalPackCount, SubTotalDeliverCount, SubTotalEtcCount, SubTotalConfirmCount

dim TotalBaljucount, TotalUpchecount, TotalTenBaljucount, TotalOffBaljucount
dim TotalNoPackCount, TotalPackCount, TotalDeliverCount, TotalEtcCount, TotalConfirmCount

dim SubPackingCount

%>

<script>

function PopOFflineBaljuPrint(baljudate, baljuid){
	var popwin = window.open('popofflinebaljuitemlist.asp?baljudate=' + baljudate + '&baljuid=' + baljuid,'popofflinebaljuitemlist' + baljuid,'width=800, height=600, resizabled=yes, scrollbars=yes');
	popwin.focus();
}

function PopOFflineBaljuPrint2(baljunum, baljuid){
	var popwin = window.open('popofflinebaljuitemlist.asp?baljunum=' + baljunum + '&baljuid=' + baljuid,'popofflinebaljuitemlist' + baljuid,'width=800, height=600, resizabled=yes, scrollbars=yes');
	popwin.focus();
}

</script>

<!-- 검색 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frm">
	<input type="hidden" name="menupos" value="<%= menupos %>">
    <input type="hidden" name="baljunum" value="<%= baljunum %>">
    <input type="hidden" name="baljuid" value="<%= baljuid %>">
    <input type="hidden" name="baljudate" value="<%= baljudate %>">
	<tr align="center" bgcolor="#FFFFFF" >
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
		<td align="left">
			<!--
        	<a href="/offline/balju/baljulist.asp">발주업데이트</a>
        	<% if baljunum<>"" then %>
        	&nbsp;&nbsp;
        	<img src="/images/icon_print02.gif" border="0" align="absbottom">&nbsp;<a href="javascript:PopOFflineBaljuPrint2('<%= baljunum %>','<%= baljuid %>');">발주서출력</a>
        	<% end if %>
        	&nbsp;&nbsp;
			-->
        	<% drawDateBox yyyy1,yyyy2,mm1,mm2,dd1,dd2 %>
            &nbsp;
		</td>
		
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button_s" value="검색" onClick="javascript:document.frm.submit();">
		</td>
	</tr>
	</form>
</table>
<!-- 검색 끝 -->

<p>

<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
	<tr>
		<td align="left">
			<input type="button" class="button" value="새로고침" onClick="javascript:document.location.reload();">
		</td>
		<td align="right">
			
		</td>
	</tr>
</table>
<!-- 액션 끝 -->

<p>


<table width="100%" align="center" cellpadding="3" cellspacing="0" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr>
		<td height="1" colspan="20" bgcolor="<%= adminColor("tablebg") %>"></td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
    	<td width=40>IDx</td>
    	<td width=140>발주일시</td>
    	<td width=100>샆아이디</td>
    	<td width=100>샆이름</td>
    	<td width=50 align="right">총발주</td>
    	<td width=50 align="right">업배</td>
    	<td width=50 align="right">텐배</td>
    	<td width=50 align="right">오프</td>
    	<td width=30></td>
    	<td width=50 align="right">상품<br>준비</td>
    	<td width=50 align="right">출고<br>준비</td>
    	<td width=50 align="right">패킹<br>완료</td>
    	<td width=50 align="right">출고<br>완료</td>
    	<td width=80>완료율</td>
    	<td align=right>출고<br>목록</td>
    </tr>
    <tr>
    	<td height="1" colspan="15" bgcolor="<%= adminColor("tablebg") %>"></td>
    </tr>
    <% for i=0 to baljuoff.FResultCount - 1 %>
    <%
    if (predate<>"") and (predate<>Left(baljuoff.FItemList(i).FBaljuDate,10)) then
            TotalBaljucount         = TotalBaljucount + SubTotalBaljucount
            TotalUpchecount         = TotalUpchecount + SubTotalUpchecount
            TotalTenBaljucount      = TotalTenBaljucount + SubTotalTenBaljucount
            TotalOffBaljucount      = TotalOffBaljucount + TotalOffBaljucount
            TotalNoPackCount        = TotalNoPackCount + SubTotalNoPackCount
            TotalPackCount          = TotalPackCount + SubTotalPackCount
            TotalDeliverCount       = TotalDeliverCount + SubTotalDeliverCount
            TotalEtcCount           = TotalEtcCount + SubTotalEtcCount
            TotalConfirmCount       = TotalConfirmCount + SubTotalConfirmCount
    %>
    <tr align=center bgcolor="#DDDDDD" >
    	<td ></td>
    	<td ></td>
    	<td ></td>
    	<td ></td>
    	<td align="right"><b><%= SubTotalBaljucount %></b></td>
    	<td align="right"><%= SubTotalUpchecount %></td>
    	<td align="right"><%= SubTotalTenBaljucount %></td>
    	<td align="right"><%= SubTotalOffBaljucount %></td>
    	<td ></td>
    	<td align="right"><font color=red><%= SubTotalNoPackCount %></font></td>
    	<td align="right"><%= SubTotalPackCount %></td>
    	<td align="right"><b><%= SubTotalDeliverCount %></b></td>
    	<td align="right"><%= SubTotalConfirmCount %></td>
    	<td >
    	<% if (SubTotalBaljucount <> 0) then %>
    		<% if ((SubTotalDeliverCount)=(SubTotalBaljucount)) then %>
    		<b><font color=red><%= CLng((SubTotalDeliverCount)/(SubTotalBaljucount)*100*100)/100 %>%</font></b>
    		<% else %>
    		<%= CLng((SubTotalDeliverCount)/(SubTotalBaljucount)*100*100)/100 %>%
    		<% end if %>
    	<% end if %>
            </td>
            <td ></td>
    </tr>
    <tr>
    	<td height="1" colspan="15" bgcolor="<%= adminColor("tablebg") %>"></td>
    </tr>
    <%
    	SubTotalBaljucount      = 0
    	SubTotalUpchecount      = 0
    	SubTotalTenBaljucount   = 0
    	SubTotalOffBaljucount   = 0
    	SubTotalNoPackCount     = 0
    	SubTotalPackCount       = 0
    	SubTotalDeliverCount    = 0
    	SubTotalEtcCount        = 0
    	SubTotalConfirmCount    = 0
    end if
    %>
    <%
    predate = Left(baljuoff.FItemList(i).FBaljudate,10)
    %>
    <%
    SubTotalBaljucount      = SubTotalBaljucount + baljuoff.FItemList(i).Ftotalbaljuno
    SubTotalUpchecount      = SubTotalUpchecount +  baljuoff.FItemList(i).Ftotalupcheno
    SubTotalTenBaljucount   = SubTotalTenBaljucount +  baljuoff.FItemList(i).Ftotaltenbaeno
    SubTotalOffBaljucount   = SubTotalOffBaljucount +  baljuoff.FItemList(i).Ftotalofflineno

    SubTotalNoPackCount     = SubTotalNoPackCount + baljuoff.FItemList(i).Ftotalnopackno
    SubTotalPackCount       = SubTotalPackCount + baljuoff.FItemList(i).Ftotalpackno
    SubTotalDeliverCount    = SubTotalDeliverCount + baljuoff.FItemList(i).Ftotaldeliverno
    SubTotalEtcCount        = SubTotalEtcCount + baljuoff.FItemList(i).Ftotaletcno
    SubTotalConfirmCount    = SubTotalConfirmCount + baljuoff.FItemList(i).Ftotalconfirmno
    %>
    <tr align="center" bgcolor="#FFFFFF">
    	<% if ((CStr(baljuoff.FItemList(i).FBaljuNum)=CStr(baljunum)) and (CStr(baljuoff.FItemList(i).FBaljuId)=CStr(baljuid))) then %>
    	<td><b><font color="#3333AA"><%= baljuoff.FItemList(i).FBaljuNum %></font></b></td>
    	<% else %>
    	<td><%= baljuoff.FItemList(i).FBaljuNum %></td>
    	<% end if %>
    	
		<!--
		<td><a href="?baljunum=<%= baljuoff.FItemList(i).FBaljuNum %>&baljuid=<%= baljuoff.FItemList(i).FBaljuId %>"><%= baljuoff.FItemList(i).FBaljuDate %></a></td>
    	-->
    	
    	<td align="left"><a href="?baljunum=<%= baljuoff.FItemList(i).FBaljuNum %>&baljuid=<%= baljuoff.FItemList(i).FBaljuId %>"><%= baljuoff.FItemList(i).FBaljuDate %></a></td>
    	<td><%= baljuoff.FItemList(i).FBaljuId %></td>
    	<td align="left"><%= baljuoff.FItemList(i).FBaljuName %></td>
    	<td align="right"><b><%= baljuoff.FItemList(i).Ftotalbaljuno %></b></td>
    	<td align="right"><%= baljuoff.FItemList(i).Ftotalupcheno %></td>
    	<td align="right"><%= baljuoff.FItemList(i).Ftotaltenbaeno %></td>
    	<td align="right"><%= baljuoff.FItemList(i).Ftotalofflineno %></td>
    	<td></td>
    	<td align="right">
    	  <% if (baljuoff.FItemList(i).Ftotalnopackno > 0) then %>
    	  <font color=red><%= baljuoff.FItemList(i).Ftotalnopackno %></font>
    	  <% else %>
    	  <%= baljuoff.FItemList(i).Ftotalnopackno %>
    	  <% end if %>

    	</td>
    	<td align="right"><%= baljuoff.FItemList(i).Ftotalpackno %></td>
    	<td align="right"><b><%= baljuoff.FItemList(i).Ftotaldeliverno %></b></td>
    	<td align="right"><%= baljuoff.FItemList(i).Ftotalconfirmno %></td>
    	<td></td>
    	<td align=right>
    	  <!--
    	  <a href="shopchulgo.asp?baljunum=<%= baljuoff.FItemList(i).FBaljuNum %>&baljuid=<%= baljuoff.FItemList(i).FBaljuId %>">-&gt;</a>
    	  -->
    	  <a href="http://logics.10x10.co.kr/offline/balju/shopchulgo.asp?baljudate=<%= Left(baljuoff.FItemList(i).FBaljuDate,10) %>&baljuid=<%= baljuoff.FItemList(i).FBaljuId %>" target="_blank">-&gt;</a>
    	</td>
    </tr>
    <tr>
    	<td height="1" colspan="15" bgcolor="<%= adminColor("tablebg") %>"></td>
    </tr>
    <% next %>
    <tr align=center bgcolor="#DDDDDD">
    	<td></td>
    	<td></td>
    	<td></td>
    	<td></td>
    	<td align="right"><b><%= SubTotalBaljucount %></b></td>
    	<td align="right"><%= SubTotalUpchecount %></td>
    	<td align="right"><%= SubTotalTenBaljucount %></td>
    	<td align="right"><%= SubTotalOffBaljucount %></td>
    	<td></td>
    	<td align="right"><%= SubTotalNoPackCount %></td>
    	<td align="right"><%= SubTotalPackCount %></td>
    	<td align="right"><b><%= SubTotalDeliverCount %></b></td>
    	<td align="right"><%= SubTotalConfirmCount %></td>
    	<td>
    	<% if (SubTotalBaljucount <> 0) then %>
    		<% if (SubTotalBaljucount = SubTotalDeliverCount) then %>
    		<b><font color=red><%= CLng((SubTotalDeliverCount)/(SubTotalBaljucount)*100*100)/100 %>%</font></b>
    		<% else %>
    		<%= CLng((SubTotalDeliverCount)/(SubTotalBaljucount)*100*100)/100 %>%
    		<% end if %>
    	<% end if %>
    	</td>
    	<td></td>
    </tr>
    <tr>
    	<td height="1" colspan="15" bgcolor="<%= adminColor("tablebg") %>"></td>
    </tr>
    <%
            TotalBaljucount         = TotalBaljucount + SubTotalBaljucount
            TotalUpchecount         = TotalUpchecount + SubTotalUpchecount
            TotalTenBaljucount      = TotalTenBaljucount + SubTotalTenBaljucount
            TotalOffBaljucount      = TotalOffBaljucount + SubTotalOffBaljucount
            TotalNoPackCount        = TotalNoPackCount + SubTotalNoPackCount
            TotalPackCount          = TotalPackCount + SubTotalPackCount
            TotalDeliverCount       = TotalDeliverCount + SubTotalDeliverCount
            TotalEtcCount           = TotalEtcCount + SubTotalEtcCount
            TotalConfirmCount       = TotalConfirmCount + SubTotalConfirmCount
    %>
    <tr align=center  bgcolor="#EEEE22" >
    	<td >Total</td>
    	<td ></td>
    	<td ></td>
    	<td ></td>
    	<td align="right"><b><%= TotalBaljucount %></b></td>
    	<td align="right"><%= TotalUpchecount %></td>
    	<td align="right"><%= TotalTenBaljucount %></td>
    	<td align="right"><%= TotalOffBaljucount %></td>
    	<td ></td>
    	<td align="right"><%= TotalNoPackCount %></td>
    	<td align="right"><%= TotalPackCount %></td>
    	<td align="right"><b><%= TotalDeliverCount %></b></td>
    	<td align="right"><%= TotalConfirmCount %></td>

    	<td >
    	<% if (TotalBaljucount > 0) then %>
    	<% if (TotalBaljucount=TotalDeliverCount) then %>
    		<font color=red><b><%= CLng((TotalDeliverCount)/(TotalBaljucount)*100*100)/100 %>%</b></font>
    	<% else %>
    	        <%= CLng((TotalDeliverCount)/(TotalBaljucount)*100*100)/100 %>%
    	<% end if %>
    	<% end if %>
    	</td>
    	<td ></td>
    </tr>
    <tr>
    	<td height="1" colspan="15" bgcolor="<%= adminColor("tablebg") %>"></td>
    </tr>
</table>



<%

set baljuoff = Nothing

%>
					
					
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dblogicsclose.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->
					