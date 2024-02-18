<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/order/baljucls_off.asp"-->
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
	date1 = dateAdd("d",-2,nowdate)
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

<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="#F3F3FF">
	<tr>
		<td width="100%" valign=top>
			<table width="100%" height="50" border="0" align="left" cellpadding="0" cellspacing="0" class="a" bgcolor="#F3F3FF">
			  <tr valign="bottom" height="10">
			    <td width="10" align="right" valign="bottom"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
			    <td valign="bottom" background="/images/tbl_blue_round_02.gif"></td>
			    <td width="10" align="left" valign="bottom"><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
			  </tr>
			  <tr valign="top" height="25">
			    <td background="/images/tbl_blue_round_04.gif"></td>
			    <td background="/images/tbl_blue_round_06.gif">
			    	<img src="/images/icon_star.gif" align="absbottom">&nbsp;<font color="red"><strong>발주리스트</strong></font>
			    	&nbsp;&nbsp;&nbsp;&nbsp;
			    	<!--<a href="/offline/balju/baljulist.asp"><img src="/images/icon_reload.gif" align="absbottom" border="0">&nbsp;새로고침</a>-->
			    	<% if baljudate<>"" then %>
			    	&nbsp;&nbsp;&nbsp;&nbsp;
			    	<!--<img src="/images/icon_print02.gif" border="0" align="absbottom">&nbsp;<a href="javascript:PopOFflineBaljuPrint('<%= baljudate %>','<%= baljuid %>');">출고대기전환</a>-->
			    	<% end if %>
			    </td>
			    <td background="/images/tbl_blue_round_05.gif"></td>
			  </tr>
			  <tr valign="top">
			    <td background="/images/tbl_blue_round_04.gif"></td>
			    <td bgcolor="#F3F3FF">
			      <br>
			      <table width="100%" border="0" cellpadding="0" cellspacing="0" class="a">
			        <form name="frm">
			        <input type="hidden" name="baljunum" value="<%= baljunum %>">
			        <input type="hidden" name="baljuid" value="<%= baljuid %>">
			        <input type="hidden" name="baljudate" value="<%= baljudate %>">
			        <tr>
			          <td width="330"><% drawDateBox yyyy1,yyyy2,mm1,mm2,dd1,dd2 %></td>
			          <td align="left"><a href="javascript:document.frm.submit()"><img src="/images/search2.gif" width="74" height="22" border="0" valign="bottom"></a></td>
			        </tr>
			        </form>
			      </table>
			    </td>
			    <td background="/images/tbl_blue_round_05.gif"></td>
			  </tr>
			  <tr height="10" valign="top">
			    <td><img src="/images/tbl_blue_round_07.gif" width="10" height="10"></td>
			    <td background="/images/tbl_blue_round_08.gif"></td>
			    <td><img src="/images/tbl_blue_round_09.gif" width="10" height="10"></td>
			  </tr>
			</table>
		</td>

	</tr>
</table>

<p>

<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="#F3F3FF">
	<tr>
		<td width="100%" valign=top>
			<table width="100%" height="50" border="0" align="left" cellpadding="0" cellspacing="0" class="a">
			  <tr valign="bottom" height="10">
			    <td width="10" align="right" valign="bottom"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
			    <td valign="bottom" background="/images/tbl_blue_round_02.gif"></td>
			    <td width="10" align="left" valign="bottom"><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
			  </tr>

			  <tr valign="top">
			    <td background="/images/tbl_blue_round_04.gif"></td>
			    <td bgcolor="#F3F3FF">
			                <br>

                                        <table width="100%" border=0 cellspacing=0 cellpadding=2 class=a bgcolor="FFFFFF">
                                        <tr align="center"  bgcolor="F3F3FF">
                                        	<td width=40>IDx</td>
                                        	<td width=80 align="left">발주일</td>
                                        	<td width=100 align="left">샆아이디</td>
                                        	<td width=100 align="left">샆이름</td>
                                        	<td width=50>상태</td>
                                        	<td width=50 align=right>총발주</td>
                                        	<td width=50 align=right>업배</td>
                                        	<td width=50 align=right>텐배</td>
                                        	<td width=50 align=right>오프</td>
                                        	<td width=15></td>
                                        	<td width=45 align=right>상품<br>준비</td>
                                        	<td width=60 align=right>출고준비<br>(Box in)</td>
                                        	<td width=45 align=right>패킹<br>완료</td>
                                        	<td width=45 align=right>출고<br>완료</td>
                                        	<td width=50>완료율</td>
                                        	<td align=right>출고<br>전환</td>
                                        </tr>
                                        <tr>
                                        	<td height="1" colspan="16" bgcolor="#CCCCCC"></td>
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
                                        	<td ></td>
                                        	<td align=right><%= SubTotalBaljucount %></td>
                                        	<td align=right><%= SubTotalUpchecount %></td>
                                        	<td align=right><%= SubTotalTenBaljucount %></td>
                                        	<td align=right><%= SubTotalOffBaljucount %></td>
                                        	<td ></td>
                                        	<td align=right><%= SubTotalNoPackCount %></td>
                                        	<td align=right><%= SubTotalPackCount %></td>
                                        	<td align=right><b><%= SubTotalDeliverCount %></b></td>
                                        	<td align=right><%= SubTotalConfirmCount %></td>
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
                                        	<td height="1" colspan="16" bgcolor="#555555"></td>
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
                                        <tr align="center">
                                        	<% if (((CStr(baljuoff.FItemList(i).FBaljuNum)=CStr(baljunum)) or (CStr(Left(baljuoff.FItemList(i).FBaljuDate,10))=CStr(baljudate))) and (CStr(baljuoff.FItemList(i).FBaljuId)=CStr(baljuid))) then %>
                                        	<td><a href="javascript:PopOFflineBaljuPrint2('<%= baljuoff.FItemList(i).FBaljuNum %>','<%= baljuid %>');"><b><font color="#3333AA"><%= baljuoff.FItemList(i).FBaljuNum %></font></b></a></td>
                                        	<% else %>
                                        	<td><a href="javascript:PopOFflineBaljuPrint2('<%= baljuoff.FItemList(i).FBaljuNum %>','<%= baljuid %>');"><%= baljuoff.FItemList(i).FBaljuNum %></a></td>
                                        	<% end if %>
						<!--
						<td><a href="?baljunum=<%= baljuoff.FItemList(i).FBaljuNum %>&baljuid=<%= baljuoff.FItemList(i).FBaljuId %>"><%= baljuoff.FItemList(i).FBaljuDate %></a></td>
                                        	-->
                                        	<td align="left"><a href="?baljudate=<%= Left(baljuoff.FItemList(i).FBaljuDate,10) %>&baljuid=<%= baljuoff.FItemList(i).FBaljuId %>"><%= Left(baljuoff.FItemList(i).FBaljuDate,10) %></a></td>
                                        	<td align="left"><%= baljuoff.FItemList(i).FBaljuId %></td>
                                        	<td align="left"><%= baljuoff.FItemList(i).FBaljuName %></td>
                                        	<td ><%= baljuoff.FItemList(i).getStateNameHTML %></td>
                                        	<td align=right><%= baljuoff.FItemList(i).Ftotalbaljuno %></td>
                                        	<td align=right><%= baljuoff.FItemList(i).Ftotalupcheno %></td>
                                        	<td align=right><%= baljuoff.FItemList(i).Ftotaltenbaeno %></td>
                                        	<td align=right><%= baljuoff.FItemList(i).Ftotalofflineno %></td>
                                        	<td></td>
                                        	<td align=right><%= baljuoff.FItemList(i).Ftotalnopackno %></td>
                                        	<td align=right><%= baljuoff.FItemList(i).Ftotalpackno %></td>
                                        	<td align=right><b><%= baljuoff.FItemList(i).Ftotaldeliverno %></b></td>
                                        	<td align=right><%= baljuoff.FItemList(i).Ftotalconfirmno %></td>
                                        	<td></td>
                                        	<td align=right>
                                        	  <a href="baljufinishoffline.asp?baljunum=<%= baljuoff.FItemList(i).FBaljuNum %>&baljuid=<%= baljuoff.FItemList(i).FBaljuId %>">-&gt;</a>
                                        	</td>
                                        </tr>
                                        <tr>
                                        	<td height="1" colspan="16" bgcolor="#CCCCCC"></td>
                                        </tr>
                                        <% next %>
                                        <tr align=center  bgcolor="#DDDDDD" >
                                        	<td ></td>
                                        	<td ></td>
                                        	<td ></td>
                                        	<td ></td>
                                        	<td ></td>
                                        	<td align=right><%= SubTotalBaljucount %></td>
                                        	<td align=right><%= SubTotalUpchecount %></td>
                                        	<td align=right><%= SubTotalTenBaljucount %></td>
                                        	<td align=right><%= SubTotalOffBaljucount %></td>
                                        	<td ></td>
                                        	<td align=right><%= SubTotalNoPackCount %></td>
                                        	<td align=right><%= SubTotalPackCount %></td>
                                        	<td align=right><b><%= SubTotalDeliverCount %></b></td>
                                        	<td align=right><%= SubTotalConfirmCount %></td>
                                        	<td >
                                        	<% if (SubTotalBaljucount <> 0) then %>
                                        		<% if (SubTotalBaljucount = SubTotalDeliverCount) then %>
                                        		<b><font color=red><%= CLng((SubTotalDeliverCount)/(SubTotalBaljucount)*100*100)/100 %>%</font></b>
                                        		<% else %>
                                        		<%= CLng((SubTotalDeliverCount)/(SubTotalBaljucount)*100*100)/100 %>%
                                        		<% end if %>
                                        	<% end if %>
                                        	</td>
                                        	<td ></td>
                                        </tr>
                                        <tr>
                                        	<td height="1" colspan="16" bgcolor="#555555"></td>
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
                                        	<td ></td>
                                        	<td align=right><%= TotalBaljucount %></td>
                                        	<td align=right><%= TotalUpchecount %></td>
                                        	<td align=right><%= TotalTenBaljucount %></td>
                                        	<td align=right><%= TotalOffBaljucount %></td>
                                        	<td ></td>
                                        	<td align=right><%= TotalNoPackCount %></td>
                                        	<td align=right><%= TotalPackCount %></td>
                                        	<td align=right><b><%= TotalDeliverCount %></b></td>
                                        	<td align=right><%= TotalConfirmCount %></td>

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
                                        	<td height="1" colspan="16" bgcolor="#555555"></td>
                                        </tr>
                                        </table>

				</td>
			    <td background="/images/tbl_blue_round_05.gif"></td>
			  </tr>
			  <tr height="10" valign="top">
			    <td><img src="/images/tbl_blue_round_07.gif" width="10" height="10"></td>
			    <td background="/images/tbl_blue_round_08.gif"></td>
			    <td><img src="/images/tbl_blue_round_09.gif" width="10" height="10"></td>
			  </tr>
			</table>
		</td>

	</tr>
</table>

<%

set baljuoff = Nothing

%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->