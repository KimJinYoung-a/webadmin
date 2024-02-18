<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/stock/ordersheetcls.asp"-->
<%
''개별 매입, 위탁 (90Code or 업체배송 등)

dim page, statecd, designer
dim baljucode, notipgo, minusjumun
dim divgubun

page        = RequestCheckVar(request("page"),9)
statecd     = RequestCheckVar(request("statecd"),9)
designer    = RequestCheckVar(request("designer"),32)
baljucode   = RequestCheckVar(request("baljucode"),9)
notipgo     = RequestCheckVar(request("notipgo"),9)
minusjumun  = RequestCheckVar(request("minusjumun"),9)
divgubun    = RequestCheckVar(request("divgubun"),9)

if page="" then page=1

dim osheet
set osheet = new COrderSheet
osheet.FCurrPage = page
osheet.Fpagesize = 20
osheet.FRectBaljuCode = baljucode
if (baljucode="") then
	osheet.FRectStatecd = statecd
	osheet.FRectTargetid = designer
	osheet.FRectNotIpgoOnly = notipgo
	osheet.FRectMinusOnly = minusjumun
	osheet.FRectDivGubun = divgubun
	
end if
osheet.FRectDivCodeUnder = "300"

osheet.GetOrderSheetList


dim i
dim totaljumunsuply, totalfixsuply, totaljumunsellcash



%>
<script language='javascript'>
function PopIpgoSheet(v){
	var popwin;
	popwin = window.open('popshopjumunsheet2.asp?idx=' + v ,'shopjumunsheet','width=740,height=600,scrollbars=yes,resizabled=yes');
	popwin.focus();
}

function ExcelSheet(v,itype){
	var popwin = window.open('popshopjumunsheet2.asp?idx=' + v + '&xl=on');
	popwin.focus();
}

function sendSMSEmail(idesigner,iidx){
	var popwin = window.open("/admin/lib/popupchejumunsms.asp?designer=" + idesigner + "&idx=" + iidx,"popupchejumunsms","width=600 height=500,scrollbars=yes,resizabled=yes");
	popwin.focus();
}
</script>

<!-- 검색 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frm" method="get" action="">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<input type="hidden" name="research" value="on">
	<tr align="center" bgcolor="#FFFFFF" >
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
		<td align="left">
			발주코드 :
		<input type="text" class="text" name="baljucode" value="<%= baljucode %>" size="8" maxlength="8" >
		브랜드 :
		<% drawSelectBoxDesignerwithName "designer", designer %>
		&nbsp;주문상태 :
		<select name="statecd" class="select">
			<option value="">전체
			<option value="0" <% if statecd="0" then response.write "selected" %> >주문접수
			<option value="1" <% if statecd="1" then response.write "selected" %> >주문확인
			<option value="2" <% if statecd="2" then response.write "selected" %> >입금대기
			<option value="5" <% if statecd="5" then response.write "selected" %> >배송준비
			<option value="7" <% if statecd="7" then response.write "selected" %> >출고완료
			<option value="8" <% if statecd="8" then response.write "selected" %> >입고대기
			<option value="9" <% if statecd="9" then response.write "selected" %> >입고완료
		</select>
		</td>
		
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button_s" value="검색" onClick="javascript:document.frm.submit();">
		</td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF" >
		<td align="left">
			<input type="checkbox" name="notipgo" <% if notipgo="on" then response.write "checked" %> >입고미처리만
	     	&nbsp;
	     	<input type="checkbox" name="minusjumun" <% if minusjumun="on" then response.write "checked" %> >마이너스주문만
	     	&nbsp;
	     	매입구분 : 
			<input type="radio" name="divgubun" value="" <% if divgubun="" then response.write "checked" %> >전체
			<!--
			<input type="radio" name="divgubun" value="j" <% if divgubun="j" then response.write "checked" %> >개별입고
			-->
			<input type="radio" name="divgubun" value="101" <% if divgubun="101" then response.write "checked" %> >개별매입
			<input type="radio" name="divgubun" value="111" <% if divgubun="111" then response.write "checked" %> >개별위탁
		</td>
	</tr>
	</form>
</table>
<!-- 검색 끝 -->

<p>



<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="15">
			검색결과 : <b><%= osheet.FTotalCount %></b>
			&nbsp;
			페이지 : <b><%= page %> / <%= osheet.FTotalpage %></b>
		</td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td width="60">주문코드</td>
		<td>공급처</td>
		<td width="90">주문처</td>
		<td width="120">구분</td>

		<td width="70">주문상태</td>
		<td width=80>주문/<br>입고요청일</td>
		<td width=60>총주문액<br>확정액(소)</td>
		<td width=60>총주문액<br>확정액(매)</td>
		<td width=50>마진</td>
		<td width=70>업체발송일</td>
		<td width="90">택배사<br>운송장번호</td>
		<td width=50>내역서</td>
	</tr>
	<% if osheet.FResultCount >0 then %>
	<% for i=0 to osheet.FResultcount-1 %>
	<%
	totaljumunsellcash = totaljumunsellcash + osheet.FItemList(i).Fjumunsellcash
'	if (osheet.FItemList(i).Ftargetid="10x10") then
'		totaljumunsuply = totaljumunsuply + osheet.FItemList(i).Fjumunsuplycash
'		totalfixsuply   = totalfixsuply + osheet.FItemList(i).Ftotalsuplycash
'	else
		totaljumunsuply = totaljumunsuply + osheet.FItemList(i).Fjumunbuycash
		totalfixsuply   = totalfixsuply + osheet.FItemList(i).Ftotalbuycash
'	end if
	%>
	<tr bgcolor="#FFFFFF">
		<td rowspan=2 align=center><a href="upchejumuninputedit.asp?idx=<%= osheet.FItemList(i).Fidx %>&opage=<%= page %>&ourl=upchejumunlist.asp"><%= osheet.FItemList(i).Fbaljucode %></a></td>
		<td rowspan=2 align=center><b><a href="javascript:PopUpcheBrandInfoEdit('<%= osheet.FItemList(i).Ftargetid %>');"><%= osheet.FItemList(i).Ftargetid %></a></b><br>(<%= osheet.FItemList(i).Ftargetname %>)</td>
		<td rowspan=2 align=center><%= osheet.FItemList(i).Fbaljuid %><br>(<%= osheet.FItemList(i).Fbaljuname %>)</td>
		<td rowspan=2 align=center><font color="<%= osheet.FItemList(i).GetDivCodeColor %>"><%= osheet.FItemList(i).GetDivCodeName %></font></td>
		<td rowspan=2 align=center>
			<font color="<%= osheet.FItemList(i).GetStateColor %>"><%= osheet.FItemList(i).GetStateName %></font>
			<br><%= osheet.FItemList(i).FAlinkCode %>
		</td>
		<td align=center><font color="#777777"><%= Left(osheet.FItemList(i).FRegdate,10) %></font></td>
		<td align=right><%= FormatNumber(osheet.FItemList(i).Fjumunsellcash,0) %></td>
		<td align="right">
		    <!-- <%= FormatNumber(osheet.FItemList(i).Fjumunsuplycash,0) %> -->
			<%= FormatNumber(osheet.FItemList(i).Fjumunbuycash,0) %>
		</td>
		<td rowspan="2" align="center">
		<% if osheet.FItemList(i).Ftotalsellcash<>0 then %>
		    <%= CLng((osheet.FItemList(i).Ftotalsellcash-osheet.FItemList(i).Ftotalbuycash)/osheet.FItemList(i).Ftotalsellcash*100*100)/100 %> %
		<% end if %>
		
		
		</td>
		<td rowspan="2" align="center"><%= Left(osheet.FItemList(i).Fbeasongdate,10) %></td>
		<td rowspan="2" align=center>
			<% if (Not osheet.FItemList(i).IsSendedSMS) and (osheet.FItemList(i).getScheduleIpgodate="") and (osheet.FItemList(i).Fstatecd="0") then %>
			<input type="button" class="button" value="SMS" onclick="sendSMSEmail('<%= osheet.FItemList(i).Ftargetid %>','<%= osheet.FItemList(i).Fidx %>')">
			<% else %>
			<a href="<%= DeliverDivTrace(Trim(osheet.FItemList(i).Fsongjangdiv)) %><%= osheet.FItemList(i).Fsongjangno %>" target=_blank>
			<%= DeliverDivCd2Nm(Trim(osheet.FItemList(i).Fsongjangdiv)) %><br><%= osheet.FItemList(i).Fsongjangno %>
			<% end if %>
		</td>
		<td rowspan="2" width=50 align=center>
			<a href="javascript:PopIpgoSheet('<%= osheet.FItemList(i).FIdx %>');"><img src="/images/iexplorer.gif" width=21 border=0></a>
			<a href="javascript:ExcelSheet('<%= osheet.FItemList(i).FIdx %>');"><img src="/images/iexcel.gif" width=21 border=0></a>
		</td>
	</tr>
	<tr bgcolor="#FFFFFF">
	    <td align="center"><%= Left(osheet.FItemList(i).Fscheduledate,10) %></td>
		<td align="right"><%= FormatNumber(osheet.FItemList(i).Ftotalsellcash,0) %></td>
		<td align="right">
			<%= FormatNumber(osheet.FItemList(i).Ftotalbuycash,0) %>
		</td>

<!-- <font color="#777777"><%= DdotFormat(osheet.FItemList(i).Fbrandlist,20) %></font> -->
	</tr>
	<% next %>
	<tr bgcolor="#FFFFFF">
		<td align="center">총계</td>
		<td colspan="4"></td>
		<td align="right"><%= formatNumber(totaljumunsellcash,0) %></td>
		<td align="right"><%= formatNumber(totaljumunsuply,0) %></td>
		<td align="right"><%= formatNumber(totalfixsuply,0) %></td>
		<td align="right"></td>
		<td colspan="3"></td>
	</tr>
	<tr bgcolor="#FFFFFF" height=20>
		<td colspan="13" align=center>
		<% if osheet.HasPreScroll then %>
			<a href="?page=<%= osheet.StartScrollPage-1 %>&statecd=<%= statecd %>&designer=<%= designer %>&notipgo=<%= notipgo %>&minusjumun=<%= minusjumun %>&divgubun=<%= divgubun %>">[pre]</a>
		<% else %>
			[pre]
		<% end if %>
	
		<% for i=0 + osheet.StartScrollPage to osheet.FScrollCount + osheet.StartScrollPage - 1 %>
			<% if i>osheet.FTotalpage then Exit for %>
			<% if CStr(page)=CStr(i) then %>
			<font color="red">[<%= i %>]</font>
			<% else %>
			<a href="?page=<%= i %>&statecd=<%= statecd %>&designer=<%= designer %>&notipgo=<%= notipgo %>&minusjumun=<%= minusjumun %>&divgubun=<%= divgubun %>">[<%= i %>]</a>
			<% end if %>
		<% next %>
	
		<% if osheet.HasNextScroll then %>
			<a href="?page=<%= i %>&statecd=<%= statecd %>&designer=<%= designer %>&notipgo=<%= notipgo %>&minusjumun=<%= minusjumun %>&divgubun=<%= divgubun %>">[next]</a>
		<% else %>
			[next]
		<% end if %>
		</td>
	</tr>
	<% else %>
	<tr bgcolor="#FFFFFF">
		<td colspan=11 align=center>[ 검색결과가 없습니다. ]</td>
	</tr>
	<% end if %>
</table>
<%
set osheet = Nothing
%>


<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->

