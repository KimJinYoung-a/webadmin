<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/classes/payment/ipkumlistcls.asp"-->
<%

dim yyyy1,mm1,dd1
dim yyyy2,mm2,dd2
dim ipkumstate,tenbank,ipkumname,page, ipkumgubun, ipkumidx
dim research

ipkumstate=request("ipkumstate")
ipkumgubun=request("ipkumgubun")
tenbank=request("tenbank")
ipkumname=request("ipkumname")
page=request("page")
research = request("research")
ipkumidx=request("ipkumidx")

if page="" then page=1


yyyy1 = request("yyyy1")
mm1 = request("mm1")
dd1 = request("dd1")
yyyy2 = request("yyyy2")
mm2 = request("mm2")
dd2 = request("dd2")

if (yyyy1="") then
	yyyy1 = Cstr(Year(now()))
	mm1 = Cstr(Month(now())-1)
	dd1 = Cstr(day(now()))
end if

if (yyyy2="") then
	yyyy2 = Cstr(Year(now()))
	mm2 = Cstr(Month(now()))
	dd2 = Cstr(day(now()))
end if

dim ipkum,i,ix
set ipkum = new IpkumChecklist

ipkum.FCurrpage=page
ipkum.FPagesize=200
ipkum.FScrollCount = 10
ipkum.ipkumstate=ipkumstate
ipkum.Ctenbank=tenbank
ipkum.ipkumname=ipkumname
ipkum.FRectIpkumGubun=ipkumgubun
ipkum.FRectIpkumIdx=ipkumidx

ipkum.yyyy1=yyyy1
ipkum.mm1=mm1
ipkum.dd1=dd1
ipkum.yyyy2=yyyy2
ipkum.mm2=mm2
ipkum.dd2=dd2

ipkum.Getipkumlist

%>
<script language='javascript'>

function scrollmove(v) {
 	document.frmipkum.page.value=v;
 	document.frmipkum.action='ipkumlist.asp';
 	document.frmipkum.submit();

}

function jsMatchOrderlist(ipkumgubun, idx) {
	var v = "pop_matchorderlist.asp?ipkumgubun=" + ipkumgubun + "&idx=" + idx;
	var popwin = window.open(v,"jsMatchOrderlist","width=1200,height=600,scrollbars=yes,resizable=yes");
	popwin.focus();
}

function cashreceiptInfo(iorderserial){
	var receiptUrl = "/cscenter/taxsheet/popCashReceipt.asp?orderserial=" + iorderserial;
	var popwin = window.open(receiptUrl,"Cashreceipt","width=500,height=400,scrollbars=yes,resizable=yes");
	popwin.focus();
}
</script>

<!-- 검색 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frmipkum" method="get" action="">
	<input type="hidden" name="showtype" value="showtype">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<input type="hidden" name="page" value="">
	<tr align="center" bgcolor="#FFFFFF" >
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
		<td align="left">
			입금구분 :
            <select class="select" name="ipkumgubun">
	            <option value="">전체
	            <option value="10x10" <% if ipkumgubun="10x10" then response.write " selected" %>>온라인
	            <option value="fingers" <% if ipkumgubun="fingers" then response.write " selected" %>>핑거스
            </select>
			&nbsp;
			확인구분 :
            <select class="select" name="ipkumstate">
	            <option value="">전체
	            <option value="1" <% if ipkumstate="1" then response.write " selected" %>>매칭실패
	            <option value="0" <% if ipkumstate="0" then response.write " selected" %>>미처리
            </select>
            &nbsp;
    		은행 :
    		<select class="select" name="tenbank">
	    		<option value="">전체
	    		<option value="조흥" <% if tenbank="조흥" then response.write " selected" %>>조흥
	    		<option value="국민" <% if tenbank="국민" then response.write " selected" %>>국민
	    		<option value="우리" <% if tenbank="우리" then response.write " selected" %>>우리
	    		<option value="하나" <% if tenbank="하나" then response.write " selected" %>>하나
	    		<option value="농협" <% if tenbank="농협" then response.write " selected" %>>농협
	    		<option value="신한" <% if tenbank="신한" then response.write " selected" %>>신한
	    		<option value="기업" <% if tenbank="기업" then response.write " selected" %>>기업
    		</select>
    		&nbsp;
            검색기간 :
    		<% DrawDateBox yyyy1,yyyy2,mm1,mm2,dd1,dd2 %>
            &nbsp;
    		입금자명 :
    		<input type="text" class="text" name=ipkumname value="<%= ipkumname %>" size=10 >
            &nbsp;
    		입금IDX :
    		<input type="text" class="text" name=ipkumidx value="<%= ipkumidx %>" size=10 >
		</td>

		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button_s" value="검색" onClick="javascript:document.frmipkum.submit();">
		</td>
	</tr>
	</form>
</table>

<p>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="16">
			검색결과 : <b><%= ipkum.FTotalcount %></b>
			&nbsp;
			페이지 : <b><%= page %> / <%= ipkum.FTotalpage %></b>
		</td>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
    	<td width="50">Idx</td>
    	<td width="50">은행</td>
    	<td width="65">날짜</td>
    	<td width="70">구분</td>
    	<td>입금자</td>
    	<td width="50">출금액</td>
    	<td width="50">입금액</td>
<!--   	<td width="50">잔액</td>	-->
    	<td>적요</td>
    	<td width="70">확인유무</td>
    	<td width="70">입금구분</td>
		<td>관련주문번호</td>
		<td>입금사유</td>
    	<td width="100">처리자</td>
    	<td width="50">방식</td>
    	<td width="50">매칭</td>
    	<td>증빙</td>
    </tr>
    <% if ipkum.FResultCount<1 then %>
    <% else %>
    <% for i=0 to ipkum.FResultCount-1 %>
    <form name="MatchOrderlist" method="post">
	<input type="hidden" name="idx" value="<%= ipkum.Fipkumitem(i).Fidx %>">
	<input type="hidden" name="searchtype01" value="on">
	<input type="hidden" name="searchtype02" value="on">
	<input type="hidden" name="searchname" value="<%= ipkum.Fipkumitem(i).Fipkumuser %>">
	<input type="hidden" name="searchprice" value="<%= ipkum.Fipkumitem(i).Fipkumsum %>">
    <tr align="center" bgcolor="#FFFFFF">
    	<td><%= ipkum.Fipkumitem(i).Fidx %></td>
    	<td><%= ipkum.Fipkumitem(i).Ftenbank %></td>
    	<td><%= left(ipkum.Fipkumitem(i).FBankdate,10) %></td>
    	<td><%= ipkum.Fipkumitem(i).Fgubun %></td>
    	<td><%= ipkum.Fipkumitem(i).Fipkumuser %></td>
    	<td align="right"><%= FormatNumber(ipkum.Fipkumitem(i).Fchulkumsum,0) %></td>
    	<td align="right"><%= FormatNumber(ipkum.Fipkumitem(i).Fipkumsum,0) %></td>
<!--   	<td align="right"><%= FormatNumber(ipkum.Fipkumitem(i).Fremainsum,0) %></td>	-->
    	<td><%= ipkum.Fipkumitem(i).Fbankname %></td>
    	<td>
    		<% if ipkum.Fipkumitem(i).Fipkumstate=0 then %>미처리
    		<% elseif ipkum.Fipkumitem(i).Fipkumstate=1 then %><font color="red">매칭실패</font>
    		<% elseif ipkum.Fipkumitem(i).Fipkumstate=7 then %><font color="blue">매칭완료</font>
    		<% end if %>
    	</td>
		<td><%= ipkum.Fipkumitem(i).Fipkumgubun %></td>
    	<td>
    		<% if ipkum.Fipkumitem(i).Fipkumstate=7 then %>
	    		<% if ipkum.Fipkumitem(i).Forderserial<>"" then %><%= ipkum.Fipkumitem(i).Forderserial %>
	    		<% else %>
	    			<% if ipkum.Fipkumitem(i).Ffinishstr<>"" then %><font color="red"><%= db2html(ipkum.Fipkumitem(i).Ffinishstr) %></font>
	    			<% else %><font color="red">미입력</font>
	    			<% end if %>
	    		<% end if %>
	    	<% end if %>
    	</td>
		<td><%= ipkum.Fipkumitem(i).FipkumCause %></td>
    	<td><%= ipkum.Fipkumitem(i).Ffinishuser %></td>
    	<td>
    		<% if ipkum.Fipkumitem(i).Fipkumstate=7 then %>
	    		<% if ipkum.Fipkumitem(i).Forderserial<>"" then %>자동
	    		<% else %><font color="red">수동</font>
	    		<% end if %>
	    	<% end if %>
    	</td>
    	<td>
    		<% if ipkum.Fipkumitem(i).Fipkumstate=1 or (left(ipkum.Fipkumitem(i).FBankdate,10) = Left(Now(), 10)) or (DateDiff("d", left(ipkum.Fipkumitem(i).FBankdate,10), Left(Now(), 10)) <= 30) then %>
    			<input type="button" class="button" value="매칭" onclick="javascript:jsMatchOrderlist('<%= ipkum.Fipkumitem(i).Fipkumgubun %>', <%= ipkum.Fipkumitem(i).Fidx %>);">
    		<% end if %>
    	</td>
    	<td>
    		<% if ipkum.Fipkumitem(i).Fpaperexist = "Y" then %>
    			<input type="button" class="button" value="증빙" onclick="javascript:cashreceiptInfo('<%= ipkum.Fipkumitem(i).Forderserial %>');">
    		<% end if %>
    	</td>
    </tr>
    </form>
    <% next %>
    <% end if %>

    <tr height="25" bgcolor="FFFFFF">
		<td colspan="16" align="center">
			<% if ipkum.HasPreScroll then %>
	    		<a href="javascript:scrollmove('<%= ipkum.StartScrollPage-1 %>');">[pre]</a>
	    	<% else %>
	    	<% end if %>
	    	<% for ix = 0 + ipkum.StartScrollPage  to ipkum.StartScrollPage + ipkum.FScrollCount - 1 %>
	    	<% if (ix > ipkum.FTotalpage) then Exit for %>
	    	<% if CStr(ix) = CStr(ipkum.FCurrPage) then %>
	    	<font color="#666666" class="verdana-xsmall"><strong><%= ix %></strong></font>
	    	<% else %>
	    	<a href="javascript:scrollmove('<%= ix %>');" class="bb"><font color="#666666"><%= ix %></font></a>
	    	<% end if %>
	    	<% next %>
	    	<% if ipkum.HasNextScroll then %>
	    	<a href="javascript:scrollmove('<%= ix %>');" class="verdana-xsmall">[next]</a>
	    	<% else %>
	    	&nbsp;
	    	<% end if %>
		</td>
	</tr>
</table>



<% set ipkum=nothing %>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
