<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  은행입출금리스트
' History : 2009.04.07 서동석 생성
'			2011.05.20 한용민 수정
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/jungsan/ipkumlistcls.asp"-->
<%

dim yyyy1,mm1,dd1 ,yyyy2,mm2,dd2 ,fromDate,toDate ,ckdate,tenbank,ipkumname,txamount,page
dim searchtype01,searchtype02,orderby,research,itype ,ipkum,i,ix
dim inoutgubun, exmatchfinish, excustomer, ex10x10, showdismatch

	ckdate = request("ckdate")
	tenbank = request("tenbank")
	page = request("page")

	ipkumname = request("ipkumname")
	txamount = request("txamount")
	searchtype01 =  request("searchtype01")
	searchtype02 =  request("searchtype02")

	orderby = request("orderby")
	research = request("research")
	page = request("page")
	itype = request("itype")
	yyyy1 = request("yyyy1")
	mm1 = request("mm1")
	dd1 = request("dd1")
	yyyy2 = request("yyyy2")
	mm2 = request("mm2")
	dd2 = request("dd2")
	inoutgubun = request("inoutgubun")
	exmatchfinish = request("exmatchfinish")
	excustomer = request("excustomer")
	ex10x10 = request("ex10x10")
	showdismatch = request("showdismatch")

if (yyyy1="") then yyyy1 = Cstr(Year(now()))
if (mm1="") then mm1 = Cstr(Month(now()))
if (dd1="") then dd1 = Cstr(day(now())-1)
if (yyyy2="") then yyyy2 = Cstr(Year(now()))
if (mm2="") then mm2 = Cstr(Month(now()))
if (dd2="") then dd2 = Cstr(day(now()))

fromDate = CStr(DateSerial(yyyy1, mm1, dd1))
toDate = CStr(DateSerial(yyyy2, mm2, dd2+1))

if page="" then page=1
IF research="" then
    orderby="on"
    ckdate="on"
	showdismatch = "on"
end if

'/핑거스 일경우
if getpart_sn("",session("ssBctId")) = "16" then
	'/어드민권한 파트장 미만 일경우
	if getlevel_sn("",session("ssBctId")) > 3 then
	    if (tenbank<>"27703783604018") or (tenbank<>"27702818201078") then
		    tenbank = "27702818201078"
		end if
	end if
end if

'/소싱파트 강제지정
if getpart_sn("",session("ssBctId")) = "21" then
	tenbank = "27703918804031"
end if

'/물류파트
if getpart_sn("",session("ssBctId")) = "9" then
    '/어드민권한 파트장 미만 일경우
	if getlevel_sn("",session("ssBctId")) > 3 then
	    tenbank = "27703783604018"
    end if
end if



set ipkum = new IpkumChecklist
	ipkum.FCurrpage=page
	ipkum.FPagesize=100
	ipkum.FScrollCount = 10
	ipkum.Fckdate = ckdate
	ipkum.Ctenbank = tenbank

	if (searchtype01 <> "") then
		ipkum.FRectJeokyo = ipkumname
	end if
	if (searchtype02 <> "") then
		ipkum.FRectTXAmmount = txamount
	end if

	ipkum.FOrderby = orderby
	ipkum.FRectInOutGubun = inoutgubun
	ipkum.FRectExcluudeMatchFinish = exmatchfinish
	ipkum.FRectExcluudeCustomer = excustomer
	ipkum.FRectExcluude10X10 = ex10x10
	ipkum.FRectShowDismatch = showdismatch

	if ckdate="on" then
		ipkum.FRectRegStart = fromDate
		ipkum.FRectRegEnd = toDate
	end if

	ipkum.GetipkumlistAccounts

%>

<script language='javascript'>

function EnDisabledDateBox(comp){
    //nothing
}
function NextPage(page){
	document.frmipkum.page.value = page;
	document.frmipkum.submit();
}

function PopPaymentlist(frmpayment){
	var url  = "pop_paymentlist_accounts.asp";
  	var title  = "PopPaymentlist";
  	var status = "toolbar=no,directories=no,scrollbars=yes,resizable=no,status=no,menubar=no,width=800, height=600, top=0,left=20";

  	window.open("", title,status);

	frmpayment.target = title;
	frmpayment.action = "pop_paymentlist_accounts.asp";
	frmpayment.method = "post";
	frmpayment.submit();
}

function PopJungsanList(bankinoutidx) {
	var popwin = window.open("/admin/offshop/etc_meachul.asp?menupos=1466&bankinoutidx=" + bankinoutidx,'PopJungsanList','width=1100, height=600, scrollbars=yes, resizable=yes');
	popwin.focus();
}


function reg() {
	frmipkum.submit();
}

function popViewMatchMemo(inoutidx) {
	var winR = window.open("popMatchMemo.asp?inoutidx=" + inoutidx,"popViewMatchMemo","width=600, height=300, resizable=yes, scrollbars=yes");
	winR.focus();
}

</script>


<!-- 검색 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frmipkum" method="get" action="">
<input type="hidden" name="showtype" value="showtype">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="page" value="">
<input type="hidden" name="research" value="on">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="3" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
	<td align="left">
		구분 :
		<select class="select" name="inoutgubun">
			<option value="">전체</option>
			<option value="1" <%if (inoutgubun = "1") then %>selected<% end if %> >출금</option>
			<option value="2" <%if (inoutgubun = "2") then %>selected<% end if %> >입금</option>
		</select>

		은행 :
		<!-- 디비에서 긁어 올것 db_order.dbo.tbl_bank_div -->
		<% Call drawSelectBoxBankList("tenbank", tenbank) %>
	</td>
	<td rowspan="3" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="검색" onClick="reg();">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
		<input type="checkbox" name="searchtype01" value="Y" <% if searchtype01<>"" then response.write "checked" %> > 적요
		<input type=text name=ipkumname value="<%= ipkumname %>" size=10>
		&nbsp;
		<input type="checkbox" name="searchtype02" value="Y" <% if searchtype02<>"" then response.write "checked" %> > 입출금액
		<input type=text name=txamount value="<%= txamount %>" size=10>
		&nbsp;
        <input type=checkbox name="ckdate" <% if ckdate="on" then  response.write "checked" %> onclick="EnDisabledDateBox(this)">검색기간
		<% DrawDateBox yyyy1,yyyy2,mm1,mm2,dd1,dd2 %>
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
		<input type="checkbox" name="orderby" <% if orderby<>"" then response.write "checked" %> > 최근일순
		&nbsp;
		<input type="checkbox" name="exmatchfinish" value="Y" <% if exmatchfinish<>"" then response.write "checked" %> > 매칭완료 제외
		&nbsp;
		<input type="checkbox" name="excustomer" value="Y" <% if excustomer<>"" then response.write "checked" %> > 고객입금 제외
		&nbsp;
		<input type="checkbox" name="ex10x10" value="Y" <% if ex10x10<>"" then response.write "checked" %> > 텐바이텐(03 0277)입금 제외
		&nbsp;
		<input type="checkbox" name="showdismatch" value="Y" <% if showdismatch<>"" then response.write "checked" %> > 매칭제외 포함
	</td>
</tr>
</form>
</table>
<!-- 검색 끝 -->
<br>
<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
<form name="formxl" method="post" action="">
<input type="hidden" name="ckdate" value="<%= ckdate %>">
<input type="hidden" name="yyyy1" value="<%= yyyy1 %>">
<input type="hidden" name="mm1" value="<%= mm1 %>">
<input type="hidden" name="dd1" value="<%= dd1 %>">
<input type="hidden" name="yyyy2" value="<%= yyyy2 %>">
<input type="hidden" name="mm2" value="<%= mm2 %>">
<input type="hidden" name="dd2" value="<%= dd2 %>">
<input type="hidden" name="searchtype01" value="<%= searchtype01 %>">
<input type="hidden" name="searchtype02" value="<%= searchtype02 %>">
<input type="hidden" name="ipkumname" value="<%= ipkumname %>">
<input type="hidden" name="txamount" value="<%= txamount %>">
<input type="hidden" name="tenbank" value="<%= tenbank %>">
<input type="hidden" name="orderby" value="<%= orderby %>">
<input type="hidden" name="itype" value="xl">
<tr>
	<td align="left">
	</td>
	<td align="right">
    	<input type="button" class="adminbutton" value="WEB" onclick="javascript:PopPaymentlist(formxl);">
    	<input type="button" class="adminbutton" value="EXCEL" onclick="javascript:PopPaymentlist(formxl);">
	</td>
</tr>
</form>
</table>
<!-- 액션 끝 -->

<!-- 리스트 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="14">
		검색결과 : <b><%= ipkum.FTotalCount %></b>
		&nbsp;
		페이지 : <b><%= page %>/ <%= ipkum.FTotalPage %></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td>IDX</td>
	<td>은행명</td>
	<td>계좌번호</td>
	<td>입출금일</td>
	<td>적요</td>
	<td>거래구분</td>
  	<td>입금금액</td>
	<td>출금금액</td>
	<td>잔액</td>
	<td>업데이트시간</td>
	<td>매칭잔액</td>
	<td>관련매출코드</td>
	<td>매칭상태</td>
	<td>비고</td>
</tr>
<% if ipkum.FResultCount > 0 then %>
<% for i=0 to ipkum.FResultCount-1 %>
<tr align="center" bgcolor="#FFFFFF" onmouseover=this.style.background="orange"; onmouseout=this.style.background="#ffffff";>
	<td><%= ipkum.Fipkumitem(i).Finoutidx       %></td>
	<td><%= ipkum.Fipkumitem(i).Fbkname         %></td>
	<td><%= ipkum.Fipkumitem(i).Fbkacctno       %></td>
	<td>
		<%= mid(ipkum.Fipkumitem(i).Fbkdate,1,4) %>-<%= mid(ipkum.Fipkumitem(i).Fbkdate,5,2) %>-<%= mid(ipkum.Fipkumitem(i).Fbkdate,7,2) %>
	</td>
	<td>
		<%= ipkum.Fipkumitem(i).Fbkjukyo        %>
	</td>
	<td><%= ipkum.Fipkumitem(i).Fbkcontent      %></td>
  	<td align="right">
  	    <% if ipkum.Fipkumitem(i).finout_gubun = "2" then %>
  			<%= FormatNumber(ipkum.Fipkumitem(i).Fbkinput,0) %>
  		<% end if %>
  	</td>
	<td align="right">
		<% if ipkum.Fipkumitem(i).finout_gubun = "1" then %>
			<%= FormatNumber(ipkum.Fipkumitem(i).Fbkinput,0) %>
		<% end if %>
	</td>
	<td align="right"><%= FormatNumber(ipkum.Fipkumitem(i).Fbkjango,0) %></td>
	<td>
		<%= mid(ipkum.Fipkumitem(i).Fbkxferdatetime,1,4) %>-<%= mid(ipkum.Fipkumitem(i).Fbkxferdatetime,5,2) %>-<%= mid(ipkum.Fipkumitem(i).Fbkxferdatetime,7,2) %>&nbsp;
		<%= mid(ipkum.Fipkumitem(i).Fbkxferdatetime,9,2) %>:<%= mid(ipkum.Fipkumitem(i).Fbkxferdatetime,11,2) %>:<%= mid(ipkum.Fipkumitem(i).Fbkxferdatetime,13,2) %>
	</td>
	<td align="right">
		<% if Not IsNull(ipkum.Fipkumitem(i).Ftotmatchprice) then %>
			<%= FormatNumber((ipkum.Fipkumitem(i).Fbkinput - ipkum.Fipkumitem(i).Ftotmatchprice),0) %>
		<% end if %>
	</td>
	<td>
		<% if (ipkum.Fipkumitem(i).Fjungsanidx <> 0) then %>
			<a href="javascript:PopJungsanList(<%= ipkum.Fipkumitem(i).Finoutidx %>)">
				<%= ipkum.Fipkumitem(i).Fjungsanidx %>
				<% if (ipkum.Fipkumitem(i).Fjungsancnt > 1) then %>
					외 <%= (ipkum.Fipkumitem(i).Fjungsancnt - 1) %>
				<% end if %>
			</a>
		<% end if %>
		<% if Not IsNull(ipkum.Fipkumitem(i).Forderserial) then %>
			<%= ipkum.Fipkumitem(i).Forderserial %>
		<% end if %>
	</td>
	<td>
		<font color="<%= ipkum.Fipkumitem(i).GetMatchStateColor %>"><%= ipkum.Fipkumitem(i).GetMatchStateName %></font>
	</td>
	<td align="left">
		<% if Not IsNull(ipkum.Fipkumitem(i).Fmatchmemo) then %>
			<a href="javascript:popViewMatchMemo(<%= ipkum.Fipkumitem(i).Finoutidx %>)"><%= Left(ipkum.Fipkumitem(i).Fmatchmemo, 10) %></a>
		<% else %>
			<a href="javascript:popViewMatchMemo(<%= ipkum.Fipkumitem(i).Finoutidx %>)">메모</a>
		<% end if %>
	</td>
</tr>
<% next %>
<tr bgcolor="#FFFFFF">
	<td colspan="14" align="center">
		<!-- 페이지 시작 -->
    	<% if ipkum.HasPreScroll then %>
    		<a href="javascript:NextPage('<%= ipkum.StarScrollPage-1 %>')">[pre]</a>
    	<% else %>
    		[pre]
    	<% end if %>

    	<% for i=0 + ipkum.StarScrollPage to ipkum.FScrollCount + ipkum.StarScrollPage - 1 %>
    		<% if i>ipkum.FTotalpage then Exit for %>
    		<% if CStr(page)=CStr(i) then %>
    		<font color="red">[<%= i %>]</font>
    		<% else %>
    		<a href="javascript:NextPage('<%= i %>')">[<%= i %>]</a>
    		<% end if %>
    	<% next %>

    	<% if ipkum.HasNextScroll then %>
    		<a href="javascript:NextPage('<%= i %>')">[next]</a>
    	<% else %>
    		[next]
    	<% end if %>
		<!-- 페이지 끝 -->
	</td>
</tr>
<% else %>
	<tr bgcolor="#FFFFFF">
		<td colspan="15" align="center" class="page_link">[검색결과가 없습니다.]</td>
	</tr>
<% end if %>
</table>

<% set ipkum=nothing %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
