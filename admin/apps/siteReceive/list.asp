<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  현장수령주문
' History : 2012.05.10 서동석 생성
'			2012.05.21 한용민 수정
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/order/jumuncls.asp"-->
<%
Dim orderserial ,searchrect, searchtype ,page ,ojumun ,ix ,ipkumdiv , reqdate
dim yyyy1,yyyy2,mm1,mm2,dd1,dd2 ,pdate, nowdate,searchnextdate,research, jumundiv
dim reqyyyy , reqmm ,reqdd , checkreqdate , cancelyn
	orderserial = requestCheckvar(request("orderserial"),16)
	searchtype  = requestCheckVar(request("searchtype"),32)
	searchrect  = requestCheckVar(request("searchrect"),32)
	yyyy1       = requestCheckVar(request("yyyy1"),4)
	mm1         = requestCheckVar(request("mm1"),2)
	dd1         = requestCheckVar(request("dd1"),2)
	yyyy2       = requestCheckVar(request("yyyy2"),4)
	mm2         = requestCheckVar(request("mm2"),2)
	dd2         = requestCheckVar(request("dd2"),2)
	jumundiv    = requestCheckVar(request("jumundiv"),10)
	page = request("page")
	ipkumdiv  = requestCheckVar(request("ipkumdiv"),3)
	reqyyyy       = requestCheckVar(request("reqyyyy"),4)
	reqmm         = requestCheckVar(request("reqmm"),2)
	reqdd         = requestCheckVar(request("reqdd"),2)
	checkreqdate         = requestCheckVar(request("checkreqdate"),2)
	cancelyn = requestCheckvar(request("cancelyn"),1)
	
if (page="") then page=1
if (jumundiv="") then jumundiv="7"      ''현장수령 주문
	
nowdate = Left(CStr(now()),10)
pdate   = Left(CStr(DateAdd("m",-1,now())),10)

if (yyyy1="") then
	yyyy1 = Left(pdate,4)
	mm1   = Mid(pdate,6,2)
	dd1   = Mid(pdate,9,2)

	yyyy2 = Left(nowdate,4)
	mm2   = Mid(nowdate,6,2)
	dd2   = Mid(nowdate,9,2)
end if

if (reqyyyy="") then
	reqyyyy = Left(nowdate,4)
	reqmm   = Mid(nowdate,6,2)
	reqdd   = Mid(nowdate,9,2)
end if

reqdate = Cdate(reqyyyy + "-" + reqmm + "-" + reqdd)
searchnextdate = Left(CStr(DateAdd("d",Cdate(yyyy2 + "-" + mm2 + "-" + dd2),1)),10)

set ojumun = new CJumunMaster
	ojumun.FRectJumunDiv = jumundiv
	ojumun.FRectRegStart = yyyy1 + "-" + mm1 + "-" + dd1
	ojumun.FRectRegEnd = searchnextdate

	if searchtype="01" then
		ojumun.FRectBuyname = searchrect
	elseif searchtype="02" then
		ojumun.FRectReqName = searchrect
	elseif searchtype="03" then
		ojumun.FRectUserID = searchrect
	elseif searchtype="04" then
		ojumun.FRectIpkumName = searchrect
	elseif searchtype="06" then
		ojumun.FRectSubTotalPrice = searchrect
	elseif searchtype="07" then
	    ojumun.FRectBuyHp = searchrect
	elseif searchtype="08" then
	    ojumun.FRectreqHp = searchrect
	elseif searchtype="09" then
	    ojumun.FRectBuyEmail = searchrect    
	end if

	ojumun.FPageSize = 30
	''ojumun.FRectIpkumDiv4 = ckipkumdiv4
	ojumun.FRectIpkumDiv2 = "on"
	ojumun.FRectOrderSerial = orderserial
	ojumun.FCurrPage = page

	if ipkumdiv = "44" then
		ojumun.FRectIpkumDiv4 = "ON"
	elseif ipkumdiv = "444" then
		ojumun.FRectIpkumDiv4before = "ON"		
	else
		ojumun.FRectIpkumdiv = ipkumdiv
	end if
	
	if checkreqdate = "on" then
		ojumun.FRectreqdate = reqdate
	end if

	if cancelyn <> "" then
		ojumun.FRectcancelyn = cancelyn
	end if

	ojumun.SearchJumunList

%>

<script language='javascript'>

function ckreqdate(){
	if (frm.checkreqdate.checked){
		frm.reqyyyy.disabled = false;
		frm.reqmm.disabled = false;
		frm.reqdd.disabled = false;
	}else{
		frm.reqyyyy.disabled = true;
		frm.reqmm.disabled = true;
		frm.reqdd.disabled = true;
	}
}

function chksubmit(){
	var frm = document.frm;
	if ((frm.orderserial.value.length>0)&&(frm.orderserial.value.length!=11)){
		alert('주문 번호를 정확히 입력하세요.');
		frm.orderserial.focus();
		frm.orderserial.select();
		return;
	}
	
	frm.submit();
}

function popSiteReceive(iorderserial){
	var popwin;
    popwin = window.open('popSiteReceive.asp?orderserial=' + iorderserial,'popSiteReceive','scrollbars=yes,resizable=yes,width=800,height=768');
    popwin.focus();
}


function NextPage(ipage){
	document.frm.page.value= ipage;
	document.frm.submit();
}

function GetOnLoad(){
	document.frm.orderserial.focus();
	document.frm.orderserial.select();

}

window.onload = GetOnLoad;

</script>


<table width="100%" align="center" cellpadding="5" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get" action="">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="page" value="">
<input type="hidden" name="research" value="on">
<tr align="center" bgcolor="#FFFFFF" >
	<td  rowspan="3" width="100" height="50" bgcolor="<%= adminColor("gray") %>">검색 조건</td>
	<td align="left">
		주문번호 :
		<input type="text" name="orderserial" value="<%= orderserial %>" size="16" maxlength="16" onKeyPress="if (event.keyCode == 13) chksubmit();">
		검색기간 :
		<% DrawDateBox yyyy1,yyyy2,mm1,mm2,dd1,dd2 %>
		수령예정일 : 
		<input type="checkbox" name="checkreqdate" onclick="ckreqdate()" <% if checkreqdate = "on" then response.write " checked" %>>
		<% DrawOneDateBoxdynamic "reqyyyy", reqyyyy, "reqmm", reqmm, "reqdd", reqdd, "", "", "", "" %>
		
		<script language="javascript">
			ckreqdate()
		</script>
	</td>
	<td  rowspan="3" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="검색" onClick="NextPage('');">
	</td>
</tr>
<tr>
    <td  bgcolor="#FFFFFF">
       검색조건 :
		<select name="searchtype">
    		<option value="">선택</option>
    		<option value="01" <% if searchtype="01" then response.write "selected" %> >구매자</option>
    		<option value="02" <% if searchtype="02" then response.write "selected" %> >수령인</option>
    		<option value="03" <% if searchtype="03" then response.write "selected" %> >아이디</option>
    		<!--<option value="04" <% if searchtype="04" then response.write "selected" %> >입금자</option>-->
    		<option value="06" <% if searchtype="06" then response.write "selected" %> >결제금액</option>
    		<option value="07" <% if searchtype="07" then response.write "selected" %> >휴대폰(구매자)</option>
    		<option value="08" <% if searchtype="08" then response.write "selected" %> >휴대폰(수령인)</option>	
    		<option value="09" <% if searchtype="09" then response.write "selected" %> >이메일(구매자)</option>
		</select>
		<input type="text" name="searchrect" value="<%= searchrect %>" size="20" maxlength="21" onkeydown="if(event.keyCode==13) NextPage('');;">
		거래상태 : 
		<select name="ipkumdiv" onchange="NextPage('');">
			<option value="" <% if ipkumdiv="" then response.write "selected" %>>주문접수이상</option>
			<option value="444" <% if ipkumdiv="444" then response.write "selected" %>>결제완료이전</option>
			<option value="44" <% if ipkumdiv="44" then response.write "selected" %>>결제완료이상</option>
			<option value="8" <% if ipkumdiv="8" then response.write "selected" %>>출고완료</option>		
		</select>
		취소여부 : 
		<select name="cancelyn" onchange="NextPage('');">
			<option value="" <% if cancelyn="" then response.write "selected" %>>전체</option>
			<option value="Y" <% if cancelyn="Y" then response.write "selected" %>>Y</option>
			<option value="N" <% if cancelyn="N" then response.write "selected" %>>N</option>		
		</select>		
    </td>
</tr>
<tr>
    <td bgcolor="#FFFFFF" align="right">
		<input type="button" value="POS" class="button_s" onClick="popSiteReceive('')">
    </td>
</tr>
</form>
</table>

<BR>

<table width="100%" cellpadding="3" cellspacing="1" class="a" bgcolor="#999999">
<tr bgcolor="#FFFFFF">
	<td colspan="11">
		총 건수 : <Font color="#3333FF"><%= FormatNumber(ojumun.FTotalCount,0) %></font>
		&nbsp;총 금액 : <Font color="#3333FF"><%= FormatNumber(ojumun.FSubTotal,0) %></font>
		&nbsp;평균객단가 : <Font color="#3333FF"><%= FormatNumber(ojumun.FAvgTotal,0) %></font>
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td colspan="11" align="right">page : <%= ojumun.FCurrPage %>/<%=ojumun.FTotalPage %></td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td>주문번호</td>
	<!-- td>국가</td -->
	<!-- td >Site</td -->
	<td>UserID</td>
	<td>구매자</td>
	<td>수령인</td>
	<td>결제금액</td>
	<td>구매총액</td>
	<td>결제방법</td>
	<td>거래상태</td>
	<td>취소<br>여부</td>
	<td>수령(예정)일</td>
	<!-- td>주문일</td -->
	<td>상세</td>
</tr>
<% if ojumun.FresultCount<1 then %>
<tr bgcolor="#FFFFFF" align="center">
	<td colspan="11">[검색결과가 없습니다.]</td>
</tr>
<% else %>

<% for ix=0 to ojumun.FresultCount-1 %>

<% if ojumun.FMasterItemList(ix).IsAvailJumun then %>
	<tr bgcolor="#FFFFFF" align="center">
<% else %>
	<tr bgcolor="silver" align="center">
<% end if %>
	<td width="100">
		<a href="javascript:popSiteReceive('<%= ojumun.FMasterItemList(ix).FOrderSerial %>');" class="zzz">
		<%= ojumun.FMasterItemList(ix).FOrderSerial %></a>
	</td>
	<!-- td  width="40"><%= CHKIIF(ojumun.FMasterItemList(ix).IsForeignDeliver,ojumun.FMasterItemList(ix).FDlvcountryCode,"") %></td -->
	<!-- td width="60"><font color="<%= ojumun.FMasterItemList(ix).SiteNameColor %>"><%= ojumun.FMasterItemList(ix).FSitename %></font></td -->
	
	<% if ojumun.FMasterItemList(ix).UserIDName<>"&nbsp;" then %>
		<td><%= ojumun.FMasterItemList(ix).UserIDName %></td>
	<% else %>
		<td><%= ojumun.FMasterItemList(ix).UserIDName %></td>
	<% end if %>
	<td><%= ojumun.FMasterItemList(ix).FBuyName %></td>
	<td><%= ojumun.FMasterItemList(ix).FReqName %></td>
	<td align="right" width="80">
		<font color="<%= ojumun.FMasterItemList(ix).SubTotalColor%>">
		<%= FormatNumber(ojumun.FMasterItemList(ix).FSubTotalPrice,0) %></font>
	</td>
	<td align="right" width="80"><%= FormatNumber(ojumun.FMasterItemList(ix).FTotalSum,0) %></td>
	<td width="80"><%= ojumun.FMasterItemList(ix).JumunMethodName %></td>
	<td width="80">
		<font color="<%= ojumun.FMasterItemList(ix).IpkumDivColor %>">
		<%= ojumun.FMasterItemList(ix).IpkumDivName %></font>
	</td>
	<td width="50">
		<font color="<%= ojumun.FMasterItemList(ix).CancelYnColor %>">
		<%= ojumun.FMasterItemList(ix).CancelYnName %></font>
	</td>
	<td width="110"><%= Left(ojumun.FMasterItemList(ix).FReqdate,10) %></td>
	<!-- td width="110"><%= Left(ojumun.FMasterItemList(ix).GetRegDate,10) %></td -->
	<td></td>
</tr>
<% next %>

<tr bgcolor="#FFFFFF">
	<td colspan="11" height="30" align="center">
	<% if ojumun.HasPreScroll then %>
		<a href="javascript:NextPage('<%= ojumun.StartScrollPage-1 %>')">[pre]</a>
	<% else %>
		[pre]
	<% end if %>

	<% for ix=0 + ojumun.StartScrollPage to ojumun.FScrollCount + ojumun.StartScrollPage - 1 %>
		<% if ix>ojumun.FTotalpage then Exit for %>
		<% if CStr(page)=CStr(ix) then %>
		<font color="red">[<%= ix %>]</font>
		<% else %>
		<a href="javascript:NextPage('<%= ix %>')">[<%= ix %>]</a>
		<% end if %>
	<% next %>

	<% if ojumun.HasNextScroll then %>
		<a href="javascript:NextPage('<%= ix %>')">[next]</a>
	<% else %>
		[next]
	<% end if %>
	</td>
</tr>
<% end if %>
</table>

<%
set ojumun = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
