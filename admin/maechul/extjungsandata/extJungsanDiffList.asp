<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbHelper.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db3open.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/outmall_function.asp"-->
<!-- #include virtual="/lib/classes/extjungsan/extjungsancls.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<%
dim research, page
dim sellsite, searchfield, searchtext, diffType
'dim yyyy, mm, yyyymm, yyyymm_prev, yyyymm_next
'dim yyyy1,mm1

Dim i

research = requestCheckvar(request("research"),10)
page 	 = requestCheckvar(request("page"),10)

' yyyy1   = requestCheckvar(request("yyyy1"),4)
' mm1     = requestCheckvar(request("mm1"),2)

sellsite		= request("sellsite")
searchfield 	= request("searchfield")
searchtext 		= Replace(Replace(request("searchtext"), "'", ""), Chr(34), "")
diffType 		= request("diffType")

if (page="") then page = 1
if (diffType="") then diffType = "S"


' if (yyyy1="") then
' 	yyyy1 = Cstr(Year(now()))
' 	mm1 = Cstr(Month(now()) - 2)
' end if

' yyyymm = yyyy1 + "-" & mm1
' yyyymm_prev = Left(DateSerial(yyyy1,(mm1 - 1), 1), 7)
' yyyymm_next = Left(DateSerial(yyyy1,(mm1 + 1), 1), 7)


Dim oCExtJungsan
set oCExtJungsan = new CExtJungsan
	oCExtJungsan.FPageSize = 100
	oCExtJungsan.FCurrPage = page
	oCExtJungsan.FRectSellSite = sellsite
	oCExtJungsan.FRectDiffType = diffType

	'oCExtJungsan.FRectYYYYMM = yyyymm
	'oCExtJungsan.FRectSearchField = searchfield
	'oCExtJungsan.FRectSearchText = searchtext

    oCExtJungsan.GetExtJungsanDiff


%>

<script language='javascript'>

function NextPage(page){
    document.frm.page.value = page;
    document.frm.submit();
}

function popExtSiteJungsanData() {
    var window_width = 500;
    var window_height = 250;

    var popwin = window.open("/admin/maechul/extjungsandata/popRegExtJungsanDataFile.asp","popExtSiteJungsanData","width=" + window_width + " height=" + window_height + " left=0 top=0 scrollbars=yes resizable=yes status=yes");

	popwin.focus();
}



function popMiMapExtjungsan(sellsite,yyyy1,mm1,dd1,yyyy2,mm2,dd2,jungsantype){
	var iurl = "/admin/maechul/extjungsandata/extJungsanDataList.asp";
	iurl += "?menupos=1652&page=1&sellsite="+sellsite;
	iurl += "&yyyy1="+yyyy1+"&mm1="+mm1+"&dd1="+dd1;
	iurl += "&yyyy2="+yyyy2+"&mm2="+mm2+"&dd2="+dd2;
	iurl += "&mimap=on&jungsantype="+jungsantype;

	var popwin = window.open(iurl,"popextJungsanDataList","width=1200 height=850 scrollbars=yes resizable=yes status=yes");

	popwin.focus();

}

function popFixedErrExtjungsanDTL(sellsite,yyyy1,mm1,jungsantype,ierrtp,iaccerrtype){
	var iurl = "/admin/maechul/extjungsandata/extJungsanFixedErrDetail.asp";
	iurl += "?menupos=1652&page=1&sellsite="+sellsite;
	iurl += "&yyyy1="+yyyy1+"&mm1="+mm1;
	iurl += "&jungsantype="+jungsantype;
	iurl += "&errtp="+ierrtp;
	iurl += "&accerrtype="+iaccerrtype;

	var popwin = window.open(iurl,"popextJungsanFixedErrDetail","width=1200 height=850 scrollbars=yes resizable=yes status=yes");

	popwin.focus();

}

function popErrExtjungsanDTL(sellsite,yyyy1,mm1,dd1,yyyy2,mm2,dd2,jungsantype,ierrtp,ionlyErrNoExists){
	var iurl = "/admin/maechul/extjungsandata/extJungsanErrDetail.asp";
	iurl += "?menupos=1652&page=1&sellsite="+sellsite;
	iurl += "&yyyy1="+yyyy1+"&mm1="+mm1+"&dd1="+dd1;
	iurl += "&yyyy2="+yyyy2+"&mm2="+mm2+"&dd2="+dd2;
	iurl += "&jungsantype="+jungsantype;
	iurl += "&errtp="+ierrtp;
	iurl += "&onlyErrNoExists="+ionlyErrNoExists;

	var popwin = window.open(iurl,"popextJungsanErrDetail","width=1200 height=850 scrollbars=yes resizable=yes status=yes");

	popwin.focus();

}

function jsExtJungsanDiffMake(sellsite,yyyymm) {
	var frm = document.frmAct;

	if (confirm(sellsite + " "+yyyymm+" (재)작성하시겠습니까?") == true) {
		frm.mode.value = "extjungsandiffmake";
		frm.sellsite.value = sellsite;
		frm.yyyymm.value = yyyymm;

		frm.submit();
	}
}


function jsExtJungsanErrMake(sellsite,yyyymm) {
	var frm = document.frmAct;

	if (confirm(sellsite + " "+yyyymm+" Err 작성하시겠습니까?") == true) {
		frm.mode.value = "extjungsanerrmake";
		frm.sellsite.value = sellsite;
		frm.yyyymm.value = yyyymm;

		frm.submit();
	}
}

function jsExtJungsanDiffMakeDetail(sellsite){
	var frm = document.frmAct;

	if (confirm(sellsite + " 누적 오차 작성하시겠습니까?") == true) {
		frm.mode.value = "extjungsanaccDetailmake";
		frm.sellsite.value = sellsite;
		frm.yyyymm.value = "";
		frm.submit();
	}
}

</script>
<link rel="stylesheet" href="/css/tpl.css" type="text/css">

<!-- 검색 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get" action="">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="page" value="">
<input type="hidden" name="research" value="on">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
	<td align="left">
		&nbsp;
		제휴몰:	<%' = getJungsanXsiteComboHTML("sellsite",sellsite,"") %>
				<% call drawOutmallSelectBox("sellsite",sellsite) %>
		<% if (FALSE) then %>
		&nbsp;
		매출월:
		<% DrawYMBox yyyy1,mm1 %>
		<% end if %>
		&nbsp;
		조회내역:

		<input type="radio" name="diffType" value="S" <% if (diffType = "S") then %>checked<% end if %> > 각각정산기준
		<input type="radio" name="diffType" value="T" <% if (diffType = "T") then %>checked<% end if %> > TEN정산기준

		<% if (FALSE) then %>
		<input type="radio" name="diffType" value="DIF" <% if (diffType = "DIF") then %>checked<% end if %> > 오차내역
		<input type="radio" name="diffType" value="TOT" <% if (diffType = "TOT") then %>checked<% end if %> > 전체내역
		<input type="radio" name="diffType" value="SUM" <% if (diffType = "SUM") then %>checked<% end if %> > 합계내역
		<% end if %>
	</td>
	<td rowspan="4" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="검색" onClick="javascript:document.frm.submit();">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" height="30">
	<td align="left">
		&nbsp;
		<% if (FALSE) then %>
		* 검색조건 :
		<select class="select" name="searchfield">
			<option value=""></option>
			<option value="OrgOrderserial" <% if (searchfield = "OrgOrderserial") then %>selected<% end if %> >원주문번호</option>
		</select>
		<input type="text" class="text" name="searchtext" size="30" value="<%= searchtext %>">
		<% end if %>
	</td>
</tr>
</form>
</table>
<!-- 검색 끝 -->

<%

if (sellsite = "") then
	Response.write "<h5>제휴몰을 선택하세요</h5>"
end if

%>

<p>

<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
<tr>
	<td align="left">

	</td>
	<td align="right">

	</td>
</tr>
</table>
<!-- 액션 끝 -->

<p>

<!-- 리스트 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="<%=CHKIIF(diffType="S","21","20")%>">
		검색결과 : <b><%= oCExtJungsan.FTotalcount %></b>
		&nbsp;
		페이지 : <b><%= page %> / <%= oCExtJungsan.FTotalPage %></b>
	</td>
</tr>
<% if diffType="S" then %>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width="70">년월</td>
	<td width="100">자사매출합<br>(정산기준)</td>
	<td width="100">자사매출(상품)</td>
	<td width="90">자사매출(배송비)</td>
	<td width="1"></td>
	<td width="100">제휴매출합<br>(제휴정산입력기준)</td>
	<td width="100">제휴매출(상품)</td>
	<td width="90">제휴매출(배송비)</td>
	<td width="1"></td>
	<td width="100">당월오차(상품)<br>(자사-제휴)</td>
	<td width="90">당월오차(배송비)<br>(자사-제휴)</td>
	<td width="1"></td>
	<td width="100">누적오차(상품)<br>(자사-제휴)</td>
	<td width="100">누적오차(배송비)<br>(자사-제휴)</td>
	<td width="120">최종업데이트</td>
	<td width="80">당월오차상세<br>(자사-제휴)</td>
	<td width="80">미반영오차<br>(자사-제휴)</td>
	<td width="80">당월Fix오차<br>(자사-제휴)</td>
	<td width="80">당월Fix오차<br>(자사-제휴)상품</td>
	<td width="90">당월Fix오차<br>(자사-제휴)상품<br>검토필요</td>
	<td>
		비고
		<% if oCExtJungsan.FresultCount>0 then %>
			<% if oCExtJungsan.FItemList(0).Fyyyymm<LEFT(now(),7) then %>
			<% if (sellsite<>"") then %>
			<br><input type="button" value="<%=LEFT(now(),7)%>작성" onClick="jsExtJungsanDiffMake('<%=sellsite%>','<%=LEFT(now(),7)%>')">
			<% end if %>
			<% end if %>
		<% end if %>
	</td>
</tr>
<% for i=0 to oCExtJungsan.FresultCount -1 %>
<tr align="right" bgcolor="FFFFFF" onmouseover=this.style.background="F1F1F1"; onmouseout=this.style.background="FFFFFF";>
	<td align="center" ><%= oCExtJungsan.FItemList(i).Fyyyymm %></td>
	<td><%= FormatNumber(oCExtJungsan.FItemList(i).FTMeachulItem+oCExtJungsan.FItemList(i).FTMeachulDLV,0) %></td>
	<td><%= FormatNumber(oCExtJungsan.FItemList(i).FTMeachulItem,0) %></td>
	<td><%= FormatNumber(oCExtJungsan.FItemList(i).FTMeachulDLV,0) %></td>
	<td></td>
	<td><%= FormatNumber(oCExtJungsan.FItemList(i).FXMeachulItem+oCExtJungsan.FItemList(i).FXMeachulDLV,0) %></td>
	<td><%= FormatNumber(oCExtJungsan.FItemList(i).FXMeachulItem,0) %></td>
	<td><%= FormatNumber(oCExtJungsan.FItemList(i).FXMeachulDLV,0) %></td>
	<td></td>
	<td><%= FormatNumber(oCExtJungsan.FItemList(i).FmonthItemDiff,0) %></td>
	<td><%= FormatNumber(oCExtJungsan.FItemList(i).FmonthdlvDiff,0) %></td>
	<td></td>
	<td><%= FormatNumber(oCExtJungsan.FItemList(i).FdiffITEMsum,0) %></td>
	<td><%= FormatNumber(oCExtJungsan.FItemList(i).FdiffDlvsum,0) %></td>

	<td align="center" ><%= LEFT(oCExtJungsan.FItemList(i).FupdDt,16) %></td>
	<td>
		<% if NOT isNULL(oCExtJungsan.FItemList(i).FMonthDiffSum) then %>
		<a href="#" onClick="popFixedErrExtjungsanDTL('<%=sellsite%>','<%= LEFT(oCExtJungsan.FItemList(i).Fyyyymm,4) %>','<%= RIGHT(oCExtJungsan.FItemList(i).Fyyyymm,2) %>','','','');return false;"><%= FormatNumber(oCExtJungsan.FItemList(i).FMonthDiffSum,0) %></a>
			<% if oCExtJungsan.FItemList(i).getSumVsDtlDiffSum<>0 then %>
				<br><font color=gray><%=FormatNumber(oCExtJungsan.FItemList(i).getSumVsDtlDiffSum,0) %></font>
			<% end if %>
		<% end if %>
	</td>
	<td>
		<a href="#" onClick="popFixedErrExtjungsanDTL('<%=sellsite%>','<%= LEFT(oCExtJungsan.FItemList(i).Fyyyymm,4) %>','<%= RIGHT(oCExtJungsan.FItemList(i).Fyyyymm,2) %>','','1','');return false;"><%= FormatNumber(oCExtJungsan.FItemList(i).FMonthnotAssignErr,0) %></a>
	</td>
	<td>
		<a href="#" onClick="popFixedErrExtjungsanDTL('<%=sellsite%>','<%= LEFT(oCExtJungsan.FItemList(i).Fyyyymm,4) %>','<%= RIGHT(oCExtJungsan.FItemList(i).Fyyyymm,2) %>','','3','');return false;"><%= FormatNumber(oCExtJungsan.FItemList(i).FMonthErrAsignSum,0) %></a>
	</td>
	<td>
		<a href="#" onClick="popFixedErrExtjungsanDTL('<%=sellsite%>','<%= LEFT(oCExtJungsan.FItemList(i).Fyyyymm,4) %>','<%= RIGHT(oCExtJungsan.FItemList(i).Fyyyymm,2) %>','C','3','');return false;"><%= FormatNumber(oCExtJungsan.FItemList(i).FMonthErrAsignItemSum,0) %></a>
	</td>
	<td>
		<a href="#" onClick="popFixedErrExtjungsanDTL('<%=sellsite%>','<%= LEFT(oCExtJungsan.FItemList(i).Fyyyymm,4) %>','<%= RIGHT(oCExtJungsan.FItemList(i).Fyyyymm,2) %>','C','3','3');return false;"><%= FormatNumber(oCExtJungsan.FItemList(i).FMonthErrAsignItemSumReqCheck,0) %></a>
	</td>
	<td align="center" >
	<% if (sellsite<>"") then %>
	<% if (LEFT(dateadd("m",-3,now()),7)<=oCExtJungsan.FItemList(i).Fyyyymm) then %>
	<input type="button" value="재작성" onClick="jsExtJungsanDiffMake('<%=sellsite%>','<%=oCExtJungsan.FItemList(i).Fyyyymm%>')">

		<% if (oCExtJungsan.FItemList(i).Fyyyymm>="2020-01") then %>
		<input type="button" value="Fix오차" onClick="jsExtJungsanErrMake('<%=sellsite%>','<%=oCExtJungsan.FItemList(i).Fyyyymm%>')">
		<% end if %>
	<% end if %>
	<% end if %>

	</td>
</tr>
<% next %>
<% else %>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width="70">년월</td>
	<td width="120">자사매출합<br>(정산기준)</td>
	<td width="110">자사매출(상품)</td>
	<td width="100">자사매출(배송비)</td>
	<td width="1"></td>
	<td width="120">제휴매출합<br>(미매핑내역)</td>
	<td width="110">제휴매출(상품)<br>(미매핑내역)</td>
	<td width="110">제휴매출(배송비)<br>(미매핑내역)</td>
	<td width="1"></td>
	<td width="120">당월매핑오차(상품)<br>(자사-제휴)</td>
	<td width="120">당월매핑오차(배송비)<br>(자사-제휴)</td>

	<td width="110"><font color="#AAAAAA">당월매핑오차<br>(상품 오매핑)</font></td>
	<td width="100"><font color="#AAAAAA">당월매핑오차<br>(배송비 오매핑)</font></td>

	<td width="110"><font color="#AAAAAA">당월매핑오차<br>(상품 수량오차)</font></td>
	<td width="100"><font color="#AAAAAA">당월매핑오차<br>(배송비 수량오차)</font></td>

	<td width="1"></td>
	<td width="120">누적매핑오차(상품)<br>(자사-제휴-미매핑)</td>
	<td width="120">누적매핑오차(배송비)<br>(자사-제휴-미매핑)</td>
	<td width="200">최종업데이트</td>
	<td>
		비고
		<% if (sellsite<>"") then %>
		<% if oCExtJungsan.FresultCount>0 then %>
			<br><input type="button" value="재작성" onClick="jsExtJungsanDiffMakeDetail('<%=sellsite%>')">
		<% end if %>
		<% end if %>
	</td>
</tr>
<% for i=0 to oCExtJungsan.FresultCount -1 %>
<tr align="right" bgcolor="FFFFFF" onmouseover=this.style.background="F1F1F1"; onmouseout=this.style.background="FFFFFF";>
	<td align="center" ><%= oCExtJungsan.FItemList(i).Fyyyymm %></td>
	<td><%= FormatNumber(oCExtJungsan.FItemList(i).FTMeachulItem+oCExtJungsan.FItemList(i).FTMeachulDLV,0) %></td>
	<td><%= FormatNumber(oCExtJungsan.FItemList(i).FTMeachulItem,0) %></td>
	<td><%= FormatNumber(oCExtJungsan.FItemList(i).FTMeachulDLV,0) %></td>
	<td></td>
	<td><a href="#" onClick="popMiMapExtjungsan('<%=sellsite%>','<%= LEFT(oCExtJungsan.FItemList(i).Fyyyymm,4) %>','<%= RIGHT(oCExtJungsan.FItemList(i).Fyyyymm,2) %>','01','<%= LEFT(oCExtJungsan.FItemList(i).Fyyyymm,4) %>','<%= RIGHT(oCExtJungsan.FItemList(i).Fyyyymm,2) %>','<%= RIGHT(dateadd("d",-1,dateadd("m",1,oCExtJungsan.FItemList(i).Fyyyymm+"-01")),2) %>','');return false;"><%= FormatNumber(oCExtJungsan.FItemList(i).FXMeachulItem+oCExtJungsan.FItemList(i).FXMeachulDLV,0) %></a></td>
	<td><a href="#" onClick="popMiMapExtjungsan('<%=sellsite%>','<%= LEFT(oCExtJungsan.FItemList(i).Fyyyymm,4) %>','<%= RIGHT(oCExtJungsan.FItemList(i).Fyyyymm,2) %>','01','<%= LEFT(oCExtJungsan.FItemList(i).Fyyyymm,4) %>','<%= RIGHT(oCExtJungsan.FItemList(i).Fyyyymm,2) %>','<%= RIGHT(dateadd("d",-1,dateadd("m",1,oCExtJungsan.FItemList(i).Fyyyymm+"-01")),2) %>','C');return false;"><%= FormatNumber(oCExtJungsan.FItemList(i).FXMeachulItem,0) %></a></td>
	<td><a href="#" onClick="popMiMapExtjungsan('<%=sellsite%>','<%= LEFT(oCExtJungsan.FItemList(i).Fyyyymm,4) %>','<%= RIGHT(oCExtJungsan.FItemList(i).Fyyyymm,2) %>','01','<%= LEFT(oCExtJungsan.FItemList(i).Fyyyymm,4) %>','<%= RIGHT(oCExtJungsan.FItemList(i).Fyyyymm,2) %>','<%= RIGHT(dateadd("d",-1,dateadd("m",1,oCExtJungsan.FItemList(i).Fyyyymm+"-01")),2) %>','D');return false;"><%= FormatNumber(oCExtJungsan.FItemList(i).FXMeachulDLV,0) %></a></td>
	<td></td>
	<td><a href="#" onClick="popErrExtjungsanDTL('<%=sellsite%>','<%= LEFT(oCExtJungsan.FItemList(i).Fyyyymm,4) %>','<%= RIGHT(oCExtJungsan.FItemList(i).Fyyyymm,2) %>','01','<%= LEFT(oCExtJungsan.FItemList(i).Fyyyymm,4) %>','<%= RIGHT(oCExtJungsan.FItemList(i).Fyyyymm,2) %>','<%= RIGHT(dateadd("d",-1,dateadd("m",1,oCExtJungsan.FItemList(i).Fyyyymm+"-01")),2) %>','C','','');return false;"><%= FormatNumber(oCExtJungsan.FItemList(i).FmonthItemDiff,0) %></a></td>
	<td><a href="#" onClick="popErrExtjungsanDTL('<%=sellsite%>','<%= LEFT(oCExtJungsan.FItemList(i).Fyyyymm,4) %>','<%= RIGHT(oCExtJungsan.FItemList(i).Fyyyymm,2) %>','01','<%= LEFT(oCExtJungsan.FItemList(i).Fyyyymm,4) %>','<%= RIGHT(oCExtJungsan.FItemList(i).Fyyyymm,2) %>','<%= RIGHT(dateadd("d",-1,dateadd("m",1,oCExtJungsan.FItemList(i).Fyyyymm+"-01")),2) %>','D','','');return false;"><%= FormatNumber(oCExtJungsan.FItemList(i).FmonthdlvDiff,0) %></a></td>

	<td><a href="#" onClick="popErrExtjungsanDTL('<%=sellsite%>','<%= LEFT(oCExtJungsan.FItemList(i).Fyyyymm,4) %>','<%= RIGHT(oCExtJungsan.FItemList(i).Fyyyymm,2) %>','01','<%= LEFT(oCExtJungsan.FItemList(i).Fyyyymm,4) %>','<%= RIGHT(oCExtJungsan.FItemList(i).Fyyyymm,2) %>','<%= RIGHT(dateadd("d",-1,dateadd("m",1,oCExtJungsan.FItemList(i).Fyyyymm+"-01")),2) %>','C','1','');return false;"><font color="#AAAAAA"><%= FormatNumber(oCExtJungsan.FItemList(i).FmonthItemDiffMapErr,0) %></font></a></td>
	<td><a href="#" onClick="popErrExtjungsanDTL('<%=sellsite%>','<%= LEFT(oCExtJungsan.FItemList(i).Fyyyymm,4) %>','<%= RIGHT(oCExtJungsan.FItemList(i).Fyyyymm,2) %>','01','<%= LEFT(oCExtJungsan.FItemList(i).Fyyyymm,4) %>','<%= RIGHT(oCExtJungsan.FItemList(i).Fyyyymm,2) %>','<%= RIGHT(dateadd("d",-1,dateadd("m",1,oCExtJungsan.FItemList(i).Fyyyymm+"-01")),2) %>','D','1','');return false;"><font color="#AAAAAA"><%= FormatNumber(oCExtJungsan.FItemList(i).FmonthdlvDiffTMapErr,0) %></font></a></td>

	<td><a href="#" onClick="popErrExtjungsanDTL('<%=sellsite%>','<%= LEFT(oCExtJungsan.FItemList(i).Fyyyymm,4) %>','<%= RIGHT(oCExtJungsan.FItemList(i).Fyyyymm,2) %>','01','<%= LEFT(oCExtJungsan.FItemList(i).Fyyyymm,4) %>','<%= RIGHT(oCExtJungsan.FItemList(i).Fyyyymm,2) %>','<%= RIGHT(dateadd("d",-1,dateadd("m",1,oCExtJungsan.FItemList(i).Fyyyymm+"-01")),2) %>','C','','on');return false;"><font color="#AAAAAA"><%= FormatNumber(oCExtJungsan.FItemList(i).FmonthItemDiffNoExists,0) %></font></a></td>
	<td><a href="#" onClick="popErrExtjungsanDTL('<%=sellsite%>','<%= LEFT(oCExtJungsan.FItemList(i).Fyyyymm,4) %>','<%= RIGHT(oCExtJungsan.FItemList(i).Fyyyymm,2) %>','01','<%= LEFT(oCExtJungsan.FItemList(i).Fyyyymm,4) %>','<%= RIGHT(oCExtJungsan.FItemList(i).Fyyyymm,2) %>','<%= RIGHT(dateadd("d",-1,dateadd("m",1,oCExtJungsan.FItemList(i).Fyyyymm+"-01")),2) %>','D','','on');return false;"><font color="#AAAAAA"><%= FormatNumber(oCExtJungsan.FItemList(i).FmonthdlvDiffTNoExists,0) %></font></a></td>

	<td></td>
	<td><%= FormatNumber(oCExtJungsan.FItemList(i).FdiffITEMsum,0) %></td>
	<td><%= FormatNumber(oCExtJungsan.FItemList(i).FdiffDlvsum,0) %></td>

	<td align="center" ><%= oCExtJungsan.FItemList(i).FupdDt %></td>
	<td align="center" ></td>
</tr>
<% next %>
<% end if %>

<tr height="25" bgcolor="FFFFFF">
	<td colspan="<%=CHKIIF(diffType="S","21","20")%>" align="center">
	<% if (FALSE) then %>
		<% if oCExtJungsan.HasPreScroll then %>
		<a href="javascript:NextPage('<%= oCExtJungsan.StartScrollPage-1 %>')">[pre]</a>
		<% else %>
			[pre]
		<% end if %>

		<% for i=0 + oCExtJungsan.StartScrollPage to oCExtJungsan.FScrollCount + oCExtJungsan.StartScrollPage - 1 %>
			<% if i>oCExtJungsan.FTotalpage then Exit for %>
			<% if CStr(page)=CStr(i) then %>
			<font color="red">[<%= i %>]</font>
			<% else %>
			<a href="javascript:NextPage('<%= i %>')">[<%= i %>]</a>
			<% end if %>
		<% next %>

		<% if oCExtJungsan.HasNextScroll then %>
			<a href="javascript:NextPage('<%= i %>')">[next]</a>
		<% else %>
			[next]
		<% end if %>
	<% end if %>
	</td>
</tr>
</table>

<form name="frmAct" method="post" action="extJungsan_process.asp">
<input type="hidden" name="mode" value="">
<input type="hidden" name="sellsite" value="">
<input type="hidden" name="yyyymm" value="">
</form>

<%
set oCExtJungsan = Nothing
%>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/db3close.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->
