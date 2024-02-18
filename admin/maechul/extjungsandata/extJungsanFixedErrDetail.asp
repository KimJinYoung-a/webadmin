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
dim sellsite, searchfield, searchtext, jungsantype, errtp, accerrtype
dim yyyy1, mm1
Dim i

sellsite = requestCheckvar(request("sellsite"),32)
research = requestCheckvar(request("research"),10)
page 	 = requestCheckvar(request("page"),10)
yyyy1   = requestCheckvar(request("yyyy1"),4)
mm1     = requestCheckvar(request("mm1"),2)
jungsantype	= requestCheckvar(request("jungsantype"),10)

errtp	   = requestCheckvar(request("errtp"),10)
accerrtype = requestCheckvar(request("accerrtype"),10)

if (page="") then page=1


if (yyyy1="") then
	BaseDay = dateadd("m",-1,now())

	yyyy1 = Cstr(Year(BaseDay))
	mm1 = Cstr(Month(BaseDay))
end if



Dim oCExtJungsan
set oCExtJungsan = new CExtJungsan
	oCExtJungsan.FPageSize = 50
	oCExtJungsan.FCurrPage = page
	oCExtJungsan.FRectSellSite = sellsite
	oCExtJungsan.FRectyyyymm = yyyy1+"-"+mm1
	oCExtJungsan.FRectJungsanType = jungsantype
	oCExtJungsan.FRectAccerrtype = accerrtype
	oCExtJungsan.FRectErrorType = errtp

    oCExtJungsan.GetExtJungsanFixedErrDetailList
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

function jsExtJungsanDiffMake(sellsite) {
	var frm = document.frmAct;
return;
	if (confirm("재작성하시겠습니까?") == true) {
		frm.mode.value = "extjungsandiffmake";
		frm.sellsite.value = sellsite;
		frm.yyyymm.value = yyyymm;

		frm.submit();
	}
}

function popByExtorderserial(iextorderserial){
	var iUrl = "/admin/maechul/extjungsandata/extJungsanMapEdit.asp?menupos=<%=menupos%>&page=1&research=on";
	iUrl += "&sellsite=<%=sellsite%>"
	iUrl += "&searchfield=extOrderserial&searchtext="+iextorderserial;
	var popwin = window.open(iUrl,"extJungsanMapEdit","width=1400,height=800,crollbars=yes,resizable=yes,status=yes");

	popwin.focus();

}

function popByOrderserial(iorderserial,iitemid,iitemoption){
    var iurl = "/admin/maechul/extjungsandata/extJungsanFixedErrDetailByOrder.asp?orderserial="+iorderserial+"&itemid="+iitemid+"&itemoption="+iitemoption;
    var popwin = window.open(iurl,'extJungsanFixedErrDetailByOrder','width=1200 height=800 scrollbars=yes resizable=yes');
    popwin.focus();

}

function popJcomment(iorderserial,iitemid,iitemoption){
    var addcmt = "";
    addcmt = prompt("정산 comment", "");
    if (addcmt == null) return;

    if (addcmt.length<1){
        alert("코멘트를 작성해주세요.");
        return;
    }

    var frm = document.frmcmt;
    frm.mode.value="addcmt";
    frm.orderserial.value=iorderserial;
    frm.itemid.value=iitemid;
    frm.itemoption.value=iitemoption;
    frm.addcomment.value=addcmt;

    frm.submit();
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
		제휴몰:	<%= getJungsanXsiteComboHTML("sellsite",sellsite,"") %>
		&nbsp;
		&nbsp;
		상품구분:
        <input type="radio" name="jungsantype" value="" <% if (jungsantype = "") then %>checked<% end if %> > 전체
		<input type="radio" name="jungsantype" value="C" <% if (jungsantype = "C") then %>checked<% end if %> > 상품
		<input type="radio" name="jungsantype" value="D" <% if (jungsantype = "D") then %>checked<% end if %> > 배송비
		&nbsp;
		&nbsp;
		|
		&nbsp;
		&nbsp;
		오차반영여부:
		<input type="radio" name="errtp" value="" <% if (errtp = "") then %>checked<% end if %> > 전체
		<input type="radio" name="errtp" value="1" <% if (errtp = "1") then %>checked<% end if %> > 미반영
		<input type="radio" name="errtp" value="3" <% if (errtp = "3") then %>checked<% end if %> > 기반영(반영월기준)
		<input type="radio" name="errtp" value="2" <% if (errtp = "2") then %>checked<% end if %> > 기반영(매출월기준)

		
        
	</td>
	<td rowspan="4" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="검색" onClick="javascript:document.frm.submit();">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" height="30">
	<td align="left">
		&nbsp;
		매출월:
		<% DrawYMBox yyyy1,mm1 %>

		&nbsp;
		&nbsp;
		|
		&nbsp;
		&nbsp;
		누적오차 : 
		<input type="radio" name="accerrtype" value="" <%=CHKIIF(accerrtype="","checked","") %> >전체
		<!-- <input type="radio" name="accerrtype" value="5" <%=CHKIIF(accerrtype="5","checked","") %> >반영오차금액0제외 -->

		<input type="radio" name="accerrtype" value="1" <%=CHKIIF(accerrtype="1","checked","") %> >누적오차금액0제외
		<input type="radio" name="accerrtype" value="2" <%=CHKIIF(accerrtype="2","checked","") %> >누적오차반영합0제외
		
		&nbsp;|&nbsp;

		<input type="radio" name="accerrtype" value="3" <%=CHKIIF(accerrtype="3","checked","") %> >누적오차수량0제외(쿠폰오차제외)
		<input type="radio" name="accerrtype" value="4" <%=CHKIIF(accerrtype="4","checked","") %> >누적오차수량0만

		<!-- <input type="radio" name="accerrtype" value="4" <%=CHKIIF(accerrtype="4","checked","") %> >누적오차수량반영합0제외 -->
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
	<% if (FALSE) then %>
		<input type="button" class="button" value="재작성(<%= sellsite %>)" onClick="jsExtJungsanDiffMake('<%= sellsite %>');">
	<% end if %>
	</td>
</tr>
</table>
<!-- 액션 끝 -->

<p>

<!-- 리스트 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="6">
		검색결과 : <b><%= oCExtJungsan.FTotalcount %></b>
		&nbsp;
		페이지 : <b><%= page %> / <%= oCExtJungsan.FTotalPage %></b>

	</td>
	<td align="center"><b><%= FormatNumber(oCExtJungsan.FdiffnoSum,0) %></b></td>
	<td align="right"><b><%= FormatNumber(oCExtJungsan.FdiffsumSum,0) %></b></td>
	<td></td>
	<td align="right"><b><%= FormatNumber(oCExtJungsan.FErrAsignSum,0) %></b></td>
    <td></td>
    <td></td>
    <td></td>
    <td></td>
	<td></td>
    <td></td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width="100">제휴몰</td>
	<td width="100">매출일</td>
	<td width="120">제휴주문번호</td>
	<td width="110">TEN주문번호</td>
	
	<td width="80">상품코드</td>
	<td width="80">옵션코드</td>
	<td width="60">오차수량</td>
	<td width="110">오차합계</td>
    <td width="80">오차반영월</td>
	<td width="80">반영오차</td>
    <td width="4"></td>
	<td width="80">누적오차수량</td>
    <td width="80">누적오차</td>
    <td width="80">누적반영오차</td>
    <td width="80">누적오차<br>+반영오차 합</td>
	<td>비고</td>
</tr>
<% for i=0 to oCExtJungsan.FresultCount -1 %>
<tr align="center" bgcolor="FFFFFF" onmouseover=this.style.background="F1F1F1"; onmouseout=this.style.background="FFFFFF";>
	<td ><%= oCExtJungsan.FItemList(i).Fsellsite %></td>
	<td ><%= oCExtJungsan.FItemList(i).Fyyyymm %></td>
	<td ><a href="#" onClick="popByExtorderserial('<%= oCExtJungsan.FItemList(i).Foutmallorderserial %>'); return false;"><%= oCExtJungsan.FItemList(i).Foutmallorderserial %></a></td>
	<td ><a href="#" onclick="popByOrderserial('<%= oCExtJungsan.FItemList(i).Forderserial %>','<%= oCExtJungsan.FItemList(i).Fitemid %>','<%= oCExtJungsan.FItemList(i).Fitemoption %>');return false;"><%= oCExtJungsan.FItemList(i).Forderserial %></a></td>
	
	<td ><%= oCExtJungsan.FItemList(i).Fitemid %></td>
	<td ><%= oCExtJungsan.FItemList(i).Fitemoption %></td>

	<td><%= FormatNumber(oCExtJungsan.FItemList(i).getDiffNo,0) %></td>
	<td align="right"><%= FormatNumber(oCExtJungsan.FItemList(i).getDiffSum,0) %></td>

	<td align="center" ><%=oCExtJungsan.FItemList(i).FErrAsignMonth%></td>
    <td align="right" >
		<% if NOT isNULL(oCExtJungsan.FItemList(i).FErrAsignSum) then %>
		<%=FormatNumber(oCExtJungsan.FItemList(i).FErrAsignSum,0)%>
		<% end if %>
	</td>
    <td align="center" ></td>
	<td align="center" >
		<% if NOT isNULL(oCExtJungsan.FItemList(i).FaccErrNoSum) then %>
		<%=FormatNumber(oCExtJungsan.FItemList(i).FaccErrNoSum,0)%>
		<% end if %>
	</td>
    <td align="right" >
		<% if NOT isNULL(oCExtJungsan.FItemList(i).FacctErrsum) then %>
		<%=FormatNumber(oCExtJungsan.FItemList(i).FacctErrsum,0)%>
		<% end if %>
	</td>
    <td align="right" >
		<% if NOT isNULL(oCExtJungsan.FItemList(i).FaccAsgnErrSum) then %>
		<%=FormatNumber(oCExtJungsan.FItemList(i).FaccAsgnErrSum,0)%>
		<% end if %>
	</td>
    <td align="right" >
		<% if NOT isNULL(oCExtJungsan.FItemList(i).FaccTTLErrSum) then %>
		<%=FormatNumber(oCExtJungsan.FItemList(i).FaccTTLErrSum,0)%>
		<% end if %>
	</td>

	<td align="center" >
	<a href="#" onClick="popJcomment('<%=oCExtJungsan.FItemList(i).Forderserial%>','<%=oCExtJungsan.FItemList(i).Fitemid%>','<%=oCExtJungsan.FItemList(i).Fitemoption%>');return false;">
	<%=CHKIIF(isNULL(oCExtJungsan.FItemList(i).Fcomment),"<img src='/images/icon_new.gif' alt='코멘트작성'>",oCExtJungsan.FItemList(i).Fcomment)%>
	</a>
	</td>
</tr>
<% next %>


<tr height="25" bgcolor="FFFFFF">
	<td colspan="16" align="center">
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
	</td>
</tr>
</table>
<%
set oCExtJungsan = Nothing
%>

<form name="frmAct" method="post" action="extJungsan_process.asp">
<input type="hidden" name="mode" value="">
<input type="hidden" name="sellsite" value="">
<input type="hidden" name="yyyymm" value="">
</form>

<form name="frmcmt" method="post" action="extJungsan_process.asp">
<input type="hidden" name="mode" value="addcmt">
<input type="hidden" name="orderserial" value="">
<input type="hidden" name="itemid" value="">
<input type="hidden" name="itemoption" value="">
<input type="hidden" name="addcomment" value="">
<input type="hidden" name="rowidx" value="">
</form>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/db3close.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->
