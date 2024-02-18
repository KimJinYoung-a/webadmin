<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/datamart/brandcssummaryclass.asp"-->

<%


dim ck_date
dim yyyy1,mm1,dd1, yyyy2,mm2,dd2
dim makerid, isupchebeasong, divcd, gubunNm, gubun02Arr
dim startdateStr, nextdateStr
dim page

ck_date     = RequestCheckVar(request("ck_date"),9)
yyyy1       = RequestCheckVar(request("yyyy1"),4)
mm1         = RequestCheckVar(request("mm1"),2)
dd1         = RequestCheckVar(request("dd1"),2)
yyyy2       = RequestCheckVar(request("yyyy2"),4)
mm2         = RequestCheckVar(request("mm2"),2)
dd2         = RequestCheckVar(request("dd2"),2)

makerid         = RequestCheckVar(request("makerid"),32)
isupchebeasong  = RequestCheckVar(request("isupchebeasong"),9)
divcd           = RequestCheckVar(request("divcd"),4)

gubunNm         = RequestCheckVar(request("gubunNm"),32)
page            = RequestCheckVar(request("page"),9)

if (yyyy1="") then yyyy1 = Cstr(Year(now()))
if (mm1="") then mm1 = Cstr(Month(now()))
if (dd1="") then dd1 = Cstr(day(now()))
if (yyyy2="") then yyyy2 = Cstr(Year(now()))
if (mm2="") then mm2 = Cstr(Month(now()))
if (dd2="") then dd2 = Cstr(day(now()))
if (page="") then page=1
if (gubunNm<>"") then gubun02Arr="'" + Replace(gubunNm,",","','") + "'"
startdateStr = yyyy1 + "-" + mm1 + "-" + dd1
nextdateStr = Left(CStr(DateAdd("d",Cdate(yyyy2 + "-" + mm2 + "-" + dd2),1)),10)

dim obrandCs
set obrandCs = new CBrandCSSummary
obrandCs.FPageSize = 50
obrandCs.FCurrPage = page
obrandCs.FRectMakerid   = makerid
obrandCs.FRectStartDate = startdateStr
obrandCs.FRectEndDate   = nextdateStr
obrandCs.FRectIsUpchebeasong = isupchebeasong
obrandCs.FRectDivCd     = divcd
obrandCs.FRectgubun02Arr = gubun02Arr

if (makerid<>"") then
    obrandCs.getBrandCsSUMList
end if

dim i
%>
<script language='javascript'>
function goPage(page){
    frm.page.value = page;
    frm.submit();
}

function fnSearch(frm) {
    frm.ck_date.disabled = false;
    frm.submit();
}


</script>
<!-- 표 상단바 시작-->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frm" method="get"   >
	<input type="hidden" name="menupos" value="<%=menupos%>">
	<input type="hidden" name="page" value="1">
	<input type="hidden" name="research" value="on">
  	<tr align="center" bgcolor="#FFFFFF" >
		<td width="50" bgcolor="<%= adminColor("gray") %>" rowspan="2">검색<br>조건</td>
		<td align="left">
			<table border="0" width="100%" cellpadding="3" cellspacing="0" class="a">
			<tr>
				<td>
    				<input type="checkbox" name="ck_date" checked disabled >
    				처리완료일 : <% DrawDateBox yyyy1,yyyy2,mm1,mm2,dd1,dd2 %>
    				&nbsp;&nbsp;
    				브랜드: <% drawSelectBoxDesignerwithName "makerid", makerid %>


				</td>
			</tr>
			</table>
        </td>
		<td  width="50" bgcolor="<%= adminColor("gray") %>" rowspan="2">
		    <input type="button" class="button_s" value="검색" onClick="fnSearch(document.frm)">
	    </td>
	</tr>
	<tr bgcolor="#FFFFFF">
	    <td align="left">
	        <table border="0" width="100%" cellpadding="3" cellspacing="0" class="a">
			<tr>
				<td>
				    배송구분:
    				<select name="isupchebeasong">
    				<option value="">전체
    				<option value="Y" <%= chkIIF(isupchebeasong="Y","selected","") %> >업체배송
    				<option value="N" <%= chkIIF(isupchebeasong="N","selected","") %> >텐배
    				</select>
    				&nbsp;&nbsp;

    				구분:
    				<select name="divcd">
    				<option value="">전체
    				<option value="A008" <%= chkIIF(divcd="A008","selected","") %> >주문취소(품절,출고지연)
    				<option value="A004" <%= chkIIF(divcd="A004","selected","") %> >반품(업체배송)
    				<option value="T012" <%= chkIIF(divcd="T012","selected","") %> >누락/서비스/맞교환
    				<option value="A000" <%= chkIIF(divcd="A000","selected","") %> >맞교환
    				<option value="A001" <%= chkIIF(divcd="A001","selected","") %> >누락재발송
    				<option value="A002" <%= chkIIF(divcd="A002","selected","") %> >서비스발송
    				</select>
    				&nbsp;&nbsp;

    				상세사유 :
    				<select name="gubunNm">
    				<option value="">전체
    				<option value="CD05,CF05" <%= chkIIF(gubunNm="CD05,CF05","selected","") %> >품절
    				<option value="CF06,CG01" <%= chkIIF(gubunNm="CE01","selected","") %> >상품불량
    				<option value="CE01,CE02" <%= chkIIF(gubunNm="CF01","selected","") %> >오발송
    				<option value="CF03,CF04,CF01" <%= chkIIF(gubunNm="CF02","selected","") %> >상품파손
    				<option value="CE04,CE03" <%= chkIIF(gubunNm="CF03,CF04","selected","") %> >상품누락
    				<option value="CF02,CG02,CG03" <%= chkIIF(gubunNm="CF06,CG01","selected","") %> >출고지연
    				<option value="CD01,CB04" <%= chkIIF(gubunNm="CD01,CB04","selected","") %> > 고객변심
    				</select>
    		    </td>
    		</tr>
    		</table>
	    </td>
	</tr>
	</form>
</table>
<!-- 표 상단바 끝-->

<p>

* 최대 100개의 상품만 표시됩니다.

<p>

<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="#BABABA">
<tr bgcolor="#FFFFFF">
    <td colspan="11" align="right">총 <%= obrandCs.FTotalCount %> 건 <%= page %>/<%= obrandCs.FTotalPage %> &nbsp;</td>
</tr>
<tr align="center" bgcolor="#E6E6E6">
	<!-- td width="60">ID</td -->
	<td width="100">구분</td>
	<td width="80">사유</td>
	<td width="50">상품코드</td>
	<td width="150">상품명</td>
	<td width="40">갯수</td>
	<td width="40">배송<br>구분</td>
	<td ></td>
</tr>
<% if obrandCs.FResultCount<1 then %>
<tr bgcolor="#FFFFFF">
    <td align="center" colspan="11">
        <% if (makerid="") then %>
            <font color="blue">[브랜드 ID를 선택하세요.]</font>
        <% else  %>
            [검색 결과가 없습니다.]
        <% end if %>
    </td>
</tr>
<% else %>
<% for i=0 to obrandCs.FREsultCount -1 %>
<tr bgcolor="#FFFFFF">
    <!--td ><%= obrandCs.FItemList(i).FID %></td -->
    <td align="center"><%= obrandCs.FItemList(i).Fdivcd_Name %></td>
    <td align="center"><%= obrandCs.FItemList(i).Fgubun02_Name %></td>
    <td align="center"><%= obrandCs.FItemList(i).FItemID %></td>
    <td ><%= obrandCs.FItemList(i).FItemName %>
    </td>
    <td align="center"><%= obrandCs.FItemList(i).FconfirmItemNo %></td>
    <td align="center"><%= obrandCs.FItemList(i).FIsUpchebeasong %></td>
    <td ></td>
</tr>
<% next %>
<% end if %>
</table>

<script language='javascript'>
function chkComp(bool){
    document.frm.yyyy1.disabled = !(bool);
    document.frm.mm1.disabled = !(bool);
    document.frm.dd1.disabled = !(bool);

    document.frm.yyyy2.disabled = !(bool);
    document.frm.mm2.disabled = !(bool);
    document.frm.dd2.disabled = !(bool);

}

function getOnload(){
    chkComp(<%=ChkIIF(ck_date="on","true","false") %>);
}

window.onload = getOnload;
</script>

<%
set obrandCs = Nothing
%>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
