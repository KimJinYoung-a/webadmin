<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/jungsan/ipkumlistcls.asp"-->
<%

dim yyyy1,mm1,dd1
dim yyyy2,mm2,dd2
dim fromDate,toDate
dim ckdate,tenbank,ipkumname,page
dim searchtype01,orderby,research,itype

ckdate = request.form("ckdate")
tenbank = request.form("tenbank")
ipkumname = request.form("ipkumname")
page = request("page")
searchtype01 =  request.form("searchtype01")
orderby = request.form("orderby")
research = request.form("research")
itype = request.form("itype")

if page="" then page=1

yyyy1 = request.form("yyyy1")
mm1 = request.form("mm1")
dd1 = request.form("dd1")
yyyy2 = request.form("yyyy2")
mm2 = request.form("mm2")
dd2 = request.form("dd2")

if (yyyy1="") then yyyy1 = Cstr(Year(now()))
if (mm1="") then mm1 = Cstr(Month(now()))
if (dd1="") then dd1 = Cstr(day(now())-1)
if (yyyy2="") then yyyy2 = Cstr(Year(now()))
if (mm2="") then mm2 = Cstr(Month(now()))
if (dd2="") then dd2 = Cstr(day(now())-1)
	
fromDate = CStr(DateSerial(yyyy1, mm1, dd1))
toDate = CStr(DateSerial(yyyy2, mm2, dd2+1))

dim ipkum,i,ix
set ipkum = new IpkumChecklist

ipkum.FCurrpage=page
ipkum.FPagesize=5000
ipkum.FScrollCount = 10
ipkum.Fckdate = ckdate
ipkum.Ctenbank = tenbank

ipkum.FSearchtype01 = searchtype01
ipkum.FOrderby = orderby
ipkum.ipkumname = ipkumname

ipkum.FRectRegStart = fromDate
ipkum.FRectRegEnd = toDate

ipkum.GetipkumlistAccounts
%>


<%
if itype="xl" then
response.ContentType = "application/vnd.ms-excel"
Response.AddHeader "Content-Disposition", "attachment; filename=&nbsp;"
end if
%>

<% if itype="xl" then %>
	<html xmlns:x="urn:schemas-microsoft-com:office:excel">
	<head>
		<meta http-equiv=Content-Type content="text/html; charset=ks_c_5601-1987">
		<style>
		.a
		{	
		vertical-align:middle;
		border:0.5pt solid black;
		font-size:9.0pt;
		}
		</style>
	</head>
	<body>
<% else %>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<script language="JavaScript" src="/js/xl.js"></script>
		<script language="JavaScript" src="/js/common.js"></script>
		<script language="JavaScript" src="/js/report.js"></script>
		<link rel="stylesheet" href="/css/scm.css" type="text/css">
	</head>
	<body>
<% end if %>

<% if ipkum.FTotalCount <= ipkum.FPagesize then %>
	
<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
	<tr align="center" bgcolor="<%= adminColor("topbar") %>">
		<td class="a" width="80">은행</td>
    	<td class="a" width="120">계좌번호</td>
    	<td class="a" width="80">검색어</td>
    	<td class="a" width="80">기간(FROM)</td>
    	<td class="a" width="80">기간(TO)</td>
    	<td class="a" width="120">정렬순서</td>
    	<td class="a" width="100">검색건수</td>
    </tr>
    <tr align="center" bgcolor="#FFFFFF">
		<td class="a"><b><%= ipkum.Fipkumitem(0).Fbkname %></b></td>
    	<td class="a"><b><%= tenbank %></b></td>
    	<td class="a"><% if searchtype01<>"" then %><b><%= ipkumname %></b><% end if %></td>
    	<td class="a"><b><%= CStr(DateSerial(yyyy1, mm1, dd1)) %></b></td>
    	<td class="a"><b><%= CStr(DateSerial(yyyy2, mm2, dd2)) %></b></td>
    	<td class="a"><% if orderby<>"" then %><b>최근일순</b><% else %><b>과거일순</b><% end if %></td>
    	<td class="a"><b><%= FormatNumber(ipkum.FTotalCount,0) %></b></td>
    </tr>
    <tr align="center" bgcolor="<%= adminColor("topbar") %>">
    	<td class="a">입출금일</td>
    	<td class="a">출금금액</td>
    	<td class="a">입금금액</td>
    	<td class="a">잔액</td>
    	<td class="a">거래구분</td>
    	<td class="a">거래내용</td>
    	<td class="a">비고(은행지점)</td>
    </tr>
    
    <% if ipkum.FResultCount<1 then %>
    <% else %>
    <% for i=0 to ipkum.FResultCount-1 %>
    <tr align="center" bgcolor="#FFFFFF">
    	<td class="a"><%= mid(ipkum.Fipkumitem(i).Fbkdate,1,4) %>-<%= mid(ipkum.Fipkumitem(i).Fbkdate,5,2) %>-<%= mid(ipkum.Fipkumitem(i).Fbkdate,7,2) %></td>
    	<td class="a" align="right">
 			<% if ipkum.Fipkumitem(i).finout_gubun = "1" then %>    	
    			<%= FormatNumber(ipkum.Fipkumitem(i).Fbkinput,0) %>
    		<% end if %>
    	</td>
	  	<td class="a" align="right">
 			<% if ipkum.Fipkumitem(i).finout_gubun = "2" then %>  	  	
	  			<%= FormatNumber(ipkum.Fipkumitem(i).Fbkinput,0) %>
    		<% end if %>
	  	</td>
    	<td class="a" align="right"><%= FormatNumber(ipkum.Fipkumitem(i).Fbkjango,0) %></td>
    	<td class="a"><%= ipkum.Fipkumitem(i).Fbkcontent      %></td>    	
    	<td class="a"><%= ipkum.Fipkumitem(i).Fbkjukyo        %></td>
    	<td class="a"><%= ipkum.Fipkumitem(i).Fbketc          %></td>
    </tr>
    <% next %>
    <% end if %>
</table>

<% else %>
 
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
	<tr align="center" bgcolor="#FFFFFF">
		<td>
			총검색건수가 <%= ipkum.FPagesize %>건을 초과하였습니다.<br>
			<%= ipkum.FPagesize %>건 이하로 검색 후 사용하세요.
		</td>
	</tr>
</table>

<% end if %>
	


<% set ipkum=nothing %> 

<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
