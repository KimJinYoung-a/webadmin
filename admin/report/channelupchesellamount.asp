<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 카테고리별브랜드매출집계
' Hieditor : 서동석 생성
'			2022.02.09 한용민 수정(구매유형 디비에서 가져오게 통합작업)
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->

<!-- #include virtual="/lib/classes/report/reportcls.asp"-->
<%
const Maxlines = 10

dim yyyy1,mm1,dd1,yyyy2,mm2,dd2, i, yyyymmdd1, yyymmdd2, fromDate,toDate, cdl,ordertype, oldlist, sitename, dispCate
dim vPurchasetype, mwdiv
	yyyy1 = request("yyyy1")
	mm1 = request("mm1")
	dd1 = request("dd1")
	yyyy2 = request("yyyy2")
	mm2 = request("mm2")
	dd2 = request("dd2")
	oldlist = request("oldlist")
	sitename = request("sitename")
	cdl = request("cdl")
	dispCate = requestCheckvar(request("disp"),16)
	ordertype = request("ordertype")
	vPurchasetype = request("purchasetype")
	mwdiv       = NullFillWith(request("mwdiv"),"")

if ordertype="" then ordertype="ea"
if (yyyy1="") then yyyy1 = Cstr(Year(now()))
if (mm1="") then mm1 = Cstr(Month(now()))
if (dd1="") then dd1 = Cstr(day(now()))
if (yyyy2="") then yyyy2 = Cstr(Year(now()))
if (mm2="") then mm2 = Cstr(Month(now()))
if (dd2="") then dd2 = Cstr(day(now()))

fromDate = DateSerial(yyyy1, mm1, dd1)
toDate = DateSerial(yyyy2, mm2, dd2+1)

dim ojumun
set ojumun = new CJumunMaster
	ojumun.FRectFromDate = fromDate
	ojumun.FRectToDate = toDate
	ojumun.FRectOrdertype = ordertype
	ojumun.FRectOldJumun = oldlist
	ojumun.FRectCD1 = cdl
	ojumun.Fsitename = sitename
	ojumun.FRectDispCate =  dispCate
	ojumun.FRectPurchasetype = vPurchasetype
	ojumun.FRectMwDiv = mwdiv

	'// 2014-08-27, skyer9
	if (DateDiff("m", fromDate, toDate)) > 1 then
		response.write "한번에 2달 이상을 검색할 수 없습니다."
		dbget.close()
		response.end
	end if
	
	ojumun.ChannelBrandSellrePort
%>
<style type="text/css">
	.tb tr {background-color:#fff; text-align:right;}
	.tb tr:nth-child(1) {text-align:right;}
	.tb tr:nth-child(2) {background-color:#EFBE00; text-align:center;}
	.tb td:first-child {text-align:left;}
</style>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript">


function TnBrandSellMount(v){
	window.open('upcheselllist2.asp?' + v, 'itemlist', 'width=700,height=660,scrollbars=yes,resizable=yes');
}

</script>

<!-- 검색 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get" action="Channelupchesellamount.asp">
<input type="hidden" name="showtype" value="showtype">
<input type="hidden" name="menupos" value="<%= menupos %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
	<td align="left">
		<input type="checkbox" name="oldlist" <% if oldlist="on" then response.write "checked" %> >6개월이전내역&nbsp;&nbsp;&nbsp;
		검색기간 :
		<% DrawDateBox yyyy1,yyyy2,mm1,mm2,dd1,dd2 %>
		&nbsp;&nbsp;
		<span style="width:230px;">사이트구분 : <% Drawsitename "sitename",sitename %></span>
		&nbsp;&nbsp;
		구매유형: 
		<% drawPartnerCommCodeBox true,"purchasetype","purchasetype",vPurchasetype,"" %>
		&nbsp;&nbsp;
		매입구분:<% Call DrawBrandMWUCombo("mwdiv",mwdiv) %>
	</td>	
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="검색" onClick="frm.submit();">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
		카테고리선택 :
		<% SelectBoxCategoryLarge cdl %>
		&nbsp;&nbsp;전시카테고리: <!-- #include virtual="/common/module/dispCateSelectBox.asp"-->
		&nbsp;&nbsp;
		<input type="radio" name="ordertype" value="ea" <% if ordertype="ea" then response.write "checked" %>>수량순
		<input type="radio" name="ordertype" value="totalprice" <% if ordertype="totalprice" then response.write "checked" %>>매출순
		<input type="radio" name="ordertype" value="totalgain" <% if ordertype="totalgain" then response.write "checked" %>>수익순
		</td>
	</td>
</tr>
</form>
</table>
<!-- 검색 끝 -->

<br>
<!-- 리스트 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="20">
		검색결과 : <b><%= ojumun.FTotalCount %></b>
		&nbsp;
		판매가 : <font color="red"><%= FormatNumber(ojumun.FTotalPrice,0) %></font>원&nbsp;
		매입가 : <font color="red"><%= FormatNumber(ojumun.FTotalBuyPrice,0) %></font>원&nbsp;
		총판매량 : <font color="red"><% =FormatNumber(ojumun.FTotalEA,0) %></font>개&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
	</td>
</tr>
<tr bgcolor="<%= adminColor("tabletop") %>" align="center">
	<td>브랜드</td>
	<td>온라인매입</td>
	<td>온라인위탁</td>
	<td>온라인업체배송</td>
	
	<% if (FALSE) then %><td width="500"></td><% end if %>
	<td >마진</td>
	<td>갯수합</td>
	<td>판매금액</td>
	<td>매입금액</td>
	<td>수익</td>
</tr>
<% if ojumun.FresultCount>0 then %>
<% for i=0 to ojumun.FresultCount-1 %>
<form action="" name="frmBuyPrc<%=i%>" method="get">			
<tr bgcolor="#ffffff" align="center">
	<td><a href="javascript:TnBrandSellMount('designer=<%= ojumun.FMasterItemList(i).Fmakerid %>&yyyy1=<% = yyyy1 %>&mm1=<% = mm1 %>&dd1=<% = dd1 %>&yyyy2=<% = yyyy2 %>&mm2=<% = mm2 %>&dd2=<% = dd2 %>&datetype=jumunil');"><%= ojumun.FMasterItemList(i).Fmakerid %></a></td>
	<td align="right">
		<% if ojumun.FMasterItemList(i).FM_margin<>"" and ojumun.FMasterItemList(i).FM_margin<>0 then %>
			<%= ojumun.FMasterItemList(i).FM_margin %>%
		<% end if %>
	</td>
	<td align="right">
		<% if ojumun.FMasterItemList(i).FW_margin<>"" and ojumun.FMasterItemList(i).FW_margin<>0 then %>
			<%= ojumun.FMasterItemList(i).FW_margin %>%
		<% end if %>
	</td>
	<td align="right">
		<% if ojumun.FMasterItemList(i).FU_margin<>"" and ojumun.FMasterItemList(i).FU_margin<>0 then %>
			<%= ojumun.FMasterItemList(i).FU_margin %>%
		<% end if %>
	</td>

	<% if (FALSE) then %>
		<td>
			<% if Not (IsNull(ojumun.FMasterItemList(i).Fselltotal)) then %>
				<div align="left"> <img src="/images/dot1.gif" height="4" width="<%= CLng((ojumun.FMasterItemList(i).Fselltotal/ojumun.maxt)*400) %>"></div><br>
				<div align="left"> <img src="/images/dot2.gif" height="4" width="<%= CLng((ojumun.FMasterItemList(i).Fsellcnt/ojumun.maxc)*400) %>"></div>
			<% end if %>
		</td>
	<% end if %>

	<td align="right"><% if ojumun.FMasterItemList(i).Fselltotal<>0 then %><%= 100-CLng(ojumun.FMasterItemList(i).Fbuytotal/ojumun.FMasterItemList(i).Fselltotal*100*100)/100 %> %<% end if %></td>
	<td align="right"><% if (ojumun.FTotalEA<>0) then %><%= ojumun.FMasterItemList(i).Fsellcnt & "건"%><% end if %></td>
	<td align="right"><% if (ojumun.FTotalPrice<>0) then %><%= FormatNumber(FormatCurrency(ojumun.FMasterItemList(i).Fselltotal),0) & "원"%><% end if %></td>
	<td align="right"><% if (ojumun.FTotalPrice<>0) then %><%= FormatNumber(FormatCurrency(ojumun.FMasterItemList(i).Fbuytotal),0) & "원"%><% end if %></td>
	<td align="right"><% if (ojumun.FTotalPrice-ojumun.FTotalBuyPrice<>0) then %><%= FormatNumber(FormatCurrency(ojumun.FMasterItemList(i).Fselltotal-ojumun.FMasterItemList(i).Fbuytotal),0) & "원"%><% end if %></td>
</tr>   
</form>
<% next %>
<% else %>
	<tr bgcolor="#FFFFFF">
		<td colspan="20" align="center" class="page_link">[검색결과가 없습니다.]</td>
	</tr>
<% end if %>

</table>

<%
'// 사이트구분로검색시작
Sub Drawsitename(selectboxname, sitename)
	dim userquery, tem_str

	response.write "<select name='" & selectboxname & "' class='select'>"
	response.write "<option value=''"
		if sitename ="" then
			response.write "selected"
		end if
	response.write ">전체</option>"

	'사용자 검색 옵션 내용 DB에서 가져오기
	userquery = " select id from [db_partner].[dbo].tbl_partner"
	userquery = userquery + " where 1=1"
	userquery = userquery + " and id <> '' and id is not null"
	userquery = userquery + " and userdiv= '999'" 		'/ and isusing='Y'
	userquery = userquery + " group by id"
	userquery = userquery + " order by id asc"

	'response.write userquery & "<Br>"
	rsget.Open userquery, dbget, 1

	if not rsget.EOF then
		do until rsget.EOF
			if Lcase(sitename) = Lcase(rsget("id")) then
				tem_str = " selected"
			end if

			response.write "<option value='" & rsget("id") & "' " & tem_str & ">" & rsget("id") & "</option>"
			tem_str = ""
			rsget.movenext
		loop
	end if
	rsget.close
	response.write "</select>"
End Sub
%>

<%
set ojumun = nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->