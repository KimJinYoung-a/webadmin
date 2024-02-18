<%@ language=vbscript %>
<% option explicit %>
<%
	Response.AddHeader "Cache-Control","no-cache"
	Response.AddHeader "Expires","0"
	Response.AddHeader "Pragma","no-cache"
%>
<%
'###########################################################
' Description : 삽별구역설정
' Hieditor : 2010.12.29 한용민 생성
'###########################################################
%>
<!-- #include virtual="/common/incSessionAdminOrShop.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/common/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/offshop/zone/zone_cls.asp"-->

<%
Dim ozone,i,page , parameter , shopid ,yyyy1,mm1,dd1,yyyy2,mm2,dd2 ,fromDate ,toDate
dim designer , sellgubun ,cdl ,cdm ,cds ,datefg ,menupos , idx ,itemid ,itemname ,searchtype
dim totrealsellprice , totitemno ,totrealmaechul , totsumrealmaechul ,totshopsuplycash ,totsuplycashmaechul
dim totsumshopsuplycash , totprofit ,totsumprofit ,totrate , totsumrate
	designer = RequestCheckVar(request("designer"),32)
	page = requestCheckVar(request("page"),10)
	shopid = requestCheckVar(request("shopid"),32)
	yyyy1 = requestCheckVar(request("yyyy1"),4)
	mm1 = requestCheckVar(request("mm1"),2)
	dd1 = requestCheckVar(request("dd1"),2)
	yyyy2 = requestCheckVar(request("yyyy2"),4)
	mm2 = requestCheckVar(request("mm2"),2)
	dd2 = requestCheckVar(request("dd2"),2)
	sellgubun = requestCheckVar(request("sellgubun"),1)
	cdl     = requestCheckVar(request("cdl"),3)
	cdm     = requestCheckVar(request("cdm"),3)
	cds     = requestCheckVar(request("cds"),3)
	datefg = requestCheckVar(request("datefg"),10)
	idx = requestCheckVar(request("idx"),10)
	menupos = requestCheckVar(request("menupos"),10)
	itemid = requestCheckVar(request("itemid"),10)
	itemname = requestCheckVar(request("itemname"),124)
	searchtype = requestCheckVar(request("searchtype"),1)

if searchtype="" then searchtype="I"
if datefg = "" then datefg = "maechul"
if sellgubun = "" then sellgubun = "S"
if (yyyy1="") then yyyy1 = Cstr(Year(now()))
if (mm1="") then mm1 = Cstr(Month(now()))
if (dd1="") then dd1 = Cstr(day(now()))
if (yyyy2="") then yyyy2 = Cstr(Year(now()))
if (mm2="") then mm2 = Cstr(Month(now()))
if (dd2="") then dd2 = Cstr(day(now()))

fromDate = DateSerial(yyyy1, mm1, dd1)
toDate = DateSerial(yyyy2, mm2, dd2+1)

if page = "" then page = 1
if (searchtype="C") and ((cdl<>"") and (cdm<>"") and (cds<>"")) then cds=""

'직영/가맹점
if (C_IS_SHOP) then

	'/어드민권한 점장 미만
	if getlevel_sn("",session("ssBctId")) > 6 then
		shopid = C_STREETSHOPID
	end if
end if

set ozone = new czone_list
	ozone.FPageSize = 100
	ozone.FCurrPage = page
	ozone.frectshopid = shopid
	ozone.FRectStartDay = fromDate
	ozone.FRectEndDay = toDate
	ozone.FRectmakerid = designer
	ozone.FRectCDL = cdl
	ozone.FRectCDM = cdm
	ozone.FRectCDN = cds
	ozone.frectdatefg = datefg
	ozone.frectidx = idx
	ozone.frectitemid = itemid
	ozone.frectitemname = itemname
	ozone.frectsellgubun = sellgubun

	'/판매상품목록
	if searchtype="I" then
		ozone.Getoffshopzone_detail

	'/카테고리합계
	else
		ozone.Getoffshopzone_detailCategory
	end if

totrealsellprice = 0
totitemno = 0
totrealmaechul =0
totsumrealmaechul = 0
totshopsuplycash = 0
totsuplycashmaechul = 0
totsumshopsuplycash = 0
totprofit = 0
totsumprofit = 0
totrate = 0
totsumrate = 0

parameter = "designer="&designer&"&shopid="&shopid&"&yyyy1="&yyyy1&"&mm1="&mm1&"&dd1="&dd1&"&yyyy2="&yyyy2&"&mm2="&mm2&"&dd2="&dd2&"&sellgubun="&sellgubun
parameter = parameter & "&datefg="&datefg&"&idx="&idx&"&menupos="&menupos&"&itemid="&itemid&"&itemname="&itemname
%>

<script language="javascript">

	function gopage(page){
		frm.page.value=page;
		frm.submit();
	}

</script>

<!-- 검색 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get" action="">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="page" value="<%= page %>">
<input type="hidden" name="idx" value="<%= idx %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
	<td align="left">
		매장:<% drawSelectBoxOffShop "shopid",shopid %>
		매출기준 :
		<% drawmaechul_datefg "datefg" ,datefg ,""%>
		<% DrawDateBox yyyy1,yyyy2,mm1,mm2,dd1,dd2 %>
		<Br>
		<input type="radio" name="searchtype" value="I" <% if searchtype="I" then response.write "checked" %> >판매상품목록
		<input type="radio" name="searchtype" value="C" <% if searchtype="C" then response.write "checked" %> >카테고리합계
		<!-- #include virtual="/common/module/categoryselectbox.asp"-->
	</td>
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="검색" onClick="gopage('');">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
		브랜드 : <% drawSelectBoxDesignerwithName "designer",designer %>
		<input type="radio" name="sellgubun" value="S" <% if sellgubun="S" then response.write " checked" %>>결제내역기준
		<input type="radio" name="sellgubun" value="N" <% if sellgubun="N" then response.write " checked" %>>현재등록내역기준
		상품코드 : <input type="text" name="itemid" value="<%=itemid %>" size=10>
		상품명 : <input type="text" name="itemname" value="<%=itemname %>" size=20>
	</td>
</tr>
</form>
</table>
<!-- 검색 끝 -->
<br>

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

<!-- 리스트 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<%
'/판매상품목록
if searchtype="I" then
%>
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="15">
			검색결과 : <b><%= ozone.FTotalCount %></b>
			&nbsp;
			페이지 : <b><%= page %>/ <%= ozone.FTotalPage %></b>
		</td>
	</tr>
	<% if ozone.FresultCount>0 then %>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td>주문번호</td>
		<td>상품번호</td>
		<td>상품명(옵션명)</td>
		<td>브랜드</td>
		<td>대카테고리<br>중카테고리<br>소카테고리</td>
		<td>판매수량</td>
		<td>실판매가</td>
		<td>총<br>실매출액</td>
		<% if NOT(C_IS_SHOP) then %>
			<td>매입가</td>
			<td>총매입가</td>
			<td>매출수익</td>
			<td>마진율</td>
		<% end if %>
		<td>그룹</td>
		<td>매대타입</td>
		<td>상세구역명</td>
	</tr>
	<%
	for i=0 to ozone.FresultCount-1

	totrealmaechul = ozone.FItemList(i).frealsellprice * ozone.FItemList(i).fitemno
	totsumrealmaechul = totsumrealmaechul + totrealmaechul
	totitemno = totitemno + ozone.FItemList(i).fitemno
	totrealsellprice = totrealsellprice + ozone.FItemList(i).frealsellprice
	totshopsuplycash = totshopsuplycash + ozone.FItemList(i).fshopsuplycash
	totsuplycashmaechul = ozone.FItemList(i).fshopsuplycash * ozone.FItemList(i).fitemno
	totsumshopsuplycash = totsumshopsuplycash + totsuplycashmaechul
	totprofit = totrealmaechul-totsuplycashmaechul
	totsumprofit = totsumprofit + (totrealmaechul-totsuplycashmaechul)

	if totsuplycashmaechul <> 0 and totrealmaechul <> 0 then
		totrate = round(100-((totsuplycashmaechul)/(totrealmaechul)*100*100)/100,1)
	else
		totrate = 0
	end if

	totsumrate = totsumrate + totrate
	%>
	<% if ozone.FItemList(i).fzonename <> "" then %>
	<tr align="center" bgcolor="#FFFFFF" onmouseover=this.style.background="#f1f1f1"; onmouseout=this.style.background='#ffffff';>
	<% else %>
	<tr align="center" bgcolor="#FFFFaa" onmouseover=this.style.background="#f1f1f1"; onmouseout=this.style.background='#FFFFaa';>
	<% end if %>
		<td align="center">
			<%= ozone.FItemList(i).forderno %>
		</td>
		<td align="center">
			<%= ozone.FItemList(i).fitemgubun %>-<%= CHKIIF(ozone.FItemList(i).fshopitemid>=1000000,Format00(8,ozone.FItemList(i).fshopitemid),Format00(6,ozone.FItemList(i).fshopitemid)) %>-<%= ozone.FItemList(i).fitemoption %>
		</td>

		<td align="center">
			<%= ozone.FItemList(i).fitemname %>
			<% if ozone.FItemList(i).fitemoptionname <> "" then %>
				(<%= ozone.FItemList(i).fitemoptionname %>)
			<% end if %>
		</td>
		<td align="center">
			<%= ozone.FItemList(i).fmakerid %>
		</td>
		<td align="center">
			<% if ozone.FItemList(i).fcdl_nm <> "" then %>
				<%= ozone.FItemList(i).fcdl_nm %>
			<% else %>
				-
			<% end if %>
			<% if ozone.FItemList(i).fcdm_nm <> "" then %>
				<br><%= ozone.FItemList(i).fcdm_nm %>
			<% else %>
				<br>-
			<% end if %>
			<% if ozone.FItemList(i).fcds_nm <> "" then %>
				<br><%= ozone.FItemList(i).fcds_nm %>
			<% else %>
				<br>-
			<% end if %>
		</td>
		<td align="center">
			<%= ozone.FItemList(i).fitemno %>
		</td>
		<td align="center">
			<%= FormatNumber(ozone.FItemList(i).frealsellprice,0) %>
		</td>
		<td align="center">
			<%= FormatNumber(totrealmaechul,0) %>
		</td>
		<% if NOT(C_IS_SHOP) then %>
			<td>
				<%= FormatNumber(ozone.FItemList(i).fshopsuplycash,0) %>
			</td>
			<td>
				<%= FormatNumber(totsuplycashmaechul,0) %>
			</td>
			<td>
				<%= FormatNumber(totprofit,0) %>
			</td>
			<td>
				<%= FormatNumber(totrate,0) %>%
			</td>
		<% end if %>
		<td align="center">
			<% if ozone.FItemList(i).fzonegroup_name = "" or isnull(ozone.FItemList(i).fzonegroup_name) then %>
				-
			<% else %>
				<%= ozone.FItemList(i).fzonegroup_name %>
			<% end if %>
		</td>
		<td align="center">
			<% if ozone.FItemList(i).fracktype = "" or isnull(ozone.FItemList(i).fracktype) then %>
				-
			<% else %>
				<%= getOffShopracktype(ozone.FItemList(i).fracktype) %>
			<% end if %>
		</td>

		<td align="center">
			<% if ozone.FItemList(i).fzonename = "" or isnull(ozone.FItemList(i).fzonename) then %>
				-
			<% else %>
				<%= ozone.FItemList(i).fzonename %>
			<% end if %>
		</td>
	</tr>
	<% next %>
	<tr align="center" bgcolor="#FFFFFF">
		<td colspan=5>
			합계
		</td>
		<td>
			<%= FormatNumber(totitemno,0) %>
		</td>
		<td>
			<%= FormatNumber(totrealsellprice,0) %>
		</td>
		<td>
			<%= FormatNumber(totsumrealmaechul,0) %>
		</td>
		<% if NOT(C_IS_SHOP) then %>
			<td>
				<%= FormatNumber(totshopsuplycash,0) %>
			</td>
			<td>
				<%= FormatNumber(totsumshopsuplycash,0) %>
			</td>
			<td>
				<%= FormatNumber(totsumprofit,0) %>
			</td>
			<td>
				<% if totsuplycashmaechul <> 0 and totrealmaechul <> 0 then %>
					<%= FormatNumber(totsumrate / ozone.fresultcount,0) %>%
				<% else %>
					0%
				<% end if %>
			</td>
		<% end if %>
		<td align="center" colspan=3></td>
	</tr>
	<% else %>
		<tr bgcolor="#FFFFFF">
			<td colspan="15" align="center" class="page_link">[검색결과가 없습니다.]</td>
		</tr>
	<% end if %>
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="15" align="center">
	       	<% if ozone.HasPreScroll then %>
				<span class="list_link"><a href="javascript:gopage('<%= ozone.StartScrollPage-1 %>');">[pre]</a></span>
			<% else %>
			[pre]
			<% end if %>
			<% for i = 0 + ozone.StartScrollPage to ozone.StartScrollPage + ozone.FScrollCount - 1 %>
				<% if (i > ozone.FTotalpage) then Exit for %>
				<% if CStr(i) = CStr(ozone.FCurrPage) then %>
				<span class="page_link"><font color="red"><b><%= i %></b></font></span>
				<% else %>
				<a href="javascript:gopage('<%= i %>');" class="list_link"><font color="#000000"><%= i %></font></a>
				<% end if %>
			<% next %>
			<% if ozone.HasNextScroll then %>
				<span class="list_link"><a href="javascript:gopage('<%= i %>');">[next]</a></span>
			<% else %>
			[next]
			<% end if %>
		</td>
	</tr>
<% else %>
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="15">
			검색결과 : <b><%= ozone.FTotalCount %></b>
		</td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td>카데고리명</font></td>
		<td width="400" align="left">
			<img src="/images/dot1.gif" height="4" width=10>총판매수량
			<br><img src="/images/dot2.gif" height="4" width=10>총실매출
		</td>
		<td>점유율</td>
		<td>총<br>판매수량</td>
		<td>총<br>실매출</td>
		<% if NOT(C_IS_SHOP) then %>
			<td>총<br>매입가합</td>
			<td>매출수익</td>
			<td>마진율</td>
		<% end if %>
	</tr>
	<% if ozone.FtotalCount>0 then %>
	<%
	for i=0 to ozone.FtotalCount-1

	totsumrealmaechul = totsumrealmaechul + ozone.FItemList(i).frealmaechul
	totitemno = totitemno + ozone.FItemList(i).fitemnosum
	totsumshopsuplycash = totsumshopsuplycash + ozone.FItemList(i).fsuplymaechul
	totprofit = ozone.FItemList(i).frealmaechul-ozone.FItemList(i).fsuplymaechul
	totsumprofit = totsumprofit + (ozone.FItemList(i).frealmaechul-ozone.FItemList(i).fsuplymaechul)

	if ozone.FItemList(i).fsuplymaechul <> 0 and ozone.FItemList(i).frealmaechul <> 0 then
		totrate = round(100-((ozone.FItemList(i).fsuplymaechul)/ozone.FItemList(i).frealmaechul*100*100)/100,1)
	else
		totrate = 0
	end if

	totsumrate = totsumrate + totrate
	%>
	<tr align="center" bgcolor="#FFFFFF" onmouseover=this.style.background="#f1f1f1"; onmouseout=this.style.background='#ffffff';>
		<td>
			<% if (ozone.FItemList(i).FCateCDm="") or (ozone.FItemList(i).FCateCDs="") then %>
				<a href="?searchtype=C&cdl=<%= ozone.FItemList(i).FCateCDL %>&cdm=<%= ozone.FItemList(i).FCateCDM %>&cds=<%= ozone.FItemList(i).FCateCDs %>&<%=parameter%>"><%= ozone.FItemList(i).FCateName %></a>
			<% else %>
				<a href="?searchtype=I&cdl=<%= ozone.FItemList(i).FCateCDL %>&cdm=<%= ozone.FItemList(i).FCateCDM %>&cds=<%= ozone.FItemList(i).FCateCDs %>&<%=parameter%>"><%= ozone.FItemList(i).FCateName %></a>
			<% end if %>
		</td>
		<td height="10" width="400">
			<% if  (ozone.FItemList(i).frealmaechul<>0) then %>
				<div align="left">
					<img src="/images/dot1.gif" height="4" width="<%= CLng((ozone.FItemList(i).frealmaechul/ozone.maxt)*400) %>">
				</div>
				<br><div align="left">
					<img src="/images/dot2.gif" height="4" width="<%= CLng((ozone.FItemList(i).fitemnosum/ozone.maxc)*400) %>">
				</div>
			<% end if %>
		</td>
		<td>
			<% if ozone.FSumTotal<>0 then %>
				<%= Clng( ((ozone.FItemList(i).frealmaechul / ozone.FSumTotal) * 10000)) / 100 %> %
			<% end if %>
		</td>
		<td><%= ozone.FItemList(i).Fitemnosum %></td>
		<td>
			<%= FormatNumber(ozone.FItemList(i).frealmaechul,0) %>
		</td>
		<% if NOT(C_IS_SHOP) then %>
			<td><%= FormatNumber(ozone.FItemList(i).fsuplymaechul,0) %></td>
			<td><%= FormatNumber(totprofit,0) %></td>
			<td>
				<%= totrate %>%
			</td>
		<% end if %>
	</tr>
	<% next %>
	<tr bgcolor="#FFFFFF" align="center">
		<td colspan=3></td>
		<td><%= FormatNumber(totitemno,0) %></td>
		<td><%= FormatNumber(totsumrealmaechul,0) %></td>
		<td><%= FormatNumber(totsumshopsuplycash,0)%></td>
		<td><%= FormatNumber(totsumprofit,0) %></td>
		<td><%= round(totsumrate/ozone.FtotalCount,0) %>%</td>
	</tr>
	<% else %>
		<tr bgcolor="#FFFFFF">
			<td colspan="15" align="center" class="page_link">[검색결과가 없습니다.]</td>
		</tr>
	<% end if %>
<% end if %>
</table>

<%
set ozone = nothing
%>
<!-- #include virtual="/common/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->