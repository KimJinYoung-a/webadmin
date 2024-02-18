<%@ language=vbscript %>
<% option explicit %>
<%
	Response.AddHeader "Cache-Control","no-cache"
	Response.AddHeader "Expires","0"
	Response.AddHeader "Pragma","no-cache"
%>
<%
'###########################################################
' Description : 조닝별매출
' Hieditor : 2010.12.29 한용민 생성
'###########################################################
%>
<!-- #include virtual="/common/incSessionAdminOrShop.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/common/lib/commonbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/offshop/zone2/zone_cls.asp"-->
<%
Dim ozone,i,page , parameter , shopid ,yyyy1,mm1,dd1,yyyy2,mm2,dd2 ,fromDate ,toDate
dim designer , sellgubun ,datefg, zoneidx, dategubun, inc3pl
dim totsellsum ,totunitmaechul , totunit
dim tmpselldate
	designer = RequestCheckVar(request("designer"),32)
	page = requestCheckVar(request("page"),10)
	shopid = requestCheckVar(request("shopid"),32)
	yyyy1 = requestCheckVar(request("yyyy1"),4)
	mm1 = requestCheckVar(request("mm1"),2)
	dd1 = requestCheckVar(request("dd1"),2)
	yyyy2 = requestCheckVar(request("yyyy2"),4)
	mm2 = requestCheckVar(request("mm2"),2)
	dd2 = requestCheckVar(request("dd2"),2)
	sellgubun = requestCheckVar(request("sellgubun"),10)
	datefg = requestCheckVar(request("datefg"),10)
	zoneidx = requestCheckVar(request("zoneidx"),10)
	dategubun = requestCheckVar(request("dategubun"),10)
    inc3pl = requestCheckVar(request("inc3pl"),32)
    
if dategubun = "" then dategubun = "G"
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

'/매장
if (C_IS_SHOP) then
	
	'//직영점일때
	if C_IS_OWN_SHOP then
		
		'/어드민권한 점장 미만
		'if getlevel_sn("",session("ssBctId")) > 6 then
			'shopid = C_STREETSHOPID		'"streetshop011"
		'end if		
	else
		shopid = C_STREETSHOPID
	end if
else
	'/업체
	if (C_IS_Maker_Upche) then
		designer = session("ssBctID")	'"7321"

	else
		if (Not C_ADMIN_USER) then
		    shopid = "X"                ''다른매장조회 막음.
		else
		end if
	end if
end if	

set ozone = new czone_list
	ozone.FPageSize = 500
	ozone.FCurrPage = page
	ozone.frectshopid = shopid
	ozone.FRectStartDay = fromDate
	ozone.FRectEndDay = toDate
	ozone.FRectmakerid = designer	
	ozone.frectdatefg = datefg
	ozone.frectsellgubun = sellgubun
	ozone.frectdategubun = dategubun
	ozone.frectzoneidx = zoneidx
	ozone.FRectInc3pl = inc3pl
	
	if shopid <> "" then
		ozone.Getoffshopzonesum_category

		if drawnewipgobrand(shopid) <> "" then
			response.write "<script language='javascript'>"
			response.write "	alert('"&shopid&" 매장에 최근 3개월내에 조닝에 설정되지 않은 신규브랜드가 있습니다\n\n"&drawnewipgobrand(shopid)&"');"
			response.write "</script>"
		end if
	end if

parameter = "shopid="&shopid&"&sellgubun="&sellgubun&"&menupos="&menupos&"&designer="&designer&"&dategubun="&dategubun&"&inc3pl="&inc3pl
%>

<script language="javascript">

//조닝별 목표 매출 등록
function regtargetmaechul(shopid,yyyy,mm,gubuntype){
	var regtargetmaechul = window.open('/common/offshop/maechul/targetmaechul/targetmaechul_sub.asp?shopid='+shopid+'&yyyy1='+yyyy+'&mm1='+mm+'&gubuntype='+gubuntype,'regtargetmaechul','width=1024,height=768,scrollbars=yes,resizable=yes');
	regtargetmaechul.focus();
}

//카테고리매출
function category_detail(zoneidx,yyyy1,mm1,dd1,yyyy2,mm2,dd2){
	var category_detail = window.open('zone_sum_category_detail.asp?zoneidx='+zoneidx+'&yyyy1='+yyyy1+'&mm1='+mm1+'&dd1='+dd1+'&yyyy2='+yyyy2+'&mm2='+mm2+'&dd2='+dd2+'&<%=parameter%>','category_detail','width=1024,height=768,scrollbars=yes,resizable=yes');
	category_detail.focus();
}

//브랜드매출
function brand_detail(zoneidx,yyyy1,mm1,dd1,yyyy2,mm2,dd2){
	var brand_detail = window.open('zone_sum_brand_detail.asp?zoneidx='+zoneidx+'&yyyy1='+yyyy1+'&mm1='+mm1+'&dd1='+dd1+'&yyyy2='+yyyy2+'&mm2='+mm2+'&dd2='+dd2+'&<%=parameter%>','brand_detail','width=1024,height=768,scrollbars=yes,resizable=yes');
	brand_detail.focus();
}

//상품매출
function item_detail(zoneidx,yyyy1,mm1,dd1,yyyy2,mm2,dd2){
	var item_detail = window.open('zone_sum_item_detail.asp?zoneidx='+zoneidx+'&yyyy1='+yyyy1+'&mm1='+mm1+'&dd1='+dd1+'&yyyy2='+yyyy2+'&mm2='+mm2+'&dd2='+dd2+'&<%=parameter%>','item_detail','width=1024,height=768,scrollbars=yes,resizable=yes');
	item_detail.focus();
}

function frmsubmit(){
	frm.submit();
}

function divch(divid,zoneidx){
	frmdiv.divid.value = divid;
	frmdiv.zoneidx.value = zoneidx;
	frmdiv.target="view";
	frmdiv.action='/admin/offshop/zone2/zone_manager_search.asp';
	frmdiv.submit();
}

</script>

<!-- 검색 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frmdiv" method="get" action="">
	<input type="hidden" name="divid">
	<input type="hidden" name="zoneidx">
</form>
<form name="frm" method="post" action="">
<input type="hidden" name="menupos" value="<%= menupos %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
	<td align="left">
		* 기간 : <% drawmaechul_datefg "datefg" ,datefg ,""%>
		<% DrawDateBox yyyy1,yyyy2,mm1,mm2,dd1,dd2 %>
		&nbsp;&nbsp;	
		<%
		'직영/가맹점
		if (C_IS_SHOP) then
		%>	
			<% if not C_IS_OWN_SHOP and shopid <> "" then %>
				* 매장 : <%=shopid%><input type="hidden" name="shopid" value="<%= shopid %>">
			<% else %>
				* 매장 : <% drawSelectBoxOffShop "shopid",shopid %>
			<% end if %>
		<% else %>
			* 매장 : <% drawSelectBoxOffShop "shopid",shopid %>
		<% end if %>
	</td>	
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="검색" onClick="frmsubmit();">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
		* 브랜드 : <% drawSelectBoxDesignerwithName "designer",designer %>
		&nbsp;&nbsp;				
		<% Call zoneselectbox(shopid,"zoneidx",zoneidx," onchange='frmsubmit();'") %>
		&nbsp;&nbsp;
		<input type="radio" name="dategubun" value="G" <% if dategubun="G" then response.write " checked" %> onclick="frmsubmit();">기간별통계
		<input type="radio" name="dategubun" value="M" <% if dategubun="M" then response.write " checked" %> onclick="frmsubmit();">월별통계
		<input type="radio" name="dategubun" value="D" <% if dategubun="D" then response.write " checked" %> onclick="frmsubmit();">일별통계
        &nbsp;&nbsp;
        <b>* 매출처구분</b>
        <% Call draw3plMeachulComboBox("inc3pl",inc3pl) %>			
	</td>
</tr>
</table>
<!-- 검색 끝 -->
<br>

<% If shopid = "" Then %>
	<center><font color="red"><b>※ ShopID(매장)를 선택하셔야 데이터가 나타납니다.</b></font></center><br>
<% End If %>

<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a">
<tr>
	<td align="left">
		<Br><font color="red">-결제내역기준</font>
		<br> &nbsp; &nbsp; 매장에서 결제가 되면, 해당 브랜드가 등록된 조닝이 저장 되며, 판매당시 등록되어져 있던 조닝을 기준으로 통계가 산출됩니다.
		<br> &nbsp; &nbsp; 그러므로 결제 당시 브랜드가 조닝에 등록되어 있지 않으면, 통계가 남지 않습니다.
		<br><font color="red">-현재등록기준</font>
		<br> &nbsp; &nbsp; 결제당시 저장된 브랜드 조닝내역과 무관하게, 현재 브랜드가 등록된 조닝을 기준으로 리스트가 보여 짐니다.
	</td>
	<td valign="bottom" align="right">
		<input type="radio" name="sellgubun" value="S" <% if sellgubun="S" then response.write " checked" %> onclick="frmsubmit();">결제내역기준
		<input type="radio" name="sellgubun" value="N" <% if sellgubun="N" then response.write " checked" %> onclick="frmsubmit();">현재등록기준		
	</td>
</tr>
</form>
</table>
<!-- 액션 끝 -->

<iframe id="view" name="view" src="" width=0 height=0 frameborder="0" scrolling="no" ></iframe>

<!-- 리스트 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="15">
		검색결과 : <b><%= ozone.FTotalCount %></b>
		※ 500건 까지 검색됩니다.
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<% if dategubun = "D" or dategubun = "M" then %>
		<td>날짜</td>
	<% end if %>	
	<td>매장</td>
	<td>조닝명</td>
	<td>매장내<br>담당자</td>
	<td>총<br>매출액</td>
	<td>총매출액<br>점유율</td>
	
	<% if C_ADMIN_USER or C_IS_OWN_SHOP then %>
		<% if not(C_IS_Maker_Upche) then %>
			<% if dategubun = "M" or dategubun = "D" then %>
				<td>달성율</td>
			<% end if %>
		<% end if %>	
	<% end if %>
		
	<td>평당<br>매출액</td>
	<td>조닝<br>크기</td>
	<td>조닝<br>점유율</td>
	<td>비고</td>
</tr>
<%
totsellsum = 0
totunitmaechul= 0 
totunit = 0	
tmpselldate = ""

if ozone.FTotalCount>0 then

for i=0 to ozone.FTotalCount-1 

if tmpselldate <> ozone.FItemList(i).fIXyyyymmdd and i <> 0 then
%>
	<tr align="center" bgcolor="#f1f1f1">
		<% if dategubun = "D" or dategubun = "M" then %>
			<td colspan=4>
				<%= tmpselldate %> 총합계
			</td>
		<% else %>
			<td colspan=3>
				<%= tmpselldate %> 총합계
			</td>
		<% end if %>		
		<td>
			<%= FormatNumber(totsellsum,0) %>
		</td>
		<td></td>
		
		<% if C_ADMIN_USER or C_IS_OWN_SHOP then %>
			<% if not(C_IS_Maker_Upche) then %>
				<% if dategubun = "M" or dategubun = "D" then %>
					<td>
					</td>	
				<% end if %>
			<% end if %>
		<% end if %>	
		
		<td>
			<%= FormatNumber(totunitmaechul,0) %>
		</td>
		<td>
			<%= totunit %>
		</td>
		<td colspan=2></td>
	</tr>
<%
	totsellsum = 0
	totunitmaechul= 0 
	totunit = 0	
	tmpselldate = ""
end if

tmpselldate = ozone.FItemList(i).fIXyyyymmdd
totsellsum = totsellsum + ozone.FItemList(i).fsellsum

if ozone.FItemList(i).fsellsum <> 0 and ozone.FItemList(i).funit <> 0 then
	totunitmaechul = totunitmaechul + (ozone.FItemList(i).fsellsum / ozone.FItemList(i).funit)
end if

totunit = totunit + ozone.FItemList(i).funit
%>

<%' if ozone.FItemList(i).fzonename <> "" or isnull(ozone.FItemList(i).fzonename) then %>
	<tr align="center" bgcolor="#FFFFFF" onmouseover=this.style.background="#f1f1f1"; onmouseout=this.style.background='#ffffff';>
<%' else %>
	<!--<tr align="center" bgcolor="silver" onmouseover=this.style.background="#f1f1f1"; onmouseout=this.style.background='silver';>-->
<%' end if %>
	<% if dategubun = "D" or dategubun = "M" then %>
		<td>
			<%= ozone.FItemList(i).fIXyyyymmdd %>
		</td>
	<% end if %>	
	<td>
		<%= ozone.FItemList(i).fshopid %>
	</td>
	<td>
		<% if ozone.FItemList(i).fzonename <> "" then %>
			<%= ozone.FItemList(i).fzonename %>
		<% else %>
			-
		<% end if %>
	</td>
	<td>
		<% if ozone.FItemList(i).fmanagershopyn = "Y" then %>
			<div name="div<%=i%>" id="div<%=i%>">
				<img src="/images/icon_search.jpg" onmouseover="javascript:divch('div<%=i%>','<%=ozone.FItemList(i).fidx%>');">
			</div>
		<% end if %>
	</td>	
	<td bgcolor="#E6B9B8">
		<%= FormatNumber(ozone.FItemList(i).fsellsum,0) %>
	</td>	
	<td>
		<%
		'/중간합계의 해당 배열값을 가져와 점유율을 계산한다
		if (split(ozone.ftmpSumTotal,",")(ozone.FItemList(i).fblock))<>0 and ozone.FItemList(i).fsellsum<>0 then
		%>
			<%= Clng( ((ozone.FItemList(i).fsellsum / (split(ozone.ftmpSumTotal,",")(ozone.FItemList(i).fblock))) * 10000)) / 100 %> %
		<% else %>
			0 %
		<% end if %>
	</td>

	<% if C_ADMIN_USER or C_IS_OWN_SHOP then %>
		<% if not(C_IS_Maker_Upche) then %>
			<% if dategubun = "M" or dategubun = "D" then %>
				<td>
					<% if ozone.FItemList(i).fsellsum <> 0 and ozone.FItemList(i).ftargetmaechul <> 0 then %>
						<%'= FormatNumber(ooffsell.FItemList(i).ftargetmaechul,0) %>
						<% response.write round(((ozone.FItemList(i).fsellsum/ozone.FItemList(i).ftargetmaechul) *100),1) %> %
					<% end if %>					

					<% if (ozone.FItemList(i).fzonename <> "" or isnull(ozone.FItemList(i).fzonename)) and ozone.FItemList(i).ftargetmaechul = 0 then %>
						<a href="javascript:regtargetmaechul('<%= ozone.FItemList(i).fshopid %>','<%= left(ozone.FItemList(i).fIXyyyymmdd,4) %>','<%= mid(ozone.FItemList(i).fIXyyyymmdd,6,2) %>','2')" onfocus='this.blur()'>
						목표매출등록</a>
					<% end if %>
				</td>
			<% end if %>
		<% end if %>
	<% end if %>
		
	<td>
		<% if ozone.FItemList(i).fsellsum <> 0 and ozone.FItemList(i).funit <> 0 then %>
			<%= FormatNumber(ozone.FItemList(i).fsellsum / ozone.FItemList(i).funit,0) %>
		<% else %>
			0
		<% end if %>
	</td>
	<td>
		<%= ozone.FItemList(i).funit %>
	</td>
	<td align="center">
		<% if ozone.FItemList(i).funit<>0 and ozone.FItemList(i).frealpyeong <> 0 then %>
			<%= Clng( ((ozone.FItemList(i).funit / ozone.FItemList(i).frealpyeong) * 10000)) / 100 %> %
		<% else %>
			0 %
		<% end if %>
	</td>		
	<td width=250>
		<% if dategubun = "G" then %>
			<input type="button" class="button" value="카테고리" onclick="category_detail('<%= ozone.FItemList(i).fidx %>','<%= yyyy1 %>','<%= mm1 %>','<%= dd1 %>','<%= yyyy2 %>','<%= mm2 %>','<%= dd2 %>');">
			<input type="button" class="button" value="브랜드" onclick="brand_detail('<%= ozone.FItemList(i).fidx %>','<%= yyyy1 %>','<%= mm1 %>','<%= dd1 %>','<%= yyyy2 %>','<%= mm2 %>','<%= dd2 %>');">
			<input type="button" class="button" value="상품상세" onclick="item_detail('<%= ozone.FItemList(i).fidx %>','<%= yyyy1 %>','<%= mm1 %>','<%= dd1 %>','<%= yyyy2 %>','<%= mm2 %>','<%= dd2 %>');">
		<% elseif dategubun = "M" then %>
			<input type="button" class="button" value="카테고리" onclick="category_detail('<%= ozone.FItemList(i).fidx %>','<%= left(ozone.FItemList(i).fIXyyyymmdd,4) %>','<%= mid(ozone.FItemList(i).fIXyyyymmdd,6,2) %>','01','<%= left(ozone.FItemList(i).fIXyyyymmdd,4) %>','<%= mid(ozone.FItemList(i).fIXyyyymmdd,6,2) %>','31');">
			<input type="button" class="button" value="브랜드" onclick="brand_detail('<%= ozone.FItemList(i).fidx %>','<%= left(ozone.FItemList(i).fIXyyyymmdd,4) %>','<%= mid(ozone.FItemList(i).fIXyyyymmdd,6,2) %>','01','<%= left(ozone.FItemList(i).fIXyyyymmdd,4) %>','<%= mid(ozone.FItemList(i).fIXyyyymmdd,6,2) %>','31');">
			<input type="button" class="button" value="상품상세" onclick="item_detail('<%= ozone.FItemList(i).fidx %>','<%= left(ozone.FItemList(i).fIXyyyymmdd,4) %>','<%= mid(ozone.FItemList(i).fIXyyyymmdd,6,2) %>','01','<%= left(ozone.FItemList(i).fIXyyyymmdd,4) %>','<%= mid(ozone.FItemList(i).fIXyyyymmdd,6,2) %>','31');">
		<% elseif dategubun = "D" then %>
			<input type="button" class="button" value="카테고리" onclick="category_detail('<%= ozone.FItemList(i).fidx %>','<%= left(ozone.FItemList(i).fIXyyyymmdd,4) %>','<%= mid(ozone.FItemList(i).fIXyyyymmdd,6,2) %>','<%= right(ozone.FItemList(i).fIXyyyymmdd,2) %>','<%= left(ozone.FItemList(i).fIXyyyymmdd,4) %>','<%= mid(ozone.FItemList(i).fIXyyyymmdd,6,2) %>','<%= right(ozone.FItemList(i).fIXyyyymmdd,2) %>');">
			<input type="button" class="button" value="브랜드" onclick="brand_detail('<%= ozone.FItemList(i).fidx %>','<%= left(ozone.FItemList(i).fIXyyyymmdd,4) %>','<%= mid(ozone.FItemList(i).fIXyyyymmdd,6,2) %>','<%= right(ozone.FItemList(i).fIXyyyymmdd,2) %>','<%= left(ozone.FItemList(i).fIXyyyymmdd,4) %>','<%= mid(ozone.FItemList(i).fIXyyyymmdd,6,2) %>','<%= right(ozone.FItemList(i).fIXyyyymmdd,2) %>');">
			<input type="button" class="button" value="상품상세" onclick="item_detail('<%= ozone.FItemList(i).fidx %>','<%= left(ozone.FItemList(i).fIXyyyymmdd,4) %>','<%= mid(ozone.FItemList(i).fIXyyyymmdd,6,2) %>','<%= right(ozone.FItemList(i).fIXyyyymmdd,2) %>','<%= left(ozone.FItemList(i).fIXyyyymmdd,4) %>','<%= mid(ozone.FItemList(i).fIXyyyymmdd,6,2) %>','<%= right(ozone.FItemList(i).fIXyyyymmdd,2) %>');">		
		<% end if %>
	</td>
</tr>
<% next %>
<tr align="center" bgcolor="#f1f1f1">
	<% if dategubun = "D" or dategubun = "M" then %>
		<td colspan=4>
			<%= tmpselldate %> 총합계
		</td>
	<% else %>
		<td colspan=3>
			<%= tmpselldate %> 총합계
		</td>
	<% end if %>		
	<td>
		<%= FormatNumber(totsellsum,0) %>
	</td>
	<td></td>
	
	<% if C_ADMIN_USER or C_IS_OWN_SHOP then %>
		<% if not(C_IS_Maker_Upche) then %>
			<% if dategubun = "M" or dategubun = "D" then %>
				<td>				
				</td>
			<% end if %>
		<% end if %>
	<% end if %>	
	
	<td>
		<%= FormatNumber(totunitmaechul,0) %>
	</td>
	<td>
		<%= totunit %>
	</td>
	<td colspan=2></td>
</tr>
<% else %>
	<tr bgcolor="#FFFFFF">
		<td colspan="15" align="center" class="page_link">[검색결과가 없습니다.]</td>
	</tr>
<% end if %>
</table>

<%
set ozone = nothing
%>
<!-- #include virtual="/common/lib/commonbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->