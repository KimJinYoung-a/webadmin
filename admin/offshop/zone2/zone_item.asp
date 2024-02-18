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
<!-- #include virtual="/common/lib/commonbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/offshop/zone2/zone_cls.asp"-->

<%
Dim ozone,i,page,isusing , parameter , shopid ,fromDate ,toDate,yyyy1,mm1,dd1,yyyy2,mm2,dd2, zonegroup ,racktype
dim designer , itemid , itemname , searchtype  ,datefg,cdl ,cdm ,cds, zoneidx, tr_color, tmp_tr, viewzone
dim zoneisusing , searchgubun
	tr_color = 0
	designer = RequestCheckVar(request("designer"),32)
	searchgubun = RequestCheckVar(request("searchgubun"),16)
	isusing = requestCheckVar(request("isusing"),1)
	page = requestCheckVar(request("page"),10)
	shopid = requestCheckVar(request("shopid"),32)
	itemid = requestCheckVar(request("itemid"),10)
	itemname = requestCheckVar(request("itemname"),124)
	yyyy1 = requestCheckVar(request("yyyy1"),4)
	mm1 = requestCheckVar(request("mm1"),2)
	dd1 = requestCheckVar(request("dd1"),2)
	yyyy2 = requestCheckVar(request("yyyy2"),4)
	mm2 = requestCheckVar(request("mm2"),2)
	dd2 = requestCheckVar(request("dd2"),2)
	zonegroup = requestCheckVar(request("zonegroup"),32)
	racktype = requestCheckVar(request("racktype"),10)
	searchtype = requestCheckVar(request("searchtype"),1)
	menupos = requestCheckVar(request("menupos"),10)
	datefg = requestCheckVar(request("datefg"),16)
	cdl     = requestCheckVar(request("cdl"),3)
	cdm     = requestCheckVar(request("cdm"),3)
	cds     = requestCheckVar(request("cds"),3)
	zoneidx = requestCheckVar(request("zoneidx"),10)
	viewzone = requestCheckVar(request("viewzone"),32)
	zoneisusing = requestCheckVar(request("zoneisusing"),1)

if searchgubun = "" then searchgubun = "M3"
if datefg = "" then datefg = "maechul"
if page = "" then page = 1
if searchtype = "" then searchtype = "M"
if (yyyy1="") then yyyy1 = Cstr(Year(now()))
if (mm1="") then mm1 = Cstr(Month(now()))
if (dd1="") then dd1 = Cstr(day(now()))
if (yyyy2="") then yyyy2 = Cstr(Year(now()))
if (mm2="") then mm2 = Cstr(Month(now()))
if (dd2="") then dd2 = Cstr(day(now()))

fromDate = DateSerial(yyyy1, mm1, dd1)
toDate = DateSerial(yyyy2, mm2, dd2+1)

'직영/가맹점
if (C_IS_SHOP) then

	'/어드민권한 점장 미만
	if getlevel_sn("",session("ssBctId")) > 6 then
		shopid = C_STREETSHOPID
	end if
end if

set ozone = new czone_list
	ozone.FPageSize = 1000
	ozone.FCurrPage = page
	ozone.frectzonegroup = zonegroup
	ozone.frectracktype = racktype
	ozone.frectisusing = isusing
	ozone.frectshopid = shopid
	ozone.frectitemid = itemid
	ozone.frectitemname = itemname
	ozone.FRectStartDay = fromDate
	ozone.FRectEndDay = toDate
	ozone.FRectmakerid = designer
	ozone.frectdatefg = datefg
	ozone.FRectCDL = cdl
	ozone.FRectCDM = cdm
	ozone.FRectCDN = cds
	ozone.FRectsearchtype = searchtype
	ozone.frectidx = zoneidx
	ozone.frectzoneisusing = zoneisusing
	ozone.frectsearchgubun = searchgubun

	if shopid <> "" then
		ozone.GetoffshopzoneitemMatch

		if drawnewipgobrand(shopid) <> "" then
			response.write "<script language='javascript'>"
			response.write "	alert('"&shopid&" 매장에 최근 3개월내에 조닝에 설정되지 않은 신규브랜드가 있습니다\n\n"&drawnewipgobrand(shopid)&"');"
			response.write "</script>"
		end if

	end if
%>

<script language="javascript">

	//선택상품 저장
	function zone_change(upfrm){

		var j = document.getElementsByName("makerid").length;
		var k = new Array();
		var m = 0;
		for(var i=0; i < j ; i++){
			if (document.getElementsByName("makerid")[i].checked == true)
			{
				k[m] = document.getElementsByName("makerid")[i].value;
				m = m+1;
			}
		}
		if(m == 0)
		{
			alert('브랜드를 선택해 주세요');
			return;
		}

		if (upfrm.shopid.value==''){
			alert('매장을 선택해 주세요');
			return;
		}
		if (upfrm.chzoneidx.value==''){
			alert('저장 하실 조닝을 선택해 주세요');
			return;
		}

		upfrm.action='zone_process.asp';
		upfrm.mode.value='zoneitemreg';
		upfrm.submit();
	}

	function gopage(page){
		frm.page.value=page;
		frm.submit();
	}

	function changeshop(shopid)
	{
		if(shopid == "")
		{
			alert("매장을 선택해 주세요.");
			return;
		}

		frm.shopid.value=shopid
		frm.submit();
	}

	function onlyviewzone(zoneidx)
	{
		frm.zoneidx.value=zoneidx
		frm.submit();
	}

	function Check_All(icomp)
	{
		var chked = "";
		if(icomp.checked)
		{
			chked = "checked";
		}

		var chk = document.getElementsByName("makerid");
		var cnt=0;
		if (cnt==0 && chk.length != 0) {
			for(i = 0; i < chk.length; i++)
			{
				chk.item(i).checked = chked;
			}
			cnt++;
		}
	}

	function divch(divid,zoneidx){
		frmdiv.divid.value = divid;
		frmdiv.zoneidx.value = zoneidx;
		frmdiv.target="view";
		frmdiv.action='/admin/offshop/zone2/zone_manager_search.asp';
		frmdiv.submit();
	}

</script>

<form name="frmdiv" method="get" action="">
	<input type="hidden" name="divid">
	<input type="hidden" name="zoneidx">
</form>
<form name="frm" method="get" action="" style="margin:0px;">

<!-- 검색 시작 -->
<table cellpadding="0" cellspacing="0" border="0" class="a">
<tr>
	<td style="padding:0 0 7px 0;">

	</td>
</tr>
</table>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="page" value="<%= page %>">
<input type="hidden" name="mode">
<input type="hidden" name="itemgubunarr" value="">
<input type="hidden" name="shopitemidarr" value="">
<input type="hidden" name="itemoptionarr" value="">
<input type="hidden" name="zoneidxarr" value="">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="2" width="60" bgcolor="<%= adminColor("gray") %>">검색조건</td>
	<td align="left">
		<table cellpadding="0" cellspacing="0" border="0" class="a">
		<tr height="25">
			<td>
				<%
				'직영/가맹점
				if (C_IS_SHOP) then
				%>
					<% if getoffshopdiv(shopid) <> "1" and shopid <> "" then %>
						* 매장 : <%=shopid%><input type="hidden" name="shopid" value="<%= shopid %>">
					<% else %>
						* 매장 : <% Call NewDrawSelectBoxDesignerwithNameAndUserDIV("shopid",shopid, "21") %>
					<% end if %>
				<% else %>
					* 매장 : <% Call NewDrawSelectBoxDesignerwithNameAndUserDIV("shopid",shopid, "21") %>
				<% end if %>
				&nbsp;&nbsp;
				* 조회기준 :
				<input type="radio" name="searchgubun" value="M3" <% if searchgubun = "M3" then response.write " checked" %> onclick="gopage('');">최근3개월판매내역
				<input type="radio" name="searchgubun" value="A" <% if searchgubun = "A" then response.write " checked" %> onclick="gopage('');">입고된전체브랜드(신규브랜드포함)
			</td>
		</tr>
		<tr height="25">
			<td>
				* 브랜드 : <% drawSelectBoxDesignerwithName "designer",designer %>
				&nbsp;&nbsp;
				* 조닝지정여부:
				<select name="zoneisusing" value="<%=zoneisusing%>" onchange="gopage('');">
					<option value="" <% if zoneisusing = "" then response.write " selected" %>>선택</option>
					<option value="Y" <% if zoneisusing = "Y" then response.write " selected" %>>Y</option>
					<option value="N" <% if zoneisusing = "N" then response.write " selected" %>>N</option>
				</select>
				&nbsp;&nbsp;
				<% Call zoneselectbox(shopid,"zoneidx",zoneidx," onchange=""gopage('');""") %>
			</td>
		</tr>
		</table>
	</td>
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="검색" onClick="gopage('');">
	</td>
</tr>
</table>
<!-- 검색 끝 -->

<br>

<% If shopid = "" Then %>
	<center><font color="red"><b>※ ShopID(매장)를 선택하셔야 데이터가 나타납니다.</b></font></center><br>
<% End If %>

<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
<tr>
	<td align="left">
		<% drawzonechange "chzoneidx","",shopid,"" %>
		<input type="button" value="선택상품저장" class="button" onclick='zone_change(frm)';>
	</td>
	<td align="right">
	</td>
</tr>
</table>

<!-- 액션 끝 -->

<!-- 리스트 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="15">
		검색결과 : <b><%= ozone.FTotalCount %></b>
		※ 1000건 까지 검색가능
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td><input type="checkbox" name="ckall" onclick="Check_All(this)"></td>
	<td>브랜드ID</td>
	<td>조닝</td>
	<td>등록일</td>
	<td>매장내<br>담당자</td>
</tr>
<% if ozone.FresultCount>0 then %>
<% for i=0 to ozone.FresultCount-1 %>

<% if ozone.FItemList(i).fzonename <> "" then %>
<%
	If tmp_tr <> ozone.FItemList(i).fzonename Then
		tr_color = tr_color + 1
	End If

	tmp_tr = ozone.FItemList(i).fzonename
%>
<tr align="center" bgcolor="<%=TrColor(tr_color)%>">
<% else %>
<tr align="center" bgcolor="#FFFFFF">
<% end if %>
	<td><input type="checkbox" name="makerid" value="<%= ozone.FItemList(i).fmakerid %>" onClick="AnCheckClick(this);"></td>
	<td align="left"><%= ozone.FItemList(i).fmakerid %></td>
	<td align="center">
		<% if ozone.FItemList(i).fzoneidx = "" or isnull(ozone.FItemList(i).fzoneidx)then %>
			-
		<% else %>
			<b><a href="javascript:" onClick="onlyviewzone('<%=ozone.FItemList(i).fzoneidx%>')">[<%= ozone.FItemList(i).fzonename %>]</a></b>
		<% end if %>
	</td>
	<td><%= ozone.FItemList(i).fregdate %></td>
	<td>
		<% if ozone.FItemList(i).fmanagershopyn = "Y" then %>
			<div name="div<%=i%>" id="div<%=i%>">
				<img src="/images/icon_search.jpg" onmouseover="javascript:divch('div<%=i%>','<%=ozone.FItemList(i).fzoneidx%>');">
			</div>
		<% end if %>
	</td>
</tr>

<% next %>

<% else %>
	<tr bgcolor="#FFFFFF">
		<td colspan="15" align="center" class="page_link">[검색결과가 없습니다.]</td>
	</tr>
<% end if %>
<iframe id="view" name="view" src="" width=0 height=0 frameborder="0" scrolling="no" ></iframe>
</table>
</form>

<%
set ozone = nothing
%>
<!-- #include virtual="/common/lib/commonbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
