<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/summaryupdatelib.asp"-->
<!-- #include virtual="/lib/classes/stock/ordersheetcls.asp"-->
<!-- #include virtual="/lib/classes/offshop/offshopitemcls.asp"-->

<%
' 사용안하는듯
response.end
%>

<script language='javascript'>
function MakeJumunByIdx(idxarr,designerid,etcstr){
	//alert(idxarr);
	//alert(designerid);
	document.dumifrm.idxarr.value=idxarr;
	document.dumifrm.designerid.value=designerid;
	document.dumifrm.etcstr.value=etcstr;
	document.dumifrm.submit();
}

function PopFranBalju2Upchebalju(frm){
	var designerid,baljuid,popwin;
	designerid = frm.designerid.value;
	baljuid = frm.baljuid.value;
	popwin = window.open('popfranbalju2upchebalju.asp?designerid=' + designerid + '&baljuid=' + baljuid  ,'franbalju2upchebalju','width=800,height=600,scrollbars=yes,status=no');
	popwin.focus();
}

function PopFranBalju2UpchebaljuByID(designerid){
    var baljuid,popwin;
	baljuid = "10x10";
	popwin = window.open('popfranbalju2upchebalju.asp?designerid=' + designerid + '&baljuid=' + baljuid  ,'franbalju2upchebalju','width=800,height=600,scrollbars=yes,resizable=yes');
	popwin.focus();
}
</script>
<%
''업체개별주문서작성

dim idxarr,designerid,statecd, etcstr
dim designer
dim includepreorderno, shortyn
dim research

idxarr      		= request("idxarr")
designerid  		= request("designerid")
statecd     		= request("statecd")
designer    		= request("designer")
etcstr      		= request("etcstr")
shortyn   			= request("shortyn")
includepreorderno   = request("includepreorderno")
research   			= request("research")

if (research = "") then
	'shortyn = "Y"
	'includepreorderno = "Y"
end if

if (includepreorderno = "Y") then
	shortyn = "Y"
end if



dim oupchejumun,iidx
dim DefaultItemMwDiv

iidx =0
if (idxarr<>"") and (designerid<>"") then
	if Right(idxarr,1)="," then idxarr=Left(idxarr,Len(idxarr)-1)
	if Right(etcstr,1)="," then etcstr=Left(etcstr,Len(etcstr)-1)

	DefaultItemMwDiv = GetDefaultItemMwdivByBrand(designerid)

	set oupchejumun = new COrderSheet
	oupchejumun.FRectIdxArr  = idxarr
	oupchejumun.FRectMakerid = designerid
	oupchejumun.FRectTargetid = designerid
	oupchejumun.FRectBaljuId = "10x10"
	oupchejumun.FRectBaljuname = "텐바이텐"
	oupchejumun.FRectReguser = session("ssBctId")
	oupchejumun.FRectRegname = session("ssBctCname")
	if (DefaultItemMwDiv="M") then
	oupchejumun.FRectdivcode = "101"
	else
	oupchejumun.FRectdivcode = "111"
	end if
	oupchejumun.FRectComment = "원주문 : " + html2db(etcstr)

	iidx = oupchejumun.MakeUpcheJumun

	set oupchejumun = Nothing

	'주문서 기준 기주문 업데이트
	PreOrderUpdateBySheetIdx(iidx)

	response.redirect "/admin/fran/upchejumuninputedit.asp?idx=" + CStr(iidx) + "&opage=1&ourl=upchejumunlist.asp"
	dbget.close()	:	response.End
end if

dim oordersheet1
set oordersheet1 = new COrderSheet
oordersheet1.FRectMakerid = designer
oordersheet1.FRectStatecd = statecd
oordersheet1.FRectBaljuId = "10x10"

oordersheet1.FRectShortYN = shortyn
oordersheet1.FRectIncludePreOrderNo = includepreorderno

oordersheet1.GetFranBalju2UpcheBaljuBrandlist

dim i
%>
<!--
<table width="800" cellspacing="1" class="a" bgcolor=#3d3d3d>
<tr bgcolor="#FFFFFF">
	<td>
		* 2월 1일 주문 부터 <br>
		매입업체경우 무조건 온라인 재고에서 나가야함.<br>
		위탁업체중 매입건->온라인재고 나가고 출고로 등록<br>
		위탁업체중 위탁건->온라인재고 나가고 출고로 등록<br>
		<br>
		이곳에서 따로 주문해야 하는경우<br>
		- 매입인데 마진이 다를경우(없앨 예정 prixe, multiple_choice, nanishow)<br>
		- 업체배송주문건.(가맹점용 개별매입, 가맹점용 개별위탁)
	</td>
</tr>
</table>
-->

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="searchfrm" method=get">
	<input type="hidden" name="research" value="on">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<tr align="center" bgcolor="#FFFFFF" >
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
		<td align="left">
		    브랜드 : <% drawSelectBoxDesignerwithName "designer", designer %>
		    &nbsp;
			<input type=radio name="statecd" value="" <% if statecd="" then response.write "checked" %> >주문접수 + 상품준비
			<input type=radio name="statecd" value="0" <% if statecd="0" then response.write "checked" %> >주문접수
			<input type=radio name="statecd" value="1" <% if statecd="1" then response.write "checked" %> >상품준비
		</td>
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button_s" value="검색" onClick="javascript:document.searchfrm.submit();">
		</td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF" >
		<td align="left">
		    부족구분 :
			<input type=checkbox name="shortyn" value="Y" <% if shortyn = "Y" then response.write "checked" %>> 재고부족만
			<input type=checkbox name="includepreorderno" value="Y" <% if includepreorderno = "Y" then response.write "checked" %>> 기주문포함부족만
		</td>
	</tr>
	</form>
</table>

<p>
<!--
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frm" method=post action="">
	<tr bgcolor="#FFFFFF">
	<% if designerid<>"" then %>
		<input type="hidden" name="designerid" value="<%= designerid %>">
		<td bgcolor="<%= adminColor("tabletop") %>" width=100>매입처</td>
		<td><%= designerid %></td>
	<% else %>
		<td bgcolor="<%= adminColor("tabletop") %>" width=100>매입처</td>
		<td>
			<% drawSelectBoxDesignerwithName "designerid",designerid %>
			&nbsp;
			<input type="button" class="button" value="가맹점 주문서로 작성" onclick="PopFranBalju2Upchebalju(frm);">
		</td>
	<% end if %>
	</tr>
	<tr bgcolor="#FFFFFF">
		<input type="hidden" name="baljuid" value="10x10">
		<td bgcolor="<%= adminColor("tabletop") %>" width=100>발주처</td>
		<td>10x10</td>
	</tr>
	</form>
</table>
-->
<p>

<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
	<tr>
		<td align="left">
			<!-- <input type="button" class="button" value="주문서작성" onClick=""> -->
		</td>
		<td align="right">
		    * 브랜드 아이디 클릭후 작성 가능
			/ 업체 반품 주문서는 이곳에서 작성 하실 수 없습니다.
		</td>
	</tr>
</table>
<!-- 액션 끝 -->

<p>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frm2" method=post action="">
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td width=20></td>
		<td width="120">브랜드ID</td>
		<td width="100">상품코드</td>
		<td width="50">이미지</td>
		<td>상품명<font color="blue">[옵션명]</font></td>
		<td width="40">배송<br>구분</td>
		<td width="40">매입<br>구분</td>
		<td width="50">OFF<br>상품준비</td>
		<td width="50">OFF<br>주문접수</td>
		<td width="50"><b>미출고<br>합계</b></td>
		<td width="50">실사<br>재고</td>
		<td width="50"><b>부족<br>수량</b></td>
		<td width="100">비고</td>
	</tr>
	<% for i=0 to oordersheet1.FResultCount -1 %>
	<tr align="center" bgcolor="#FFFFFF">
		<td><input type=checkbox name="cksel" onClick="AnCheckClick(this);"></td>
		<input type=hidden name="idx" value="">
		<td><a href="javascript:PopFranBalju2UpchebaljuByID('<%= oordersheet1.FItemList(i).FMakerid %>');"><%= oordersheet1.FItemList(i).FMakerid %></a></td>
		<td><%= oordersheet1.FItemList(i).FItemGubun %>-<%= CHKIIF(oordersheet1.FItemList(i).FItemId>=1000000,Format00(8,oordersheet1.FItemList(i).FItemId),Format00(6,oordersheet1.FItemList(i).FItemId)) %>-<%= oordersheet1.FItemList(i).FItemoption %></td>
		<td></td>
		<td align="left">
			<%= oordersheet1.FItemList(i).FItemName %>
				&nbsp;
			<% if oordersheet1.FItemList(i).FItemoption<>"0000" then %>
				<font color="blue"><%= oordersheet1.FItemList(i).FItemOptionname %></font>
			<% end if %>
		</td>
		<td><%= oordersheet1.FItemList(i).GetDeliverTypeString %></td>
		<td><%= oordersheet1.FItemList(i).GetMWDivString %></td>
		<td><%= oordersheet1.FItemList(i).FCount %></td>
		<td><%= oordersheet1.FItemList(i).FJupsuCount %></td>
		<td><b><%= oordersheet1.FItemList(i).FCount + oordersheet1.FItemList(i).FJupsuCount %></b></td>
		<td><%= oordersheet1.FItemList(i).Frealstock %></td>
		<td><b><%= oordersheet1.FItemList(i).Frealstock - oordersheet1.FItemList(i).FCount - oordersheet1.FItemList(i).FJupsuCount %></b></td>
		<td>
			<% if ((Not IsNull(oordersheet1.FItemList(i).FreipgoMayDate)) and (Left(oordersheet1.FItemList(i).FreipgoMayDate, 10) >= Left(DateAdd("m", -3, now()), 10) ) ) then %>
				<%= Left(oordersheet1.FItemList(i).FreipgoMayDate, 10) %><br>
			<% end if %>
			<% if oordersheet1.FItemList(i).Fpreorderno<>0 then %>
				기주문:
				<% if oordersheet1.FItemList(i).Fpreorderno<>oordersheet1.FItemList(i).Fpreordernofix then response.write "</br>" + CStr(oordersheet1.FItemList(i).Fpreorderno) + "->" %>
					<%= oordersheet1.FItemList(i).Fpreordernofix %>
			<% end if %>
		</td>
	</tr>
	<% next %>
	</form>
</table>
<%
set oordersheet1 = nothing
%>
<form name="dumifrm" method=post action="">
<input type="hidden" name="idxarr" value="">
<input type="hidden" name="designerid" value="">
<input type="hidden" name="etcstr" value="">
</form>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
