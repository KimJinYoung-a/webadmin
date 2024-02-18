<%@  codepage="65001" language="VBScript" %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/items/overseas/overseasCls.asp"-->

<%
Dim oitem, i, page, itemid, itemname, makerid, cdl, cdm, cds, vCountryCd, sellyn, usingyn, useyn, limityn,actionURL, discountKey
Dim sitename, currencyunit, onlineBasePrc

	page = request("page")
	vCountryCd	= request("countrycd")
	itemid		= request("itemid")
	itemname	= request("itemname")
	makerid		= request("makerid")
	sellyn		= request("sellyn")
	usingyn		= request("usingyn")
	cdl = request("cdl")
	cdm = request("cdm")
	cds = request("cds")
	useyn = request("useyn")
	limityn = request("limityn")
    sitename = request("sitename")
    currencyunit = request("currencyunit")
    discountKey = request("discountKey")
    actionURL	= "/admin/etc/kaffa/sale/saleitemProc.asp"
'기본값
if (page = "") then page = 1
'if (vCountryCd = "") then vCountryCd = "kr"

set oitem = new COverSeasItem
	oitem.FPageSize         = 30
	oitem.FCurrPage         = page
	oitem.FRectCountryCd	= vCountryCd
	oitem.FRectMakerid      = makerid
	oitem.FRectSellYN       = sellyn
	oitem.FRectIsUsing      = usingyn
	oitem.FRectItemid       = itemid
	oitem.FRectItemName     = itemname
	oitem.FRectCate_Large   = cdl
	oitem.FRectCate_Mid     = cdm
	oitem.FRectCate_Small   = cds
	oitem.FRectLimitYN		= limityn
	oitem.FRectuseyn = useyn
    oitem.FRectSitename = Sitename
    oitem.FRectcurrencyunit = currencyunit

	If sitename <> "" Then
		oitem.GetOverSeasItemList
	else
		response.write "<script type='text/javascript'>"
		response.write "	alert('사이트를 선택하세요');"
		response.write "</script>"
	End If

onlineBasePrc = 0
%>

<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=UTF-8">
<script language="JavaScript" src="/js/xl.js"></script>
<script language="JavaScript" src="/js/common.js"></script>
<script language="JavaScript" src="/js/report.js"></script>
<script language="JavaScript" src="/cscenter/js/cscenter.js"></script>
<script language="JavaScript" src="/js/calendar.js"></script>

<link rel="stylesheet" href="/css/scm.css" type="text/css">

<script language='javascript'>

function PopMenuHelp(menupos){
	var popwin = window.open('/designer/menu/help.asp?menupos=' + menupos,'PopMenuHelp_a','width=900, height=600, scrollbars=yes,resizable=yes');
	popwin.focus();
}

function PopMenuEdit(menupos){
	var popwin = window.open('/admin/menu/pop_menu_edit.asp?mid=' + menupos,'PopMenuEdit','width=600, height=400, scrollbars=yes,resizable=yes');
	popwin.focus();
}

function PopItemContent(iitemid){
	var popwin = window.open('/admin/itemmaster/overseas/popItemContent.asp?itemid=' + iitemid +'&sitename=<%=sitename%>','itemWeightEdit','width=1024,height=768,scrollbars=yes,resizable=yes')
	popwin.focus();
}

function NextPage(ipage){
	document.frm.page.value= ipage;
	document.frm.submit();
}
//전체 선택
function jsChkAll(){
var frm;
frm = document.frm2;
	if (frm.chkAll.checked){
	   if(typeof(frm.chkitem) !="undefined"){
	   	   if(!frm.chkitem.length){
		   	 	frm.chkitem.checked = true;
		   }else{
				for(i=0;i<frm.chkitem.length;i++){
					frm.chkitem[i].checked = true;
			 	}
		   }
	   }
	} else {
	  if(typeof(frm.chkitem) !="undefined"){
	  	if(!frm.chkitem.length){
	   	 	frm.chkitem.checked = false;
	   	}else{
			for(i=0;i<frm.chkitem.length;i++){
				frm.chkitem[i].checked = false;
			}
		}
	  }

	}

}
function SelectItems(sType){
var frm;
var itemcount = 0;
frm = document.frm2;
frm.sType.value = sType;   //전체선택 or 선택상품 여부 구분

	if (sType == "sel"){
		 if(typeof(frm.chkitem) !="undefined"){
	   	   	if(!frm.chkitem.length){
	   	   		if(!frm.chkitem.checked){
	   	   			alert("선택한 상품이 없습니다. 상품을 선택해 주세요");
	   	   			return;
	   	   		}
	   	   		 frm.itemidarr.value = frm.chkitem.value;
	   	   		 itemcount = 1;
	   	    }else{
	   	    	for(i=0;i<frm.chkitem.length;i++){
	   	    		if(frm.chkitem[i].checked) {
	   	    			if (frm.itemidarr.value==""){
	   	    			 frm.itemidarr.value =  frm.chkitem[i].value;
	   	    			}else{
	   	    			 frm.itemidarr.value = frm.itemidarr.value + "," +frm.chkitem[i].value;
	   	    			}
	   	    		}
	   	    		itemcount = frm.chkitem.length;
	   	    	}

	   	    	if (frm.itemidarr.value == ""){
	   	    		alert("선택한 상품이 없습니다. 상품을 선택해 주세요");
	   	   			return;
	   	    	}
	   	    }
	   	  }else{
	   	  	alert("추가할 상품이 없습니다.");
	   	  	return;
	   	  }
	}
	frm.target = "FrameCKP";
	frm.action = "<%=actionURL%>";
	frm.itemcount.value = itemcount;
	frm.submit();
	frm.itemidarr.value = "";
	frm.itemcount.value = 0;
	opener.history.go(0);
	//window.close();
}
</script>

<% if session("sslgnMethod")<>"S" then %>
	<!-- USB키 처리 시작 (2008.06.23;허진원) -->
	<OBJECT ID='MaGerAuth' WIDTH='' HEIGHT=''	CLASSID='CLSID:781E60AE-A0AD-4A0D-A6A1-C9C060736CFC' codebase='/lib/util/MaGer/MagerAuth.cab#Version=1,0,2,4'></OBJECT>
	<script language="javascript" src="/js/check_USBToken.js"></script>
	<!-- USB키 처리 끝 -->
<% end if %>

</head>
<body bgcolor="#F4F4F4" onload="checkUSBKey()">
<div id="calendarPopup" style="position: absolute; visibility: hidden; z-index: 2;"></div>

<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a">
<tr>
	<td>
		<table width="100%" align="center" cellpadding="1" cellspacing="0" class="a" bgcolor="#999999">
			<tr>
				<td align="right" style="border-bottom:1px solid #999999" bgcolor="#F4F4F4">
					<a href="Javascript:PopMenuEdit('1491');"><img src="/images/icon_chgauth.gif" border="0" valign="bottom"></a>
					<a href="Javascript:PopMenuHelp('1491');"><img src="/images/icon_help.gif" border="0" valign="bottom"></a>
				</td>
			</tr>
		</table>
	</td>
</tr>
</table>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method=get>
<input type="hidden" name="menupos" value="<%= Request("menupos") %>">
<input type="hidden" name="page" >
<input type="hidden" name="discountKey" value="<%=discountKey%>">
<tr align="center" bgcolor="<%= adminColor("topbar") %>" >
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
	<td align="left">
		* 브랜드 :<%	drawSelectBoxDesignerWithName "makerid", makerid %>
		&nbsp;
		<!-- #include virtual="/common/module/categoryselectbox_utf8.asp"-->
		<p>
		* 상품코드 :
		<input type="text" class="text" name="itemid" value="<%= itemid %>" size="30" maxlength="100" onKeyPress="if (event.keyCode == 13) document.frm.submit();">(쉼표로 복수입력가능)
		&nbsp;
		* 상품명 :
		<input type="text" class="text" name="itemname" value="<%= itemname %>" size="32" maxlength="32">
	</td>
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="검색" onClick="javascript:document.frm.submit();">
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("topbar") %>" >
	<td align="left">
		* 판매:
	   <select class="select" name="sellyn">
		   <option value="">전체</option>
		   <option value="Y"  <%=CHKIIF(sellyn="Y","selected","")%>>판매</option>
		   <option value="S"  <%=CHKIIF(sellyn="S","selected","")%>>일시품절</option>
		   <option value="N"  <%=CHKIIF(sellyn="N","selected","")%>>품절</option>
		   <option value="YS"  <%=CHKIIF(sellyn="YS","selected","")%>>판매+일시품절</option>
	   </select>
     	&nbsp;&nbsp;
     	* 상품사용여부:
	   <select class="select" name="usingyn">
		   <option value="">전체</option>
		   <option value="Y"  <%=CHKIIF(usingyn="Y","selected","")%>>사용함</option>
		   <option value="N"  <%=CHKIIF(usingyn="N","selected","")%>>사용안함</option>
	   </select>
	   &nbsp;&nbsp;
     	* 한정여부:
	   <select class="select" name="limityn">
		   <option value="">전체</option>
		   <option value="Y"  <%=CHKIIF(limityn="Y","selected","")%>>Y</option>
		   <option value="N"  <%=CHKIIF(limityn="N","selected","")%>>N</option>
	   </select>
	   &nbsp;&nbsp;
	   * 해외상품사용여부 : <% drawSelectBoxUsingYN "useyn", useyn %>
     	<p>
		<b><font color="blue">
	    * 사이트 : <% drawSelectboxMultiSiteSitename "sitename", sitename, " onchange='NextPage("""");'" %>
	    </font></b>
	    &nbsp;&nbsp;
	    <% if sitename<>"" then %>
	    	* 화폐 : <% drawSelectBoxsitecurrencyunit sitename, "currencyunit", currencyunit, " onchange='NextPage("""");'" %>
	    <% end if %>
	</td>
</tr>
</form>
</table>

<br>
<!-- 표 중간바 시작-->
<table width="100%" align="center" cellpadding="1" cellspacing="1" class="a">
<tr valign="bottom">
    <td align="left">
    </td>
    <td align="right">
    </td>
</tr>
</table>
<!-- 표 중간바 끝-->

<table width="100%" height="40" align="center" cellpadding="3" cellspacing="1" class="a" border="0">
	<tr>
		<td  valign="bottom">
			<input type="button" value="선택상품 추가" onClick="SelectItems('sel')" class="button">
		</td>
	</tr>
</table>

<form name="frm2" method="post">
<input type="hidden" name="page" >
<input type="hidden" name="sType" >
<input type="hidden" name="itemidarr" >
<input type="hidden" name="itemcount" value="0">
<input type="hidden" name="mode" value="I">
<input type="hidden" name="discountKey" value="<%=discountKey%>">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="20">
		<table width="100%" cellpadding="0" cellspacing="0" class="a">
		<tr>
			<td>
				검색결과 : <b><%= oitem.FTotalCount%></b>
				&nbsp;
				페이지 : <b><%= page %> /<%=  oitem.FTotalpage %></b>
			</td>
			<td align="right">
			</td>
		</tr>
		</table>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td align="center"><input type="checkbox" name="chkAll" onClick="jsChkAll();"></td>
	<td width="60">itemID</td>
	<td width=50> 이미지</td>
	<td width="100">브랜드ID</td>
	<td>해외<Br>상품명</td>
	<td>온라인<Br>상품명</td>

	<% if currencyunit<> "" then %>
		<td width="60">해외<Br>판매가(@)</td>
		<td width="60">해외<Br>판매가(원)</td>
		<td width="60">해외<Br>배수(계산)</td>
	<% end if %>

	<td width="60">해외<Br>사용</td>
	<td width="60">온라인<Br>판매가</td>
	<td width="60">온라인<Br>매입가</td>
	<td width="50">온라인<Br>판매여부</td>
	<td width="50">온라인<Br>사용여부</td>
	<td width="50">온라인<Br>한정여부</td>
	<td width="50">언어<br>사용여부</td>
	<td width="60">상품<br>무게</td>
</tr>
<% if oitem.FresultCount<1 then %>
<tr bgcolor="#FFFFFF">
	<td colspan="20" align="center">[검색결과가 없습니다.]</td>
</tr>
<% end if %>
<% if oitem.FresultCount > 0 then %>
<% for i=0 to oitem.FresultCount-1 %>
<tr class="a" height="25" bgcolor="#FFFFFF">
	<td align="center">
		<input type="checkbox" name="chkitem" value="<%= oitem.FItemList(i).FItemId %>">
	</td>
	<td align="center">
		<a href="http://www.10x10.co.kr/shopping/category_prd.asp?itemid=<%= oitem.FItemList(i).Fitemid %>" target="_blank" title="미리보기">
		<%= oitem.FItemList(i).Fitemid %></a>
		</td>
	<td align="center"><img src="<%= oitem.FItemList(i).FSmallImage %>" width="50" height="50" border="0"></td>
	<td align="left"><%= oitem.FItemList(i).Fmakerid %></td>
	<td align="left"><% =oitem.FItemList(i).Fitemname %></td>
	<td align="left"><% =oitem.FItemList(i).Fitemname10x10 %></td>

	<% if currencyunit<> "" then %>
		<%
		if oitem.FItemList(i).flinkpricetype="2" then
			onlineBasePrc = oitem.FItemList(i).fforeignorgprice
		else
			onlineBasePrc = oitem.FItemList(i).Fsellcash
		end if
		%>
		<td align="right"><%= oitem.FItemList(i).forgprice %></td>
		<td align="right">
			<%= oitem.FItemList(i).fwonprice %>
		</td>
		<td align="right">
			<%
			if onlineBasePrc<>0 and onlineBasePrc<>"" and oitem.FItemList(i).fwonprice<>0 and oitem.FItemList(i).fwonprice<>"" then
				response.write round( oitem.FItemList(i).fwonprice/onlineBasePrc ,2)
			else
				response.write 0
			end if
			%>
		</td>
	<% end if %>

	<td align="center"><%= oitem.FItemList(i).fsiteisusing %></td>
	<td align="right">
	<%
		'Response.Write "<a href=""javascript:editItemPriceInfo('" & oitem.FItemList(i).Fitemid & "')"" title='판매가 및 공급가 설정'>" & FormatNumber(oitem.FItemList(i).Forgprice,0) & "</a>"
		Response.Write "" & FormatNumber(oitem.FItemList(i).Forgprice,0) & ""
		'할인가
		if oitem.FItemList(i).Fsailyn="Y" then
			Response.Write "<br><font color=#F08050>(할)" & FormatNumber(oitem.FItemList(i).Fsailprice,0) & "</font>"
		end if
		'쿠폰가
		if oitem.FItemList(i).FitemCouponYn="Y" then
			Select Case oitem.FItemList(i).FitemCouponType
				Case "1"
					'Response.Write "<br><font color=#5080F0>(쿠)" & FormatNumber(oitem.FItemList(i).Forgprice*((100-oitem.FItemList(i).FitemCouponValue)/100),0) & "</font>"
				Case "2"
					'Response.Write "<br><font color=#5080F0>(쿠)" & FormatNumber(oitem.FItemList(i).Forgprice-oitem.FItemList(i).FitemCouponValue,0) & "</font>"
			end Select
		end if
	%>
	</td>
	<td align="center"><%= FormatNumber(oitem.FItemList(i).Fbuycash,0) %></td>
	<td align="center"><%= fnColor(oitem.FItemList(i).Fsellyn,"yn") %></td>
	<td align="center"><%= fnColor(oitem.FItemList(i).Fisusing,"yn") %></td>
	<td align="center"><%= fnColor(oitem.FItemList(i).Flimityn,"yn") %></td>
	<td align="center"><%= fnColor(oitem.FItemList(i).fuseyn,"yn") %></td>
	<td align="center"><%= FormatNumber(oitem.FItemList(i).FitemWeight,0) %>g</td>
</tr>
<% next %>

<tr height="25" bgcolor="FFFFFF">
	<td colspan="20" align="center">
		<% if oitem.HasPreScroll then %>
		<a href="javascript:NextPage('<%= oitem.StartScrollPage-1 %>')">[pre]</a>
		<% else %>
			[pre]
		<% end if %>

		<% for i=0 + oitem.StartScrollPage to oitem.FScrollCount + oitem.StartScrollPage - 1 %>
			<% if i>oitem.FTotalpage then Exit for %>
			<% if CStr(page)=CStr(i) then %>
			<font color="red">[<%= i %>]</font>
			<% else %>
			<a href="javascript:NextPage('<%= i %>')">[<%= i %>]</a>
			<% end if %>
		<% next %>

		<% if oitem.HasNextScroll then %>
			<a href="javascript:NextPage('<%= i %>')">[next]</a>
		<% else %>
			[next]
		<% end if %>
	</td>
</tr>
<% end if %>
</table>
</form>
<% set oitem = nothing %>
<iframe name="FrameCKP" src="about:blank" frameborder="0" width="600" height="200"></iframe>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->