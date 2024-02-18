<%@ codepage="65001" language="VBScript" %>
<% option explicit %>
<% response.Charset="UTF-8" %>
<% session.codePage = 65001 %>
<%
'####################################################
' Description :  온라인 해외판매상품
' History : 2013.05.06 한용민 생성
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->

<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=UTF-8">
<script language="JavaScript" src="/js/xl.js"></script>
<script language="JavaScript" src="/js/common.js"></script>
<script language="JavaScript" src="/js/report.js"></script>
<script language="JavaScript" src="/cscenter/js/cscenter.js"></script>
<script language="JavaScript" src="/js/calendar.js"></script>

<link rel="stylesheet" href="/css/scm.css" type="text/css">
<link href="/js/jqueryui/css/jquery-ui.css" rel="stylesheet">
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript" src="/js/jqueryui/jquery-ui-1.10.2.custom.min.js"></script>
<script language='javascript'>

function PopMenuHelp(menupos){
	var popwin = window.open('/designer/menu/help.asp?menupos=' + menupos,'PopMenuHelp_a','width=900, height=600, scrollbars=yes,resizable=yes');
	popwin.focus();
}

function PopMenuEdit(menupos){
	var popwin = window.open('/admin/menu/pop_menu_edit.asp?mid=' + menupos,'PopMenuEdit','width=600, height=400, scrollbars=yes,resizable=yes');
	popwin.focus();
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
<!-- #include virtual="/lib/classes/items/overseas/overseasCls.asp"-->
<!-- #include virtual="/lib/classes/items/new_itemcls.asp"-->
<%
Dim vItemID, vCountryCd, oitem, cOverSeas, vItemName, vItemContent, vItemCopy, vOriginListImage, vOriginItemName, MayMultiple
dim vOriginMakerID, vOriginOrgPrice, vOriginSellCash, vItemSource, vItemSize, vMakerName, vSourceArea, useyn, isexistsmultilang, makerid
dim oprice, oitemoption, keywords, offcontractinfo, isoffcontract, i, areaCode11st, tmpcheckedyn
Dim vsitename, sitecountrylang, vMultiLang
Dim vSiteisusing, wonprice
	vItemID = Request("itemid")
    vSitename = Request("sitename")
	vMultiLang = request("ml")

tmpcheckedyn="N"
isoffcontract = false
isexistsmultilang = false
	
'/대표언어팩 가져옴
sitecountrylang = getcountrylang(vsitename)

if not isarray(sitecountrylang) then
	Response.Write "<script>alert('언어팩이 지정되어 있지 않습니다.\n[ON]해외상품관리>>해외환율관리에서 등록하고 사용하세요.');window.close()</script>"
	session.codePage = 949 : dbget.close() : Response.End
else
	if ubound(sitecountrylang,2)="0" then
		vMultiLang = sitecountrylang(1,0)
	end if
	'response.write ubound(sitecountrylang,2)
end if

If vMultiLang = "" Then
	vMultiLang = "KR"
End If
If vSitename = "ITSWEB" Then
	vMultiLang = "ITSWEB"
End If
If vSitename = "11STMY" Then
	vMultiLang = "EN"
End If

If vItemID <> "" Then
	set cOverSeas = new COverSeasItem
		cOverSeas.FRectItemID = vItemID
		cOverSeas.FRectSitename = vSitename
		cOverSeas.FRectMultiLanguage = vMultiLang
		cOverSeas.GetOverSeasTargetItem

		vItemName = cOverSeas.FOneItem.Fitemname
		vItemContent = cOverSeas.FOneItem.Fitemcontent
		vItemCopy = cOverSeas.FOneItem.Fitemcopy
		vItemSource = cOverSeas.FOneItem.Fitemsource
		vItemSize = cOverSeas.FOneItem.Fitemsize
		vMakerName = cOverSeas.FOneItem.Fmakername
		vSourceArea = cOverSeas.FOneItem.Fsourcearea
		useyn = cOverSeas.FOneItem.fuseyn
		keywords = cOverSeas.FOneItem.fkeywords
		areaCode11st = cOverSeas.FOneItem.FareaCode11st
        vCountryCd     = cOverSeas.FOneItem.fcountrycd
        vSiteisusing   = cOverSeas.FOneItem.fSiteisusing

		if useyn <> "" then
			isexistsmultilang=true
		end if
	set cOverSeas = Nothing

	set oitem = new CItemInfo
		oitem.FRectItemID = vItemID
		oitem.GetOneItemInfo

		if vMultiLang = "KR" then
			vItemName = NullFillWith(vItemName,oitem.FOneItem.Fitemname)
		end if

		vOriginListImage = oitem.FOneItem.FListImage
		vOriginItemName = oitem.FOneItem.FItemName
		vOriginMakerID = oitem.FOneItem.FMakerid
		makerid = oitem.FOneItem.fmakerid
		vOriginSellCash = oitem.FOneItem.FSellcash
		vOriginOrgPrice = oitem.FOneItem.FOrgPrice
		vItemSource = NullFillWith(vItemSource,oitem.FOneItem.Fitemsource)
		vItemSize = NullFillWith(vItemSize,oitem.FOneItem.Fitemsize)
		vMakerName = NullFillWith(vMakerName,oitem.FOneItem.Fmakername)
		vSourceArea = NullFillWith(vSourceArea,oitem.FOneItem.Fsourcearea)
	set oitem = Nothing
Else
	Response.Write "<script>alert('잘못된 경로입니다.');window.close()</script>"
	session.codePage = 949 : dbget.close() : Response.End
End IF

if useyn="" or isnull(useyn) then useyn="Y"
if vSiteisusing="" or isnull(vSiteisusing) then vSiteisusing="Y"

'/가격
set oprice = new COverSeasItem
	oprice.FRectItemID = vItemID
	oprice.FRectSitename = vSitename

	if vItemID<>"" then
		oprice.GetOverSeasItemprice
	end if

'/옵션
If vSitename = "11STMY" Then
	set oitemoption = new COverSeasItem
		oitemoption.FRectItemID = vItemID
		If vItemID <> "" Then
			oitemoption.getItem11STMYOptionInfo
		End If
Else
	set oitemoption = new COverSeasItem
	oitemoption.FRectItemID = vItemID

	If isexistsmultilang Then	
		if vItemID<>"" then
			oitemoption.frectCountryCd = vCountryCd
			oitemoption.GetOverSeasItemOptionList
		end if
	Else
		if vItemID<>"" then
			oitemoption.frectCountryCd = vMultiLang
			oitemoption.GetOverSeasItemOptionList
		end if
	End IF
End If

'//계약사항
offcontractinfo = getitemshopcontractinfo("'7'", "", makerid)
if isarray(offcontractinfo) then isoffcontract=true

%>

<script type="text/javascript">

function goMultiLng(ml){
	location.href='/admin/itemmaster/overseas/popitemcontent.asp?itemid=<%=vItemID%>&sitename=<%=vSitename%>&ml='+ml;
}

function autocapypaste(){
	var OriginItemName="<%= replace(vOriginItemName, """","'") %>";

	document.frmreg.itemname.value = OriginItemName;
}

function goChangeLang(a){
	document.location.href = "<%=CurrURL()%>?countrycd="+a+"&itemid=<%=vItemID%>";
}

//바코드관리
function upcheManageCode(itemcode){
	var popupcheManageCode = window.open('/admin/stock/popUpcheManageCode.asp?itemcode=' + itemcode,'popupcheManageCode','width=550,height=400,resizable=yes,scrollbars=yes');
	popupcheManageCode.focus();
}

//저장
function goSubmit(){
	var tmpcountrycd='';
	
	<% 'if vSitename="ITSWEB" or vSitename="WSLWEB" then %>
		<% 'if not(isoffcontract) then %>
			//alert('오프라인 계약이 없습니다. 계약을 등록해 주세요.');
			//return;
		<% 'end if %>
	<% 'end if %>

	var countrycd = document.getElementsByName("countrycd")
	for(var i=0; i < countrycd.length;i++){
		if (countrycd[i].checked){
			tmpcountrycd = countrycd[i].value
		}
	}
	if(tmpcountrycd == ""){
		alert("언어팩을 선택하세요.");
		return;
	}
	if(document.frmreg.itemname.value == ""){
		alert("상품명을 입력하세요.");
		document.frmreg.itemname.focus();
		return;
	}

	<% If vSitename="WSLWEB" Then %>
//		if(document.frmreg.itemsource.value == ""){
//			alert("재료를 입력하세요.");
//			document.frmreg.itemsource.focus();
//			return;
//		}
		if(document.frmreg.sourcearea.value == ""){
			alert("원산지를 입력하세요.");
			document.frmreg.sourcearea.focus();
			return;
		}
	<% End If %>

	<% If vsitename = "11STMY" Then %>
		if(document.frmreg.areaCode11st.value == ""){
			alert("원산지를 입력하세요.");
			document.frmreg.areaCode11st.focus();
			return;
		}
	<% End If %>

	<% 'if (ucase(vCountryCd)="KR") then %>
//		if(document.frmreg.itemcontent.value == ""){
//			alert("간략설명을 입력하세요.");
//			document.frmreg.itemcontent.focus();
//			return;
//		}
	<% 'end if %>

	//가격체크
	for (var i=0; i < frmreg.pricecount.value; i++){
		<% if C_ADMIN_AUTH then %>
			if((eval("frmreg.orgprice"+i).value=="")){
				alert(eval("frmreg.currencyUnit"+i).value + "의 해외판매가 입력하세요");
				eval("frmreg.orgprice"+i).focus();
				return;
			}

			if((eval("frmreg.wonprice"+i).value=="")){
				alert(eval("frmreg.currencyUnit"+i).value + "의 원화를 입력하세요");
				eval("frmreg.wonprice"+i).focus();
				return;
			}
		<% else %>
			if((eval("frmreg.orgprice"+i).value=="")||(eval("frmreg.orgprice"+i).value=="0")){
				alert(eval("frmreg.currencyUnit"+i).value + "의 해외판매가 입력하세요");
				eval("frmreg.orgprice"+i).focus();
				return;
			}

			if((eval("frmreg.wonprice"+i).value=="")||(eval("frmreg.wonprice"+i).value=="0")){
				alert(eval("frmreg.currencyUnit"+i).value + "의 원화를 입력하세요");
				eval("frmreg.wonprice"+i).focus();
				return;
			}
		<% end if %>
	}

	if(document.frmreg.useyn.value == ""){
		alert("언어팩 사용여부를 선택하세요.");
		document.frmreg.useyn.focus();
		return;
	}

	if(document.frmreg.Siteisusing.value == ""){
		alert("사이트 사용여부를 선택하세요.");
		document.frmreg.Siteisusing.focus();
		return;
	}
	
	if(confirm("사이트("+ frmreg.sitename.value +") 언어팩("+ tmpcountrycd +") 내용을 저장 하시겠습니까?")){
		document.frmreg.submit();
	}
}

//해외가격 수정시 계산
function CheckThisByForeign(tmpi){
    var upfrm = document.frmreg;

    var onlineBasePrc=0;
    var vOriginSellCash = <%=vOriginSellCash%>;
    var vOriginOrgPrice = <%=vOriginOrgPrice%>;

    if (eval("upfrm.linkPriceType"+tmpi).value=="2"){
        onlineBasePrc = vOriginOrgPrice;
    }else{
        onlineBasePrc = vOriginSellCash;
    }

    var exchangeRate = eval("upfrm.exchangeRate"+tmpi).value;		//환율
    var orgprice     = eval("upfrm.orgprice"+tmpi).value;		//해외가격

    eval("upfrm.wonprice"+tmpi).value = Math.round((orgprice) * exchangeRate).toFixed(0);  //원화가격
    eval("upfrm.multiplerate"+tmpi).value = eval("upfrm.wonprice"+tmpi).value/onlineBasePrc;
}

//선택한 원화계산
function CheckThismrate(tmpi){
	var upfrm = document.frmreg;

    var onlineBasePrc=0;
    var vOriginSellCash = <%=vOriginSellCash%>;
    var vOriginOrgPrice = <%=vOriginOrgPrice%>;

    if (eval("upfrm.linkPriceType"+tmpi).value=="2"){
        onlineBasePrc = vOriginOrgPrice;
    }else{
        onlineBasePrc = vOriginSellCash;
    }

    //var orgprice = eval("upfrm.orgprice"+tmpi).value;		//해외소비자가
    var exchangeRate = eval("upfrm.exchangeRate"+tmpi).value;		//환율
    var multiplerate = eval("upfrm.multiplerate"+tmpi).value;		//배수
    //var multiplerate = Math.round(eval("upfrm.multiplerate"+tmpi).value).toFixed(2);		//배수

	eval("upfrm.wonprice"+tmpi).value = Math.round((onlineBasePrc) * multiplerate).toFixed(0);  //원화가격
	//eval("upfrm.orgprice"+tmpi).value = Math.round((onlineBasePrc) * multiplerate/exchangeRate*100).toFixed(2)/100;  //해외가격 차후 소수점 체크
	eval("upfrm.orgprice"+tmpi).value = Math.ceil( (onlineBasePrc * multiplerate/exchangeRate)*2 ) / 2;    //해외가격
}

//등록 안된 상품 가격 전체 셋팅
function CheckThismrate_auto(){
    var upfrm = document.frmreg;

    var exchangeRate = 0;
	var multiplerate = 0;
    var onlineBasePrc=0;
    var vOriginSellCash = <%=vOriginSellCash%>;
    var vOriginOrgPrice = <%=vOriginOrgPrice%>;
	var pricecnt = '<%= oprice.FResultCount %>';

	for (var i=0; i<pricecnt; i++){
		//상품 가격 미등록 상태
		if (eval("upfrm.pricenotreg"+i).value=="o"){
		    onlineBasePrc=0;
		    if (eval("upfrm.linkPriceType"+i).value=="2"){
		        onlineBasePrc = vOriginOrgPrice;
		    }else{
		        onlineBasePrc = vOriginSellCash;
		    }

			exchangeRate = 0;
			multiplerate = 0;
		    exchangeRate = eval("upfrm.exchangeRate"+i).value;		//환율
		    multiplerate = eval("upfrm.multiplerate"+i).value;		//배수
		    //multiplerate = Math.round(eval("upfrm.multiplerate"+i).value).toFixed(2);		//배수반올림을 하면 안됨. 1.5나 1.7이 2.0이 되버림

			eval("upfrm.wonprice"+i).value = Math.round((onlineBasePrc) * multiplerate).toFixed(0);  //원화가격
			//eval("upfrm.orgprice"+i).value = Math.round((onlineBasePrc) * multiplerate/exchangeRate*100).toFixed(2)/100;  //해외가격 차후 소수점 체크
			eval("upfrm.orgprice"+i).value = Math.ceil( (onlineBasePrc * multiplerate/exchangeRate)*2 ) / 2;  //해외가격
		}
	}
}

$(function(){
  	$(".rdoUsing").buttonset().children().next().attr("style","font-size:11px;");
  	$(".rdopriceisusing").buttonset().children().next().attr("style","font-size:11px;");

	//첫로딩시 상품가격 셋팅
	CheckThismrate_auto()
	//setTimeout("CheckThismrate_auto()",500)
});

</script>

<form name="frmreg" method="post" action="/admin/itemmaster/overseas/itemContentProc.asp" style="margin:0px;">
<input type="hidden" name="itemid" value="<%=vItemID%>">
<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<tr>
	<td bgcolor="#FFFFFF" colspan="2">
		<table width="100%" border="0" class="a">
		<tr>
			<td width="100"><img src="<%=vOriginListImage%>" width="100" height="100"></td>
			<td valign="top">
				<table width="100%" border="0" class="a">
				<tr>
					<td height="23">상품명 : <%=vOriginItemName%>&nbsp;&nbsp;&nbsp;<input type="button" value="상품명 입력란에 넣기" class="button" style="width:130px;" onClick="autocapypaste();"></td>
				</tr>
				<tr>
					<td height="23">상품코드 : <%=vItemID%> - [<a href="http://www.10x10.co.kr/shopping/category_prd.asp?itemid=<%=vItemID%>" target="_blank">상품상세보기페이지</a>]</td>
				</tr>
				<tr>
					<td height="23">브랜드ID : <%=vOriginMakerID%></td>
				</tr>
				<tr>
					<td height="23">소비자가 : <%=FormatNumber(vOriginOrgPrice,0)%> / 판매가 : <%=FormatNumber(vOriginSellCash,0)%></td>
				</tr>
				</table>
			</td>
		</tr>
		</table>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td nowrap width=100>사이트</td>
	<td bgcolor="#FFFFFF" align="left">
		<% ''drawSelectboxMultiSiteSitename "sitename", vsitename, " onChange='goChangeSite(this.value);'" %>
		<%= getMultiSiteSitenameByCode(vsitename) %>
		<input type="hidden" name="sitename" value="<%= vsitename %>">

	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td nowrap width=100>* 언어 SET</td>
	<td bgcolor="#FFFFFF" align="left">
		<%
		dim tmpcountrycd
		%>
		<% if isarray(sitecountrylang) then %>
			<%
			for i = 0 to ubound(sitecountrylang,2)

			tmpcheckedyn = "N"
			if ucase(vMultiLang)=ucase(sitecountrylang(1,i)) then
				tmpcheckedyn = "Y"
			end if
			%>
				<input type="radio" name="countrycd" value="<%= sitecountrylang(1,i) %>" onclick="goMultiLng('<%= sitecountrylang(1,i) %>');return false;" <% if tmpcheckedyn = "Y" then response.write " checked" %>><%= sitecountrylang(1,i) %>
			<%
			tmpcheckedyn = "N"

			next
			%>
		<% end if %>

		<%
		'/사용안함
		if false then
		%>
			<% if not(isexistsmultilang) then %>
				<input type="radio" name="countrycd" value="X" <% if (vCountryCd="" or isnull(vCountryCd)) then response.write " checked" %>>언어팩사용안함
			<% end if %>
		<% end if %>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td nowrap width=100>* 상품명</td>
	<td bgcolor="#FFFFFF" align="left"><input type="text" class="text" name="itemname" value="<%=vItemName%>" size="95" maxlangth="60"></td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td nowrap width=100>간략설명</td>
	<td bgcolor="#FFFFFF" align="left"><textarea name="itemcontent" cols="71" rows="16"><%=vItemContent%></textarea></td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td nowrap width=100>상품카피</td>
	<td bgcolor="#FFFFFF" align="left"><input type="text" class="text" name="itemcopy" value="<%=vItemCopy%>" size="95" maxlangth="250"></td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td nowrap width=100>재료</td>
	<td bgcolor="#FFFFFF" align="left"><input type="text" class="text" name="itemsource" value="<%=vItemSource%>" size="95" maxlangth="128"></td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td nowrap width=100>크기</td>
	<td bgcolor="#FFFFFF" align="left"><input type="text" class="text" name="itemsize" value="<%=vItemSize%>" size="95" maxlangth="128"></td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td nowrap width=100>제조사</td>
	<td bgcolor="#FFFFFF" align="left"><input type="text" class="text" name="makername" value="<%=vMakerName%>" size="95" maxlangth="64"></td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td nowrap width=100>* 원산지</td>
	<td bgcolor="#FFFFFF" align="left"><input type="text" class="text" name="sourcearea" value="<%=vSourceArea%>" size="95" maxlangth="128"></td>
</tr>

<% If vsitename = "11STMY" Then %>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td nowrap width=100>11번가<br>원산지 코드</td>
		<td bgcolor="#FFFFFF" align="left">
			<select class="select" name="areaCode11st">
				<option value="">-Choice-</option>
				<option value="1449" <%= chkiif(areaCode11st="1449", "selected", "" )%> >KOREA</option>
				<option value="1287" <%= chkiif(areaCode11st="1287", "selected", "" )%> >CHINA</option>
				<option value="1294" <%= chkiif(areaCode11st="1294", "selected", "" )%> >TAIWAN</option>
				<option value="1357" <%= chkiif(areaCode11st="1357", "selected", "" )%> >GERMANY</option>
				<option value="1399" <%= chkiif(areaCode11st="1399", "selected", "" )%> >GUATEMALA</option>
				<option value="1450" <%= chkiif(areaCode11st="1450", "selected", "" )%> >HONGKONG</option>
				<option value="1284" <%= chkiif(areaCode11st="1284", "selected", "" )%> >INDONESIA</option>
				<option value="1285" <%= chkiif(areaCode11st="1285", "selected", "" )%> >JAPAN</option>
				<option value="1354" <%= chkiif(areaCode11st="1354", "selected", "" )%> >NETHERLAND</option>
				<option value="1394" <%= chkiif(areaCode11st="1394", "selected", "" )%> >PORTUGAL</option>
				<option value="1430" <%= chkiif(areaCode11st="1430", "selected", "" )%> >CHILE</option>
				<option value="1379" <%= chkiif(areaCode11st="1379", "selected", "" )%> >SLOVAKIA</option>
				<option value="1378" <%= chkiif(areaCode11st="1378", "selected", "" )%> >SPAIN</option>
				<option value="1293" <%= chkiif(areaCode11st="1293", "selected", "" )%> >THAILAND</option>
				<option value="1335" <%= chkiif(areaCode11st="1335", "selected", "" )%> >UGANDA</option>
				<option value="1386" <%= chkiif(areaCode11st="1386", "selected", "" )%> >UK</option>
				<option value="1265" <%= chkiif(areaCode11st="1265", "selected", "" )%> >VIETNAM</option>
				<option value="1405" <%= chkiif(areaCode11st="1405", "selected", "" )%> >USA</option>
				<option value="1250" <%= chkiif(areaCode11st="1250", "selected", "" )%> >EUROPE</option>
				<option value="1271" <%= chkiif(areaCode11st="1271", "selected", "" )%> >SINGAPORE</option>
				<option value="1283" <%= chkiif(areaCode11st="1283", "selected", "" )%> >INDIA</option>
				<option value="1297" <%= chkiif(areaCode11st="1297", "selected", "" )%> >PAKISTAN</option>
				<option value="1316" <%= chkiif(areaCode11st="1316", "selected", "" )%> >MAURITIUS</option>
				<option value="1362" <%= chkiif(areaCode11st="1362", "selected", "" )%> >LITUANIA</option>
				<option value="1376" <%= chkiif(areaCode11st="1376", "selected", "" )%> >SWEDEN</option>
				<option value="1387" <%= chkiif(areaCode11st="1387", "selected", "" )%> >AUSTRIA</option>
				<option value="1389" <%= chkiif(areaCode11st="1389", "selected", "" )%> >ITALY</option>
				<option value="1390" <%= chkiif(areaCode11st="1390", "selected", "" )%> >CZECH</option>
				<option value="1393" <%= chkiif(areaCode11st="1393", "selected", "" )%> >TURKEY</option>
				<option value="1395" <%= chkiif(areaCode11st="1395", "selected", "" )%> >POLAND</option>
				<option value="1396" <%= chkiif(areaCode11st="1396", "selected", "" )%> >FRANCE</option>
				<option value="1397" <%= chkiif(areaCode11st="1397", "selected", "" )%> >FINLAND</option>
				<option value="1404" <%= chkiif(areaCode11st="1404", "selected", "" )%> >MEXICO</option>
				<option value="1417" <%= chkiif(areaCode11st="1417", "selected", "" )%> >CANADA</option>
				<option value="1441" <%= chkiif(areaCode11st="1441", "selected", "" )%> >AUSTRALIA</option>
				<option value="1259" <%= chkiif(areaCode11st="1259", "selected", "" )%> >MALAYSIA</option>
				<option value="1259" <%= chkiif(areaCode11st="1298", "selected", "" )%> >PHILIPPINES</option>
				<option value="1255" <%= chkiif(areaCode11st="1255", "selected", "" )%> >NEPAL</option>
			</select>
		</td>
	</tr>
<% End If %>

<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td nowrap width=100>키워드</td>
	<td bgcolor="#FFFFFF" align="left"><input type="text" class="text" name="keywords" value="<%=keywords%>" size="95" maxlangth="128"></td>
</tr>
<%
i=0

If oprice.FResultCount > 0 Then
%>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td nowrap width=100>
			* 가격
		</td>
		<td bgcolor="#FFFFFF" align="left">
			<table cellpadding="3" cellspacing="1" border="0" class="a" width="100%" bgcolor="<%= adminColor("tabletop") %>" align="center">
			<%
			For i=0 To oprice.FResultCount - 1
	
		    MayMultiple = 0
	
			if (oprice.FITemList(i).fwonprice=0 or oprice.FITemList(i).fwonprice="") and (oprice.FITemList(i).forgprice<>0) then
				wonprice = round((oprice.FITemList(i).forgprice * oprice.FItemlist(i).fexchangeRate) * oprice.FItemlist(i).fmultiplerate,0)
		    elseif (oprice.FITemList(i).fwonprice=0 or oprice.FITemList(i).fwonprice="") and (oprice.FITemList(i).forgprice=0) then ''최초 디폴트값 으로 설정할 경우
	            wonprice = oprice.FItemList(i).fwonprice '' 눌르게끔
			else
				wonprice = oprice.FItemList(i).fwonprice
			end if

			if oprice.FItemlist(i).flinkPriceType=1 then
				if wonprice=0 or vOriginSellCash=0 then
				MayMultiple=1
				else
		        MayMultiple = round(wonprice/vOriginSellCash,2)
				end if
		    elseif oprice.FItemlist(i).flinkPriceType=2 then
				if wonprice=0 or vOriginSellCash=0 then
				MayMultiple=1
				else
		        MayMultiple = round(wonprice/vOriginOrgPrice,2)
				end if
		    end if
			%>
			<tr align="center">
				<td bgcolor="#FFFFFF">
					<input type="hidden" name="currencyUnit<%= i %>" value="<%=oprice.FITemList(i).fcurrencyUnit%>" />
					화폐:<%=oprice.FITemList(i).fcurrencyUnit%>
				</td>
				<td bgcolor="#FFFFFF">
					해외판매가:<input type="text" name="orgprice<%= i %>" size=10 value="<%= oprice.FITemList(i).forgprice %>" size=4 maxlength=10 onKeyup="CheckThisByForeign(<%= i %>)" />
				</td>
				<td bgcolor="#FFFFFF">
					원화:<input class="text_ro" type="text" name="wonprice<%= i %>" size=10 value="<%= wonprice %>" size=4 maxlength=10 readonly />
				</td>
				<td bgcolor="#FFFFFF">
					환율:<%= oprice.FItemlist(i).fexchangeRate %> / <%= oprice.FItemlist(i).getlinkPriceTypeName %> 대비 <%= oprice.FItemlist(i).fmultiplerate%> 배
					<input type="hidden" name="exchangeRate<%= i %>" value="<%= oprice.FItemlist(i).fexchangeRate %>" />
				</td>
				<td bgcolor="#FFFFFF">
					* 가격등록 : <%= CHKIIF(oprice.FItemList(i).FNotReg="o" ,"<font color=red><b>미등록</b></font>","<font color=blue><b>등록</b></font>") %>
					<input type="hidden" name="pricenotreg<%= i %>" value="<%= oprice.FItemList(i).FNotReg %>" />
				</td>
				<td bgcolor="#FFFFFF">
				    <input type="hidden" name="linkPriceType<%= i %>" value="<%= oprice.FItemlist(i).flinkPriceType %>" />
					배수:<input type="text" name="multiplerate<%= i %>" value="<%= CHKIIF(MayMultiple<>0,MayMultiple,oprice.FItemlist(i).fmultiplerate) %>" size=1 maxlength=3 onKeyup="CheckThismrate(<%= i %>)">
					<input type="button" value="재계산" onClick="CheckThismrate(<%= i %>)">
				</td>
			</tr>
			<%
			next
			%>
			</table>
		</td>
	</tr>
<% end if %>
<input type="hidden" name="pricecount" value="<%=i%>" />
<!-- #include virtual="/lib/BarcodeFunction.asp"-->
<%
i = 0
%>
<tr>
	<td width="100" bgcolor="<%= adminColor("tabletop") %>" align="center">옵션</td>
	<td valign="top" bgcolor="#FFFFFF">
		※ 처음등록시 옵션이 있는 상품일 경우, 디폴트로 옵션이 보여지고 옵션등록 여부가 미등록상태로 나타납니다. 수정이 필요한 경우에만 고치시면 됩니다.<br><br>
		<table cellpadding="3" cellspacing="1" border="0" class="a" bgcolor="<%= adminColor("tabletop") %>">
		<%
		If oitemoption.FResultCount > 0 Then
		%>
			<% For i=0 To oitemoption.FResultCount - 1 %>
				<tr>
					<td bgcolor="#FFFFFF" align="center">
						<input type="hidden" name="itemoption<%=i%>" value="<%= oitemoption.FITemList(i).FItemOption %>" /><%= oitemoption.FITemList(i).FItemOption %>
	
						<% if oitemoption.FItemList(i).Fitemoption="0000" then %>
							* 옵션없음
							<input type="hidden" name="optiontypename<%=i%>" value="<%= oitemoption.FITemList(i).FOptionTypeName %>" />
							<input type="hidden" name="optionname<%=i%>" value="<%= oitemoption.FITemList(i).FOptionName %>" />
							<input type="hidden" name="optisusing<%=i%>" value="<%= oitemoption.FITemList(i).FOptIsUsing %>" />
						<% else %>
							<input type="text" name="optiontypename<%= i %>" value="<%= oitemoption.FITemList(i).FOptionTypeName %>" size="10" />
							<input type="text" name="optionname<%= i %>" value="<%= oitemoption.FITemList(i).FOptionName %>" size="30" />
							<span class="rdoUsing">
								<input type="radio" name="optisusing<%= i %>" id="rdoUsing<%= i %>_1" value="Y" <%= CHKIIF(oitemoption.FITemList(i).FOptIsUsing="Y","checked","") %> /><label for="rdoUsing<%= i %>_1">사용</label>
								<input type="radio" name="optisusing<%= i %>" id="rdoUsing<%= i %>_2" value="N" <%= CHKIIF(oitemoption.FITemList(i).FOptIsUsing="N","checked","") %> /><label for="rdoUsing<%= i %>_2">사용안함</label>
							</span>
							* 옵션등록 : <%= CHKIIF(oitemoption.FItemList(i).FNotReg="o" OR not(isexistsmultilang),"<font color=red><b>미등록</b></font>","<font color=blue><b>등록</b></font>") %>
						<% end if %>
	
						&nbsp;&nbsp;* 업체코드 :
						<% if oitemoption.FItemList(i).fupchemanagecode <> "" then %>
							<a href="javascript:upcheManageCode('<%= BF_MakeTenBarcode(oitemoption.FItemList(i).fitemgubun, oitemoption.FItemList(i).Fitemid, oitemoption.FItemList(i).Fitemoption) %>')" onfocus="this.blur()">
							<%= oitemoption.FItemList(i).fupchemanagecode %></a>
						<% else %>
							<a href="javascript:upcheManageCode('<%= BF_MakeTenBarcode(oitemoption.FItemList(i).fitemgubun, oitemoption.FItemList(i).Fitemid, oitemoption.FItemList(i).Fitemoption) %>')" onfocus="this.blur()">
							미등록</a>
						<% end if %>
	
						<% if vCountryCd="ITSWEB" or vCountryCd="WSLWEB" then %>
							&nbsp;&nbsp;* 오프상품사용 : <%= oitemoption.FItemList(i).foff_isusing %>
						<% end if %>
					</td>
				</tr>
			<% Next %>
		<% else %>
			<tr>
				<td bgcolor="#FFFFFF" align="center">
					&nbsp;&nbsp;&nbsp;&nbsp;- 업체코드 :
					<a href="javascript:upcheManageCode('<%= BF_MakeTenBarcode("10", vItemID, "0000") %>')" onfocus="this.blur()">
					미등록</a>
				</td>
			</tr>
		<% End IF %>
		</table>
	</td>
</tr>
<input type="hidden" name="optioncount" value="<%=i%>" />

<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td nowrap width=100>* 언어팩사용여부</td>
	<td bgcolor="#FFFFFF" align="left">
		<% drawSelectBoxisusingYN "useyn", useyn, "" %>
	</td>
</tr>

<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td nowrap width=100>* 사이트사용여부</td>
	<td bgcolor="#FFFFFF" align="left">
		<% drawSelectBoxisusingYN "Siteisusing", vSiteisusing, "" %>
	</td>
</tr>

<% if vSitename="ITSWEB" or vSitename="WSLWEB" then %>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td nowrap width=100> 오프계약정보</td>
		<td bgcolor="#FFFFFF" align="left">
			<table cellpadding="3" cellspacing="1" width="100%" border="0" class="a" bgcolor="<%= adminColor("tabletop") %>">
			<% if isarray(offcontractinfo) then %>
			<%
			i=0
			For i=0 To ubound(offcontractinfo,2)
			%>
			<tr>
				<td bgcolor="#FFFFFF">
					매장명 : <%= offcontractinfo(1,i) %>
				</td>
			</tr>
			<%
			next
			%>
			<% else %>
				<tr>
					<td bgcolor="#FFFFFF">
						계약정보 없음
					</td>
				</tr>
			<% end if %>
			</table>
		</td>
	</tr>
<% end if %>

<tr align="center" bgcolor="#FFFFFF">
	<td colspan=2>
		<input type="button" onClick="goSubmit(); return false;" value="저장" class="button">
	</td>
</tr>

</table>
</form>

<%
set oprice = nothing
set oitemoption = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<%
session.codePage = 949
%>