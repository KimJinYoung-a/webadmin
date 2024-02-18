<%@  codepage="65001" language="VBScript" %>
<% option explicit %>
<%
session.codePage = 65001
%>
<%
'####################################################
' Description :  온라인 해외판매상품
' History : 2013.05.06 한용민 생성
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead_utf8.asp"-->
<!-- #include virtual="/lib/function_utf8.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/items/overseas/overseasCls.asp"-->

<%
Dim oitem, i, page, itemid, itemname, makerid, cdl, cdm, cds, vCountryCd, sellyn, usingyn, limityn, danjongyn
dim sitecountrylang, vlangGubun, weightYn, sellcash1, sellcash2, sortDiv, siteisusing
Dim sitename, vpriceArrv, vpriceArrk, v, k, reloading, vcountryLangCDArrv, vcountryLangCDArrk, mwdiv, overSeaYn
	page = requestCheckvar(request("page"),10)
	vCountryCd	= requestCheckvar(request("countrycd"),32)
	vlangGubun	= requestCheckvar(request("langGubun"),32)
	itemid      = requestCheckvar(request("itemid"),255)
	itemname	= requestCheckvar(request("itemname"),64)
	makerid		= requestCheckvar(request("makerid"),32)
	sellyn		= requestCheckvar(request("sellyn"),1)
	usingyn		= requestCheckvar(request("usingyn"),1)
	cdl = requestCheckvar(request("cdl"),3)
	cdm = requestCheckvar(request("cdm"),3)
	cds = requestCheckvar(request("cds"),3)
	limityn = requestCheckvar(request("limityn"),1)
    sitename = requestCheckvar(request("sitename"),32)
    reloading		= requestCheckvar(request("reloading"),2)
	danjongyn   = requestCheckvar(request("danjongyn"),10)
	mwdiv		= requestCheckvar(request("mwdiv"),2)
	overSeaYn	= requestCheckvar(request("overSeaYn"),1)
	weightYn	= requestCheckvar(request("weightYn"),1)
	sellcash1	= requestCheckvar(request("sellcash1"),10)
	sellcash2	= requestCheckvar(request("sellcash2"),10)
	sortDiv		= requestCheckvar(request("sortDiv"),16)
	siteisusing		= requestCheckvar(request("siteisusing"),1)

if (page = "") then page = 1
if sitename="" then sitename="WSLWEB"
'if (vCountryCd = "") then vCountryCd = "o"			'2022-06-17 김진영 수정..전체추가
if reloading="" and sellyn="" then sellyn="YS"
if reloading="" and usingyn="" then usingyn="Y"
if reloading<>"ON" and overSeaYn="" then overSeaYn="Y"
if sortDiv="" then sortDiv="new"

if itemid<>"" then
	dim iA ,arrTemp,arrItemid
  itemid = replace(itemid,chr(13),"")
	arrTemp = Split(itemid,chr(10))

	iA = 0
	do while iA <= ubound(arrTemp)
		if Trim(arrTemp(iA))<>"" and isNumeric(Trim(arrTemp(iA))) then
			arrItemid = arrItemid & Trim(arrTemp(iA)) & ","
		end if
		iA = iA + 1
	loop

	if len(arrItemid)>0 then
		itemid = left(arrItemid,len(arrItemid)-1)
	else
		if Not(isNumeric(itemid)) then
			itemid = ""
		end if
	end if
end if

set oitem = new COverSeasItem
	oitem.FPageSize         = 30
	oitem.FCurrPage         = page
	oitem.FRectCountryCd	= vCountryCd
	oitem.FRectLangGubun	= vlangGubun
	oitem.FRectMakerid      = makerid
	oitem.FRectSellYN       = sellyn
	oitem.FRectIsUsing      = usingyn
	oitem.FRectItemid       = itemid
	oitem.FRectItemName     = itemname
	oitem.FRectCate_Large   = cdl
	oitem.FRectCate_Mid     = cdm
	oitem.FRectCate_Small   = cds
	oitem.FRectLimitYN		= limityn
    oitem.FRectSitename = Sitename
	oitem.FRectDanjongyn    = danjongyn
	oitem.FRectMWDiv        = mwdiv
	oitem.FRectIsOversea	= overSeaYn
	oitem.FRectIsWeight		= weightYn
	oitem.FRectSellcash1	= sellcash1
	oitem.FRectSellcash2	= sellcash2
	oitem.FRectsortDiv	= sortDiv
	oitem.FRectsiteisusing	= siteisusing

	If sitename <> "" Then
		oitem.GetForeignItemList
	else
		response.write "<script type='text/javascript'>"
		response.write "	alert('사이트를 선택하세요');"
		response.write "</script>"
	End If

'/대표언어팩 가져옴
sitecountrylang = getcountrylang(Sitename)
%>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type='text/javascript'>

function num_check(gb){
	if(gb == "1"){
		if(isNaN(document.frm.sellcash1.value) == true)
		{
			alert("숫자만 입력해주세요.");
			document.frm.sellcash1.value = "";
			document.frm.sellcash1.focus();
		}
	}else{
		if(isNaN(document.frm.sellcash2.value) == true)
		{
			alert("숫자만 입력해주세요.");
			document.frm.sellcash2.value = "";
			document.frm.sellcash2.focus();
		}
	}
}

function PopItemContent(iitemid){
	var popwin = window.open('/admin/itemmaster/overseas/popItemContent.asp?itemid=' + iitemid +'&sitename=<%=sitename%>','itemWeightEdit','width=1280,height=800,scrollbars=yes,resizable=yes')
	popwin.focus();
}

function NextPage(ipage){
	frm.page.value= ipage;
	frm.action='';
	frm.submit();
}

function uploadlanguage() {
	document.domain = "10x10.co.kr";
	var popwin = window.open('/common/item/foreign/popitem_foreign_language_excelupload.asp','addreg','width=1280,height=400,scrollbars=yes,resizable=yes');
	popwin.focus();
}

function downloadlanguage() {
	alert('1만건까지 다운로드 가능. 로딩중 기다려 주세요.');
	frm.action='/common/item/foreign/popitem_foreign_language_exceldownload.asp';
	frm.target='view';
	frm.submit();
	frm.action='';
	frm.target='';
	return false;
}

function CheckClick(identikey){
	for (i=0; i< frmArr.check.length; i++){
		if (frmArr.check[i].value==identikey){
			frmArr.check[i].checked=true
		}
	}
}

function totalCheck(){
	var f = document.frmArr;
	var objStr = "check";
	var chk_flag = true;
	for(var i=0; i<f.elements.length; i++) {
		if(f.elements[i].name == objStr) {
			if(!f.elements[i].checked) {
				chk_flag = f.elements[i].checked;
				break;
			}
		}
	}

	for(var i=0; i < f.elements.length; i++) {
		if(f.elements[i].name == objStr) {
			if(chk_flag) {
				f.elements[i].checked = false;
			} else {
				f.elements[i].checked = true;
			}
		}
	}
}

function checkallsiteisusing(){
    var pass = false;
	var PrdCode="";
	var selectsiteisusing="";

	if (frmArr.allsiteisusing.value=="") {
	    alert('일괄적용하실 사이트사용여부의 기준을 선택하세요.');
	    frmArr.allsiteisusing.focus();
	    return false;
	}

    $('input[name="check"]:checkbox:checked').each(function () {
		pass = true;
        prdCode = $(this).val();
		selectsiteisusing = "selectsiteisusing_"+prdCode;

		if (frmArr.allsiteisusing.value=="Y"){
			$('input:radio[name='+selectsiteisusing+']:input[value="Y"]').prop("checked", true);
		}else{
			$('input:radio[name='+selectsiteisusing+']:input[value="N"]').prop("checked", true);
		}
    });

	if (!pass) {
		alert('선택 아이템이 없습니다.');
		return;
	}
}

function regArr(){
	var pass = false;
	var passcount = 0;
	var selectsiteisusingval="";
	var siteisusing="";

    $('input[name="check"]:checkbox:checked').each(function () {
		pass = true;
    });

	if (!pass) {
		alert('선택 아이템이 없습니다.');
		return;
	}

	var PrdCode="";
	for (i=0; i< frmArr.check.length; i++){
		if (frmArr.check[i].checked == true){
			PrdCode = frmArr.check[i].value;

			if ( eval("frmArr.selectsiteisusing_" + PrdCode)[0].checked==true ){
				selectsiteisusingval="Y"
			} else if ( eval("frmArr.selectsiteisusing_" + PrdCode)[1].checked==true ){
				selectsiteisusingval="N"
			} else {
				alert('선택하신 상품중에 사이트사용여부가 미선택 되어 있는 상품이 있습니다.');
				return false;
			}
			$("#siteisusing_"+PrdCode).val(selectsiteisusingval);
		}
	}

	var ret = confirm('저장 하시겠습니까?');
	if (ret){
		frmArr.target="view";
		frmArr.mode.value = "arrmodi";
		frmArr.action = "/admin/itemmaster/overseas/overseasItemProcess.asp";
		frmArr.submit();
	}
}

</script>

<div id="calendarPopup" style="position: absolute; visibility: hidden; z-index: 2;"></div>

<form name="frm" method="get" style="margin:0px;">
<input type="hidden" name="menupos" value="<%= Request("menupos") %>">
<input type="hidden" name="reloading" value="ON">
<input type="hidden" name="page" >

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center">
	<td rowspan="4" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
	<td align="left" bgcolor="#ffffff">
		<table border="0" cellpadding="0" cellspacing="0" class="a">
		<tr>
			<td style="white-space:nowrap;">* 온라인브랜드 : <%	drawSelectBoxDesignerWithName "makerid", makerid %></td>
			<td style="white-space:nowrap;padding-left:5px;">* 온라인상품명 : <input type="text" class="text" name="itemname" value="<%= itemname %>" size="32" maxlength="32"></td>
			<td style="white-space:nowrap;padding-left:5px;">* 온라인상품코드 :</td> 
			<td style="white-space:nowrap;" rowspan="2">
				<textarea rows="3" cols="10" name="itemid" id="itemid"><%=replace(itemid,",",chr(10))%></textarea>
			</td>
		</tr> 
		<tr>
			<td style="white-space:nowrap;" colspan=2>* 온라인관리<!-- #include virtual="/common/module/categoryselectbox_utf8.asp"--></td> 
			<td></td>
			<td></td>
		</tr>
		</table>
	</td>
	<td rowspan="4" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="검색" onClick="javascript:document.frm.submit();">
	</td>
</tr>
<tr align="center">
	<td align="left" bgcolor="#ffffff">
		* 온라인판매여부 : <% drawSelectBoxSellYN "sellyn", sellyn %>
     	&nbsp;
     	* 온라인사용여부: <% drawSelectBoxisusingYN "usingyn", usingyn, " onchange='NextPage("""");'" %>
		&nbsp;
		* 온라인한정여부: <% drawSelectBoxisusingYN "limityn", limityn, " onchange='NextPage("""");'" %>
		&nbsp;
		* 온라인단종여부: <% drawSelectBoxDanjongYN "danjongyn", danjongyn %>
		&nbsp;
     	* 온라인거래구분:<% drawSelectBoxMWU "mwdiv", mwdiv %>
	</td>
</tr>
<tr align="center">
	<td align="left" bgcolor="#ffffff">
     	* 온라인판매가 :
     	<input type="text" class="text" name="sellcash1" value="<%=sellcash1%>" size="10" onkeyUp="num_check('1')">
     	~<input type="text" class="text" name="sellcash2" value="<%=sellcash2%>" size="10" onkeyUp="num_check('2')">
		&nbsp;
     	* 온라인해외배송여부 : <% drawSelectBoxisusingYN "overSeaYn", overSeaYn, " onchange='NextPage("""");'" %>
		&nbsp;
     	* 온라인무게등록여부 : <% drawSelectBoxisusingYN "weightYn", weightYn, " onchange='NextPage("""");'" %>
	</td>
</tr>
<tr align="center">
	<td align="left" bgcolor="#ffffff">
		<strong><font color="red">* 사이트 : </font></strong>
	    <% drawSelectboxMultiSiteSitename "sitename", sitename, " onchange='NextPage("""");'" %>
	    &nbsp;
	    <strong><font color="red">* 상품(언어팩)등록여부 : </font></strong>
		<select name="countrycd" class="select" onchange="NextPage('')">
			<option value="" <%= CHKIIF(vcountrycd="","selected","") %>>전체</option>
			<option value="x" <%= CHKIIF(vcountrycd="x","selected","") %>>미등록만</option>
			<option value="o" <%= CHKIIF(vcountrycd="o","selected","") %>>등록만</option>
		</select>
		&nbsp;
		* 언어팩 구분 : 
		<select name="langGubun" class="select">
			<option value="" >선택</option>
		<%
			If isarray(sitecountrylang) Then
				For i = 0 to ubound(sitecountrylang,2)
		%>
					<option value="<%= sitecountrylang(1,i) %>" <%= CHKIIF(vlangGubun=sitecountrylang(1,i),"selected","") %>><%= sitecountrylang(1,i) %></option>
		<%
				Next
			End If
		%>
		</select>
	    <!--* 화폐 : <% 'drawSelectBoxsitecurrencyunit sitename, "currencyunit", currencyunit, " onchange='NextPage("""");'" %>-->
		<!--* 해외상품사용여부 : <% 'drawSelectBoxUsingYN "useyn", useyn %>-->
		&nbsp;
		* 사이트사용여부 : <% drawSelectBoxUsingYN "siteisusing", siteisusing %>
	</td>
</tr>
</table>

<br>
<!-- 표 중간바 시작-->
<table width="100%" align="center" cellpadding="1" cellspacing="1" class="a">
<tr valign="bottom">
    <td align="left">
		<input class="button" type="button" value="선택상품 일괄수정" onClick="regArr();">
    </td>
    <td align="right">
    	<input type="button" onclick="uploadlanguage(); return false;" value="EN언어팩일괄업로드" class="button">
    	&nbsp;
    	<input type="button" onclick="downloadlanguage(); return false;" value="EN언어팩일괄다운로드" class="button">
    </td>
</tr>
</table>
<!-- 표 중간바 끝-->

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
				* 정렬 :
				<select name="sortDiv" class="select" onchange="NextPage('')">
					<option value="" <% if sortDiv="" then Response.Write "selected" %>>-선택-</option>
					<option value="new" <% if sortDiv="new" then Response.Write "selected" %>>신상품순</option>
					<option value="best" <% if sortDiv="best" then Response.Write "selected" %>>인기상품순</option>
					<option value="min" <% if sortDiv="min" then Response.Write "selected" %>>낮은가격순</option>
					<option value="hi" <% if sortDiv="hi" then Response.Write "selected" %>>높은가격순</option>
					<option value="hs" <% if sortDiv="hs" then Response.Write "selected" %>>높은할인율순</option>
					<option value="weightup" <% if sortDiv="weightup" then Response.Write "selected" %>>상품무게높은순</option>
					<option value="weightdown" <% if sortDiv="weightdown" then Response.Write "selected" %>>상품무게낮은순</option>
				</select>
			</td>
		</tr>
		</table>
	</td>
</tr>
</form>

<form name="frmArr" id="frmArr" method="post" action="" style="margin:0px;">
<input type="hidden" name="sitename" value="<%= sitename %>" >
<input type="hidden" name="mode" value="" >
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width=20><input type="checkbox" name="ckall" id="ckall" onclick="totalCheck()"></td>	
	<td width="60">itemID</td>
	<td width=50> 이미지</td>
	<td width="100">브랜드ID</td>
	<td width="320">언어팩_해외상품명</td>
	<td>온라인<Br>상품명</td>
	<td width="80">
		사이트<Br>사용여부
		<br>
		<% drawSelectBoxisusingYN "allsiteisusing","","" %>
		<br><input class="button" type="button" value="선택적용" onClick="checkallsiteisusing();">
	</td>
	<td width="60">온라인<Br>판매가</td>
	<td width="60">온라인<Br>매입가</td>
	<td width="50">온라인<Br>판매여부</td>
	<td width="50">온라인<Br>사용여부</td>
	<td width="50">온라인<Br>한정여부</td>
	<td width="50">온라인<Br>단종여부</td>
	<td width="320">화폐별_해외가격</td>
	<td width="60">상품<br>무게</td>
	<td width="50">비고</td>
</tr>
<% if oitem.FresultCount<1 then %>
<tr bgcolor="#FFFFFF">
	<td colspan="20" align="center">[검색결과가 없습니다.]</td>
</tr>
<% end if %>
<% if oitem.FresultCount > 0 then %>
<% for i=0 to oitem.FresultCount-1 %>
<tr class="a" height="25" bgcolor="#FFFFFF" align="center">
	<td><input type="checkbox" name="check" id="check_<%= oitem.FItemList(i).Fitemid %>" value="<%= oitem.FItemList(i).Fitemid %>" onClick="AnCheckClick(this);"></td>
	<td align="center">
		<a href="http://www.10x10.co.kr/shopping/category_prd.asp?itemid=<%= oitem.FItemList(i).Fitemid %>" target="_blank" title="미리보기">
		<%= oitem.FItemList(i).Fitemid %></a>
		</td>
	<td align="center"><img src="<%= oitem.FItemList(i).FSmallImage %>" width="50" height="50" border="0"></td>
	<td align="left"><%= oitem.FItemList(i).Fmakerid %></td>
	<td align="left">
		<table cellpadding="3" cellspacing="1" border="0" class="a" width="100%" bgcolor="<%= adminColor("tabletop") %>">

		<% if oitem.FItemList(i).fmultilangcnt < 1 then %>
			<tr align="left">
				<td bgcolor="#FFFFFF" align="left" colspan=2>
					<font color="red">언어팩미지정</font>
				</td>
			</tr>
		<% end if %>

		<%
		vcountryLangCDArrv=""
		if oitem.FItemList(i).fcountryLangCDarr<>"" and not(isnull(oitem.FItemList(i).fcountryLangCDarr)) then
			vcountryLangCDArrv = split(oitem.FItemList(i).fcountryLangCDarr, "|^|")
		end if

		if isarray(vcountryLangCDArrv) then
		For v = LBound(vcountryLangCDArrv) To UBound(vcountryLangCDArrv)
			vcountryLangCDArrk = split(vcountryLangCDArrv(v), "|*|")
		%>
			<tr align="left">
				<td bgcolor="#FFFFFF" width=60>
					언어팩:<%= vcountryLangCDArrk(0) %>
				</td>
				<td bgcolor="#FFFFFF" align="left">
					<%= vcountryLangCDArrk(1) %>
				</td>
			</tr>
		<% next %>
		<% end if %>
		</table>
	</td>
	<td align="left"><% =oitem.FItemList(i).Fitemname10x10 %></td>
	<td align="center">
		<input type="hidden" name="siteisusing_<%= oitem.FItemList(i).Fitemid %>" id="siteisusing_<%= oitem.FItemList(i).Fitemid %>" value="" >
		<input type="radio" name="selectsiteisusing_<%= oitem.FItemList(i).Fitemid %>" value="Y" <% if oitem.FItemList(i).fsiteisusing="Y" then response.write " checked" %> onClick="CheckClick(<%= oitem.FItemList(i).Fitemid %>);">Y
		<input type="radio" name="selectsiteisusing_<%= oitem.FItemList(i).Fitemid %>" value="N" <% if oitem.FItemList(i).fsiteisusing="N" or oitem.FItemList(i).fsiteisusing="" or isnull(oitem.FItemList(i).fsiteisusing) then response.write " checked" %> onClick="CheckClick(<%= oitem.FItemList(i).Fitemid %>);">N
	</td>
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
	<td align="right"><%= FormatNumber(oitem.FItemList(i).Fbuycash,0) %></td>
	<td align="center"><%= fnColor(oitem.FItemList(i).Fsellyn,"yn") %></td>
	<td align="center"><%= fnColor(oitem.FItemList(i).Fisusing,"yn") %></td>
	<td align="center"><%= fnColor(oitem.FItemList(i).Flimityn,"yn") %></td>
	<td align="center"><%= fnColor(oitem.FItemList(i).Fdanjongyn,"dj") %></td>
	<td align="left">
		<table cellpadding="3" cellspacing="1" border="0" class="a" width="100%" bgcolor="<%= adminColor("tabletop") %>">
		<%
		if oitem.FItemList(i).fpricearr<>"" and not(isnull(oitem.FItemList(i).fpricearr)) then
			vpriceArrv = split(oitem.FItemList(i).fpricearr, "|^|")
		end if

		if isarray(vpriceArrv) then
		For v = LBound(vpriceArrv) To UBound(vpriceArrv)
			vpriceArrk = split(vpriceArrv(v), "|*|")
		%>
			<tr align="left">
				<td bgcolor="#FFFFFF" width=60>
					화폐:<%= vpriceArrk(0) %>
				</td>
				<td bgcolor="#FFFFFF">
					해외판매가:<%= replace(vpriceArrk(1),".00","") %>
				</td>
				<td bgcolor="#FFFFFF">
					원화:<%= FormatNumber(vpriceArrk(2),0) %>
				</td>
				<td bgcolor="#FFFFFF">
					환율:<%= FormatNumber(vpriceArrk(3),0) %>
				</td>
			</tr>
		<%
		next
		end if
		%>
		</table>
	</td>
	<td align="center"><%= FormatNumber(oitem.FItemList(i).FitemWeight,0) %>g</td>
    <td>
    	<input type="button" onClick="PopItemContent( '<%= oitem.FItemList(i).Fitemid %>');" value="수정" class="button">
    </td>
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

<% IF application("Svr_Info")="Dev" THEN %>
	<iframe id="view" name="view" src="" width="100%" height="300" frameborder="0" scrolling="yes"></iframe>
<% else %>
	<iframe id="view" name="view" src="" width="0" height="0" frameborder="0" scrolling="yes"></iframe>
<% end if %>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<%
	set oitem = nothing
	session.codePage = 949
 %>