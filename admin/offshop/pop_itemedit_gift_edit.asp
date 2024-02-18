<%@ language=vbscript %>
<% option explicit %>
<%
	Response.AddHeader "Cache-Control","no-cache"
	Response.AddHeader "Expires","0"
	Response.AddHeader "Pragma","no-cache"
%>
<%
'###########################################################
' Description : 사은품 등록
' Hieditor : 2013.01.15 이상구 생성
'			2017.04.13 한용민 수정(보안관련처리)
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/common/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/offshop/offshopitemcls.asp"-->
<!-- #include virtual="/lib/classes/partners/partnerusercls.asp"-->

<%
dim itemgubun,itemid, itemoption, barcode ,i ,makerid ,ioffitem ,opartner ,ooffontract ,IsOnlineItem
dim editmode , CenterMwDiv ,offList ,offSmall ,OnlineSailYn , IsDirectIpchulContractExistsBrand
dim shopitemname ,shopitemoptionname ,cd1 ,cd2 ,cd3 ,cd1_name ,cd2_name ,cd3_name ,orgsellprice ,shopitemprice
dim shopsuplycash ,shopbuyprice ,isusing ,vatinclude ,extbarcode ,imageList ,offmain ,OnlineOrgprice
dim OnlineBuycash, mwDiv ,OnlineSellcash ,regdate ,updt
	makerid = requestCheckVar(request("makerid"),32)
	barcode	  = requestCheckVar(request("barcode"),32)

editmode = FALSE

'//수정일경우
if barcode <> "" and not(isnull(barcode)) then
	editmode = TRUE

	itemgubun = Left(barcode,2)
	itemid	  = CLng(Mid(barcode,3,6))
	itemoption = Right(barcode,4)

	set ioffitem  = new COffShopItem
		ioffitem.FRectItemgubun = itemgubun
		ioffitem.FRectItemId = itemid
		ioffitem.FRectItemOption = itemoption
		ioffitem.GetOffOneItem

	if ioffitem.FResultCount > 0 then
		makerid = ioffitem.FOneItem.Fmakerid
		Barcode = ioffitem.FOneItem.GetBarcode
		shopitemname = ioffitem.FOneItem.Fshopitemname
		shopitemoptionname = ioffitem.FOneItem.Fshopitemoptionname
		cd1 = ioffitem.FOneItem.FCateCDL
		cd2 = ioffitem.FOneItem.FCateCDM
		cd3 = ioffitem.FOneItem.FCateCDS
		cd1_name = ioffitem.FOneItem.FCateCDLName
		cd2_name = ioffitem.FOneItem.FCateCDMName
		cd3_name = ioffitem.FOneItem.FCateCDSName
		orgsellprice = ioffitem.FOneItem.FShopItemOrgprice
		shopitemprice = ioffitem.FOneItem.Fshopitemprice
		shopsuplycash = ioffitem.FOneItem.Fshopsuplycash
		shopbuyprice = ioffitem.FOneItem.Fshopbuyprice
		ItemGubun = ioffitem.FOneItem.FItemGubun
		isusing = ioffitem.FOneItem.Fisusing
		CenterMwDiv = ioffitem.FOneItem.FCenterMwDiv
		vatinclude = ioffitem.FOneItem.Fvatinclude
		extbarcode = ioffitem.FOneItem.Fextbarcode
		imageList = ioffitem.FOneItem.FimageList
		offmain = ioffitem.FOneItem.FOffImgMain
		offList = ioffitem.FOneItem.FOffImgList
		offSmall = ioffitem.FOneItem.FOffImgSmall
		OnlineSailYn = ioffitem.FOneItem.FOnlineSailYn
		OnlineOrgprice = ioffitem.FOneItem.FOnlineOrgprice
		OnlineBuycash = ioffitem.FOneItem.FOnlineBuycash
		mwDiv = ioffitem.FOneItem.FmwDiv
		OnlineSellcash = ioffitem.FOneItem.FOnlineSellcash
		regdate = ioffitem.FOneItem.Fregdate
		updt = ioffitem.FOneItem.Fupdt

		if left(Barcode,2) <> "80" and left(Barcode,2) <> "85" then
			response.write "<script type='text/javascript'>"
			response.write "	alert('잘못된 접근입니다.');"
			response.write "</script>"
			dbget.close()	:	response.end
		end if
	else
		response.write "<script type='text/javascript'>"
		response.write "	alert('해당되는 상품이 없습니다');"
		'response.write "	self.close();"
		response.write "</script>"
		dbget.close()	:	response.end
	end if

	IsOnlineItem = (itemgubun="10")

'/신규등록
else
	if makerid <> "" then
		CenterMwDiv = GetDefaultItemMwdivByBrand(makerid)

		shopitemprice = "0"
		orgsellprice = "0"
	end if
end if

set opartner = new CPartnerUser
    opartner.FRectDesignerID = makerid

    if makerid <> "" then
    	opartner.GetOnePartnerNUser
    else
		opartner.FResultCount = 0
	end if

set ooffontract = new COffContractInfo
    ooffontract.FRectDesignerID = makerid

    if makerid <> "" then
		ooffontract.GetPartnerOffContractInfo
	end if

function drawOffContractBrandChangeEvent(selectBoxName,selectedId)
   dim tmp_str,query1
   %><select class="select" name="<%=selectBoxName%>" onchange="ChangeBrand(this)">
     <option value='' <%if selectedId="" then response.write " selected"%>>선택</option><%
   query1 = " select c.userid, c.socname_kor"
   query1 = query1 & " from [db_user].[dbo].tbl_user_c c with (nolock)"
   query1 = query1 & " join [db_shop].[dbo].tbl_shop_designer s with (nolock)"
   query1 = query1 & " 		on s.shopid='streetshop000'"
   query1 = query1 & " where c.userid = s.makerid"
   query1 = query1 & " order by c.userid"

	'response.write query1 & "<Br>"
	rsget.CursorLocation = adUseClient
	rsget.Open query1, dbget, adOpenForwardOnly, adLockReadOnly

   if  not rsget.EOF  then
       rsget.Movefirst

       do until rsget.EOF
           if Lcase(selectedId) = Lcase(rsget("userid")) then
               tmp_str = " selected"
           end if
           response.write("<option value='"&rsget("userid")&"' "&tmp_str&">"&rsget("userid")&" ["&db2html(rsget("socname_kor"))&"]</option>")
           tmp_str = ""
           rsget.MoveNext
       loop
   end if
   rsget.close
   response.write("</select>")
End Function

if vatinclude = "" then vatinclude = "Y"
if isusing = "" then isusing = "Y"
'C_IS_SHOP = TRUE
%>

<script type='text/javascript'>

//신규등록때 브랜드 선택
function ChangeBrand(comp){
	location.href="?makerid=" + comp.value;
}

//저장
function EditItem(frm){
	var tmpitemgubuncheck = '';
	<% if editmode then %> var editmode = true; <% else %> var editmode = false; <% end if %>

	//상품구분 선택값 체크
	if (editmode){
		tmpitemgubuncheck = frm.itemgubun.value;
	}else{
		var itemgubun = document.getElementsByName("itemgubun");
		for(var i=0; i < itemgubun.length ; i++){
			if (itemgubun[i].checked){
				tmpitemgubuncheck = frm.itemgubun[i].value;
			}
		}
	}

	if (!editmode){
		if (tmpitemgubuncheck == ''){
			alert('상품구분을 선택하세요.');
			return;
		}
	}

	if (frm.shopitemname.value.length<1){
		alert('상품명을 입력하세요.');
		frm.shopitemname.focus();
		return;
	} else {
		// 특수문자 제거
		frm.shopitemname.value = frm.shopitemname.value.replace(/['"\\\|]/gi, "");
	}

	if (editmode){
	    if (frm.orgsellprice.value.length<1){
			alert('소비자가를 입력하세요.');
			frm.orgsellprice.focus();
			return;
		}
	}

	if (frm.shopitemprice.value.length<1){
		alert('판매가를 입력하세요.');
		frm.shopitemprice.focus();
		return;
	}

	if (frm.shopsuplycash.value.length<1){
		alert('매입가를 입력하세요.');
		frm.shopsuplycash.focus();
		return;
	}

	if (editmode != true) {
        frm.itemoptioncode2.value = "";
        frm.itemoptioncode3.value = "";

		var optiont = "";
		var optionv = "";
		var optioncnt = 0;
		if (frm.useoptionyn[0].checked == true) {
			if (tmpitemgubuncheck == "80") {
				alert("오프라인 사은품에 대해 옵션을 등록할 수 없습니다.\n\n테스트 진행중!!");
				return;
			}

			for (var i = 0; i < frm.etcOpt.length; i++) {
				// 특수문자 제거
				frm.etcOpt[i].value = frm.etcOpt[i].value.replace(/['"\\\|]/gi, "");

				if (frm.etcOpt[i].value != "") {
					optioncnt = optioncnt + 1;
					var s = "0000" + optioncnt;

					optiont += (frm.etcOpt[i].value + "|");
					optionv += s.substring(s.length - 4) + "|";
				}
			}

			if (optioncnt < 2) {
				alert("옵션은 두개 이상이어야 합니다.");
				return;
			}
		}

		frm.itemoptioncode2.value = optionv;
        frm.itemoptioncode3.value = optiont;
	}

	if (frm.shopitemprice.value > 0){
		alert("사은품은 판매가가 0이하여야 합니다.");
		frm.shopitemprice.focus();
		return;
	} if (editmode){
		if (frm.orgsellprice.value > 0){
			alert("사은품은 소비자가 0이하여야 합니다.");
			frm.orgsellprice.focus();
			return;
		}
	} if (editmode){
		if (frm.shopitemname.value.match(/^\[사은품\] /) == null) {
			alert("사은품 문구는 삭제할 수 없습니다.");
			return;
		}
	} if (!editmode){
		if (frm.shopitemname.value.match(/사은품/) != null) {
			alert("사은품 문구는 상품명에 자동입력됩니다. 사은품 문구를 지우세요.");
			return;
		}
	}

	if (!editmode){
		if (((frm.shopsuplycash.value!=0)||(frm.shopbuyprice.value!=0))){
			if (!confirm('!! 매장에 판매하는 경우에만 매장공급가를 입력 하셔야 합니다. \n\n계속 하시겠습니까?')){
				return;
			}
		}
	}

	if (editmode){
		if (frm.tmpoffmain.value.length<1 && frm.file1.value.length<1){
			alert('이미지를 입력해 주세요 - 필수 사항입니다.');
			frm.file1.focus();
			return;
		}
	}else{
		if (frm.file1.value.length<1){
			alert('이미지를 입력해 주세요 - 필수 사항입니다.');
			frm.file1.focus();
			return;
		}
	}

	var ret = 0;
	for (i=0; i< document.getElementsByName("centermwdiv").length; i++){
		if (document.getElementsByName("centermwdiv")[i].checked == true){
			ret = ret + 1;
		}
	}
	if (ret == 0){
		alert("센터 매입 구분을 선택 하세요.");
		return;
	}

    if ((!frm.vatinclude[0].checked)&&(!frm.vatinclude[1].checked)){
        alert('과세 구분을 선택 하세요.');
		frm.vatinclude[0].focus();
		return;
    }

	if (confirm('저장 하시겠습니까?')){
		if (frm.shopitemname.value.match(/사은품/) == null) {
			frm.shopitemname.value = "[사은품] " + frm.shopitemname.value;
		}

		frm.submit();
	}
}

function PopUpcheInfo(v){
	window.open("/admin/lib/popbrandinfoonly.asp?designer=" + v,"popupcheinfo","width=640 height=580 scrollbars=yes resizable=yes");
}

function editOffDesinger(shopid,designerid){
	var popwin = window.open("/admin/lib/popshopupcheinfo.asp?shopid=" + shopid + "&designer=" + designerid,"popshopupcheinfo","width=640 height=540");
	popwin.focus();
}

function TnCheckOptionYN(frm){
	if (frm.useoptionyn[0].checked == true) {
	    // 옵션사용
		document.all.optlist.style.display="";
		document.all.optname.style.display="none";

	} else {
	    // 옵션없음
		document.all.optlist.style.display="none";
		document.all.optname.style.display="";
    }
}

function InsertOptionWithGubun(ioptTypeName, ft, fv) {
	var frm = document.frmedit;

	//옵션값이 같은것이 있으면 skip ,전용옵션인경우 제외
	if (fv!="0000"){
		for (i=0;i<frm.realopt.length;i++){
			if (frm.realopt[i].value==fv){
				return;
			}
		}
	}

    frm.optTypeNm.value = ioptTypeName;
	frm.elements['realopt'].options[frm.realopt.options.length] = new Option(ft, fv);
}

function popNormalOptionAdd() {
	popwin = window.open('/common/module/normalitemoptionadd.asp' ,'popNormalOptionAdd','width=540,height=260,scrollbars=yes,resizable=yes');
	popwin.focus();
}

// 선택된 옵션 삭제
function delItemOptionAdd()
{
	var frm = document.frmedit;
	var sidx = frm.realopt.options.selectedIndex;

	if(sidx<0){
		alert("삭제할 옵션을 선택해주십오.");
	}else{
	    for(i=0; i<frm.realopt.options.length; i++){
    		if(frm.realopt.options[i].selected){
    			frm.realopt.options[i] = null;
    			i=i-1;
    		}
    	}

		if (frm.realopt.options.length<1){
		    frm.optTypeNm.value = '';
		}
	}
}

// 카테고리등록
function editCategory(cdl,cdm,cdn){
	var param = "cdl=" + cdl + "&cdm=" + cdm + "&cdn=" + cdn ;

	popwin = window.open('/common/module/categoryselect.asp?' + param ,'editcategory','width=700,height=400,scrollbars=yes,resizable=yes');
	popwin.focus();
}

//카테고리 셋팅
function setCategory(cd1,cd2,cd3,cd1_name,cd2_name,cd3_name){
	var frm = document.frmedit;
	frm.cd1.value = cd1;
	frm.cd2.value = cd2;
	frm.cd3.value = cd3;
	frm.cd1_name.value = cd1_name;
	frm.cd2_name.value = cd2_name;
	frm.cd3_name.value = cd3_name;
}

</script>

<!-- 리스트 시작 -->
>>사은품 등록
<form name="frmedit" method="post" action="<%=uploadImgUrl%>/linkweb/offshop/item/itemedit_off.asp" enctype="MULTIPART/FORM-DATA" style="margin:0px;">
<input type="hidden" name="editmode" value="<%=editmode%>">
<input type="hidden" name="makerid" value="<%= makerid %>">
<input type="hidden" name="barcode" value="<%=barcode%>">
<input type="hidden" name="offmain" value="<%=offmain%>">
<input type="hidden" name="itemid" value="<%= itemid %>">
<input type="hidden" name="itemoption" value="<%= itemoption %>">
<input type="hidden" name="itemoptioncode2">
<input type="hidden" name="itemoptioncode3">
<input type="hidden" name="regtype" value="giftitem">

<input type="hidden" name="cd1" value="">
<input type="hidden" name="cd2" value="">
<input type="hidden" name="cd3" value="">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">

<% if NOT(editmode) then %>
<tr bgcolor="<%= adminColor("pink") %>">
	<td width="100" height="30">브랜드ID</td>
	<td bgcolor="#FFFFFF">
		<% drawSelectBoxDesignerwithName "imakerid",makerid  %> ※신규 등록하실 사은품의 브랜드를 선택해 주세요.
	</td>
</tr>
<% if (makerid = "") or (opartner.FResultCount < 1) then %>
<tr bgcolor="<%= adminColor("pink") %>">
	<td colspan="2" bgcolor="#FFFFFF" align="center">
		<input type="button" class="button" value="검색" onclick="ChangeBrand(document.frmedit.imakerid);">
	</td>
</tr>
<% end if %>
<%
end if

'// 브랜드 선택이 없을경우 노출하지 않고, 무조건 브랜드 선택하도록..
if makerid = "" then dbget.close() : response.write "</table>" : response.end

'// 잘못된 브랜드
if opartner.FResultCount < 1 then
	response.write "<script>alert('잘못된 브랜드입니다.');</script>"
	dbget.close() : response.write "</table>" : response.end
end if

%>

<tr bgcolor="<%= adminColor("pink") %>" height="30">
	<td width=100>브랜드계약정보</td>
	<td bgcolor="#FFFFFF">
		<a href="javascript:PopUpcheInfo('<%= makerid %>');"><%= makerid %></a> (<%= opartner.FOneItem.Fsocname_kor %>,<%= opartner.FOneItem.FCompany_name %>)
	</td>
</tr>
<% if (editmode) then %>
<input type="hidden" name="itemgubun" value="<%=itemgubun%>">
<tr bgcolor="<%= adminColor("tabletop") %>" height="30">
	<td width="100">상품코드</td>
	<td bgcolor="#FFFFFF">
		<%= Barcode %>
		<%if left(Barcode,2) = "10" then %>
			온라인공용상품
		<% elseif left(Barcode,2) = "90" then %>
			오프라인전용상품
		<% elseif left(Barcode,2) = "95" then %>
			가맹점개별매입판매상품
		<% elseif left(Barcode,2) = "85" then %>
			ON사은품
		<% elseif left(Barcode,2) = "80" then %>
			OFF사은품
		<% elseif left(Barcode,2) = "70" then %>
			소모품
		<% end if %>
		<br><font color="#AAAAAA">(85ON사은품, 80OFF사은품)</font>
	</td>
</tr>
<% else %>
<tr bgcolor="<%= adminColor("tabletop") %>">
	<td width=100 height="30">상품구분</td>
	<td bgcolor="#FFFFFF">
		<input type="radio" name="itemgubun" value="85" <% if itemgubun = "85" then response.write " checked" %>>ON사은품(85)
		<input type="radio" name="itemgubun" value="80" <% if itemgubun = "80" then response.write " checked" %> disabled>OFF사은품(80)
	</td>
</tr>
<% end if %>
<tr bgcolor="<%= adminColor("tabletop") %>">
	<td height="30">상품명</td>
	<td bgcolor="#FFFFFF">
		<input type="text" class="text" name="shopitemname" value="<%= shopitemname %>" size="80" maxlength="90">
		<br>※ 상품명에 "[사은품]" 문구가 자동으로 붙습니다.
	</td>
</tr>
<% if NOT(editmode) then %>
<tr bgcolor="<%= adminColor("tabletop") %>" height="30">
	<td height="30">옵션구분</td>
	<td bgcolor="#FFFFFF">
		<label><input type="radio" name="useoptionyn" value="Y" onClick="TnCheckOptionYN(this.form);" disabled>옵션사용함</label>&nbsp;&nbsp;
		<label><input type="radio" name="useoptionyn" value="N" onClick="TnCheckOptionYN(this.form);" checked>옵션사용안함</label>
	</td>
</tr>
<!----- 단일 옵션 DIV ----->
<tr bgcolor="<%= adminColor("tabletop") %>" id="optname" height="30">
    <td height="30">옵션명</td>
  	<td bgcolor="#FFFFFF" align="left">
		<input type="text" class="text" name="shopitemoptionname" value="<%= shopitemoptionname %>" size="40" maxlength="40">
  	</td>
</tr>
<!----- 단일 옵션 DIV ----->
<tr bgcolor="<%= adminColor("tabletop") %>" id="optlist" style="display:none" height="30">
    <td height="30">옵션 설정</td>
  	<td bgcolor="#FFFFFF" align="left">

		<table width="440" border="0" cellspacing="1" cellpadding="2" align="left" class="a"  bgcolor="#3d3d3d" >
		<% for i = 1 to 10 %>
		<tr bgcolor="#FFFFFF" align="center">
			<td>옵션명 <%= i %> </td>
			<td align="center"><input type="text" class="text" name="etcOpt" size="20" maxlength="20"></td>
		</tr>
		<% next %>
		</table>

  	</td>
</tr>
<% else %>
<tr bgcolor="<%= adminColor("tabletop") %>" height="30">
	<td>옵션구분</td>
	<td bgcolor="#FFFFFF">
		<input type="text" class="text" name="shopitemoptionname" value="<%= shopitemoptionname %>" size="40" maxlength="40">
	</td>
</tr>
<% end if %>
<tr bgcolor="<%= adminColor("tabletop") %>" >
	<td>소비자가</td>
	<td bgcolor="#FFFFFF">
		<input type="text" class="text_ro" name="orgsellprice" value="<%= orgsellprice %>" size=8 maxlength=9 readonly>
	</td>
</tr>
<tr bgcolor="<%= adminColor("tabletop") %>" >
	<td>판매가</td>
	<td bgcolor="#FFFFFF">
		<input type="text" class="text_ro" name="shopitemprice" value="<%= shopitemprice %>" size=8 maxlength=9 readonly>
	</td>
</tr>
<tr bgcolor="<%= adminColor("tabletop") %>" >
	<td>매입가</td>
	<td bgcolor="#FFFFFF">
		<input type="text" name="shopsuplycash" value="<%= shopsuplycash %>" size=8 maxlength=9 class="input_right"> ※설정 않으면 매입 정산안함
	</td>
</tr>
<tr bgcolor="<%= adminColor("tabletop") %>" >
	<td>매장공급가</td>
	<td bgcolor="#FFFFFF">
		<input type="text" name="shopbuyprice" value="<%= shopbuyprice %>" size=8 maxlength=9 class="input_right" > ※설정 않으면 매장 정산안함
	</td>
</tr>
<tr bgcolor="<%= adminColor("tabletop") %>">
	<td>사용유무</td>
	<td bgcolor="#FFFFFF">
		<% if isusing = "Y" then %>
		<input type=radio name=isusing value="Y" checked >사용함
		<input type=radio name=isusing value="N">사용안함
		<% else %>
		<input type=radio name=isusing value="Y"  >사용함
		<input type=radio name=isusing value="N" checked >사용안함
		<% end if %>
	</td>
</tr>
<tr bgcolor="<%= adminColor("tabletop") %>">
	<td>센터매입구분</td>
	<td bgcolor="#FFFFFF">
		<%
		' 신규등록시에는 무조건 위탁으로 셋팅.	2023.06.23 이문재이사님 요청
		if not(editmode) then
		%>
			<input type="radio" name="centermwdiv" value="W" checked >위탁
		<% else %>
			<input type="radio" name="centermwdiv" value="W" <%= ChkIIF(centermwdiv="W","checked","") %> >위탁
			<input type="radio" name="centermwdiv" value="M" <%= ChkIIF(centermwdiv="M","checked","") %> >매입
		<% end if %>
	</td>
</tr>
<tr bgcolor="<%= adminColor("tabletop") %>">
	<td>과세구분</td>
	<td bgcolor="#FFFFFF">
		<input type="radio" name="vatinclude" value="Y" <%= ChkIIF(vatinclude = "Y","checked","") %>  >과세
		<input type="radio" name="vatinclude" value="N" <%= ChkIIF(vatinclude = "N","checked","") %> > <font color="<%= ChkIIF(vatinclude = "N","blue","#000000") %>">면세</font>
	</td>
</tr>
<tr bgcolor="<%= adminColor("tabletop") %>">
	<td>이미지</td>
	<td bgcolor="#FFFFFF">
		<% if IsOnlineItem then %>
			<img src="<%= imageList %>" width="50" height="50">
		<% else %>
			<input type="file" name="file1" class="button" size=20 >
			<Br>※ 기본 이미지는 반드시 400x400 , jpg 파일로 올려주시기 바랍니다.
			<Br>※ 400x400 이미지를 저장 하시면, 자동으로 100x100 , 50x50 이 생성 됩니다.
			<input type="hidden" name="tmpoffmain" value="<%= offmain %>">
   				<% IF offmain <> "" THEN %>
	   				<BR><img src="<%=offmain%>" border="0" width=400 height=400> 400x400
   				<% END IF %>
   				<% if offlist <> "" then %>
   					<BR><img src="<%=offlist%>" border="0" width=100 height=100> 100x100
   				<% end if %>
   				<% if offsmall <> "" then %>
   					<BR><img src="<%=offsmall%>" border="0" width=50 height=50> 50x50
   				<% end if %>
		<% end if %>
	</td>
</tr>
<% if editmode then %>
<tr bgcolor="<%= adminColor("tabletop") %>">
	<td width=100>등록일</td>
	<td bgcolor="#FFFFFF"><%= regdate %></td>
</tr>
<tr bgcolor="<%= adminColor("tabletop") %>">
	<td width=100>최종수정일</td>
	<td bgcolor="#FFFFFF"><%= updt %></td>
</tr>
<% end if %>
<tr bgcolor="#FFFFFF">
	<td colspan="2" align=center>
		<input type="button" class="button" value="<% if editmode then %>수정<% else %>신규저장<% end if %>" onclick="EditItem(frmedit)">
	</td>
</tr>
</table>
</form>

<%
set opartner = Nothing
set ooffontract = Nothing
set ioffitem = Nothing
%>
<!-- #include virtual="/common/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->