<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/admin/etc/aboutpet/aboutpetcls.asp"-->
<%
Dim makerid, itemid, itemname, bestOrd, sellyn, limityn, sailyn, onlyValidMargin, isMadeHand, isOption, infoDiv, morningJY
Dim bestOrdMall, aboutpetGoodNo, extsellyn, ExtNotReg, isReged, MatchCate, optAddPrcRegTypeNone, notinmakerid, notinitemid, deliverytype, mwdiv, exctrans
Dim expensive10x10, diffPrc, aboutpetYes10x10No, aboutpetNo10x10Yes, reqEdit, reqExpire, failCntExists, priceOption, isSpecialPrice
Dim page, i, research
Dim oaboutpet, isextusing, cisextusing, rctsellcnt
dim startsell, stopsell
Dim startMargin, endMargin
page    				= request("page")
research				= request("research")
itemid  				= request("itemid")
makerid					= request("makerid")
itemname				= request("itemname")
bestOrd					= request("bestOrd")
bestOrdMall				= request("bestOrdMall")
sellyn					= request("sellyn")
limityn					= request("limityn")
sailyn					= request("sailyn")
onlyValidMargin			= request("onlyValidMargin")
startMargin				= request("startMargin")
endMargin				= request("endMargin")
isMadeHand				= request("isMadeHand")
isOption				= request("isOption")

infoDiv					= request("infoDiv")
morningJY				= request("morningJY")
extsellyn				= request("extsellyn")
aboutpetGoodNo			= request("aboutpetGoodNo")
ExtNotReg				= request("ExtNotReg")
isReged					= request("isReged")
MatchCate				= request("MatchCate")
expensive10x10			= request("expensive10x10")
diffPrc					= request("diffPrc")
aboutpetYes10x10No		= request("aboutpetYes10x10No")
aboutpetNo10x10Yes		= request("aboutpetNo10x10Yes")
reqEdit					= request("reqEdit")
reqExpire				= request("reqExpire")
failCntExists			= request("failCntExists")
optAddPrcRegTypeNone	= request("optAddPrcRegTypeNone")
notinmakerid			= request("notinmakerid")
priceOption				= request("priceOption")
isSpecialPrice          = request("isSpecialPrice")
deliverytype			= request("deliverytype")
mwdiv					= request("mwdiv")
startsell				= requestCheckVar(request("startsell"), 1)
stopsell				= requestCheckVar(request("stopsell"), 1)
notinitemid				= requestCheckVar(request("notinitemid"), 1)
exctrans				= requestCheckVar(request("exctrans"), 1)
isextusing				= requestCheckVar(request("isextusing"), 1)
cisextusing				= requestCheckVar(request("cisextusing"), 1)
rctsellcnt				= requestCheckVar(request("rctsellcnt"), 1)

If page = "" Then page = 1
If sellyn="" Then sellyn = "Y"
''기본조건 등록예정이상
If (research = "") Then
	ExtNotReg = "D"
	MatchCate = ""
	onlyValidMargin = "Y"
	bestOrd = "on"
	sellyn = "Y"

	if (stopsell = "Y") then
		'// 판매중지 대상 상품목록
		ExtNotReg = "D"
		sellyn = "N"
		extsellyn = "Y"
		aboutpetYes10x10No = "on"
		onlyValidMargin = ""
	elseif (startsell = "Y") then
		'// 판매전환 대상 상품목록
		ExtNotReg = "D"
		sellyn = "Y"
		extsellyn = "N"
		onlyValidMargin = "Y"
		notinmakerid = "N"
		notinitemid = "N"
		aboutpetNo10x10Yes = "on"
	end if
End If

If (session("ssBctID")="kjy8517") Then
'	itemid = ""

End If

'텐바이텐 상품코드 엔터키로 검색되게
If itemid<>"" then
	Dim iA, arrTemp, arrItemid
	itemid = replace(itemid,",",chr(10))
	itemid = replace(itemid,chr(13),"")
	arrTemp = Split(itemid,chr(10))
	iA = 0
	Do While iA <= ubound(arrTemp)
		If Trim(arrTemp(iA))<>"" then
			If Not(isNumeric(trim(arrTemp(iA)))) then
				Response.Write "<script language=javascript>alert('[" & arrTemp(iA) & "]은(는) 유효한 상품코드가 아닙니다.');history.back();</script>"
				dbget.close()	:	response.End
			Else
				arrItemid = arrItemid & trim(arrTemp(iA)) & ","
			End If
		End If
		iA = iA + 1
	Loop
	itemid = left(arrItemid,len(arrItemid)-1)
End If
'옥션 상품코드 엔터키로 검색되게
If aboutpetGoodNo <> "" then
	Dim iA2, arrTemp2, arraboutpetGoodNo
	aboutpetGoodNo = replace(aboutpetGoodNo,",",chr(10))
	aboutpetGoodNo = replace(aboutpetGoodNo,chr(13),"")
	arrTemp2 = Split(aboutpetGoodNo,chr(10))
	iA2 = 0
	Do While iA2 <= ubound(arrTemp2)
		If Trim(arrTemp2(iA2))<>"" then
			arraboutpetGoodNo = arraboutpetGoodNo& "'"& trim(arrTemp2(iA2)) & "',"
		End If
		iA2 = iA2 + 1
	Loop
	aboutpetGoodNo = left(arraboutpetGoodNo,len(arraboutpetGoodNo)-1)
End If

Set oaboutpet = new Caboutpet
	oaboutpet.FCurrPage					= page
If (session("ssBctID")="kjy8517") Then
	oaboutpet.FPageSize					= 100

Else
	oaboutpet.FPageSize					= 50
End If
	oaboutpet.FRectCDL					= request("cdl")
	oaboutpet.FRectCDM					= request("cdm")
	oaboutpet.FRectCDS					= request("cds")
	oaboutpet.FRectItemID				= itemid
	oaboutpet.FRectItemName				= itemname
	oaboutpet.FRectSellYn				= sellyn
	oaboutpet.FRectLimitYn				= limityn
	oaboutpet.FRectSailYn				= sailyn
'	oaboutpet.FRectonlyValidMargin		= onlyValidMargin
	oaboutpet.FRectStartMargin			= startMargin
	oaboutpet.FRectEndMargin				= endMargin
	oaboutpet.FRectMakerid				= makerid
	oaboutpet.FRectaboutpetGoodNo			= aboutpetGoodNo
	oaboutpet.FRectMatchCate				= MatchCate
	oaboutpet.FRectIsMadeHand			= isMadeHand
	oaboutpet.FRectIsOption				= isOption
	oaboutpet.FRectIsReged				= isReged
	oaboutpet.FRectNotinmakerid			= notinmakerid
	oaboutpet.FRectNotinitemid			= notinitemid
	oaboutpet.FRectExcTrans				= exctrans
	oaboutpet.FRectPriceOption			= priceOption
	oaboutpet.FRectIsSpecialPrice       	= isSpecialPrice
	oaboutpet.FRectDeliverytype			= deliverytype
	oaboutpet.FRectMwdiv					= mwdiv
	oaboutpet.FRectIsextusing			= isextusing
	oaboutpet.FRectCisextusing			= cisextusing
	oaboutpet.FRectRctsellcnt			= rctsellcnt

	oaboutpet.FRectExtNotReg				= ExtNotReg
	oaboutpet.FRectExpensive10x10		= expensive10x10
	oaboutpet.FRectdiffPrc				= diffPrc
	oaboutpet.FRectaboutpetYes10x10No		= aboutpetYes10x10No
	oaboutpet.FRectaboutpetNo10x10Yes		= aboutpetNo10x10Yes
	oaboutpet.FRectExtSellYn				= extsellyn
	oaboutpet.FRectInfoDiv				= infoDiv
	oaboutpet.FRectFailCntOverExcept		= ""
	oaboutpet.FRectFailCntExists			= failCntExists
	oaboutpet.FRectReqEdit				= reqEdit
If (bestOrd = "on") Then
    oaboutpet.FRectOrdType = "B"
ElseIf (bestOrdMall = "on") Then
    oaboutpet.FRectOrdType = "BM"
End If

	oaboutpet.getaboutpetRegedItemList		'그 외 리스트
%>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script language='javascript'>
//크롬 업데이트로 alert 수정..2021-07-26
function systemAlert(message){
	alert(message);
}
window.addEventListener("message", (event) => {
    var data = event.data;
    if (typeof(window[data.action]) == "function") {
        window[data.action].call(null, data.message);
    } },
false);
//크롬 업데이트로 alert 수정..2021-07-26 끝

function goPage(pg){
    frm.page.value = pg;
    frm.submit();
}

function aboutpetUpdateProcess() {
	var chkSel=0;
	var v1, v2, v3, v4, v5;
	v1 = "";
	v2 = "";
	v3 = "";
	v4 = "";
	v5 = "";

	var cnt = $("input[name=cksel]:checkbox:checked").length;
	if (cnt == 1 && typeof(frmSvArr.cksel.length)=='undefined')  {
		chkSel++;
		v1 = v1 + frmSvArr.cksel.value + '||';
		v2 = v2 + frmSvArr.regedItemName.value + '||';
		v3 = v3 + frmSvArr.regedOptionname.value + '||';
		v4 = v4 + frmSvArr.regedItemprice.value + '||';
		v5 = v5 + frmSvArr.aboutpetSellYn.value + '||';
	}else{
		try {
			if(frmSvArr.cksel.length>1) {
				for(var i=0;i<frmSvArr.cksel.length;i++) {
					if(frmSvArr.cksel[i].checked) {
						chkSel++;
						v1 = v1 + frmSvArr.cksel[i].value + '||';
						v2 = v2 + frmSvArr.regedItemName[i].value + '||';
						v3 = v3 + frmSvArr.regedOptionname[i].value + '||';
						v4 = v4 + frmSvArr.regedItemprice[i].value + '||';
						v5 = v5 + frmSvArr.aboutpetSellYn[i].value + '||';
					}
				}
			}else {
				if(frmSvArr.cksel.checked) chkSel++;
			}
			if(chkSel<=0) {
				alert("선택한 상품이 없습니다.");
				return;
			}
		}
		catch(e) {
			alert(e);
			alert("상품이 없습니다.");
			return;
		}
	}

    if (confirm('선택하신 ' + chkSel + '개 상품을 수정 하시겠습니까?\n\n※반드시 aboutpet에 수정된 것 확인 후 수정하셔야 합니다.')){
		$("#itemarr").val(v1);
		$("#regedItemNamearr").val(v2);
		$("#regedOptionnamearr").val(v3);
		$("#regedItempricearr").val(v4);
		$("#aboutpetSellYnarr").val(v5);

        document.frmArr.target = "xLink";
		document.frmArr.action = "/admin/etc/aboutpet/proc_aboutpet.asp"
        document.frmArr.submit();
    }
}
function popMayEditList(){
    var popwin = window.open('/admin/etc/aboutpet/popMayEditList.asp','popMayEditList','width=900,height=700,scrollbars=yes,resizable=yes');
    popwin.focus();
}
function btnOk(v){
	if (v == 'itemname'){
		$("input[name=regedItemName]").val( $("#copyitemname").val() );
	}else if(v == 'sellyn'){
		$("select[name=aboutpetSellYn]").val( $("select[name=copysellyn]").val() );
	}else if(v=='price'){
		$("input[name=regedItemprice]").val( $("#copyprice").val() );
	}
}
</script>


<form name="frmArr" method=post>
	<input type="hidden" id= "itemarr" name="itemarr" value="" />
	<input type="hidden" id= "regedItemNamearr" name="regedItemNamearr" value="" />
	<input type="hidden" id= "regedOptionnamearr" name="regedOptionnamearr" value="" />
	<input type="hidden" id= "regedItempricearr" name="regedItempricearr" value="" />
	<input type="hidden" id= "aboutpetSellYnarr" name="aboutpetSellYnarr" value="" />
	<input type="hidden" id= "cmdparam" name="cmdparam" value="I" />
</form>


<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get" action="">
<input type="hidden" name="page" value="">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="research" value="on">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="3" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
	<td align="left">
		브랜드&nbsp;&nbsp;&nbsp; : <% drawSelectBoxDesignerwithName "makerid",makerid %>&nbsp;
		&nbsp;
		<a href="https://po.aboutpet.co.kr/login/loginView.do" target="_blank">aboutpet Admin바로가기</a>
		<%
			If (session("ssBctID")="kjy8517") OR (session("ssBctID")="icommang") OR (session("ssBctID")="hrkang97") OR (session("ssBctID")="as2304") Then
				response.write "<font color='GREEN'>[ tenbyten | cube101010! ]</font>"
			End If
		%>
		<br>
		상품코드 : <textarea rows="2" cols="20" name="itemid" id="itemid"><%=replace(itemid,",",chr(10))%></textarea>
		&nbsp;
		상품명: <input type="text" class="text" name="itemname" value="<%= itemname %>" size="32" maxlength="32">
		&nbsp;
		<br>
		<!-- #include virtual="/common/module/categoryselectbox.asp"-->
		<br>
	</td>
	<td rowspan="3" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="검색" onClick="javascript:document.frm.submit();">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
		<!-- #include virtual="/admin/etc/incsearch1.asp"-->
	</td>
</tr>
</form>
</table>

<p />

<form name="frmReg" method="post" action="aboutpetitem.asp" style="margin:0px;">
<input type="hidden" name="mode" value="">
<input type="hidden" name="menupos" value="<%= menupos %>">
<table width="100%" class="a" bgcolor="#FFFFFF">
<tr>
	<td align="left" valign="top">
		<input class="button" type="button" id="btnSellYn" value="수정요망List" onClick="popMayEditList();">
	</td>
	<td align="right" valign="top">
		선택상품
		<input class="button" type="button" id="btnSellYn" value="수정" onClick="aboutpetUpdateProcess();">
	</td>
</tr>
</table>
</form>
<br>
<!-- 리스트 시작 -->
<form name="frmSvArr" method="post" onSubmit="return false;" action="" style="margin:0px;">
<input type="hidden" name="mode" value="">
<input type="hidden" name="cmdparam" value="">
<input type="hidden" name="delitemid" value="">
<input type="hidden" name="chgSellYn" value="">
<input type="hidden" name="chgStatItemCode" value="">
<input type="hidden" name="ckLimit">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="19">
		검색결과 : <b><%= FormatNumber(oaboutpet.FTotalCount,0) %></b>
		&nbsp;
		페이지 : <b> <%= FormatNumber(page,0) %> / <%= FormatNumber(oaboutpet.FTotalPage,0) %></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width="20"><input type="checkbox" name="chkAll" onClick="fnCheckAll(this.checked,frmSvArr.cksel);"></td>
	<td width="50">이미지</td>
	<td width="60">텐바이텐<br>상품번호</td>
	<td width="60">텐바이텐<br>옵션번호</td>
	<td>브랜드<br>상품명</td>
	<td>aboutpet<br>상품명<br>
		<input type="text" class="text" name="copyitemname" id="copyitemname" value="" size="20" >
		<input type="button" class="button" value="일괄적용" onclick="btnOk('itemname');">
	</td>
	<td>텐바이텐<br>옵션명</td>
	<td>aboutpet<br>옵션명</td>
	<td width="140">aboutpet등록일</td>
	<td width="70">텐바이텐<br>판매가</td>
	<td width="70">품절여부</td>
	<td width="70">aboutpet<br>가격<br>
		<input type="text" class="text" name="copyprice" id="copyprice" value="" size="10" >
		<input type="button" class="button" value="일괄적용" onclick="btnOk('price');"></td>
	<td width="70">aboutpet<br>판매여부<br />
		<select class="select" name="copysellyn" onchange="btnOk('sellyn');">
			<option value="Y">Y</option>
			<option value="N">N</option>
		</select>
	</td>
	<td width="70">텐바이텐<br>마진</td>
	<td width="70">주문제작<br>여부</td>
</tr>
<% For i=0 to oaboutpet.FResultCount - 1 %>
<tr align="center" bgcolor="#FFFFFF">
	<td><input type="checkbox" name="cksel" onClick="AnCheckClick(this);"  value="<%= oaboutpet.FItemList(i).FIdx %>"></td>
	<td><img src="<%= oaboutpet.FItemList(i).Fsmallimage %>" width="50"></td>
	<td align="center">
		<a href="<%=vwwwUrl%>/<%=oaboutpet.FItemList(i).FItemID%>" target="_blank"><%= oaboutpet.FItemList(i).FItemID %></a>
		<%= oaboutpet.FItemList(i).getLimitHtmlStr %>
	</td>
	<td align="center"><%= oaboutpet.FItemList(i).FRegeditemoption %></td>
	<td align="left"><%= oaboutpet.FItemList(i).FMakerid %> <%= oaboutpet.FItemList(i).getDeliverytypeName %><br><%= oaboutpet.FItemList(i).FItemName %></td>
	<td align="left">
		<input type="text" class="text" size="40" id="regedItemName<%= i %>" name="regedItemName" value="<%= oaboutpet.FItemList(i).FRegedItemName %>"></td>
	<td align="left"><%= oaboutpet.FItemList(i).FOptionname %></td>
	<td align="left">
		<input type="text" name="regedOptionname" id="regedOptionname<%= i %>" value="<%= oaboutpet.FItemList(i).FRegedoptionname %>">
	</td>
	<td align="center"><%= oaboutpet.FItemList(i).FaboutpetRegdate %></td>
	<td align="right">
		<% If oaboutpet.FItemList(i).FSaleYn = "Y" Then %>
			<strike><%= FormatNumber(oaboutpet.FItemList(i).FOrgPrice,0) %></strike><br>
			<font color="#CC3333"><%= FormatNumber(oaboutpet.FItemList(i).FSellcash,0) %></font>
		<% Else %>
			<%= FormatNumber(oaboutpet.FItemList(i).FSellcash,0) %>
		<% End If %>
		<% If oaboutpet.FItemList(i).Foptaddprice > 0 Then  %>
			<br />
			<font color="yellowgreen"> (+<%= FormatNumber(oaboutpet.FItemList(i).Foptaddprice,0) %>) </font>
		<% End If %>
	</td>
	<td align="center">
	<%
		If oaboutpet.FItemList(i).IsSoldOut Then
			If oaboutpet.FItemList(i).FSellyn = "N" Then
	%>
		<font color="red">품절</font>
	<%
			Else
	%>
		<font color="red">일시<br>품절</font>
	<%
			End If
		End If
	%>
	</td>
	<td align="center">
		<input type="text" class="text" id="regedItemprice<%= i %>" name="regedItemprice" size="10" value="<%= oaboutpet.FItemList(i).FaboutpetPrice %>">
	</td>
	<td align="center">
		<select class="select" name="aboutpetSellYn">
			<option value="Y" <%= chkiif(oaboutpet.FItemList(i).FaboutpetSellYn="Y", "selected", "") %>>Y</option>
			<option value="N" <%= chkiif(oaboutpet.FItemList(i).FaboutpetSellYn="N", "selected", "") %>>N</option>
		</select>
	</td>
	<td align="center">
	<%
		If oaboutpet.FItemList(i).Fsellcash = 0 Then
		elseif (oaboutpet.FItemList(i).FSaleYn="Y") Then
	%>
		<% if (oaboutpet.FItemList(i).FOrgPrice<>0) then %>
		<strike><%= CLng(10000-oaboutpet.FItemList(i).FOrgSuplycash/oaboutpet.FItemList(i).FOrgPrice*100*100)/100 & "%" %></strike><br>
		<% end if %>
		<font color="#CC3333"><%= CLng(10000-oaboutpet.FItemList(i).Fbuycash/oaboutpet.FItemList(i).Fsellcash*100*100)/100 & "%" %></font>
	<%
		else
			response.write CLng(10000-oaboutpet.FItemList(i).Fbuycash/oaboutpet.FItemList(i).Fsellcash*100*100)/100 & "%"
		end if
	%>
	</td>
	<td align="center">
	<%
		If oaboutpet.FItemList(i).FItemdiv = "06" OR oaboutpet.FItemList(i).FItemdiv = "16" Then
			response.write "<font color='green'>주문제작</font>"
		End If
	%>
	</td>
</tr>
<% Next %>
<tr height="20">
    <td colspan="18" align="center" bgcolor="#FFFFFF">
        <% if oaboutpet.HasPreScroll then %>
		<a href="javascript:goPage('<%= oaboutpet.StartScrollPage-1 %>');">[pre]</a>
    	<% else %>
    		[pre]
    	<% end if %>

    	<% for i=0 + oaboutpet.StartScrollPage to oaboutpet.FScrollCount + oaboutpet.StartScrollPage - 1 %>
    		<% if i>oaboutpet.FTotalpage then Exit for %>
    		<% if CStr(page)=CStr(i) then %>
    		<font color="red">[<%= i %>]</font>
    		<% else %>
    		<a href="javascript:goPage('<%= i %>');">[<%= i %>]</a>
    		<% end if %>
    	<% next %>

    	<% if oaboutpet.HasNextScroll then %>
    		<a href="javascript:goPage('<%= i %>');">[next]</a>
    	<% else %>
    		[next]
    	<% end if %>
    </td>
</tr>
</table>
</form>
<iframe name="xLink" id="xLink" frameborder="0" width="100%" height="300"></iframe>
<% SET oaboutpet = nothing %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
