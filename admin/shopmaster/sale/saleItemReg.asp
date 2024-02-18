<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  할인 상품 관리
' History : 2008.04.08 정윤정 생성
'           2013.06.21 허진원 / 할인율 표시 및 경고문 추가
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/util/datelib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/classes/items/itemsalecls.asp"-->
<!-- #include virtual="/admin/eventmanage/common/event_function.asp"-->
<%
Dim sCode, clsSale,clsSaleItem
Dim sTitle,isRate, isMargin, isStatus,eCode, egCode, dSDay, dEDay, isUsing, dOpenDay,isMValue, smargin
Dim acURL
Dim iTotCnt, arrList,intLoop
Dim iPageSize, iCurrpage ,iDelCnt
Dim iStartPage, iEndPage, iTotalPage, ix,iPerCnt
dim makerid, sailyn,invalidmargin, sRectItemidArr 

sCode = requestCheckVar(Request("sC"),10)
makerid =  requestCheckVar(Request("makerid"),32)
sailyn	=  requestCheckVar(Request("sailyn"),1)
invalidmargin=  requestCheckVar(Request("invalidmargin"),1)
sRectItemidArr=  requestCheckVar(Request("sRectItemidArr"),400)

acURL =Server.HTMLEncode("/admin/shopmaster/sale/saleitemProc.asp?sC="&sCode)

if sRectItemidArr<>"" then
	dim iA ,arrTemp,arrItemid
	sRectItemidArr = replace(sRectItemidArr,",",chr(10)) 
	sRectItemidArr = replace(sRectItemidArr,chr(13),"") 
	arrTemp = Split(sRectItemidArr,chr(10))

	iA = 0
	do while iA <= ubound(arrTemp) 
		if trim(arrTemp(iA))<>"" then 
			'상품코드 유효성 검사(2008.08.05;허진원)
			if Not(isNumeric(trim(arrTemp(iA)))) then
				Response.Write "<script language=javascript>alert('[" & arrTemp(iA) & "]은(는) 유효한 상품코드가 아닙니다.');history.back();</script>"
				dbget.close()	:	response.End
			else
				arrItemid = arrItemid & trim(arrTemp(iA)) & ","
			end if
		end if
		iA = iA + 1
	loop

	sRectItemidArr = left(arrItemid,len(arrItemid)-1)
end if

'마진형태에 따른 매입가 생성-------------------------------------------------------
Function fnSetSaleSupplyPrice(ByVal MarginType, ByVal MarginValue, ByVal orgPrice, ByVal orgSupplyPrice, ByVal salePrice)
	Dim orgMRate
	if orgPrice <>0 then '원 마진율
		orgMRate = 100-fix(orgSupplyPrice/orgPrice*10000)/100
	end if

	SELECT CASE MarginType
		Case 1	'동일마진
			fnSetSaleSupplyPrice = salePrice- fix(salePrice*(orgMRate/100))
		Case 2	'업체부담
			fnSetSaleSupplyPrice = salePrice-(orgPrice-orgSupplyPrice)
		Case 3	'반반부담
			fnSetSaleSupplyPrice = orgSupplyPrice- fix((orgPrice-salePrice)/2)
		Case 4	'10x10부담
			fnSetSaleSupplyPrice = orgSupplyPrice
		Case 5	'직접설정
			fnSetSaleSupplyPrice = salePrice - fix(salePrice*(MarginValue/100))
	END SELECT
End Function
'-----------------------------------------------------------------------------------
 
'할인 기본정보
set clsSale = new CSale
clsSale.FSCode  = sCode
clsSale.fnGetSaleConts

sTitle 		= clsSale.FSName
isRate 		= clsSale.FSRate
isMargin 	= clsSale.FSMargin
eCode 		= clsSale.FECode
egCode		= clsSale.FEGroupCode
dSDay 		= clsSale.FSDate
dEDay 		= clsSale.FEDate
isStatus 	= clsSale.FSStatus
isUsing     = clsSale.FSUsing
dOpenDay	= clsSale.FOpenDate
isMValue	= clsSale.FSMarginValue
set clsSale = nothing

iCurrpage = Request("iC")	'현재 페이지 번호
IF iCurrpage = "" THEN	iCurrpage = 1
iPageSize = 20		'한 페이지의 보여지는 열의 수
iPerCnt = 10		'보여지는 페이지 간격

'할인 상품정보
set clsSaleItem = new CSaleItem
clsSaleItem.FCPage = iCurrpage
clsSaleItem.FPSize = iPageSize
clsSaleItem.FSCode = sCode
clsSaleItem.FRectMakerid = makerid
clsSaleItem.FRectsailyn = sailyn
clsSaleItem.FRectinvalidmargin =invalidmargin
clsSaleItem.FRectItemidArr = sRectItemidArr 
arrList = clsSaleItem.fnGetSaleItemList
iTotCnt = clsSaleItem.FTotCnt	'전체 데이터  수

iTotalPage 	=  int((iTotCnt-1)/iPageSize) +1  '전체 페이지 수

'동기간내 상품쿠폰 정보 접수
Dim arrItemCoupon, iclp
arrItemCoupon = clsSaleItem.fnGetCouponListBySaleInfo

'공통코드 값 배열로 한꺼번에 가져온 후 값 보여주기
Dim arrsalemargin, arrsalestatus
arrsalemargin = fnSetCommonCodeArr("salemargin",False)
arrsalestatus= fnSetCommonCodeArr("salestatus",False)
%>
<script language="JavaScript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript">
<!--
// 페이지 이동
function jsGoPage(iP){
	location.href="saleItemReg.asp?menupos=<%=menupos%>&sC=<%=sCode%>&iC="+iP;
}

// 새상품 추가 팝업
function addnewItem(eC,egC){
		var popwin;
		if ( eC > 0 ){
			popwin = window.open("/admin/eventmanage/common/pop_eventitem_addinfo.asp?acURL=<%=acURL%>&eC="+eC+"&egC="+egC, "popup_item", "width=800,height=500,scrollbars=yes,resizable=yes");
		}else{
			popwin = window.open("/admin/itemmaster/pop_itemAddInfo.asp?acURL=<%=acURL%>&PR=S", "popup_item", "width=800,height=500,scrollbars=yes,resizable=yes");
		}
		popwin.focus();
}

//전체 선택
function SelectCk(opt){
	var bool = opt.checked;
	AnSelectAllFrame(bool)
}


function CkDisPrice(){
	CkDisOrOrg(true);
}

function CkOrgPrice(){
	CkDisOrOrg(false);
}

//원가 할인가 적용
function CkDisOrOrg(isDisc){
	var frm;
	var pass = false;

	for (var i=0;i<document.forms.length;i++){
		frm = document.forms[i];
		if (frm.name.substr(0,9)=="frmBuyPrc") {
			pass = ((pass)||(frm.cksel.checked));
		}
	}

	var ret;

	if (!pass) {
		alert('선택 아이템이 없습니다.');
		return;
	}


	for (var i=0;i<document.forms.length;i++){
		frm = document.forms[i];
		if (frm.name.substr(0,9)=="frmBuyPrc") {
			if (frm.cksel.checked){
				if(isDisc==true){
					frm.iDSPrice.value = frm.saleprice.value;
					frm.iDBPrice.value = frm.salesupplyprice.value;
					frm.iDSMargin.value= frm.salemargin.value;
					frm.saleItemStatus.value = 7;
				}else{
					frm.iDSPrice.value = frm.orgPrice.value;
					frm.iDBPrice.value = frm.orgSupplyPrice.value;
					frm.iDSMargin.value= frm.orgMarginValue.value;
					frm.saleItemStatus.value = 9;
				}
			}
			reCALbyPrice(frm.itemid.value);
		}
	}
}

//선택상품 저장
function saveArr(){
	var frm;
	var pass = false;
	var ovPer = 0;

	for (var i=0;i<document.forms.length;i++){
		frm = document.forms[i];
		if (frm.name.substr(0,9)=="frmBuyPrc") {
			pass = ((pass)||(frm.cksel.checked));
		}
	}

	var ret;

	if (!pass) {
		alert('선택 아이템이 없습니다.');
		return;
	}

	frmarr.itemid.value = "";
	frmarr.sailyn.value = "";
	frmarr.iDSPrice.value ="";
	frmarr.iDBPrice.value ="";


	for (var i=0;i<document.forms.length;i++){
		frm = document.forms[i];
		if (frm.name.substr(0,9)=="frmBuyPrc") {
			if (frm.cksel.checked){
				//check Not AvaliValue
				if (!IsDigit(frm.iDSPrice.value)){
					alert('숫자만 가능합니다.');
					frm.iDSPrice.focus();
					return;
				}

				if (frm.iDSPrice.value<1){
					alert('금액을 정확히 입력하세요.');
					frm.iDSPrice.focus();
					return;
				}

				if (!IsDigit(frm.iDBPrice.value)){
					alert('숫자만 가능합니다.');
					frm.iDBPrice.focus();
					return;
				}

				if (frm.iDBPrice.value<1){
					alert('금액을 정확히 입력하세요.');
					frm.iDBPrice.focus();
					return;
				}

				if(Math.round((frm.orgPrice.value-frm.iDSPrice.value)/frm.orgPrice.value*100)>=50) {
					ovPer++;
				}

				frmarr.itemid.value = frmarr.itemid.value + frm.itemid.value + ","
				//if (frm.sailyn[0].checked){
					//frmarr.sailyn.value = frmarr.sailyn.value + "Y" + ","
				//}else{
					//frmarr.sailyn.value = frmarr.sailyn.value + "N" + ","
				//}
				frmarr.iDSPrice.value = frmarr.iDSPrice.value + frm.iDSPrice.value + ","
				frmarr.iDBPrice.value = frmarr.iDBPrice.value + frm.iDBPrice.value + ","
				frmarr.saleItemStatus.value = frmarr.saleItemStatus.value + frm.saleItemStatus.value+","

			}
		}
	}

	if(ovPer>0) {
		if(!confirm('!!!\n\n\n선택 상품중에 할인율이 매우 높은 상품(50%+)이 있습니다!\n\n입력하신 내용이 맞습니까?\n\n')) {
			return;
		}
	}

	var ret = confirm('저장 하시겠습니까?');

	if (ret){
		frmarr.submit();
	}

}

function delArr(){
	var frm;
	var pass = false;

	for (var i=0;i<document.forms.length;i++){
		frm = document.forms[i];
		if (frm.name.substr(0,9)=="frmBuyPrc") {
			pass = ((pass)||(frm.cksel.checked));
		}
	}

	var ret;

	if (!pass) {
		alert('선택 아이템이 없습니다.');
		return;
	}

	frmdel.itemid.value = "";

	for (var i=0;i<document.forms.length;i++){
		frm = document.forms[i];
		if (frm.name.substr(0,9)=="frmBuyPrc") {
			if (frm.cksel.checked){
				frmdel.itemid.value = frmdel.itemid.value + frm.itemid.value + ","
			}
		}
	}

	var ret = confirm('삭제하시겠습니까?');

	if (ret){
		frmdel.submit();
	}

}

// 마진율 재계산
function reCALbyPrice(fid) {
	var frm = document["frmBuyPrc_" + fid];
	if(frm.iDSPrice.value>0) {
		frm.iDSMargin.value = Math.round(((frm.iDSPrice.value-frm.iDBPrice.value)/frm.iDSPrice.value)*100);
	} else {
		frm.iDSMargin.value = 0;
	}

	//할인율 표시
	var iorgPrice = frm.orgPrice.value;
	var isailprice = frm.iDSPrice.value;
	var isalePercent = Math.round((iorgPrice-isailprice)/iorgPrice*100);

	if(isalePercent>=50) {
		document.getElementById("lyrSpct"+fid).style.color="#EE0000";
		document.getElementById("lyrSpct"+fid).style.fontWeight="bold";
	} else {
		document.getElementById("lyrSpct"+fid).style.color="#000000";
		document.getElementById("lyrSpct"+fid).style.fontWeight="normal";
	}
	document.getElementById("lyrSpct"+fid).innerHTML = isalePercent + "%";

}

// 매입가 재계산
function reCALbyMargin(fid) {
	var frm = document["frmBuyPrc_" + fid];
	if(frm.iDSMargin.value>0) {
		frm.iDBPrice.value = Math.round(frm.iDSPrice.value*(1-(frm.iDSMargin.value/100)));
	} else {
		frm.iDBPrice.value = frm.iDSPrice.value;
	}
}
//-->
</script>
<table width="100%" border="0" align="center" cellpadding="3" cellspacing="0" class="a">
<tr> 
	<td width="100%">
		<table  border="0"  width="100%" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
		<tr>
			<td align="center" bgcolor="<%= adminColor("tabletop") %>" width="100">할인코드</td>
			<td bgcolor="#FFFFFF" ><%=sCode%></td>
			<td align="center" bgcolor="<%= adminColor("tabletop") %>"  width="100">할인명</td>
			<td bgcolor="#FFFFFF"  ><%=sTitle%></td>
		</tr>
		<tr>	
			<td align="center" bgcolor="<%= adminColor("tabletop") %>"   >이벤트코드(그룹)</td>
			<td bgcolor="#FFFFFF"  ><%If eCode > 0 THEN%><%=eCode%><%If egCode > 0 THEN%>(<%=egCode%>)<%END IF%><%END IF%>&nbsp;</td>
			<td align="center" bgcolor="<%= adminColor("tabletop") %>" >상태</td>
			<td bgcolor="#FFFFFF" ><%=fnGetCommCodeArrDesc(arrsalestatus,isStatus)%></td>
		</tr>
		<tr>	
			<td align="center" bgcolor="<%= adminColor("tabletop") %>"  >기간</td>
			<td bgcolor="#FFFFFF" colspan="3"><%=dSDay%> ~ <%=dEDay%></td>
		</tr>
		</table>
	</td>
</tr>
<tr>
	<td>
		<form name="frmSearch" method="get" action="">
			<input type=hidden name=menupos value="<%=menupos%>">
			<input type=hidden name=sC value="<%=sCode%>">
			<input type=hidden name=iC value="<%=iCurrpage%>">
		<table width="100%" border="0" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>"> 
			<tr>
				<td width="100" bgcolor="#EEEEEE" align="center">검색조건</td>
				<td bgcolor="#ffffff">
					<table   border="0"  cellpadding="3" cellspacing="1" class="a">
					<tr>
						<td width="300"> 브랜드: 
					   	<% drawSelectBoxDesignerWithName "makerid",makerid %> 
						</td> 
						<td>상품코드:</td>
						<td rowspan="2" bgcolor="#FFFFFF"><textarea name="sRectItemidArr" rows="3" cols="10"><%=replace(sRectItemidArr,",",chr(10))%></textarea> </td>  
					</tr> 	
					<tr>
						<td colspan="3"  bgcolor="#FFFFFF">
					    	<input type="checkbox" name="sailyn" value="Y" <% if sailyn="Y" then response.write "checked" %> >세일중인 상품 검색
				            &nbsp;<input type="checkbox" name="invalidmargin" value="Y" <% if invalidmargin="Y" then response.write "checked" %> >마진부족(or 역마진) 상품 검색
				       	</td> 
					</tr> 
				</table>
				</td>
				<td  width="120" bgcolor="#EEEEEE" align="center">
					 <input type="button" class="button" value="등록된 상품 검색" onclick="document.frmSearch.submit();">
				</td> 
			</tr>
		</table>
		</form>
	</td>
</tr>
<tr>
	<td>
		<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" border=0>
		<form name=frmdummi>
		<input type="hidden" name="menupos" value="<%=menupos%>">
		<tr height="40" valign="bottom">
			<td align="left"><input type=button value="선택상품수정" onClick="saveArr()" class="button">
			<!--<input type=button value="선택상품삭제" onClick="delArr()" class="button">		-->
			</td>
			<td align="right">
			할인율: <font color="blue"><%=isRate%>%</font>, 마진구분: <%=fnGetCommCodeArrDesc(arrsalemargin,isMargin)%><%IF isMargin = 5 THEN%>,&nbsp;할인마진율: <font color="blue"><%=isMValue%>%</font> <%END IF%>
			<input type="button" value="할인적용" onClick="CkDisPrice();" class="button">
			<input type="button" value="원가적용" onClick="CkOrgPrice();" class="button">
			&nbsp;&nbsp;
			<input type="button" value="새상품 추가" onclick="addnewItem(<%=eCode%>,<%=egCode%>);" class="button">
			</td>
		</tr>
		</form>
		</table>
	</td>
</tr>
<tr>
	<td colspan="2">
		<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
		<tr bgcolor="#FFFFFF">
			<td colspan="17" align="left">검색결과 : <b><%=iTotCnt%></b>&nbsp;&nbsp;페이지 : <b><%=iCurrpage%> / <%=iTotalPage%></b></td>
		</tr>
		<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
			<td><input type="checkbox" name="ck_all" onclick="SelectCk(this)"></td>
				<td align="center">상품ID</td>
				<td align="center" >이미지</td>
				<td align="center">브랜드</td>
				<td align="center">상품명</td>
				<td align="center">계약<br>구분</td>
				<td align="center">할인상태</td>
				<td align="center">현재<br>판매가</td>
				<td align="center">현재<br>매입가</td>
				<td align="center">현재<br>마진율</td>

				<td align="center">원<br>판매가</td>
				<td align="center">원<br>매입가</td>
				<td align="center">원<br>마진율</td>

				<td align="center">할인율</td>
				<td align="center">할인<br>판매가</td>
				<td align="center">할인<br>매입가</td>
				<td align="center">할인<br>마진율</td>
		</tr>
		<%	Dim mSPrice, mSBPrice, iSaleMargin, iOrgMargin, iSalePercent
			Dim cpSP, cpSB, cpSM, strCpDesc, strCpList
			iSaleMargin=0
			iOrgMargin = 0
		IF isArray(arrList) THEN
			For intLoop = 0 To UBound(arrList,2)
			mSPrice  =arrList(13,intLoop) - (arrList(13,intLoop)*(isRate/100))
			mSBPrice = fnSetSaleSupplyPrice(isMargin,isMValue,arrList(13,intLoop),arrList(14,intLoop),mSPrice)
			if mSPrice<>0 then iSaleMargin =  100-fix(mSBPrice/mSPrice*10000)/100
			 if arrList(13,intLoop)<>0 then iOrgMargin= 100-fix(arrList(14,intLoop)/arrList(13,intLoop)*10000)/100
			 iSalePercent = ((arrList(13,intLoop)-arrList(2,intLoop))/arrList(13,intLoop))*100

			cpSP=0: cpSB=0: cpSM=0: strCpDesc="": strCpList=""
			if isArray(arrItemCoupon) then

				for icLp=0 to ubound(arrItemCoupon,2)
					if cStr(arrItemCoupon(4,icLp))=cStr(arrList(1,intLoop)) then
						'상품쿠폰판매가
						Select Case arrItemCoupon(1,icLp)
							Case "1"
								cpSP = mSPrice- CLng(arrItemCoupon(2,icLp)*mSPrice/100)
							Case "2"
								cpSP = mSPrice- arrItemCoupon(2,icLp)
							Case Else
								cpSP = mSPrice
						End Select
						'상품쿠폰매입가
						cpSB = arrItemCoupon(5,icLp)
						'상품쿠폰마진
						if cpSB>0 then cpSM = formatNumber(100-fix(cpSB/cpSP*10000)/100,0)

						strCpList = strCpList & "<tr align='center' onclick=""window.open('/admin/shopmaster/itemcouponlist.asp?menupos=786&research=on&iSerachType=1&sSearchTxt=" & arrItemCoupon(0,icLp) & "')"">" &_
								"<td>[" & arrItemCoupon(0,icLp) & "]</td>" &_
								"<td>" & arrItemCoupon(3,icLp) & "</td>" &_
								"<td>" & FormatNumber(cpSP,0) & "원</td>" &_
								"<td>" & FormatNumber(cpSB,0) & "원</td>" &_
								"<td " & chkIIF(cpSM<=5,"style='color:#ee0000;'","") & ">" & FormatNumber(cpSM,0) & "%</td>" &_
								"<td>" & left(arrItemCoupon(6,icLp),10) & "</td>" &_
								"<td>" & left(arrItemCoupon(7,icLp),10) & "</td>" &_
								"</tr>"
					end if
				next

				if strCpList<>"" then
					strCpDesc = "<div><font color=darkgreen style='cursor:pointer;' onmouseover=""$(this).find('div').show()"" onmouseout=""$(this).find('div').hide()"">상품쿠폰 ▶" &_
							"<div style='display:none;position:absolute;border:1px solid #C0C0C0;padding:5px;background-color:#FFFFFF;margin:-10px -20px;'>" &_
							"<table width='600' border='0' cellpadding='3' cellspacing='1' class='a'>" &_
							"<tr><td colspan='7' align='left'><strong>할인기간중 진행되는 쿠폰</strong></td></tr>" &_
							"<tr align='center' bgcolor='#F0F0F0'>" &_
							"<td colspan='2'>쿠폰명</td>" &_
							"<td>쿠폰할인가</td>" &_
							"<td>쿠폰매입가</td>" &_
							"<td>쿠폰할인마진</td>" &_
							"<td>시작일</td>" &_
							"<td>종료일</td>" &_
							"</tr>" &_
							strCpList &_
							"</table>" &_
							"</div></font></div>"
				end if

			end if
			%>
			<form name="frmBuyPrc_<%=arrList(1,intLoop)%>" >
			<input type=hidden name="itemid" value="<%=arrList(1,intLoop)%>">
			<input type=hidden name="saleprice" value="<%=mSPrice%>">
			<input type=hidden name="salesupplyprice" value="<%=mSBPrice%>">
			<input type=hidden name="salemargin" value="<%=iSaleMargin%>">
			<input type=hidden name="orgPrice" value="<%=arrList(13,intLoop)%>">
			<input type=hidden name="orgSupplyPrice" value="<%=arrList(14,intLoop)%>">
			<input type=hidden name="orgMarginValue" value="<%=iOrgMargin%>">
			<input type=hidden name="saleItemStatus" value="<%=arrList(4,intLoop)%>">
		 <tr align="center" bgcolor=<%IF cint(arrList(4,intLoop)) = 8 THEN%>"#B3B3B3"<%ELSE%>"#FFFFFF"<%END IF%>>
			    <td><input type="checkbox" name="cksel" onClick="AnCheckClick(this);"></td>
			    <td><%=arrList(1,intLoop)%></td>
			    <td><%IF arrList(9,intLoop) <> "" THEN%><img src="http://webimage.10x10.co.kr/image/small/<%=GetImageSubFolderByItemid(arrList(1,intLoop))%>/<%=arrList(9,intLoop)%>"><%END IF%></td>
			    <td><%=db2html(arrList(7,intLoop))%></td>
			    <td align="left">&nbsp;<%=db2html(arrList(8,intLoop))%></td>
			    <td><%= fnColor(arrList(17,intLoop),"mw") %></td>
			    <td>
			    	<%= fnColor(arrList(10,intLoop),"yn") %>&nbsp;<%IF arrList(4,intLoop) = 6 THEN%><font color="blue"><%END IF%><%=fnGetCommCodeArrDesc(arrsalestatus,arrList(4,intLoop))%>
			    	<%=chkIIF(strCpDesc>"",strCpDesc,"")%>
			    </td>

			    <td><%=formatnumber(arrList(11,intLoop),0)%></td>
			    <td><%=formatnumber(arrList(12,intLoop),0)%></td>
			    <td><% if arrList(11,intLoop)<>0 then %>
					<%= 100-fix(arrList(12,intLoop)/arrList(11,intLoop)*10000)/100 %>%
					<% end if %>  
				</td>


			    <td><%=formatnumber(arrList(13,intLoop),0)%></td>
			    <td><%=formatnumber(arrList(14,intLoop),0)%></td>
			    <td><%=iOrgMargin%>%</td>

				<td id="lyrSpct<%=arrList(1,intLoop)%>" style="<%=chkIIF(iSalePercent>=50,"color:#EE0000;font-weight:bold;","")%>"><%=formatnumber(iSalePercent,0)%>%</td>
			<%IF cint(arrList(4,intLoop)) = 8 or  cint(arrList(4,intLoop)) = 9 THEN%>
				<td><input type="text" name="iDSPrice" size="6" maxlength="9" value="0" style="text-align:right;" onkeyup="reCALbyPrice('<%=arrList(1,intLoop)%>')"><br><%=arrList(2,intLoop)%></td>
			    <td><input type="text" name="iDBPrice" size="6" maxlength="9" value="0" style="text-align:right;" onkeyup="reCALbyPrice('<%=arrList(1,intLoop)%>')"><br><%=arrList(3,intLoop)%></td>
			    <td><input type="text" name="iDSMargin" value="0" style="text-align:right;" size="4" onkeyup="reCALbyMargin('<%=arrList(1,intLoop)%>')">%<br><%  if arrList(2,intLoop)<>0 then smargin= 100-fix(arrList(3,intLoop)/arrList(2,intLoop)*10000)/100 	%></td>
			<%ELSE%>
			    <td><input type="text" name="iDSPrice" size="6" maxlength="9" value="<%=arrList(2,intLoop)%>" style="text-align:right;" onkeyup="reCALbyPrice('<%=arrList(1,intLoop)%>')"></td>
			    <td><input type="text" name="iDBPrice" size="6" maxlength="9" value="<%=arrList(3,intLoop)%>" style="text-align:right;" onkeyup="reCALbyPrice('<%=arrList(1,intLoop)%>')"></td>
			    <td><%  if arrList(2,intLoop)<>0 then smargin= 100-fix(arrList(3,intLoop)/arrList(2,intLoop)*10000)/100 	%> 
					<input type="text" name="iDSMargin" value="<%=smargin%>" style=text-align:right;" size="4" onkeyup="reCALbyMargin('<%=arrList(1,intLoop)%>')">%
			    </td>
			<%END IF%>
		</tr>
		</form>
		<%	next %>
		<% else %>
		<tr>
			<td colspan="17" bgcolor="#ffffff" align="center">등록된 내역이 없습니다.</td>
		</tr>
		<%
		END IF%>
		<!-- 페이징처리 -->
			<%
			iStartPage = (Int((iCurrpage-1)/iPerCnt)*iPerCnt) + 1

			If (iCurrpage mod iPerCnt) = 0 Then
				iEndPage = iCurrpage
			Else
				iEndPage = iStartPage + (iPerCnt-1)
			End If
			%>
			<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
			    <tr valign="bottom" height="25">
			        <td valign="bottom" align="center">
			         <% if (iStartPage-1 )> 0 then %><a href="javascript:jsGoPage(<%= iStartPage-1 %>)" onfocus="this.blur();">[pre]</a>
					<% else %>[pre]<% end if %>
			        <%
						for ix = iStartPage  to iEndPage
							if (ix > iTotalPage) then Exit for
							if Cint(ix) = Cint(iCurrpage) then
					%>
						<a href="javascript:jsGoPage(<%= ix %>)" class="menu_link3" onfocus="this.blur();"><font color="00abdf"><strong><%=ix%></strong></font></a>
					<%		else %>
						<a href="javascript:jsGoPage(<%= ix %>)" class="menu_link3" onfocus="this.blur();"><%=ix%></a>
					<%
							end if
						next
					%>
			    	<% if Cint(iTotalPage) > Cint(iEndPage)  then %><a href="javascript:jsGoPage(<%= ix %>)" onfocus="this.blur();">[next]</a>
					<% else %>[next]<% end if %>
			        </td>
			        <td  width="50" align="right"><a href="saleList.asp?menupos=<%=menupos%>"><img src="/images/icon_list.gif" border="0"></a></td>
			    </tr>
		</table>
	</td>
</tr>
</table>
<form name="frmarr" method="post" action="saleItemPRoc.asp">
<input type="hidden" name="mode" value="U">
<input type="hidden" name="menupos" value="<%=menupos%>">
<input type="hidden" name="sC" value="<%=sCode%>">
<input type="hidden" name="iC" value="<%=iCurrpage%>">
<input type="hidden" name="itemid" value="">
<input type="hidden" name="sailyn" value="">
<input type="hidden" name="iDSPrice" value="">
<input type="hidden" name="iDBPrice" value="">
<input type="hidden" name="saleItemStatus" value="">
<input type="hidden" name="saleStatus" value="<%=isStatus%>">
</form>
<form name="frmdel" method="post" action="saleItemPRoc.asp">
<input type="hidden" name="mode" value="D">
<input type="hidden" name="sC" value="<%=sCode%>">
<input type="hidden" name="itemid" value="">
</form>
<%
set clsSaleItem = nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->