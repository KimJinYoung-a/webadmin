<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  상시할인 리스트
' History : 2018. 01.09
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/util/datelib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/classes/items/itemalltimesalecls.asp"--> 
<!-- #include virtual="/admin/eventmanage/common/event_function.asp"-->
<!-- #include virtual="/admin/lib/incPageFunction.asp" -->
<%
dim iCurrpage, iTotCnt, iTotPage,iPageSize,iPerCnt
dim CATSale, arrList, intLoop
dim makerid,rdoSale, itemid, invalidmargin
dim dispcate ,couponyn
 dim iSalePercent,isort
 
iCurrpage = requestCheckVar(Request("iC"),10)	'현재 페이지 번호
makerid =  requestCheckVar(Request("makerid"),32)	
rdoSale =  requestCheckVar(Request("rdoSale"),1)	 
itemid =  requestCheckVar(Request("itemid"),1024)
dispCate = requestCheckvar(request("disp"),16)
couponyn		= requestCheckvar(request("couponyn"),1)
isort= requestCheckvar(request("isort"),1)

invalidmargin=  requestCheckVar(Request("invalidmargin"),1)

	IF iCurrpage = "" THEN
		iCurrpage = 1
	END IF
	iPageSize = 30		'한 페이지의 보여지는 열의 수
	iPerCnt = 10		'보여지는 페이지 간격
	if rdoSale = "" then rdoSale ="1"
	if isort = "" then isort ="1"
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

set CATSale = new ClsAllTimeSale
CATSale.FPSize = iPageSize
CATSale.FCPage = iCurrpage
CATSale.FRectMakerid = makerid
CATSale.FRectSale = rdoSale 
CATSale.FRectDispcate	= dispCate
CATSale.FRectitemid = itemid 
CATSale.FRectcouponyn = couponyn
CATSale.FRectSort = isort
CATSale.FRectinvalidmargin =invalidmargin
arrList = CATSale.fnGetItemList
iTotCnt = CATSale.FTotCnt
set CATSale = nothing

 iTotPage 	=  int((iTotCnt-1)/iPageSize) +1  '전체 페이지 수
%>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript">
//전체 선택
function SelectCk(opt){
	var bool = opt.checked;
	AnSelectAllFrame(bool)
}


//원가적용
function jsSetOrgPrice(){
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

	frmarr.itemid.value = "";
	
	for (var i=0;i<document.forms.length;i++){
		frm = document.forms[i];
		if (frm.name.substr(0,9)=="frmBuyPrc") {
			if (frm.cksel.checked){ 
				frmarr.itemid.value = frmarr.itemid.value + frm.cksel.value + ","   
			}
		}
	}
 
	 if(confirm("선택한 상품을 원가적용 하시겠습니까? 이벤트 할인이 있는 경우 모든 이벤트 할인은 종료됩니다. ")){
	 		frmarr.hidM.value ="O";
	 		frmarr.submit();
		}
}

//상시할인 등록
function jsSetSale(){
	var frm;
	var pass = false;
	var ovPer = 0;
	var ovLimitPer = 0;

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

				if(frm.mwdiv.value!='M') {
					var limitMarPrc = frm.orgsuplycash.value-((frm.orgPrice.value-frm.iDSPrice.value)*0.5);
					var limitMarPer = (frm.iDSPrice.value-limitMarPrc)/frm.iDSPrice.value*100;
					if(parseInt(limitMarPrc)>parseInt(frm.iDBPrice.value)) {
						ovLimitPer++;
					}
				}


				frmarr.itemid.value = frmarr.itemid.value + frm.cksel.value + "," 
				frmarr.iDSPrice.value = frmarr.iDSPrice.value + frm.iDSPrice.value + ","
				frmarr.iDBPrice.value = frmarr.iDBPrice.value + frm.iDBPrice.value + "," 

			}
		}
	}

	if(ovPer>0) {
		if(!confirm('!!!\n\n\n선택 상품중에 할인율이 매우 높은 상품(50%+)이 있습니다!\n\n입력하신 내용이 맞습니까?\n\n')) {
			return;
		}
	} 

	if(ovLimitPer>0) {
		if(!confirm('!!!\n\n\n선택 상품중에 업체 할인 분담율이 50%를 넘는 상품이 있습니다!\n\n입력하신 내용이 맞습니까?\n\n')) {
			return;
		}
	} 

	 if(confirm("선택한 상품을 상시할인 하시겠습니까? 이벤트 할인이 걸려있는 경우 모든 이벤트 할인은 종료됩니다.")){
	 		frmarr.hidM.value ="S";
	 		frmarr.submit();
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
 	frm.cksel.checked = true;
}

// 매입가 재계산
function reCALbyMargin(fid) {
	var frm = document["frmBuyPrc_" + fid];
	if(frm.iDSMargin.value>0) {
		frm.iDBPrice.value = Math.round(frm.iDSPrice.value*(1-(frm.iDSMargin.value/100)));
	} else {
		frm.iDBPrice.value = frm.iDSPrice.value;
	}
	frm.cksel.checked = true;
}

// 새상품 추가 엑셀 팝업
function jsRegExcel(){
	var popwin;
	popwin = window.open("/admin/shopmaster/alltimesale/popRegFile.asp", "popup_item", "width=500,height=230,scrollbars=yes,resizable=yes");
	popwin.focus();
}

function jstrSort(vsorting){

	var tmpSorting = document.getElementById("img"+vsorting);
	document.frmSearch.isort.value= vsorting;
//	
//	if (-1 < tmpSorting.src.indexOf("_alpha")){
//		frm.isort.value= vsorting;
//	}else if (-1 < tmpSorting.src.indexOf("_bot")){
//		frm.isort.value= vsorting ;
//	}else{
//		frm.isort.value= vsorting;
//	}
	document.frmSearch.submit();
}
</script>
<form name="frmarr" method="post" action="procATSale.asp">
	<input type="hidden" name="menupos" value="<%=menupos%>">
	<input type="hidden" name="hidM" value="">
	<input type="hidden" name="itemid" value=""> 
	<input type="hidden" name="iDSPrice" value="">
	<input type="hidden" name="iDBPrice" value="">	
</form>
<table width="100%" border="0" align="left" class="a" cellpadding="3" cellspacing="1"  > 
<tr>
	<td>
		<form name="frmSearch" method="get" action="">
			<input type=hidden name=menupos value="<%=menupos%>"> 
			<input type=hidden name=iC value="1"> 
			<input type="hidden" name="isort" value="<%= isort %>">
		<table width="100%" border="0" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
			<tr>
				<td width="100" bgcolor="#EEEEEE" align="center">검색조건</td>
				<td bgcolor="#ffffff">
					<table   border="0"  cellpadding="3" cellspacing="1" class="a">
					<tr>
						<td > 브랜드:
					   	<% drawSelectBoxDesignerWithName "makerid",makerid %>
						</td> 
							<td style="padding-left:20px;">상품코드:</td>
						<td rowspan="2" bgcolor="#FFFFFF"><textarea name="itemid" rows="3" cols="10"><%=replace(itemid,",",chr(10))%></textarea> </td>
					</tr>
					<tr>
						<td  colspan="2">
							전시카테고리: <!-- #include virtual="/common/module/dispCateSelectBox.asp"-->
						</td>
					</tr>
					<tr>
						<td>할인:  <input type="radio" name="rdoSale" value="0" <%if rdoSale ="0" then%>checked<%end if%>>전체
							<input type="radio" name="rdoSale" value="1" <%if rdoSale ="1" then%>checked<%end if%>>상시
							<input type="radio" name="rdoSale" value="2" <%if rdoSale ="2" then%>checked<%end if%>>이벤트
							<input type="radio" name="rdoSale" value="3" <%if rdoSale ="3" then%>checked<%end if%>>이벤트+상시
							<input type="radio" name="rdoSale" value="9" <%if rdoSale ="9" then%>checked<%end if%>>할인안함 
								 	&nbsp;&nbsp;쿠폰: <% drawSelectBoxCouponYN "couponyn", couponyn %> 
							</td>
							  
				 	<td  bgcolor="#FFFFFF">
					      <input type="checkbox" name="invalidmargin" value="Y" <% if invalidmargin="Y" then response.write "checked" %> >마진부족(or 역마진) 상품 검색
				    </td>
				  
				  </tr> 
				</table>
				</td>
				<td  width="120" bgcolor="#EEEEEE" align="center">
					 <input type="button" class="button" value="검색" style="width:100px;height:50px;" onclick="document.frmSearch.submit();">
				</td>
			</tr>
		</table>
		</form>
	</td>
</tr>
<tr>
	<td>
		<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" border=0> 
		<input type="hidden" name="menupos" value="<%=menupos%>">
		<tr height="40" valign="bottom">
			<td align="right">
				<input type=button value="선택 원가 적용" onClick="jsSetOrgPrice()" class="button"  style="height:30px;width:100px;">
				<input type=button value="선택 상시할인 적용" onClick="jsSetSale()" class="button" style="height:30px;width:150px;"> 
				&nbsp;&nbsp;&nbsp;
				<input type=button value="엑셀등록 상시할인 적용" onClick="jsRegExcel()" class="button"  style="height:30px;width:150px;">
			</td> 
		</tr> 
		</table>
	</td>
</tr>
<tr>
	<td colspan="2">  
		<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
		<tr bgcolor="#FFFFFF">
			<td colspan="17" align="left">검색결과 : <b><%=formatnumber(iTotCnt,0)%></b>&nbsp;&nbsp;페이지 : <b><%=iCurrpage%> / <%=formatnumber(iTotPage,0)%></b></td>
		</tr>
		<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
				<td><input type="checkbox" name="chkAll" onClick="SelectCk(this);"></td>
				<td onClick="jstrSort('2'); return false;" style="cursor:pointer;">상품ID
						<img src="/images/list_lineup_bot<%=CHKIIF(isort="2" ,"_on","")%>.png" id="img2">
					</td>
				<td >이미지</td>
				<td>브랜드</td>
				<td>상품명</td>
				<td>계약구분</td>
				<td>할인상태</td>
				<td>판매가</td>
				<td>매입가</td>
				<td>마진율</td> 
				<td>할인율</td> 
				<td>할인 판매가</td>
				<td>할인 매입가</td>
				<td>할인 마진율</td> 
				<td>할인 구분</td> 
				<td onClick="jstrSort('1'); return false;" style="cursor:pointer;">마지막변경일
					<img src="/images/list_lineup_bot<%=CHKIIF(isort="1","_on","")%>.png" id="img1">
					</td> 
		</tr> 
		<%if isArray(arrList) then
				for intLoop = 0 To uBound(arrList,2)
				if (arrList(7,intLoop)=0) then
				    iSalePercent = 0
				else
		 		    iSalePercent = (1-(CDBl(arrList(9,intLoop))/CDBl(arrList(7,intLoop))))*100
		 	    end if
			%>
		<form name="frmBuyPrc_<%=arrList(0,intLoop)%>" >	
			<input type=hidden name="orgPrice" value="<%=arrList(7,intLoop)%>">
			<input type=hidden name="orgsuplycash" value="<%=arrList(8,intLoop)%>">
			<input type=hidden name="mwdiv" value="<%=arrList(11,intLoop)%>">
		<tr bgcolor="#ffffff" align="center">
				<td><input type="checkbox" name="cksel" value="<%= arrList(0,intLoop)%>" <%=CHKIIF(arrList(0,intLoop)=0,"disabled","") %>></td>
				<td><a href="<%=wwwURL%>/shopping/category_prd.asp?itemid=<%= arrList(0,intLoop) %>" target="_blank" title="미리보기">		<%=arrList(0,intLoop)%></a></td>   
				<td ><% if ((Not IsNULL(arrList(3,intLoop))) and (arrList(3,intLoop)<>"")) then %><img src="<%= webImgUrl & "/image/small/" + GetImageSubFolderByItemid(arrList(0,intLoop)) + "/"  + arrList(3,intLoop)%>"><%end if%></td>
				<td><%=arrList(1,intLoop)%></td>
				<td><%=arrList(2,intLoop)%></td>
				<td><%=fnColor(arrList(11,intLoop),"mw") %></td>
				<td><%=fnColor(arrList(4,intLoop),"yn")%></td>
				<td><%=FormatNumber(arrList(7,intLoop),0)%>
						<% 		'할인가(할인율=(소비자가-할인가)/소비자가*100) 
						if arrList(4,intLoop) ="Y" then %>
						    <% if (arrList(7,intLoop)<>0) then %>
						<br><font color=#F08050>(할)<%=FormatNumber(arrList(9,intLoop),0)%></font>
						    <% end if %>
						<% end if %>
						<%'쿠폰가 
						if arrList(16,intLoop)="Y" then
						 
							Select Case arrList(17,intLoop)
								Case "1" '% 쿠폰
						%>
							<br><font color=#5080F0>(쿠)<%=FormatNumber(arrList(5,intLoop)-(CDBl(arrList(18,intLoop)*arrList(5,intLoop)/100)),0)%></font>  
						<%
								Case "2" '원 쿠폰
						%>		
							<br><font color=#5080F0>(쿠)<%=FormatNumber(arrList(5,intLoop),0)%></font>
						<%			
							end Select
						end if	
						%> 
				</td>
				<td><%=FormatNumber(arrList(8,intLoop),0)%>
					<% '할인가
						if arrList(4,intLoop) ="Y" then
					%>		
						 <br><font color=#F08050><%=FormatNumber(arrList(10,intLoop),0) %></font> 
					<%
						end if
						'쿠폰가
						if  arrList(16,intLoop)="Y" then
							if arrList(17,intLoop)="1" or arrList(17,intLoop)="2" then
								if  arrList(19,intLoop)=0 or isNull(arrList(19,intLoop)) then
									Response.Write "<br><font color=#5080F0>" & FormatNumber(arrList(6,intLoop),0) & "</font>"
								else
									Response.Write "<br><font color=#5080F0>" & FormatNumber(arrList(19,intLoop),0) & "</font>"
								end if
							end if
						end if
				%>
				</td>
				<td>
					<%=fnPercent(arrList(8,intLoop),arrList(7,intLoop),1)%>
					<%
						'할인가
						if arrList(4,intLoop) ="Y"  then
							Response.Write "<br><font color=#F08050>" & fnPercent(arrList(10,intLoop),arrList(9,intLoop),1) & "</font>"
						end if
						'쿠폰가
						if arrList(16,intLoop)="Y" then
							Select Case  arrList(17,intLoop)
								Case "1"
									if arrList(19,intLoop)=0 or isNull(arrList(19,intLoop)) then
										Response.Write "<br><font color=#5080F0>" & fnPercent(arrList(6,intLoop),arrList(5,intLoop)-(CDBl(arrList(18,intLoop)*arrList(5,intLoop)/100)),1) & "</font>"
									else
										Response.Write "<br><font color=#5080F0>" & fnPercent(arrList(19,intLoop),arrList(5,intLoop)-(CDBl(arrList(18,intLoop)*arrList(5,intLoop)/100)),1) & "</font>"
									end if
								Case "2"
									if arrList(19,intLoop)=0 or isNull(arrList(19,intLoop)) then
										Response.Write "<br><font color=#5080F0>" & fnPercent(arrList(6,intLoop),arrList(5,intLoop),1) & "</font>"
									else
										Response.Write "<br><font color=#5080F0>" & fnPercent(arrList(19,intLoop),arrList(5,intLoop),1) & "</font>"
									end if
							end Select 
					end if
				%>
				</td> 
				<td id="lyrSpct<%=arrList(0,intLoop)%>" style="<%=chkIIF(iSalePercent>=50,"color:#EE0000;font-weight:bold;","")%>"><%=formatnumber(iSalePercent,0)%> %</td>
				
				<!--<d><input type="text" name="isPro" value="<%if arrList(9,intLoop) > 0 then%><%=replace(fnPercent(arrList(9,intloop),arrList(7,intLoop),0),"%","")%><%end if%>" style="text-align:right;width:50px;" class="text">%</td>  -->
				<td><input type="text" name="iDSPrice" value="<%=arrList(9,intLoop)%>" style="text-align:right;width:100px;" class="text" onkeyup="reCALbyPrice('<%=arrList(0,intLoop)%>')"></td>
				<td><input type="text" name="iDBPrice" value="<%=arrList(10,intLoop)%>" style="text-align:right;width:100px;" class="text" onkeyup="reCALbyPrice('<%=arrList(0,intLoop)%>')"></td>
				<td><input type="text" name="iDSMargin" value="<%if arrList(10,intLoop) > 0 then%><%=replace(fnPercent(arrList(10,intloop),arrList(9,intLoop),0),"%","")%><%end if%>" style="text-align:right;width:50px;" class="text" onkeyup="reCALbyMargin('<%=arrList(0,intLoop)%>')">%</td>  
				<td><%if arrList(4,intLoop) ="Y" then%>
						<% if   isNull(arrList(20,intLoop)) or (not isNull(arrList(20,intLoop)) and  arrList(24,intLoop) ="Y") then %>  
						<font color="blue">상시
						<%if arrList(24,intLoop) ="Y" then%>
						[ <%if not isNull(arrList(22,intLoop)) and arrlist(22,intLoop) <> "" then %><%=FormatNumber(arrList(22,intLoop),0)%><%end if%>/ <%if not isNull(arrList(23,intLoop)) and arrlist(23,intLoop) <> "" then %><%=formatnumber(arrList(23,intLoop),0)%><%end if%> ]
					  <%end if %>
					  <br></font>
					  <%end if %>
					<% if not isNull(arrList(20,intLoop)) then %> 
							<a href="/admin/shopmaster/sale/saleReg.asp?sC=<%=arrList(20,intLoop)%>&menupos=290&sRectitemidArr=<%=arrList(21,intLoop)%>" target="_blank"><font color="red">이벤트(<%=arrList(20,intLoop)%>)</font></a>  
					<%end if%>			
				<%end if%>					
				</td>
				<td><%=arrList(25,intLoop)%></td>
		</tr>	
	</form>
		<%  next
		end if %>
	</table>
</td>
</tr>
</table>
<!-- 페이징처리 --> 
<table width="100%" cellpadding="10">
	<tr>
		<td align="center">  
 			<%sbDisplayPaging "iC", iCurrPage, iTotCnt, iPageSize, 10,menupos %>
		</td>
	</tr>
</table>