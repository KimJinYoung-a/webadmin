<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' History : 2008.04.04 정윤정 생성
'           2010.07.05 허진원 - 쿠폰할인 조건 추가
'						2013.12.24 정윤정 - 상품코드 검색 콤마연결에서 엔터로 변경
' Description : 상품 추가 - 할인, 사은품 상품등록에 사용
'				input - actionURL(db 처리에 필요한 파라미터까지 포함) ex.acURL = "/admin/eventmanage/event/eventitem_process.asp?eC=1234"
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/items/itemcls_2008.asp"-->
<%
dim target, actionURL
dim itemid, itemname, makerid, sellyn, usingyn, danjongyn, deliverytype, limityn, vatyn, sailyn, couponyn, mwdiv,defaultmargin, keyword , sortDiv
dim cdl, cdm, cds , dispCate
dim reAct, ptype
dim page, paraRoad , sCode, notdeal, smode

actionURL 	= Replace(ReplaceRequestSpecialChar(request("acURL")),"||","&")

If (session("ssBctID")="areum531") Then				'2018-01-04 조아름 요청, 검색 아이템수 증가요청
	itemid      = requestCheckvar(request("itemid"),1255)
Else
	itemid      = requestCheckvar(request("itemid"),255)
End If

'itemid      = requestCheckvar(request("itemid"),255)
itemname    = requestCheckvar(request("itemname"),64)
makerid     = requestCheckvar(request("makerid"),32)
sellyn      = requestCheckvar(request("sellyn"),2)
usingyn     = requestCheckvar(request("usingyn"),1)
danjongyn   = requestCheckvar(request("danjongyn"),2)
limityn     = requestCheckvar(request("limityn"),2)
sailyn      = requestCheckvar(request("sailyn"),1)
deliverytype= requestCheckvar(request("deliverytype"),1)
mwdiv       = requestCheckvar(request("mwdiv"),2)
couponyn		= requestCheckvar(request("couponyn"),1)
defaultmargin = requestCheckvar(request("defaultmargin"),10)
keyword			= requestCheckvar(request("keyword"),512)
sortDiv			= requestCheckvar(request("sortDiv"),10)
paraRoad	= requestCheckvar(request("PR"),1)
sCode		= requestCheckvar(request("sC"),10)
reAct       = requestCheckvar(request("reAct"),1)
cdl = requestCheckvar(request("cdl"),10)
cdm = requestCheckvar(request("cdm"),10)
cds = requestCheckvar(request("cds"),10)
dispCate = requestCheckvar(request("disp"),16)
ptype= requestCheckvar(request("ptype"),8)
page = requestCheckvar(request("page"),10)
notdeal = requestCheckvar(request("notdeal"),1)
smode = requestCheckvar(request("smode"),1)

if (page="") then page=1
if sailyn="" and instr(actionURL,"saleitem")>0 and reAct = "" then sailyn="N"			'할인페이지에서 검색된거라면 기본값: 할인안함(쿠폰도 동일)
if couponyn="" and instr(actionURL,"saleitem")>0 and reAct = ""  then couponyn="N"
'if sellyn = "" then sellyn ="Y"
if itemid<>"" then
	dim iA ,arrTemp,arrItemid

	itemid = replace(itemid,chr(13),"") '상품코드검색 엔터로(2013.12.24)
	arrTemp = Split(itemid,chr(10))

	iA = 0
	do while iA <= ubound(arrTemp)

		if trim(arrTemp(iA))<>"" then
			'상품코드 유효성 검사(2008.08.04;허진원)
			if Not(isNumeric(trim(arrTemp(iA)))) then
				Response.Write "<script language=javascript>alert('[" & arrTemp(iA) & "]은(는) 유효한 상품코드가 아닙니다.');history.back();</script>"
				dbget.close()	:	response.End
			else
				arrItemid = arrItemid & trim(arrTemp(iA)) & ","
			end if
		end if
		iA = iA + 1
	loop

	if arrItemid <> "" then
		itemid = left(arrItemid,len(arrItemid)-1)
	else
		itemid = ""
	end if
end if


'==============================================================================
dim oitem

set oitem = new CItem

oitem.FPageSize         = 30
oitem.FCurrPage         = page
oitem.FRectMakerid      = makerid
oitem.FRectItemid       = itemid
oitem.FRectItemName     = itemname
oitem.FRectKeyword		= keyword

oitem.FRectSellYN       = sellyn
oitem.FRectIsUsing      = usingyn
oitem.FRectDanjongyn    = danjongyn
oitem.FRectLimityn      = limityn
oitem.FRectMWDiv        = mwdiv
oitem.FRectDeliveryType = deliverytype
oitem.FRectSailYn       = sailyn
oitem.FRectCouponYn		= couponyn

oitem.FRectCate_Large   = cdl
oitem.FRectCate_Mid     = cdm
oitem.FRectCate_Small   = cds
oitem.FRectDispCate		= dispCate
oitem.FRectSortDiv = SortDiv
if notdeal = "Y" then
oitem.FRectDealYn="Y"
else
oitem.FRectDealYn="N"
end if
oitem.GetItemList

dim i


%>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script language="javascript">
<!--
function jsSerach(){
	var frm;
	frm = document.frm;
	frm.target = "_self";
	frm.action ="pop_itemAddInfo.asp";
	frm.submit();
}

function SelectItems(sType){
var frm;
var itemcount = 0;
frm = document.frmItem;
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
	}else{
		if(typeof(frm.chkitem) !="undefined"){
			itemcount = "<%= oitem.FTotalCount%>";
		  if(confirm("<%= oitem.FTotalCount%>건의 검색된 모든 상품을 추가하시겠습니까?")){
		  	if(itemcount > 1000) {
		  		alert("상품은 최대 1000건까지 가능합니다. 조건을 다시 설정해주세요 ");
		  		return;
		  	}
			frm.itemidarr.value = document.frm.itemid.value;
		  }else{
		  	return;
		  }
		}else{
		 	alert("추가할 상품이 없습니다.");
	   	  	return;
		}
	}
	<% if smode="E" then %>
	if (opener.document.itemreg.itemid.value!=""){
		opener.document.itemreg.itemid.value=opener.document.itemreg.itemid.value+","+frm.itemidarr.value;
	}
	else{
		opener.document.itemreg.itemid.value=frm.itemidarr.value;
	}
	<% else %>
	opener.document.itemreg.itemid.value=frm.itemidarr.value;
	<% end if %>
	window.close();
}

//전체 선택
function jsChkAll(){
var frm;
frm = document.frmItem;
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


// 재고현황 팝업
function PopItemStock(itemid){
	var popwin = window.open("/admin/stock/itemcurrentstock.asp?menupos=709&itemid=" + itemid,"popitemstocklist","width=1000 height=600 scrollbars=yes resizable=yes");
	popwin.focus();
}

// 페이지 이동
function NextPage(ipage){
	document.frm.page.value= ipage;
	document.frm.target = "_self";
	document.frm.action ="pop_itemAddInfo.asp";
	document.frm.submit();
}

//-->
</script>
<!-- 검색 시작 -->
<form name="frm" method="post">
	<input type="hidden" name="page" >
	<input type="hidden" name="sType" >
	<input type="hidden" name="itemidarr" >
	<input type="hidden" name="itemcount" value="0">
	<input type="hidden" name="mode" value="I">
	<input type="hidden" name="acURL" value="<%=actionURL%>">
	<input type="hidden" name="defaultmargin" value="<%=defaultmargin%>">
	<input type="hidden" name="PR" value="<%=paraRoad%>">
	<input type="hidden" name="sC" value="<%=sCode%>">
	<input type="hidden" name="ptype" value="<%=ptype%>">
	<input type="hidden" name="reAct" value="1">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr align="center" bgcolor="<%= adminColor("topbar") %>" >
		<td align="left">
			<table border="0" cellpadding="1" cellspacing="0" class="a">
				<tr>
					<td style="white-space:nowrap;">브랜드: <%	drawSelectBoxDesignerWithName "makerid", makerid %></td>
					<td style="white-space:nowrap;padding-left:5px;">상품명: <input type="text" class="text" name="itemname" value="<%= itemname %>" size="32" maxlength="20"></td>
					<td style="white-space:nowrap;padding-left:5px;">상품코드:</td>
					<td style="white-space:nowrap;" rowspan="2"><textarea rows="3" cols="10" name="itemid" id="itemid"><%=replace(itemid,",",chr(10))%></textarea> </td>
				</tr>
			  <tr>
			  	<td style="white-space:nowrap;"> <!-- #include virtual="/common/module/categoryselectbox.asp"--></td>
			    <td style="white-space:nowrap;padding-left:5px;" colspan="2">전시카테고리 : <!-- #include virtual="/common/module/dispCateSelectBox.asp"--></td>
			  </tr>
	 		<tr>
	 			<td colspan="4">검색키워드 : <input type="text" class="text" name="keyword" value="<%=keyword%>" size="40"><font color="gray" size="2">(주의:느릴수있습니다.)</font></td>
	 		</tr>
	 	</table>
		</td>
		<td rowspan="2" width="30" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button_s" value="검색" onClick="javascript:jsSerach();">
		</td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("topbar") %>" >
		<td align="left">
			판매:<% drawSelectBoxSellYN "sellyn", sellyn %>

	     	사용:<% drawSelectBoxUsingYN "usingyn", usingyn %>

	     	단종:<% drawSelectBoxDanjongYN "danjongyn", danjongyn %>

	     	한정:<% drawSelectBoxLimitYN "limityn", limityn %>

	     	계약:<% drawSelectBoxMWU "mwdiv", mwdiv %><br>

	     	할인: <% drawSelectBoxSailYN "sailyn", sailyn %>

	     	쿠폰: <% drawSelectBoxCouponYN "couponyn", couponyn %>

	     	배송:<% drawBeadalDiv "deliverytype",deliverytype %>

			베스트: <% drawSelectBoxIsBestSorting "sortDiv", sortDiv%>

			<input type="checkbox" name="notdeal" value="Y"<% if notdeal="Y" then response.write " checked" %>> 딜상품 제외
		</td>
	</tr>
</table>
</form>
<form name="frmItem" method="post">
	<input type="hidden" name="page" >
	<input type="hidden" name="sType" >
	<input type="hidden" name="itemidarr" >
	<input type="hidden" name="itemcount" value="0">
	<input type="hidden" name="mode" value="I">
	<input type="hidden" name="acURL" value="<%=actionURL%>">
	<input type="hidden" name="defaultmargin" value="<%=defaultmargin%>">
	<input type="hidden" name="sC" value="<%=sCode%>">
	<input type="hidden" name="ptype" value="<%=ptype%>">
<table width="100%" height="40" align="center" cellpadding="3" cellspacing="1" class="a" border="0">
	<tr>
		<td  valign="bottom">
				<input type="button" value="선택상품 추가" onClick="SelectItems('sel')" class="button">
		</td>
	</tr>
</table>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>" valign="top" border="0">
<tr  bgcolor="#FFFFFF">
	<td colspan="13">
	검색결과 : <b><%= oitem.FTotalCount%></b>
	&nbsp;
	페이지 : <b><%= page %> /<%=  oitem.FTotalpage %></b>
	</td>
</tr>

<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td align="center"><input type="checkbox" name="chkAll" onClick="jsChkAll();"></td>
	<td align="center">상품ID</td>
	<td align="center">이미지</td>
	<td align="center">브랜드</td>
	<td align="center">상품명</td>
	<td align="center">판매가</td>
	<td align="center">매입가</td>
	<td align="center" nowrap>배송<br>구분</td>
	<td align="center" nowrap>계약<br>구분</td>
	<td align="center" nowrap>판매<br>여부</td>
	<td align="center" nowrap>사용<br>여부</td>
	<td align="center" nowrap>한정<br>여부</td>
	<td align="center" nowrap>재고<br>현황</td>
</tr>
<% if oitem.FresultCount<1 then %>
    <tr bgcolor="#FFFFFF" >
    	<td colspan="13" align="center">[검색결과가 없습니다.]</td>
    </tr>
<% end if %>
<% if oitem.FresultCount > 0 then %>
    <% for i=0 to oitem.FresultCount-1 %>
	<tr class="a" height="25" bgcolor="#FFFFFF">
	<td  align="center"><input type="checkbox" name="chkitem" value="<%= oitem.FItemList(i).FItemId %>"></td>
	<td align="center"><A href="http://www.10x10.co.kr/shopping/category_prd.asp?itemid=<%= oitem.FItemList(i).FItemId %>" target="_blank"><%= oitem.FItemList(i).FItemId %></a></td>
	<td align="center"><%IF oitem.FItemList(i).FSmallImage <> "" THEN%><img src="<%= oitem.FItemList(i).FSmallImage %>" width="50" height="50" border=0 alt=""><%END IF%></td>
		<td align="center"><% =oitem.FItemList(i).Fmakerid %></td>
	<td>&nbsp;<% =oitem.FItemList(i).Fitemname %></td>
	<td align="center"><%
			Response.Write FormatNumber(oitem.FItemList(i).Forgprice,0)
			'할인가
			if oitem.FItemList(i).Fsailyn="Y" then
				Response.Write "<br><font color=#F08050>(할)" & FormatNumber(oitem.FItemList(i).Fsailprice,0) & "</font>"
			end if
			'쿠폰가
			if oitem.FItemList(i).FitemCouponYn="Y" then
				Select Case oitem.FItemList(i).FitemCouponType
					Case "1"
						Response.Write "<br><font color=#5080F0>(쿠)" & FormatNumber(oitem.FItemList(i).Forgprice*((100-oitem.FItemList(i).FitemCouponValue)/100),0) & "</font>"
					Case "2"
						Response.Write "<br><font color=#5080F0>(쿠)" & FormatNumber(oitem.FItemList(i).Forgprice-oitem.FItemList(i).FitemCouponValue,0) & "</font>"
				end Select
			end if
		%></td>
	<td align="center"><%
			Response.Write FormatNumber(oitem.FItemList(i).Forgsuplycash,0)
			'할인가
			if oitem.FItemList(i).Fsailyn="Y" then
				Response.Write "<br><font color=#F08050>" & FormatNumber(oitem.FItemList(i).Fsailsuplycash,0) & "</font>"
			end if
			'쿠폰가
			if oitem.FItemList(i).FitemCouponYn="Y" then
				if oitem.FItemList(i).FitemCouponType="1" or oitem.FItemList(i).FitemCouponType="2" then
					if oitem.FItemList(i).Fcouponbuyprice=0 or isNull(oitem.FItemList(i).Fcouponbuyprice) then
						Response.Write "<br><font color=#5080F0>" & FormatNumber(oitem.FItemList(i).Forgsuplycash,0) & "</font>"
					else
						Response.Write "<br><font color=#5080F0>" & FormatNumber(oitem.FItemList(i).Fcouponbuyprice,0) & "</font>"
					end if
				end if
			end if
		%></td>
	<td align="center"><%=fnColor(oitem.FItemList(i).IsUpcheBeasong(),"delivery")%></td>
	<td align="center"><%= fnColor(oitem.FItemList(i).Fmwdiv,"mw") %></td>
	<td align="center">
	<%= fnColor(oitem.FItemList(i).Fsellyn,"yn") %>
	</td>
	<td align="center">
	<%= fnColor(oitem.FItemList(i).Fisusing,"yn") %>
	</td>
	<td align="center"><%= fnColor(oitem.FItemList(i).Flimityn,"yn") %></td>
	<td align="center" nowrap>
	<a href="javascript:PopItemStock('<%= oitem.FItemList(i).FItemId %>')" title="재고현황 팝업">[보기]</a><br>
	<%IF oitem.FItemList(i).IsSoldOut() THEN%>
		<img src="http://webadmin.10x10.co.kr/images/soldout_s.gif" width="30" height="12">
<%END IF%>
	</td>
</tr>
<% next %>
<tr>
	<td colspan="13" align="center" bgcolor="#FFFFFF">
	<!-- 페이징처리 -->
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
<div style="padding:5px;text-align:right;font-size:8pt">Ver1.0  lastupdate: 2013.12.24 </div>
<iframe name="FrameCKP" src="about:blank" frameborder="0" width="600" height="200"></iframe>
<%
 set oitem = nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->