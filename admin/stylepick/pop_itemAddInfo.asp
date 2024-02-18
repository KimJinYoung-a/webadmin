<%@ language=vbscript %>
<% option explicit %>
<%
	Response.AddHeader "Cache-Control","no-cache"
	Response.AddHeader "Expires","0"
	Response.AddHeader "Pragma","no-cache"
%>
<%
'###########################################################
' Description : 스타일픽 관리
' Hieditor : 2011.04.06 한용민 생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/items/itemcls_2008.asp"-->
<!-- #include virtual="/lib/classes/stylepick/stylepick_cls.asp"-->
<%
dim target ,page ,cdl, cdm, cds ,i , cd1,cd2,cd3
dim itemid, itemname, makerid, sellyn, usingyn, danjongyn, deliverytype, limityn, vatyn, sailyn, couponyn, mwdiv,defaultmargin
	cd1 = request("cd1")
	cd2 = request("cd2")
	cd3 = request("cd3")	
	itemid      = request("itemid")
	itemname    = request("itemname")
	makerid     = request("makerid")
	sellyn      = request("sellyn")
	usingyn     = request("usingyn")
	danjongyn   = request("danjongyn") 
	mwdiv       = request("mwdiv")
	limityn     = request("limityn") 
	sailyn      = request("sailyn")
	couponyn	= request("couponyn")
	defaultmargin = request("defaultmargin")
	deliverytype       = request("deliverytype")
	cdl = request("cdl")
	cdm = request("cdm")
	cds = request("cds")
	page = request("page")

if (page="") then page=1

if itemid<>"" then
	dim iA ,arrTemp,arrItemid

	arrTemp = Split(itemid,",")

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

	itemid = left(arrItemid,len(arrItemid)-1)
end if

dim oitem
set oitem = new CItem
oitem.FPageSize         = 50
oitem.FCurrPage         = page
oitem.FRectMakerid      = makerid
oitem.FRectItemid       = itemid
oitem.FRectItemName     = itemname
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
oitem.GetItemList	
%>

<script language="javascript">

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
	frm = document.frm;
	frm.sType.value = sType;   //전체선택 or 선택상품 여부 구분
			
	if (frm.cd1.value==''){			
		alert('상품을 등록 하실 스타일 카테고리를 선택하세요');
		frm.cd1.focus();
		return;
	}

	if (frm.cd2.value==''){			
		alert('상품을 등록 하실 분류 카테고리를 선택하세요');
		frm.cd2.focus();
		return;
	}
	
	//if (frm.cd3.value==''){			
	//	alert('상품을 등록 하실 카테고리를 선택하세요');
	//	frm.cd3.focus();
	//	return;
	//}

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
		  if(confirm(itemcount +"건의 검색된 모든 상품을 추가하시겠습니까?")){
		  	if(itemcount > 1000) {
		  		alert("상품은 최대 1000건까지 가능합니다. 조건을 다시 설정해주세요");
		  		return;
		  	}
			frm.itemidarr.value = frm.itemid.value;
			
		  }else{
		  	return;
		  }
		}else{
		 	alert("추가할 상품이 없습니다.");
	   	  	return;
		}	
	}
	
	frm.target = "FrameCKP";
	frm.action = "/admin/stylepick/stylepick_item_process.asp";
	frm.itemcount.value = itemcount;
	frm.submit();
}

//전체 선택
function jsChkAll(){	
var frm;
frm = document.frm;
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

</script>

<!-- 검색 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="post">	
<input type="hidden" name="page" >
<input type="hidden" name="sType" >
<input type="hidden" name="itemidarr" >
<input type="hidden" name="itemcount" value="0">
<input type="hidden" name="mode" value="itemadd">	
<input type="hidden" name="defaultmargin" value="<%=defaultmargin%>">
<tr align="center" bgcolor="<%= adminColor("topbar") %>" >
	<td rowspan="2" width="30" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
	<td align="left">
		브랜드 :<%	drawSelectBoxDesignerWithName "makerid", makerid %>
		&nbsp;
		<!-- #include virtual="/common/module/categoryselectbox.asp"-->
		<br>
		상품코드 :
		<input type="text" class="text" name="itemid" value="<%= itemid %>" size="40" maxlength="100" onKeyPress="if (event.keyCode == 13) document.frm.submit();">			
		&nbsp;
		상품명 :
		<input type="text" class="text" name="itemname" value="<%= itemname %>" size="32" maxlength="20">			
		<div style="font-size:11px; color:gray;padding-left:60px;">(쉼표로 복수입력가능)</div>
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
     	 
     	계약:<% drawSelectBoxMWU "mwdiv", mwdiv %>
     	
     	할인: <% drawSelectBoxSailYN "sailyn", sailyn %>

     	쿠폰: <% drawSelectBoxCouponYN "couponyn", couponyn %>
     	
     	<br>배송:<% drawBeadalDiv "deliverytype",deliverytype %>
	</td>
</tr>    
</table>

<table width="100%" height="40" align="center" cellpadding="3" cellspacing="1" class="a" border="0">	
<tr>
	<td  valign="bottom">				
		스타일:<% Drawcategory "cd1",cd1," onchange='jsSerach();'","CD1" %>
		<% if cd1 <> "" then %>
			분류:<% Drawcategory "cd2",cd2," onchange='jsSerach();'","CD2" %>
		<% end if %>
		<% if cd1 <> "" and cd2 <> "" then %>
			<%' Drawcategory "cd3",cd3," onchange='jsSerach();'","CD3" %>
		<% end if %>
		<input type="button" value="선택상품 추가" onClick="SelectItems('sel')" class="button">
		<!-- saleItemProc.asp 전체선택 추가 오류 있는듯함. -->
		<input type="button" value="전체선택 추가" onClick="SelectItems('all')" class="button" disabled >		
	</td>				
</tr>
</table> 
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>" valign="top" border="0">
<tr  bgcolor="#FFFFFF">
	<td colspan="20">
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
	<td align="center">
		<%
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
		%>
	</td>
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
		<%IF oitem.FItemLiwebadmin.10x10.co.kr() THEN%>
			<img src="http://scm.10x10.co.kr/images/soldout_s.gif" width="30" height="12">
		<%END IF%>
	</td>
</tr>
<% next %>
<tr>
	<td colspan="20" align="center" bgcolor="#FFFFFF">
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
<% else %>
<tr bgcolor="#FFFFFF" >
	<td colspan="20" align="center">[검색결과가 없습니다.]</td>
</tr>
<% end if %>
<tr bgcolor="#FFFFFF" >
	<td colspan="20" align="left">		
		<input type="button" value="선택상품 추가" onClick="SelectItems('sel')" class="button">
	</td>	
</tr>
</form>
</table>
<iframe name="FrameCKP" src="about:blank" frameborder="0" width="600" height="200"></iframe>
<%
	set oitem = nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
