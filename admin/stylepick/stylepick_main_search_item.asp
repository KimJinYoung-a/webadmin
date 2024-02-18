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
' Hieditor : 2011.04.07 한용민 생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/stylepick/stylepick_cls.asp"-->

<%
Dim cd1, cd2 ,cd3,i,page,isusing ,oitem,deliverytype,sailyn,couponyn ,num
dim makerid,itemid , itemname,sellyn,danjongyn,limityn,mwdiv ,defaultmargin , SortMet
	num = request("num")
	SortMet = request("SortMet")
	cd1 = request("cd1")
	cd2 = request("cd2")
	cd3 = request("cd3")	
	isusing = request("isusing")	
	itemid      = request("itemid")
	itemname    = request("itemname")
	makerid     = request("makerid")
	sellyn      = request("sellyn")
	danjongyn   = request("danjongyn") 
	mwdiv       = request("mwdiv")
	limityn     = request("limityn") 
	sailyn      = request("sailyn")
	couponyn	= request("couponyn")
	defaultmargin = request("defaultmargin")
	deliverytype       = request("deliverytype")	
	menupos = request("menupos")
	page = request("page")
	if page = "" then page = 1
	if isusing = "" then isusing = "Y"
		
'//상품 리스트
set oitem = new cstylepick
	oitem.FPageSize = 50
	oitem.FCurrPage = page
	oitem.FRectSortDiv      = SortMet
	oitem.FRectMakerid      = makerid
	oitem.FRectItemid       = itemid
	oitem.FRectItemName     = itemname
	oitem.FRectSellYN       = sellyn
	oitem.FRectDanjongyn    = danjongyn
	oitem.FRectLimityn      = limityn
	oitem.FRectMWDiv        = mwdiv
	oitem.FRectDeliveryType = deliverytype
	oitem.FRectSailYn       = sailyn
	oitem.FRectCouponYn		= couponyn	
	oitem.frectcd1 = cd1
	oitem.frectcd2 = cd2
	oitem.frectcd3 = cd3
	oitem.frectisusing = isusing
	oitem.GetItemList()
%>

<script language="javascript">

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

function jsSerach(){
	var frm;
	frm = document.frm;
	frm.target = "_self";
	frm.action ="stylepick_main_search_item.asp";
	frm.submit();
}

// 페이지 이동
function NextPage(ipage){
	document.frm.page.value= ipage;
	document.frm.target = "_self";
	document.frm.action ="stylepick_main_search_item.asp";
	document.frm.submit();
}

//새상품추가
function choiceitem(itemid){
	opener.eval('document.all.divsub'+<%=num%>).innerHTML = "기획전코드 & 상품코드 : <input type='text' name='gubunvalue' value='"+itemid+"' size=10 maxlength=10>";
	self.close();
}

</script>

<!-- 검색 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="post">
<input type="hidden" name="num" value="<%= num %>">
<input type="hidden" name="page" >
<input type="hidden" name="sType" >
<input type="hidden" name="itemidxarr">
<input type="hidden" name="itemcount" value="0">
<input type="hidden" name="mode">	
<input type="hidden" name="defaultmargin" value="<%=defaultmargin%>">
<input type="hidden" name="menupos" value="<%= menupos %>">
<tr align="center" bgcolor="<%= adminColor("topbar") %>" >
	<td rowspan="2" width="30" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
	<td align="left">				
		<input type="hidden" name="cd1" value="<%= cd1 %>">
		<input type="hidden" name="cd3" value="<%= cd3 %>">
		분류:<% Drawcategory "cd2",cd2," onchange='jsSerach();'","CD2" %>	
		삭제:<% drawSelectBoxUsingYN "isusing", isusing %>
	</td>	
	<td rowspan="2" width="30" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="검색" onClick="javascript:jsSerach();">
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("topbar") %>" >
	<td align="left">
		판매:<% drawSelectBoxSellYN "sellyn", sellyn %>     	      	
     	단종:<% drawSelectBoxDanjongYN "danjongyn", danjongyn %>     	 
     	한정:<% drawSelectBoxLimitYN "limityn", limityn %>     	 
     	계약:<% drawSelectBoxMWU "mwdiv", mwdiv %>     	
     	할인:<% drawSelectBoxSailYN "sailyn", sailyn %>
     	쿠폰:<% drawSelectBoxCouponYN "couponyn", couponyn %>     	
     	배송:<% drawBeadalDiv "deliverytype",deliverytype %>
		<br>브랜드 :<%	drawSelectBoxDesignerWithName "makerid", makerid %>
		상품코드 :
		<input type="text" class="text" name="itemid" value="<%= itemid %>" size="40" maxlength="100" onKeyPress="if (event.keyCode == 13) document.frm.submit();">
		(쉼표로 복수입력가능)
		<br>상품명 :
		<input type="text" class="text" name="itemname" value="<%= itemname %>" size="32" maxlength="20">			
	</td>
</tr>    
</table>
<br>	
<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
<tr>
	<td align="left">
	</td>
	<td align="right">	
	</td>
</tr>
</table>
<!-- 액션 끝 -->

<!-- 리스트 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>" valign="top" border="0">
<tr bgcolor="#FFFFFF">
	<td colspan="20">
		<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a">
		<tr>
			<td align="left">
				검색결과 : <b><%= oitem.FTotalCount%></b>
				&nbsp;
				페이지 : <b><%= page %> /<%=  oitem.FTotalpage %></b>				
			</td>
			<td align="right">
				정렬:<% Drawsort "SortMet" ,SortMet ," onchange='jsSerach();'" %>				
			</td>			
		</tr>
		</table>
	</td>
	
</tr>
		
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td><input type="checkbox" name="chkAll" onClick="jsChkAll();"></td>
	<td>스타일</td>
	<td>분류</td>
	<td>상품ID</td>
	<td>이미지</td>
	<td>브랜드</td>
	<td>상품명</td>
	<td>판매가</td>
	<td>매입가</td>
	<td nowrap>배송<br>구분</td>	
	<td nowrap>계약<br>구분</td>
	<td nowrap>판매<br>여부</td>	
	<td nowrap>사용<br>여부</td>	
	<td nowrap>한정<br>여부</td>	
	<td nowrap>재고<br>현황</td>
	<td>비고</td>
</tr>
<% if oitem.FresultCount > 0 then %>
<% for i=0 to oitem.FresultCount-1 %>
<% if oitem.FItemList(i).fisusing = "Y" then %>
<tr align="center" bgcolor="#FFFFFF" onmouseover=this.style.background="orange"; onmouseout=this.style.background='ffffff';>
<% else %>    
<tr align="center" bgcolor="#FFFFaa" onmouseover=this.style.background="orange"; onmouseout=this.style.background='FFFFaa';>
<% end if %>
	<td align="center">
		<input type="checkbox" name="chkitem" value="<%= oitem.FItemList(i).FItemidx %>">
	</td>
	<td align="center">
		<%= oitem.FItemList(i).fcd1name %> (<%= oitem.FItemList(i).fcd1 %>)
	</td>
	<td align="center">
		<%= oitem.FItemList(i).fcd2name %> (<%= oitem.FItemList(i).fcd2 %>)
	</td>
	<td align="center"><A href="http://www.10x10.co.kr/shopping/category_prd.asp?itemid=<%= oitem.FItemList(i).FItemId %>" target="_blank"><%= oitem.FItemList(i).FItemId %></a></td>
	<td align="center"><%IF oitem.FItemList(i).FSmallImage <> "" THEN%><img src="<%= oitem.FItemList(i).FSmallImage %>" width="50" height="50" border=0 alt=""><%END IF%></td>
		<td align="center"><% =oitem.FItemList(i).Fmakerid %></td>
	<td>&nbsp;<% =oitem.FItemList(i).Fitemname %></td>
	<td align="center">
		<%
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
		%>
	</td>
	<td align="center"><%
			Response.Write FormatNumber(oitem.FItemList(i).Forgsuplycash,0)
			'할인가
			if oitem.FItemList(i).Fsailyn="Y" then
				Response.Write "<br><font color=#F08050>" & FormatNumber(oitem.FItemList(i).Fsailsuplycash,0) & "</font>"
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
		<a href="javascripwebadmin.10x10.co.kr'<%= oitem.FItemList(i).FItemId %>')" title="재고현황 팝업">[보기]</a><br>
		<%IF oitem.FItemList(i).IsSoldOut() THEN%>
			<img src="http://scm.10x10.co.kr/images/soldout_s.gif" width="30" height="12">
		<%END IF%>
	</td>
	<td align="center">
		<input type="button" class="button" value="선택" onclick="choiceitem('<%= oitem.FItemList(i).fitemid %>');">		
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
<tr bgcolor="#FFFFFF">
	<td colspan="20" align="center">[검색결과가 없습니다.]</td>
</tr>
<% end if %>
</form>
</table>

<% set oitem = nothing %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->

