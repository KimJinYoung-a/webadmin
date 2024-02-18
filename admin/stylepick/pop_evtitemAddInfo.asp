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
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/stylepick/stylepick_cls.asp"-->

<%
Dim cd1,i,page,isusing ,oitem,deliverytype,sailyn,couponyn,menupos ,evtidx ,CD2
dim makerid,itemid , itemname,sellyn,danjongyn,limityn,mwdiv ,defaultmargin , SortMet
dim cdl ,cdm ,cds , overlap
	overlap = request("overlap")
	cdl = request("cdl")
	cdm = request("cdm")
	cds = request("cds")
	evtidx      = request("evtidx")
	cd1 = request("cd1")
	cd2 = request("cd2")
	SortMet = request("SortMet")			
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
	isusing = "Y"
	if overlap = "" then overlap = "notoverlap"
		
'//상품 리스트
set oitem = new cstylepick
	oitem.FPageSize = 50
	oitem.FCurrPage = page
	oitem.FRectCate_Large   = cdl
	oitem.FRectCate_Mid     = cdm
	oitem.FRectCate_Small   = cds
	oitem.FRectSortDiv      = SortMet
	oitem.FRectMakerid      = makerid
If itemid <> "" Then
	If IsNumeric(itemid) = "False" Then
		rw "<script>alert('상품코드는 숫자만 입력하세요');location.replace('/admin/stylepick/pop_evtitemAddInfo.asp');</script>"
		response.end
	End If
End If	
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
	oitem.frectisusing = isusing
	oitem.frectoverlap = overlap
	oitem.GeteventitemList()
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

function SelectItemsadd(){	
	var frm;
	var itemcount = 0;
	frm = document.frm;

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
						frm.itemidarr.value = frm.chkitem[i].value;				
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
	
	frm.action = "/admin/stylepick/stylepick_event_process.asp";
	frm.mode.value = "evtitemadd";
	frm.target="view";
	frm.submit();
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
	frm.action ="pop_evtitemAddInfo.asp";
	frm.submit();
}

// 페이지 이동
function NextPage(ipage){
	document.frm.page.value= ipage;
	document.frm.target = "_self";
	document.frm.action ="pop_evtitemAddInfo.asp";
	document.frm.submit();
}

</script>

<!-- 검색 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="post">
<input type="hidden" name="page">
<input type="hidden" name="sType">
<input type="hidden" name="itemidarr">
<input type="hidden" name="itemcount" value="0">
<input type="hidden" name="mode">
<input type="hidden" name="evtidx" value="<%=evtidx%>">
<input type="hidden" name="defaultmargin" value="<%=defaultmargin%>">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="cd1" value="<%= cd1 %>">
<tr align="center" bgcolor="<%= adminColor("topbar") %>" >
	<td rowspan="2" width="30" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
	<td align="left">
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
		<!-- #include virtual="/common/module/categoryselectbox.asp"-->
		<br>
		분류:<% Drawcategory "cd2",cd2," onchange='jsSerach();'","CD2" %>
		<input type="radio" name="overlap" value="all" <% if overlap="all" then response.write " checked"%>>전상품
		<input type="radio" name="overlap" value="notoverlap" <% if overlap="notoverlap" then response.write " checked"%>>동일스타일에오픈이전인상품제외
	</td>
</tr>    
</table>
	
<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
<tr>
	<td align="left">		
		<input type="button" value="선택상품추가" onClick="SelectItemsadd()" class="button">
		<font color="red">※ "[ON]StylePick>>StylePick 상품관리" 에서 먼저 해당 카테고리에 상품을 넣으셔야 상품이 보입니다.</font>
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
</tr>
<% if oitem.FresultCount > 0 then %>
<% for i=0 to oitem.FresultCount-1 %>
<% if oitem.FItemList(i).fisusing = "Y" then %>
<tr align="center" bgcolor="#FFFFFF" onmouseover=this.style.background="orange"; onmouseout=this.style.background='ffffff';>
<% else %>    
<tr align="center" bgcolor="#FFFFaa" onmouseover=this.style.background="orange"; onmouseout=this.style.background='FFFFaa';>
<% end if %>
	<td align="center">
		<input type="checkbox" name="chkitem" value="<%= oitem.FItemList(i).FItemid %>">
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
		<a href="javascript:PopItemStock('<%= oitem.FItemList(i).FItemId %>')" title="재고현황 팝업">[보기]</a><br>
		<%IF oitem.FItemList(i).IsSoldOut() THEN%>
			<img src="http://webadmin.10x10.co.kr/images/soldout_s.gif" width="30" height="12">
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
<tr bgcolor="#FFFFFF">
	<td colspan="20" align="center">[검색결과가 없습니다.]</td>
</tr>
<% end if %>
</form>
</table>
<iframe id="view" name="view" width=300 width=300 frameborder=0 scrolling="no"></iframe>
<% set oitem = nothing %>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->

