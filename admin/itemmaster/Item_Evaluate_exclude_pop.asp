<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  상품고시
' History : 2013.12.11 한용민 생성
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/items/itemcls_2008.asp"-->
<%
dim target, cdl, cdm, cds, page, oitem, i, sailyn, couponyn, mwdiv,defaultmargin, keyword
dim itemid, itemname, makerid, sellyn, usingyn, danjongyn, deliverytype, limityn, vatyn, mode
dim dispCate, itemexists, reload, infodivYn, infodiv
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
	keyword		= request("keyword")
	cdl = request("cdl")
	cdm = request("cdm")
	cds = request("cds")
	page = request("page")
	mode = request("mode")
	dispCate = requestCheckvar(request("disp"),16)
	itemexists = requestCheckvar(request("itemexists"),1)
	reload = request("reload")
	infodiv  = request("infodiv")
	infodivYn  = requestCheckvar(request("infodivYn"),10)

if mode="" then mode="regitem"
If infodiv <> "" Then
	infodivYn = "Y"	
End If
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

if reload="" and itemexists="" then itemexists="N"

set oitem = new CItem
	oitem.FPageSize         = 100
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
	oitem.FRectitemexists	= itemexists
	oitem.FRectInfodivYn    = infodivYn
	oitem.FRectInfodiv    = infodiv	
	oitem.GetItem_Evaluate_exclude
%>

<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script language="javascript">

function jsSerach(){
	var frm;
	frm = document.frm;
	frm.target = "_self";
	frm.action ="Item_Evaluate_exclude_pop.asp";
	frm.submit();
}

function SelectItems(){	
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
	
	//frm.target = "FrameCKP";
	frm.action = "/admin/itemmaster/Item_Evaluate_exclude_process.asp";
	frm.itemcount.value = itemcount;
	frm.submit();
	frm.itemidarr.value = "";
	frm.itemcount.value = 0;	
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

// 페이지 이동
function NextPage(ipage){
	document.frm.page.value= ipage;
	document.frm.target = "_self";
	document.frm.action ="Item_Evaluate_exclude_pop.asp";
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
<input type="hidden" name="reload" value="ON">
<input type="hidden" name="mode" value="<%= mode %>">
<input type="hidden" name="defaultmargin" value="<%=defaultmargin%>">
<tr align="center" bgcolor="<%= adminColor("topbar") %>" >
	<td rowspan="2" width="30" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
	<td align="left">
		* 브랜드 :
		<%	drawSelectBoxDesignerWithName "makerid", makerid %>
		&nbsp;&nbsp;
		* 상품코드 :
		<input type="text" class="text" name="itemid" value="<%= itemid %>" size="40" maxlength="100" onKeyPress="if (event.keyCode == 13) document.frm.submit();">(쉼표로 복수입력가능)
		<p>
		* 상품명 :
		<input type="text" class="text" name="itemname" value="<%= itemname %>" size="32" maxlength="20">	
		&nbsp;&nbsp;		
		* 검색키워드 : <input type="text" class="text" name="keyword" value="<%=keyword%>" size="40"><font color="gray" size="2">(주의:느릴수있습니다.)</font>
		<p>
		<span style="white-space:nowrap;">* 관리 <!-- #include virtual="/common/module/categoryselectbox.asp"--></span>
		<p>
		<span style="white-space:nowrap;">* 전시카테고리 : <!-- #include virtual="/common/module/dispCateSelectBox.asp"--></span>
     	<p>
     	<span style="white-space:nowrap;">* 품목정보입력여부 :
     	<select class="select" name="infodivYn">
	        <option value="">전체</option>
	        <option value="N" <%= CHKIIF(infodivYn="N","selected","") %> >입력이전</option>
	        <option value="Y" <%= CHKIIF(infodivYn="Y","selected","") %> >입력완료</option>
        </select></span>
        &nbsp;&nbsp;
		<span style="white-space:nowrap;">* 품목 : <% drawSelectBoxinfodiv "infodiv", infodiv, "" %></span>		
	</td>
	
	<td rowspan="2" width="30" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="검색" onClick="javascript:jsSerach();">
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("topbar") %>" >
	<td align="left">
		* 판매:<% drawSelectBoxSellYN "sellyn", sellyn %>
		&nbsp;&nbsp;
     	* 사용:<% drawSelectBoxUsingYN "usingyn", usingyn %>
		&nbsp;&nbsp;
     	* 단종:<% drawSelectBoxDanjongYN "danjongyn", danjongyn %>
		&nbsp;&nbsp;
     	* 한정:<% drawSelectBoxLimitYN "limityn", limityn %>
		&nbsp;&nbsp;
     	* 계약:<% drawSelectBoxMWU "mwdiv", mwdiv %>
		&nbsp;&nbsp;
     	* 할인: <% drawSelectBoxSailYN "sailyn", sailyn %>
		<p>
     	* 쿠폰: <% drawSelectBoxCouponYN "couponyn", couponyn %>
		&nbsp;&nbsp;
     	* 배송:<% drawBeadalDiv "deliverytype",deliverytype %>
		&nbsp;&nbsp;
		* 상품등록여부 : 
		<% drawSelectBoxisusingYN "itemexists", itemexists,"" %>
	</td>
</tr>
</table>

<table width="100%" height="40" align="center" cellpadding="3" cellspacing="1" class="a" border="0">	
<tr>
	<td  valign="bottom">				
		<input type="button" value="선택상품 추가" onClick="SelectItems()" class="button">
	</td>
</tr>
</table>
<iframe name="FrameCKP" src="about:blank" frameborder="0" width="0" height="0"></iframe>

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
	<td align="center" nowrap>배송<br>구분</td>	
	<td align="center" nowrap>계약<br>구분</td>
	<td align="center" nowrap>판매<br>여부</td>	
	<td align="center" nowrap>사용<br>여부</td>	
	<td align="center" nowrap>한정<br>여부</td>	
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
	<td align="center"><%=fnColor(oitem.FItemList(i).IsUpcheBeasong(),"delivery")%></td>
	<td align="center"><%= fnColor(oitem.FItemList(i).Fmwdiv,"mw") %></td>
	<td align="center">
	<%= fnColor(oitem.FItemList(i).Fsellyn,"yn") %>
	</td>
	<td align="center">
	<%= fnColor(oitem.FItemList(i).Fisusing,"yn") %>
	</td>
	<td align="center"><%= fnColor(oitem.FItemList(i).Flimityn,"yn") %></td>
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
</form>
</table>

<%
 set oitem = nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
