<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/items/itemcls_2008.asp"-->
<!-- #include virtual="/lib/classes/event/eventManageCls_V3.asp"-->
<%
dim target,gubun
dim eCode, cEGroup,arrGroup,intLoop,egcode
dim itemid, itemname, makerid, sellyn, usingyn, danjongyn, mwdiv, limityn, vatyn, sailyn,deliverytype, keyword,couponyn, itemdiv
dim cdl, cdm, cds , dispCate
dim page
Dim sortDiv
dim eChannel, itemnum

eCode 			= requestCheckvar(request("eC"),10)
egcode		= requestCheckvar(request("egcode"),10)
itemid      = requestCheckvar(request("itemid"),255)
itemname    = requestCheckvar(request("itemname"),64)
makerid     = requestCheckvar(request("makerid"),32)
sellyn      = requestCheckvar(request("sellyn"),2)
usingyn     = requestCheckvar(request("usingyn"),1)
danjongyn   = requestCheckvar(request("danjongyn"),2) 
limityn     = requestCheckvar(request("limityn"),2) 
sailyn      = requestCheckvar(request("sailyn"),1)  
deliverytype= requestCheckvar(request("deliverytype"),1)
sortDiv 		= requestCheckvar(request("sortDiv"),5)
keyword			= requestCheckvar(request("keyword"),512)
eChannel    = requestCheckvar(request("eCh"),1)   
couponyn		= requestCheckvar(request("couponyn"),1)
itemdiv		= requestCheckvar(request("itemdiv"),2)
cdl = requestCheckvar(request("cdl"),10)
cdm = requestCheckvar(request("cdm"),10)
cds = requestCheckvar(request("cds"),10)
dispCate = requestCheckvar(request("disp"),16)

page = requestCheckvar(request("page"),10)
itemnum 			= requestCheckvar(request("itemnum"),10)

if (page="") then page=1
'if sellyn = "" then sellyn ="Y"
if itemid<>"" then
	dim iA ,arrTemp,arrItemid
	itemid = replace(itemid,chr(13),"")
	arrTemp = Split(itemid,chr(10))

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

	itemid = left(arrItemid,len(arrItemid)-1)
end if
 
	if sortDiv="" then sortDiv="new"	'정렬방법 기본값

'--이벤트 그룹
 set cEGroup = new ClsEventGroup
	cEGroup.FECode = eCode 
	cEGroup.FEChannel = eChannel	
	arrGroup = cEGroup.fnGetEventItemGroup		
 set cEGroup = nothing


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
oitem.FRectSailYn       = sailyn
oitem.FRectDeliveryType = deliverytype

oitem.FRectCate_Large   = cdl
oitem.FRectCate_Mid     = cdm
oitem.FRectCate_Small   = cds
oitem.FRectDispCate		= dispCate
oitem.FRectSortDiv = SortDiv
oitem.FRectCouponYn		= couponyn
oitem.FRectDealYn		= "N"
oitem.FRectItemDiv		= itemdiv
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
	frm.action ="pop_additemlist.asp";
	frm.submit();
}

function SelectItems(itemid,itemname,itemimg){
	var iteminfo =itemname + " <img src='" + itemimg + "' width='50' height='50' border=0 alt=''>";
	opener.document.frmcontents.itemid<%=itemnum%>.value=itemid;
	$("#iteminfo<%=itemnum%>",opener.document).html(iteminfo);
	window.close();
}

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
	document.frm.action ="pop_additemlist.asp";
	document.frm.submit();
}

//-->
</script>
<!-- 검색 시작 -->
<form name="frm" method=get>	
	<input type="hidden" name="eC" value="<%=eCode%>">
	<input type="hidden" name="page" >
	<input type="hidden" name="sType" >
	<input type="hidden" name="itemidarr" >
	<input type="hidden" name="mode" value="I">
	<input type="hidden" name="eCh" value="<%=eChannel%>">
	<input type="hidden" name="egcode" value="<%=egcode%>">
	<input type="hidden" name="itemnum" value="<%=itemnum%>">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" border="0" > 	
<tr>
    <td>    
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>"> 
	<tr align="center" bgcolor="<%= adminColor("topbar") %>" >
		<td rowspan="2" width="60" bgcolor="<%= adminColor("gray") %>">검색 조건</td>
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
	     		     	    	    
	     	할인 <% drawSelectBoxSailYN "sailyn", sailyn %>
	     	
	     	쿠폰: <% drawSelectBoxCouponYN "couponyn", couponyn %>
	     	
			배송:<% drawBeadalDiv "deliverytype",deliverytype %>

			상품구분:<% drawSelectBoxItemDivDeal "itemdiv",itemdiv %>
		</td>
	</tr>    
</table>
</td>
</tr>
<tr>
    <td>
<table width="100%" align="center"    cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr  bgcolor="#FFFFFF">
	<td colspan="9">
	검색결과 : <b><%= oitem.FTotalCount%></b>
	&nbsp;
	페이지 : <b><%= page %> /<%=  oitem.FTotalpage %></b>
	</td>
	<td colspan="3">
		<select name="sortDiv" onchange="this.form.submit();" class="select">
		<option value="new" <% IF sortDiv="new" Then response.write "selected" %> >신상품순</option>
		<option value="cashH" <% IF sortDiv="cashH" Then response.write "selected" %>>높은가격순</option>
		<option value="cashL" <% IF sortDiv="cashL" Then response.write "selected" %>>낮은가격순</option>
		<option value="best" <% IF sortDiv="best" Then response.write "selected" %>>베스트순</option>
		</select>
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
	<td align="center">배송</td>	
	<td align="center">판매여부</td>	
	<td align="center">사용여부</td>	
	<td align="center">한정여부</td>	
	<td align="center">재고현황</td>
</tr>
<% if oitem.FresultCount<1 then %>
    <tr bgcolor="#FFFFFF">
    	<td colspan="12" align="center">[검색결과가 없습니다.]</td>
    </tr>
<% end if %>
<% if oitem.FresultCount > 0 then %>
    <% for i=0 to oitem.FresultCount-1 %>
	<tr class="a" height="25" bgcolor="#FFFFFF"  onClick="SelectItems(<%= oitem.FItemList(i).FItemId %>,'<% =oitem.FItemList(i).Fitemname %>','<%= oitem.FItemList(i).FSmallImage %>')">
	<td  align="center"><input type="checkbox" name="chkitem" value="<%= oitem.FItemList(i).FItemId %>"></td>
	<td align="center"><A href="http://www.10x10.co.kr/shopping/category_prd.asp?itemid=<%= oitem.FItemList(i).FItemId %>" target="_blank"><%= oitem.FItemList(i).FItemId %></a></td>
	<td align="center"><%IF oitem.FItemList(i).FSmallImage <> "" THEN%><img src="<%= oitem.FItemList(i).FSmallImage %>" width="50" height="50" border=0 alt=""><%END IF%></td>
		<td align="center"><% =oitem.FItemList(i).Fmakerid %></td>
	<td><%If oitem.FItemList(i).Fitemdiv="21" Then%><font color="#0000ff">[딜]</font><%End If%>&nbsp;<% =oitem.FItemList(i).Fitemname %></td>
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
	<td align="center">
	<%= fnColor(oitem.FItemList(i).Fsellyn,"yn") %>
	</td>
	<td align="center">
	<%= fnColor(oitem.FItemList(i).Fisusing,"yn") %>
	</td>
	<td align="center"><%= fnColor(oitem.FItemList(i).Flimityn,"yn") %></td>
	<td align="center">
	<a href="javascript:PopItemStock('<%= oitem.FItemList(i).FItemId %>')" title="재고현황 팝업">[보기]</a><br>
	<%IF oitem.FItemList(i).IsSoldOut() THEN%>
		<img src="http://webadmin.10x10.co.kr/images/soldout_s.gif" width="30" height="12">
<%END IF%>
	</td>
</tr>
<% next %>
<tr>
	<td colspan="12" align="center" bgcolor="#FFFFFF">
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
</table>
</td>
</tr>
</table> 
</form>
<% end if %>
<iframe name="FrameCKP" src="" frameborder="0" width="600" height="400"></iframe>
<%
 set oitem = nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->