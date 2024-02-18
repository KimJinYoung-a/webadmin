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
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
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
'if sailyn="" then sailyn="N"			'할인페이지에서 검색된거라면 기본값: 할인안함(쿠폰도 동일)
'if couponyn="" then couponyn="N"
'if sellyn = "" then sellyn ="Y"
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

'if reload="" and itemexists="" then itemexists="N"
itemexists="Y"

set oitem = new CItem
	oitem.FPageSize         = 50
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

	function regitem(mode){
		var regitem = window.open('/admin/itemmaster/Item_Evaluate_exclude_pop.asp?mode='+mode,'regitem','width=1024,height=768,scrollbars=yes,resizable=yes');
		regitem.focus();
	}

	// 페이지 이동
	function goPage(pg){
		document.frm.page.value=pg;
		document.frm.submit();
	}

	// 선택된 항목 삭제/복구
	function doedit(mode){
		var i, chk=0;
		var frm = document.frm_list;

		if (frm.Eval_excludeitemid.length){
			for(i=0;i<frm.Eval_excludeitemid.length;i++){
				if(frm.Eval_excludeitemid[i].checked){
					chk++;
				}
			}
		} else {
			if(frm.Eval_excludeitemid.checked){
				chk++;
			}
		}

		if(chk==0){
			alert("상품을 적어도 한개이상 선택해주십시요.");
			return;
		} else {
			if(confirm("선택하신 " + chk + "개의  항목을 모두 삭제 하시겠습니까?")){
				frm.mode.value=mode;
				frm.action="/admin/itemmaster/Item_Evaluate_exclude_process.asp";
				frm.submit();
			} else {
				return;
			}
		}
	}

	//전체 선택
	function jsChkAll(){	
	var frm;
	frm = document.frm_list;
		if (frm.chkAll.checked){			      
		   if(typeof(frm.Eval_excludeitemid) !="undefined"){
		   	   if(!frm.Eval_excludeitemid.length){
			   	 	frm.Eval_excludeitemid.checked = true;	   	 
			   }else{
					for(i=0;i<frm.Eval_excludeitemid.length;i++){
						frm.Eval_excludeitemid[i].checked = true;
				 	}		
			   }	
		   }	
		} else {	  
		  if(typeof(frm.Eval_excludeitemid) !="undefined"){
		  	if(!frm.Eval_excludeitemid.length){
		   	 	frm.Eval_excludeitemid.checked = false;	  
		   	}else{
				for(i=0;i<frm.Eval_excludeitemid.length;i++){
					frm.Eval_excludeitemid[i].checked = false;
				}	
			}		
		  }	
		}
	}

</script>

<!-- 상단 검색폼 시작 -->
<table width="100%" align="center" celiadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get" action="">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="page" value="">
<input type="hidden" name="reload" value="ON">
<tr align="center" bgcolor="#FFFFFF" >
	<td width="80" bgcolor="<%= adminColor("gray") %>">검색조건</td>
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
		<p>
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
		<!--&nbsp;&nbsp;
		* 상품등록여부 : -->
		<%' drawSelectBoxisusingYN "itemexists", itemexists,"" %>		
	</td>
	<td width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="submit" class="button_s" value="검색">
	</td>
</tr>
</form>
</table>
<!-- 검색 끝 -->

<Br>

<!-- 액션 시작 -->
<table width="100%" align="center" celiadding="0" cellspacing="0" class="a" style="padding:10 0 0 0;">
<form name="frm_list" method="get" action="">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="mode" value="">
<input type="hidden" name="page" value="<%=page%>">
<tr>
	<td align="left">		
		[ON]상품관리>>상품수정 에서 <font color="red">품목(의료기기,식품(농수산물),가공식품,건강기능식품/체중조절식품)</font>에 해당되는 상품은, 하루에 한번 새벽에 이곳에 자동 저장됩니다.
	</td>
	<td align="right">
		<% if oitem.FResultCount>0 then %>
			<input type="button" value="선택삭제" onClick="doedit('delitem')" class="button">
			&nbsp;
		<% end if %>
			
		<input type="button" value="신규등록" onClick="regitem('regitem')" class="button">
	</td>
</tr>
</table>
<!-- 액션 끝 -->

<table width="100%" align="center" celiadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="20">
		검색결과 : <b><%=FormatNumber(oitem.FTotalCount,0)%></b>
		&nbsp;
		페이지 : <b><%= page %>/<%=FormatNumber(oitem.Ftotalpage,0)%></b>
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
<% if oitem.FResultCount=0 then %>
	<tr align="center">
		<td colspan="20" height="30" bgcolor="#FFFFFF">등록(검색)된 내역이 없습니다.</td>
	</tr>
<%
else

for i=0 to oitem.FResultCount - 1
%>
<tr align="center" bgcolor="#FFFFFF">
	<td  align="center"><input type="checkbox" name="Eval_excludeitemid" value="<%= oitem.FItemList(i).fEval_excludeitemid %>"></td>
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
<tr height="25" bgcolor="FFFFFF">
	<td colspan="20" align="center" bgcolor="#FFFFFF">
		<!-- 페이징처리 -->
		<% if oitem.HasPreScroll then %>
			<a href="javascript:goPage('<%= oitem.StartScrollPage-1 %>')">[pre]</a>
		<% else %>
			[pre]
		<% end if %>

		<% for i=0 + oitem.StartScrollPage to oitem.FScrollCount + oitem.StartScrollPage - 1 %>
			<% if i>oitem.FTotalpage then Exit for %>
			<% if CStr(page)=CStr(i) then %>
			<font color="red">[<%= i %>]</font>
			<% else %>
			<a href="javascript:goPage('<%= i %>')">[<%= i %>]</a>
			<% end if %>
		<% next %>

		<% if oitem.HasNextScroll then %>
			<a href="javascript:goPage('<%= i %>')">[next]</a>
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