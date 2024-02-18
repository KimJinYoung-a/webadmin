<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' History : 2010.09.28 한용민 생성
' Description : 이벤트상품 추가
'				input - actionURL(db 처리에 필요한 파라미터까지 포함) ex.acURL = "/admin/eventmanage/event/eventitem_process.asp?eC=1234"
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/academy/lib/academy_function.asp"-->
<!-- #include virtual="/academy/lib/classes/Event_cls.asp"-->

<%
dim cEvtItem , eCode, egCode , sCode ,cdl, cdm, cds ,target, actionURL ,iPageSize, iCurrpage ,iDelCnt
Dim iTotCnt, arrList,intLoop ,iStartPage, iEndPage, iTotalPage, ix,iPerCnt
dim itemid, itemname, makerid, sellyn, usingyn, deliverytype, limityn, vatyn, sailyn, couponyn, mwdiv
	sCode		= RequestCheckvar(request("sC"),10)
	eCode		= RequestCheckvar(request("eC"),10)
	egCode      = RequestCheckvar(request("egC"),10)
	actionURL	= request("acURL")
	itemid      = RequestCheckvar(request("itemid"),10)
	itemname    = RequestCheckvar(request("itemname"),64)
	makerid     = RequestCheckvar(request("makerid"),32)
	sellyn      = RequestCheckvar(request("sellyn"),1)
	usingyn     = RequestCheckvar(request("usingyn"),1)	
	mwdiv       = RequestCheckvar(request("mwdiv"),10)
	limityn     = RequestCheckvar(request("limityn"),1)
	sailyn      = RequestCheckvar(request("sailyn"),1)
	couponyn	= RequestCheckvar(request("couponyn"),1)
	deliverytype       = RequestCheckvar(request("deliverytype"),10)
	cdl = RequestCheckvar(request("cdl"),10)
	cdm = RequestCheckvar(request("cdm"),10)
	cds = RequestCheckvar(request("cds"),10)
	iCurrpage = RequestCheckvar(Request("iC"),10)	'현재 페이지 번호
	
	IF iCurrpage = "" THEN	iCurrpage = 1			
	if sailyn="" and instr(actionURL,"saleitem")>0 then sailyn="N"			'할인페이지에서 검색된거라면 기본값: 할인안함(쿠폰도 동일)
	if couponyn="" and instr(actionURL,"saleitem")>0 then couponyn="N"
	iPageSize = 20		'한 페이지의 보여지는 열의 수
	iPerCnt = 10	
	'if sellyn = "" then sellyn ="Y"

if itemid<>"" then
	dim iA ,arrTemp,arrItemid

	arrTemp = Split(itemid,",")

	iA = 0
	do while iA <= ubound(arrTemp)

		if arrTemp(iA)<>"" then
			arrItemid = arrItemid & arrTemp(iA) & ","
		end if
		iA = iA + 1
	loop

	itemid = left(arrItemid,len(arrItemid)-1)
end if

'==============================================================================
set cEvtItem = new ClsEvent	
	cEvtItem.FCPage = iCurrpage
	cEvtItem.FPSize = iPageSize	
	cEvtItem.FECode = eCode			
	cEvtItem.FESGroup = egCode			
	
	cEvtItem.FRectMakerid      = makerid
	cEvtItem.FRectItemid       = itemid
	cEvtItem.FRectItemName     = itemname

	cEvtItem.FRectSellYN       = sellyn
	cEvtItem.FRectIsUsing      = usingyn	
	cEvtItem.FRectLimityn      = limityn
	cEvtItem.FRectMWDiv        = mwdiv
	cEvtItem.FRectDeliveryType = deliverytype
	cEvtItem.FRectSailYn       = sailyn
	cEvtItem.FRectCouponYn	   = couponyn

	cEvtItem.FRectCate_Large   = cdl
	cEvtItem.FRectCate_Mid     = cdm
	cEvtItem.FRectCate_Small   = cds	
				
 	arrList = cEvtItem.fnGetEventItem 		
 	iTotCnt = cEvtItem.FTotCnt	'전체 데이터  수

	iTotalPage 	=  int((iTotCnt-1)/iPageSize) +1  '전체 페이지 수		
%>

<script language="javascript">

function jsSerach(){
	var frm;
	frm = document.frm;
	frm.target = "_self";
	frm.action ="pop_eventitem_addinfo.asp";
	frm.submit();
}

// 페이지 이동
function jsGoPage(iP){
		document.frm.iC.value = iP;		
		document.frm.submit();	
}

function SelectItems(sType){	
var frm;
var itemcount = 0;
frm = document.frm;
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
			itemcount = "<%=iTotCnt%>";
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
	
	//frm.target = opener.name;
	frm.target = "FrameCKP";
	frm.action = "<%=actionURL%>";
	frm.itemcount.value = itemcount;
	frm.submit();
	frm.itemidarr.value = "";
	frm.itemcount.value = 0;	
	opener.history.go(0);	
	//window.close();
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

</script>

<!-- 검색 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="post">	
<input type="hidden" name="iC" >
<input type="hidden" name="sC" value="<%=sCode%>">
<input type="hidden" name="eC" value="<%=eCode%>">
<input type="hidden" name="egC" value="<%=egCode%>">
<input type="hidden" name="itemcount" value="0">
<input type="hidden" name="sType" >
<input type="hidden" name="itemidarr" >
<input type="hidden" name="mode" value="I">
<input type="hidden" name="acURL" value="<%=actionURL%>">
<tr align="center" bgcolor="<%= adminColor("topbar") %>" >
	<td rowspan="2" width="30" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
	<td align="left">
		<!-- include virtual="/common/module/categoryselectbox.asp"-->
		<br>
		브랜드 : <% drawSelectBoxLecturer "makerid", makerid %>
		상품코드 :
		<input type="text" class="text" name="itemid" value="<%= itemid %>" size="40" maxlength="100" onKeyPress="if (event.keyCode == 13) document.frm.submit();">			
		<br>상품명 : 
		<input type="text" class="text" name="itemname" value="<%= itemname %>" size="32" maxlength="20">(쉼표로 복수입력가능)
	</td>
	
	<td rowspan="2" width="30" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="검색" onClick="javascript:jsSerach();">
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("topbar") %>" >
	<td align="left">
		판매:<% drawSelectBoxSellYN "sellyn", sellyn %>
     	 
     	사용:<% drawSelectBoxUsingYN "usingyn", usingyn %>         	     	
     	 
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
			<input type="button" value="선택상품 추가" onClick="SelectItems('sel')" class="button">
			<input type="button" value="전체선택 추가" onClick="SelectItems('all')" class="button">
	</td>				
</tr>
</table> 
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr bgcolor="#FFFFFF">
	<td colspan="15" align="left">검색결과 : <b><%=iTotCnt%></b>&nbsp;&nbsp;페이지 : <b><%=iCurrpage%> / <%=iTotalPage%></b></td>
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
<%IF isArray(arrList) THEN 
	For intLoop = 0 To UBound(arrList,2)
%>
<tr class="a" height="25" bgcolor="#FFFFFF">
	<td  align="center"><input type="checkbox" name="chkitem" value="<%=arrList(0,intLoop)%>"></td>
	<td align="center"><A href="<%=wwwFingers%>/diyshop/shop_prd.asp?itemid=<%=arrList(0,intLoop)%>" target="_blank"><%=arrList(0,intLoop)%></a></td>
	<td align="center">
		<% if (Not IsNull(arrList(12,intLoop)) ) and (arrList(12,intLoop)<>"") then %>
			<img src="<%=imgFingers%>/diyItem/webimage/small/<%=GetImageSubFolderByItemid(arrList(0,intLoop))%>/<%=arrList(12,intLoop)%>">			
		<%end if%>
	</td>
	<td><%=db2html(arrList(3,intLoop))%></td>
	<td align="left">&nbsp;<%=db2html(arrList(4,intLoop))%></td>
	<td align="center">
		<%
			Response.Write FormatNumber(arrList(7,intLoop),0)
			'할인가
			if arrList(18,intLoop)="Y" then
				Response.Write "<br><font color=#F08050>(할)" & FormatNumber(arrList(9,intLoop),0) & "</font>"
			end if
			'쿠폰가
			if arrList(22,intLoop)="Y" then
				Select Case arrList(23,intLoop)
					Case "1"
						Response.Write "<br><font color=#5080F0>(쿠)" & FormatNumber(arrList(7,intLoop)*((100-arrList(24,intLoop))/100),0) & "</font>"
					Case "2"
						Response.Write "<br><font color=#5080F0>(쿠)" & FormatNumber(arrList(7,intLoop)-arrList(24,intLoop),0) & "</font>"
				end Select
			end if
		%>
	</td>
	<td align="center"><%
			Response.Write FormatNumber(arrList(8,intLoop),0)
			'할인가
			if arrList(18,intLoop)="Y" then
				Response.Write "<br><font color=#F08050>" & FormatNumber(arrList(10,intLoop),0) & "</font>"
			end if
			'쿠폰가
			if arrList(22,intLoop)="Y" then
				if arrList(23,intLoop)="1" or arrList(23,intLoop)="2" then
					if arrList(25,intLoop)=0 or isNull(arrList(25,intLoop)) then
						Response.Write "<br><font color=#5080F0>" & FormatNumber(arrList(8,intLoop),0) & "</font>"
					else
						Response.Write "<br><font color=#5080F0>" & FormatNumber(arrList(25,intLoop),0) & "</font>"
					end if
				end if
			end if
		%></td>
		<td align="center"><%= fnColor(cEvtItem.IsUpcheBeasong(arrList(15,intLoop)),"delivery")%></td>    	
		<td align="center"><%= fnColor(arrList(26,intLoop),"mw") %></td>
		<td align="center"><%= fnColor(arrList(14,intLoop),"yn") %></td>
		<td align="center"><%= fnColor(arrList(19,intLoop),"yn") %></td>
		<td align="center"><%= fnColor(arrList(16,intLoop),"yn") %></td>    
		<td align="center" nowrap>
		<!--<a href="javascript:PopItemStock('<%=arrList(0,intLoop)%>')" title="재고현황 팝업">[보기]</a>-->
		<% if cEvtItem.IsSoldOut(arrList(14,intLoop),arrList(16,intLoop),arrList(20,intLoop),arrList(21,intLoop)) then %>
					<br><img src="http://webadmin.10x10.co.kr/images/soldout_s.gif" width="30" height="12">
		<% end if %>
	</td>
</tr>
 <%	Next
ELSE
%>
<tr  align="center" bgcolor="#FFFFFF">
	<td colspan="15">등록된 내용이 없습니다.</td>
</tr>	
<%END IF%>
</table>
<!-- 페이징처리 -->
<%		
iStartPage = (Int((iCurrpage-1)/iPerCnt)*iPerCnt) + 1	

If (iCurrpage mod iPerCnt) = 0 Then																
	iEndPage = iCurrpage
Else								
	iEndPage = iStartPage + (iPerCnt-1)
End If	
%>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" >
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
    </tr>			  
</form>    
</table>
<iframe name="FrameCKP" src="about:blank" frameborder="0" width="300" height="300"></iframe>

<%
	set cEvtItem = nothing
%>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->