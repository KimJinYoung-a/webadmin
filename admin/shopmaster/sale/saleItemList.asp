<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  할인 상품 리스트
' History : 2008.04.08 정윤정 생성
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/util/datelib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/items/itemsalecls.asp"-->
<!-- #include virtual="/admin/eventmanage/common/event_function.asp"-->
<%
Dim clsSaleItem
dim makerid, itemid ,sPSale, cdl, cdm, cds
Dim sSalestatus, sItemSale,research
Dim iTotCnt, arrList,intLoop
Dim iPageSize, iCurrpage ,iDelCnt
Dim iStartPage, iEndPage, iTotalPage, ix,iPerCnt
Dim dispCate

itemid      = requestCheckvar(request("itemid"),255) 
makerid     = requestCheckvar(request("makerid"),32)
 
cdl = requestCheckvar(request("cdl"),10)
cdm = requestCheckvar(request("cdm"),10)
cds = requestCheckvar(request("cds"),10)
dispCate = requestCheckvar(request("disp"),16)

research =  Request("research")
sSalestatus 	= Request("salestatus")
sItemSale	= Request("selItemStatus")
'if sSalestatus = "" and research = "" then sSalestatus = 6
'if sItemSale = "" and research = "" then sItemSale = 6
iCurrpage = Request("iC")	'현재 페이지 번호

IF iCurrpage = "" THEN	iCurrpage = 1			
iPageSize = 20		'한 페이지의 보여지는 열의 수
iPerCnt = 10		'보여지는 페이지 간격

	if itemid<>"" then
	dim iA ,arrTemp,arrItemid 
	itemid = replace(itemid,chr(13),"")  
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

	itemid = left(arrItemid,len(arrItemid)-1)
end if	
	
set clsSaleItem = new CSaleItem
clsSaleItem.FCPage 	= iCurrpage
clsSaleItem.FPSize 	= iPageSize	
clsSaleItem.FItemid = itemid
clsSaleItem.FBrand 	= makerid
clsSaleItem.FRectCate_Large = cdl
clsSaleItem.FRectCate_Mid	= cdm
clsSaleItem.FRectCate_Small	= cds
clsSaleItem.FRectDispCate		= dispCate
clsSaleItem.FRectSaleStatus = sSalestatus
clsSaleItem.FRectItemSaleStatus = sItemSale
arrList = clsSaleItem.fnGetSaleOnItemList
iTotCnt = clsSaleItem.FTotCnt	'전체 데이터  수
 
iTotalPage 	=  int((iTotCnt-1)/iPageSize) +1  '전체 페이지 수 


'공통코드 값 배열로 한꺼번에 가져온 후 값 보여주기
	Dim  arrsalestatus	
	arrsalestatus = fnSetCommonCodeArr("salestatus",False)
	
	Dim sParm
	sParm = "itemid="&itemid&"&makerid="&makerid&"&cdl=" &cdl&"&cdm=" &cdm&"&cds=" &cds&"&disp="&dispCate 
%>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript">
 
	//페이징처리
	function jsGoPage(iP){
		document.frm.iC.value = iP;
		document.frm.submit();
	}
	
	//전체 선택
function SelectCk(opt){
	var bool = opt.checked;
	AnSelectAll(document.frmL,bool)
}

 
 //원가 할인가 적용
function CkOrgPrice(){
 if (confirm("원 가격으로 적용하시습니까?")){
	var arritem = "";
	var arrscode = "";
	 
 if (document.frmL.cksel.length==undefined){ 
     if(document.frmL.cksel.checked){ 
        arritem = document.frmL.cksel.value;
        arrscode = document.frmL.hidSC.value;
      }  
 }else{  
	for (var i=0;i<document.frmL.cksel.length;i++){ 
		 if(document.frmL.cksel[i].checked){
		    if (arritem == ""){
		        arritem = document.frmL.cksel[i].value;
		        arrscode = document.frmL.hidSC[i].value;
		    }else{    
		         arritem = arritem + "," + document.frmL.cksel[i].value;
		         arrscode =arrscode + ","+ document.frmL.hidSC[i].value;
		    }
		} 	 
	}
	  
 }
	if (arritem=="") {
		alert('선택 아이템이 없습니다.');
		return;
	}

    document.frmL.arrsalecode.value = arrscode;
    document.frmL.arrItemid.value = arritem;
    document.frmL.mode.value = "R"
    document.frmL.submit();
  }
}	
</script>
<!---- 검색 ---->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">	
	<form name="frm" method="get"  action="saleItemList.asp">
	<input type="hidden" name="menupos" value="<%=menupos%>">
	<input type="hidden" name="research" value="Y">
	<input type="hidden" name="iC">
  	<tr align="center" bgcolor="#FFFFFF" >
		<td width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
		<td align="left">
			<table class="a">
				<tr>
					<td>브랜드 : <%	drawSelectBoxDesignerWithName "makerid", makerid %>		</td>
					<td>
					    <% if (FALSE) then %><!-- 2016/04/15 by eastone -->
					    관리<!-- #include virtual="/common/module/categoryselectbox.asp"-->
					    전시<!-- #include virtual="/common/module/dispCateSelectBox.asp"-->
					    <% end if %>
					    </td> 
					<td>&nbsp;상품코드 :</td>
					<td rowspan="2"><textarea rows="3" cols="10" name="itemid" id="itemid"><%=replace(itemid,",",chr(10))%></textarea> </td>
				</tr>
				<tr>
					<td colspan="3">마스터 할인상태:
						<select name="salestatus" class="select" >
						<option value="">전체</option> 
						<option value="0"  <%if sSalestatus ="0" then%>selected<%end if%>>등록대기</option> 
						<option value="7"  <%if sSalestatus ="7" then%>selected<%end if%>>할인예정</option> 
						<option value="6"  <%if sSalestatus ="6" then%>selected<%end if%>>할인중</option> 
						<option value="9"  <%if sSalestatus ="9" then%>selected<%end if%>>할인중(종료예정)</option> 
						<option value="8"  <%if sSalestatus ="8" then%>selected<%end if%>>종료</option> 
						</select>
			 			&nbsp;&nbsp;
						상품 할인상태:
						<select name="selItemStatus" class="select"> <!--// 6-오픈, 7-오픈요청, 8-종료,9-종료요청-->
							<option value="">전체</option>
							<option value="7" <%if sItemSale ="7" then%>selected<%end if%>>할인예정</option>
							<option value="6" <%if sItemSale ="6" then%>selected<%end if%>>할인중</option>
							<option value="9" <%if sItemSale ="9" then%>selected<%end if%>>할인중(종료예정)</option> 
							<option value="8" <%if sItemSale ="8" then%>selected<%end if%>>할인종료</option>
						</select> 
					</td>	
				</tr>
			</table> 
		</td>
		<td  width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button_s" value="검색" onClick="javascript:document.frm.submit();">
		</td>
	</tr>	
</table>
</form>
<form name="frmL" method="post" action="saleItemProc.asp?<%=sParm%>">
    <input type="hidden" name="menupos" value="<%=menupos%>"> 
	<input type="hidden" name="iC">
	<input type="hidden" name="arrItemid"> 
	<input type="hidden" name="arrsalecode">
	<input type="hidden" name="mode" value=""> 
<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1">
<tr>
	<td align="right">		
	    <!--<input type="button" value="원 가격적용" onClick="CkOrgPrice();" class="button" <%if  itemid = "" then%>disabled<%end if%>>-->
	    <!--
	    <input type="button" value="2008 리뉴얼이전 할인상품목록" class="button" onClick="location.href='discountitemlist.asp'">
	    -->
	</td>
</tr>
</table>
<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<tr bgcolor="#FFFFFF">
	<td colspan="22" align="left">검색결과 : <b><%=iTotCnt%></b>&nbsp;&nbsp;페이지 : <b><%=iCurrpage%> / <%=iTotalPage%></b></td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
    <td><input type="checkbox" name="ck_all" onclick="SelectCk(this)"></td>
	<td>할인코드</td>
	<td>이벤트코드<br>(그룹)</td> 
	<td>마스터할인상태</td>
	<td>상품ID</td>
	<td>이미지</td>				
	<td>브랜드</td>
	<td>상품명</td>
	<td>계약<br>구분</td>
	<td>시작일</td>
	<td>종료일</td>
	<td>상품할인상태</td> 
	<td>현재<br>판매가</td>
	<td>현재<br>매입가</td>
	<td>현재<br>마진율</td> 
	<td>원<br>판매가</td>
	<td>원<br>매입가</td>
	<td>원<br>마진율</td> 
	<td>할인율</td>
	<td>할인<br>판매가</td>
	<td>할인<br>매입가</td>
	<td>할인<br>마진율</td>		
</tr>
 
<%IF isArray(arrList) THEN%>
	<%For intLoop = 0 To UBound(arrList,2) %>
<tr bgcolor="#FFFFFF"  align="center">	
    <input type="hidden" name="hidSC" value="<%=arrList(0,intLoop)%>">
      <td><input type="checkbox" name="cksel" value="<%=arrList(1,intLoop)%>" onClick="AnCheckClick(this);"></td>
	<td align="center"><a href="/admin/shopmaster/sale/saleReg.asp?sC=<%=arrList(0,intLoop)%>&menupos=290"><%=arrList(0,intLoop)%></a></td>		    				    	
	<td align="center"><%IF arrList(22,intLoop) >0 THEN%><%=arrList(22,intLoop)%><%IF arrList(23,intLoop) >0 THEN%>(<%=arrList(23,intLoop)%>)<%END IF%><%END IF%></td>		    				    	
    <td> <%IF arrList(26,intLoop) = 6 THEN%><font color="red"><%END IF%><%=fnGetCommCodeArrDesc(arrsalestatus,arrList(26,intLoop))%></td>
    <td align="center"><%=arrList(1,intLoop)%></td>		    				    	
	<td align="center"><%IF arrList(9,intLoop) <> "" THEN%><img src="http://webimage.10x10.co.kr/image/small/<%=GetImageSubFolderByItemid(arrList(1,intLoop))%>/<%=arrList(9,intLoop)%>"><%END IF%></td>	
	<td align="center"><%=db2html(arrList(7,intLoop))%></td>
	<td align="left">&nbsp;<%=db2html(arrList(8,intLoop))%></td>			    
	<td align="center"><%= fnColor(arrList(17,intLoop),"mw") %></td>
	<td align="center"><%=arrList(24,intLoop)%></td>
	<td align="center"><%=arrList(25,intLoop)%></td> 
	<td align="center"><%= fnColor(arrList(10,intLoop),"yn") %>&nbsp;<%IF arrList(4,intLoop) = 6 THEN%><font color="red"><%END IF%><%=fnGetCommCodeArrDesc(arrsalestatus,arrList(4,intLoop))%></td>
 
	<td align="center"><%IF arrList(10,intLoop) ="Y" THEN%><font color="red"><%END IF%><%=formatnumber(arrList(11,intLoop),0)%></td>
	<td align="center"><%IF arrList(10,intLoop) ="Y" THEN%><font color="red"><%END IF%><%=formatnumber(arrList(12,intLoop),0)%></td>
	<td align="center"><%IF arrList(10,intLoop) ="Y" THEN%><font color="red"><%END IF%>
		<% if arrList(11,intLoop)<>0 then %>
		<%= 100-fix(arrList(12,intLoop)/arrList(11,intLoop)*10000)/100 %>%
		<% end if %>	</td>
		
	<td align="center"> 
		<%=formatnumber(arrList(13,intLoop),0)%> 
		<%IF arrList(4,intLoop) = 6 THEN%>
		 <% IF arrList(27,intLoop) ="Y" Then%>
		<br/><font color="#F08050">(<%=formatnumber((arrList(13,intLoop)-arrList(28,intLoop))/arrList(13,intLoop)*100,0) %>%할)<%=formatnumber(arrList(28,intLoop),0)%></font>
		 <% END IF%>
		<%ELSEIF arrList(4,intLoop) <> 8 and arrList(10,intLoop) ="Y" THEN%>
		<br/><font color="#F08050">(<%=formatnumber((arrList(13,intLoop)-arrList(15,intLoop))/arrList(13,intLoop)*100,0) %>%할)<%=formatnumber(arrList(15,intLoop),0)%></font>
		<%END IF%>
	</td>
	<td align="center">
		<%=formatnumber(arrList(14,intLoop),0)%>
		<%IF arrList(4,intLoop) = 6 THEN%>
		 <% IF arrList(27,intLoop) ="Y" Then%>
		<br/><font color="#F08050"><%=formatnumber(arrList(29,intLoop),0)%></font>
		 <% END IF%>
		<%ELSEIF arrList(4,intLoop) <> 8 and arrList(10,intLoop) ="Y" THEN%>
		<br/><font color="#F08050"><%=formatnumber(arrList(16,intLoop),0)%></font>
		<%END IF%> 
	</td>
	<td align="center"><% if arrList(13,intLoop)<>0 then %>
		<%= 100-fix(arrList(14,intLoop)/arrList(13,intLoop)*10000)/100 %>%
		<% end if %>	
		<%IF arrList(4,intLoop) = 6 THEN%>
		 <% IF arrList(27,intLoop) ="Y" Then%>
		<br/><font color="#F08050"><%=100-fix(arrList(29,intLoop)/arrList(28,intLoop)*10000)/100%>%</font>
		 <% END IF%>
		<%ELSEIF arrList(4,intLoop) <> 8 and arrList(10,intLoop) ="Y" THEN%>
		<br/><font color="#F08050"><%=100-fix(arrList(16,intLoop)/arrList(15,intLoop)*10000)/100%>%</font>
		<%END IF%>	
	</td> 
	<td><% if arrList(13,intLoop)<>0 then %><%=formatnumber(((arrList(13,intLoop)-arrList(2,intLoop))/arrList(13,intLoop))*100,0)%>%<%end if%></td>	
	<td align="center"><%IF arrList(10,intLoop) ="Y" THEN%><font color="red"><%END IF%><%if arrList(4,intLoop) = 8 then%><font color="Gray"><%end if%><%=formatnumber(arrList(2,intLoop),0)%></td>
	<td align="center"><%IF arrList(10,intLoop) ="Y" THEN%><font color="red"><%END IF%><%if arrList(4,intLoop) = 8 then%><font color="Gray"><%end if%><%=formatnumber(arrList(3,intLoop),0)%></td>
	<td align="center"><%IF arrList(10,intLoop) ="Y" THEN%><font color="red"><%END IF%><%if arrList(4,intLoop) = 8 then%><font color="Gray"><%end if%><% if arrList(2,intLoop)<>0 then %>
		<%= 100-fix(arrList(3,intLoop)/arrList(2,intLoop)*10000)/100 %>%
		<% end if %></td>
	</tr>	
	<%Next%>
<%ELSE%>
<tr bgcolor="#FFFFFF">
	<td colspan="22" align="center">등록된 내역이 없습니다.</td>
</tr>
<%END IF%>
</table>
</form>
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
				if Cdbl(ix) = Cdbl(iCurrpage) then
		%>
			<a href="javascript:jsGoPage(<%= ix %>)" class="menu_link3" onfocus="this.blur();"><font color="00abdf"><strong><%=ix%></strong></font></a>
		<%		else %>
			<a href="javascript:jsGoPage(<%= ix %>)" class="menu_link3" onfocus="this.blur();"><%=ix%></a>
		<%
				end if
			next
		%>
    	<% if Cdbl(iTotalPage) > Cdbl(iEndPage)  then %><a href="javascript:jsGoPage(<%= ix %>)" onfocus="this.blur();">[next]</a>
		<% else %>[next]<% end if %>
        </td>
        <td  width="50" align="right"><a href="saleList.asp?menupos=<%=menupos%>"><img src="/images/icon_list.gif" border="0"></a></td>			        
    </tr>	  
</table> 
<%
set clsSaleItem = nothing
%>
<!---- /검색 ---->
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->