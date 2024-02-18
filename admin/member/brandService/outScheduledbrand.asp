<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbSTSopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/partners/outBrandCls.asp"-->
<!-- #include virtual="/lib/classes/displaycate/displaycateCls.asp"-->
<%
Dim i
Dim page : page = ReQuestCheckvar(request("page"),10)
Dim research : research = ReQuestCheckvar(request("research"),10)
Dim makerid : makerid = ReQuestCheckvar(request("makerid"),32)
Dim dispcate1 : dispcate1 = ReQuestCheckvar(request("dispcate1"),3)
Dim outstatus : outstatus = ReQuestCheckvar(request("outstatus"),10)

if (page="") then page=1
if (research="" and outstatus="") then outstatus="4"

Dim oOutBrand, isJustView : isJustView = false
SET oOutBrand = new COutBrand
    oOutBrand.FPageSize=50
    oOutBrand.FCurrPage=page
    oOutBrand.FRectMakerid  = makerid
    oOutBrand.FRectOutbrandStatus = outstatus
	oOutBrand.FRectDispCate1 = dispcate1
	if (LEN(outstatus)>1) then
		if (outstatus="999") then
			oOutBrand.getCompanyClosedAndSellitemExistsBrandList
		else
			oOutBrand.FRectPreDay = outstatus
			oOutBrand.getOutBrandCheckList

			
		end if
		isJustView = true
		
	else
    	oOutBrand.getOutBrandScheduledList
	end if
%>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script> 
<script type="text/javascript">
function goPage(pg){
    frm.page.value = pg;
    frm.submit();
}

function fnBrandItemDelayProc(imakerid,isubidx){
	if (confirm("브랜드 ("+imakerid+") 상품 정리예정일을 90일 연장 하시겠습니까? (산정일기준+90 day)")){
		document.frmact.mode.value		= "delay90"; // "delay30"
		document.frmact.makerid.value	= imakerid;
		document.frmact.subidx.value	= isubidx;
		document.frmact.submit();
	}
}

function fnBrandItemKillProc(imakerid,isubidx){
	if (confirm("브랜드 ("+imakerid+") 상품을 품절처리 하시겠습니까?")){
		document.frmact.mode.value		= "soldoutitems"
		document.frmact.makerid.value	= imakerid
		document.frmact.subidx.value	= isubidx
		document.frmact.submit();
	}
}
</script>

<!-- 표 상단바 시작-->
<form name="frm" method="get" action="">
	<input type="hidden" name="page" value="1">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<input type="hidden" name="research" value="on">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">  
	<tr align="center" bgcolor="F4F4F4">
	    <td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
        <td bgcolor="#FFFFFF" align="left">
            <% if (FALSE) then %>
        	기준년월 <% DrawYMBox yyyy1,mm1 %>
        	&nbsp;&nbsp;
            <% end if %>
        	브랜드ID : <% drawSelectBoxDesignerwithName "makerid",makerid  %>
        	&nbsp;&nbsp;
        	
			브랜드 대표전시카테고리 : <%= fnStandardDispCateSelectBox(1,"", "dispcate1", dispcate1, "")%></div>
        </td>
        <td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
    		<input type="button" class="button_s" value="검색" onClick="document.frm.submit();">
    	</td>
	</tr> 
	<tr>
	    <td bgcolor="#FFFFFF" align="left">
	       상태 :
		   <input type="radio" name="outstatus" value="" <%=CHKIIF(outstatus="","checked","") %> >전체
		   <input type="radio" name="outstatus" value="4" <%=CHKIIF(outstatus="4","checked","") %> >정리예정+연장
		   
		   <input type="radio" name="outstatus" value="0" <%=CHKIIF(outstatus="0","checked","") %> >정리예정
		   <input type="radio" name="outstatus" value="3" <%=CHKIIF(outstatus="3","checked","") %> >연장
		   <input type="radio" name="outstatus" value="7" <%=CHKIIF(outstatus="7","checked","") %> >정리완료

		   &nbsp;&nbsp;
		   |
		   &nbsp;&nbsp;
		   ( 
		   검토용 :
		   <input type="radio" name="outstatus" value="92365" <%=CHKIIF(outstatus="92365","checked","") %> >3개월~1년간 판매, 신상↓

		   <input type="radio" name="outstatus" value="365" <%=CHKIIF(outstatus="365","checked","") %> >1년간 판매, 신상↓
		   <input type="radio" name="outstatus" value="183" <%=CHKIIF(outstatus="183","checked","") %> >6개월 판매, 신상↓
		   <input type="radio" name="outstatus" value="92" <%=CHKIIF(outstatus="92","checked","") %> >3개월 판매, 신상↓

		   &nbsp;

		   <input type="radio" name="outstatus" value="999" <%=CHKIIF(outstatus="999","checked","") %> >폐업/휴업업체
		   )
	    </td>
	</tr>
</table>
</form>
<!-- 표 상단바 끝--> 
<p>
<table width="100%" border="0" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>" class=a>
<tr>
    <td colspan="15" bgcolor="#FFFFFF" height="30"> 
    		검색결과 : <b> <%=formatnumber(oOutBrand.FTotalCount,0)%></b> (<%=formatnumber(oOutBrand.FMayTotalpreSellitemNo,0)%>)
			&nbsp;
			페이지 : <b><%=formatnumber(page,0)%>/ <%=formatnumber(oOutBrand.FTotalPage,0)%></b> 
   </td>
</tr>
<tr align="center"  bgcolor="<%= adminColor("tabletop") %>">
	<% if (outstatus="999") then %>
		<td width="100">브랜드ID</td>
		<td width="30">-</td>
		<td width="100">전시Cate</td>
		<td>조회일</td>
		<td>상태</td>
		<td>조회일 판매중인<br>상품수량</td>
		<td>그룹코드</td>
		<td>사업자번호</td>
		<td>상품목록</td>
		<td>연장요청</td>
		<td>일괄처리이동</td>
	<% else %>
		<td width="100">브랜드ID</td>
		<td width="30">차수</td>
		<td width="100">전시Cate</td>
		<td>산정일</td>
		<td>산정기간</td>
		<td>산정일 판매중인<br>상품수량</td>
		<td>기간중 등록<br>상품수량</td>
		<td>기간중 판매된<br>수량</td>
		<td>기간<br>매출</td>
		
		<td>정리예정일</td>
		<td width="50">남은일수</td>
		<td>상태</td>
		<td>상품목록</td>
		<td>연장요청</td>
		<td>바로정리</td>
	<% end if %>
</tr>

<% if oOutBrand.FResultCount<1 then %>
	<tr bgcolor="#FFFFFF" >
		<td colspan="25" align=center>[검색결과가 없습니다.]</td>
	</tr>
<% else %>
	<% for i=0 to oOutBrand.FResultCount -1 %>
	<% if oOutBrand.FItemList(i).Fmakerid="Y"	then %>
	<tr bgcolor="#FFFFFF" align=center>
	<% else %>
	<tr bgcolor="#FFFFFF" align=center>
	<% end if %>
	<% if (outstatus="999") then %>   
		<td><%= oOutBrand.FItemList(i).FMakerid %></td>
		<td></td>
		<td><%= oOutBrand.FItemList(i).Fdispcate1Name %></td>
        <td><%= oOutBrand.FItemList(i).Fregdate %></td>
		<td><%= oOutBrand.FItemList(i).getScoExpireStatText %></td>
		<td><%= FormatNumber(oOutBrand.FItemList(i).FpreSellitemNo,0) %></td>
		<td><%= oOutBrand.FItemList(i).FGroupid %></td>
		<td><%= oOutBrand.FItemList(i).FSocCompanyNo %></td>
		<td>
		<a href="/admin/itemmaster/itemlist.asp?menupos=594&page=1&makerid=<%= oOutBrand.FItemList(i).Fmakerid %>&sellyn=Y" target="_xpireitems">[상품목록]</a>
		</td>
		<td>-</td>
		<td>
		<a href="/admin/shopmaster/itemviewset.asp?menupos=24&makerid=<%= oOutBrand.FItemList(i).Fmakerid %>&sellyn=Y&mwdiv=U" target="_xpireitems2">[업배]</a>
		&nbsp;
		<a href="/admin/shopmaster/itemviewset.asp?menupos=24&makerid=<%= oOutBrand.FItemList(i).Fmakerid %>&sellyn=Y&mwdiv=W" target="_xpireitems2">[위탁]</a>
		</td>

	<% else %>
		<td><%= oOutBrand.FItemList(i).FMakerid %></td>
		<td><%= oOutBrand.FItemList(i).Fsubidx %></td>
		<td><%= oOutBrand.FItemList(i).Fdispcate1Name %></td>
        <td><%= oOutBrand.FItemList(i).Fregdate %></td>
		<td>- <%= oOutBrand.FItemList(i).Fpreday %> D</td>
		<td><%= FormatNumber(oOutBrand.FItemList(i).FpreSellitemNo,0) %></td>
		<td><%= oOutBrand.FItemList(i).FpreRegedItemno %></td>
		<td><%= oOutBrand.FItemList(i).FpreSellitemNoSum %></td>
		<td align=right><%= FormatNumber(oOutBrand.FItemList(i).FpreSellCostSum,0) %></td>
		
		<% if (isJustView) then %>
		<td>-</td>
		<td>-</td>
		<td>-</td>
		<td>
		<a href="/admin/itemmaster/itemlist.asp?menupos=594&page=1&makerid=<%= oOutBrand.FItemList(i).Fmakerid %>&sellyn=Y" target="_xpireitems">[상품목록]</a>
		</td>
		<td>
			-
		</td>
		<td>
			-
		</td>
		<% else %>
        <td><%= oOutBrand.FItemList(i).FoutScheduledate %></td>
		<td><%= oOutBrand.FItemList(i).getRemainDate %></td>
		<td><%= oOutBrand.FItemList(i).getOutbrandStatusHtml %></td>
		<td>
		<a href="/admin/itemmaster/itemlist.asp?menupos=594&page=1&makerid=<%= oOutBrand.FItemList(i).Fmakerid %>&sellyn=Y" target="_xpireitems">[상품목록]</a>
		</td>
		<td>
			<% if oOutBrand.FItemList(i).IsActionDelayAvailState() then %>
		    <input type="button" value="연장요청" onClick="fnBrandItemDelayProc('<%= oOutBrand.FItemList(i).Fmakerid %>','<%= oOutBrand.FItemList(i).Fsubidx %>');" class="button">
        	<% end if %>
		</td>
		<td>
			<% if oOutBrand.FItemList(i).IsActionFinAvailState() then %>
		    <input type="button" value="바로정리" onClick="fnBrandItemKillProc('<%= oOutBrand.FItemList(i).Fmakerid %>','<%= oOutBrand.FItemList(i).Fsubidx %>');" class="button">
        	<% end if %>
		</td>
		<% end if %>
	<% end if %>
	</tr>
	<% next %>
<% end if %>

 
 <!-- 페이징처리 --> 
<tr height="20">
    <td colspan="15" align="center" bgcolor="#FFFFFF">
        <% if oOutBrand.HasPreScroll then %>
		<a href="javascript:goPage('<%= oOutBrand.StartScrollPage-1 %>');">[pre]</a>
    	<% else %>
    		[pre]
    	<% end if %>

    	<% for i=0 + oOutBrand.StartScrollPage to oOutBrand.FScrollCount + oOutBrand.StartScrollPage - 1 %>
    		<% if i>oOutBrand.FTotalpage then Exit for %>
    		<% if CStr(page)=CStr(i) then %>
    		<font color="red">[<%= i %>]</font>
    		<% else %>
    		<a href="javascript:goPage('<%= i %>');">[<%= i %>]</a>
    		<% end if %>
    	<% next %>

    	<% if oOutBrand.HasNextScroll then %>
    		<a href="javascript:goPage('<%= i %>');">[next]</a>
    	<% else %>
    		[next]
    	<% end if %>
    </td>
</tr>
</table>
<%
set oOutBrand = Nothing
%>
<form name="frmact" method="post" target="iifrsubmit" action="outscheduledbrand_Process.asp">
<input type=hidden name="mode">
<input type=hidden name="makerid">
<input type=hidden name="subidx">
</form>
<iframe width="600" height="100" id="iifrsubmit" name="iifrsubmit" ></iframe>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbSTSclose.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->