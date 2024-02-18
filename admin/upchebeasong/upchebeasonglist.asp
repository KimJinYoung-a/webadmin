<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/order/upchebeasongcls.asp"-->
<%

dim yyyy1,yyyy2,mm1,mm2,dd1,dd2
dim nowdate,searchnextdate
dim makerid,dateback, cdl
dim cknodate,page, detailstate

nowdate = Left(CStr(now()),10)
makerid = requestCheckVar(request("makerid"),32)

yyyy1   = requestCheckVar(request("yyyy1"),4)
mm1     = requestCheckVar(request("mm1"),2)
dd1     = requestCheckVar(request("dd1"),2)
yyyy2   = requestCheckVar(request("yyyy2"),4)
mm2     = requestCheckVar(request("mm2"),2)
dd2     = requestCheckVar(request("dd2"),2)

detailstate   = requestCheckVar(request("detailstate"),9)
cdl         = requestCheckVar(request("cdl"),3)
cknodate    = requestCheckVar(request("cknodate"),16)
page        = requestCheckVar(request("page"),9)

if (page="") then page=1

if (yyyy1="") then
	yyyy1 = Left(nowdate,4)
	mm1   = Mid(nowdate,6,2)
	dd1   = Mid(nowdate,9,2)
	yyyy2 = yyyy1
	mm2   = mm1
	dd2   = dd1

    dateback = DateSerial(yyyy1,mm2, dd2-7)

    yyyy1 = Left(dateback,4)
    mm1   = Mid(dateback,6,2)
    dd1   = Mid(dateback,9,2)
end if

searchnextdate = Left(CStr(DateAdd("d",Cdate(yyyy2 + "-" + mm2 + "-" + dd2),1)),10)




dim ojumun

set ojumun = new CBaljuMaster

if cknodate="" then
	ojumun.FRectRegStart = yyyy1 + "-" + mm1 + "-" + dd1
	ojumun.FRectRegEnd = searchnextdate
end if


ojumun.FRectDesignerID = makerid
ojumun.FPageSize = 50
ojumun.FCurrPage = page
ojumun.FRectCDL  = cdl
ojumun.FRectDetailState = detailstate

if (makerid<>"")  then
    ojumun.getUpchebeasongList
end if

dim ix,iy
%>
<script language='javascript'>
function changecontent(){
    //nothing
}

function ViewOrderDetail(frm){
	//var popwin;
    //popwin = window.open('','orderdetail');
    frm.target = '_blank';
    frm.action="/admin/ordermaster/viewordermaster.asp"
	frm.submit();

}

function ViewItem(itemid){
window.open("http://www.10x10.co.kr/shopping/category_prd.asp?itemid=" + itemid,"sample");
}

function NextPage(ipage){
	document.frm.page.value= ipage;
	document.frm.submit();
}

</script>


<!-- 검색 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frm" method="get" action="">
	<input type="hidden" name="page" value="1">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<tr align="center" bgcolor="#FFFFFF" >
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
		<td align="left">
			브랜드 : <% drawSelectBoxDesignerwithName "makerid", makerid %>
			&nbsp;
			검색기간 : <% DrawDateBox yyyy1,yyyy2,mm1,mm2,dd1,dd2 %>
		</td>
		
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button_s" value="검색" onClick="javascript:document.frm.submit();">
		</td>
	</tr>
	<tr bgcolor="#FFFFFF" >
	    <td>
	        카테고리 : <% DrawSelectBoxCategoryLarge "cdl",cdl %>
			&nbsp;&nbsp;
	        상태 : 
			<select name="detailstate">
			<option value="">전체(출고완료포함)
			<option value="NOT7" <%= ChkIIF(detailstate="NOT7","selected","") %> >미출고전체(결제완료)
			<option value="0" <%= ChkIIF(detailstate="0","selected","") %> >결제완료(미통보)
			<option value="2" <%= ChkIIF(detailstate="2","selected","") %> >주문통보(업체통보)
			<option value="3" <%= ChkIIF(detailstate="3","selected","") %> >상품준비(주문확인)
			<option value="MOO" <%= ChkIIF(detailstate="MOO","selected","") %> >무통장접수
			</select>
			&nbsp;
			
			<!--
			상품코드 :
			<input type="text" class="text" name="" value="" size="32"> (쉼표로 복수입력가능)
			-->
		</td>
	</tr>
	</form>
</table>
<!-- 검색 끝 -->

<p>

<!-- 리스트 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="15">
			검색결과 : <b><% = ojumun.FTotalCount %></b>
			&nbsp;
			페이지 : <b><%= page %> / <%= ojumun.FTotalpage %></b>
		</td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td>브랜드ID</td>
		<td width="70">주문번호</td>
		<td width="60">주문자</td>
		<td width="60">수령인</td>
		<td width="50">상품코드</td>
		<td>상품명<font color="blue">[옵션명]</font></td>
		<td width="40">수량</td>
<!--	<td width="40">취소<br>구분</td>	없음 .-->
<!--	<td width="60">주문일</td>	-->
<!--	<td width="60">결제일</td>	-->
		<td width="60">주문통보일<br>(기준일)</td>
		<td width="60">주문확인일</td>
		<td width="60">출고일</td>
		<td width="40">소요<br>일수</td>
		<td width="60">진행상태</td>
	</tr>
	<% if ojumun.FresultCount<1 then %>
	<tr bgcolor="#FFFFFF">
		<td colspan="13" align="center">[검색결과가 없습니다.]
		<% if (makerid="") then %>
		<br> <font color='red'>브랜드를 먼저 선택 하세요.</font>
		<% end if %>
		</td>
	</tr>
<% else %>
	<% for ix=0 to ojumun.FresultCount-1 %>
	<form name="frmBuyPrc_<%= ix %>" method="post" >
	<input type="hidden" name="orderserial" value="<%= ojumun.FMasterItemList(ix).FOrderSerial %>">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<% if ojumun.FMasterItemList(ix).IsAvailJumun then %>
	<tr class="a" align="center" bgcolor="FFFFFF">
	<% else %>
	<tr class="gray" align="center" bgcolor="FFFFFF">
	<% end if %>
		<td><%= ojumun.FMasterItemList(ix).FMakerid %></td>
		<td><a href="javascript:ViewOrderDetail(frmBuyPrc_<%= ix %>)" class="zzz"><%= ojumun.FMasterItemList(ix).FOrderSerial %></a></td>
		<td><%= ojumun.FMasterItemList(ix).FBuyname %></td>
		<td><%= ojumun.FMasterItemList(ix).FReqname %></td>
		<td><%= ojumun.FMasterItemList(ix).FItemid %></td>
		<td align="left">
			<a href="javascript:ViewItem(<% =ojumun.FMasterItemList(ix).FItemid  %>)"><%= ojumun.FMasterItemList(ix).FItemname %></a>
				<% if (ojumun.FMasterItemList(ix).FItemoption<>"") then %>
					<font color="blue">[<%= ojumun.FMasterItemList(ix).FItemoption %>]</font>
				<% end if %>
		</td>
		<td><%= ojumun.FMasterItemList(ix).FItemcnt %></td>
<!--			<td>
		<% if ojumun.FMasterItemList(ix).FCancelYn <> "Y" then %>
		&nbsp;
		<% else %>
		<font color="red">주문취소</font>
		<% end if %>
		</td>	-->
<!--	<td><%= FormatDateTime(ojumun.FMasterItemList(ix).FRegdate,2) %></td>	-->
<!--	<td></td>	-->
		<td><%= Left(ojumun.FMasterItemList(ix).Fbaljudate,10) %></td>
		<td><%= Left(ojumun.FMasterItemList(ix).Fupcheconfirmdate,10) %></td>
		<td><%= Left(ojumun.FMasterItemList(ix).FUpcheBeasongDate,10) %></td>
		<td><%= ojumun.FMasterItemList(ix).getBeasongDPlusDateStr %></td>
		<td>
		    <% if (detailstate="MOO") then %>
		    
		    <% else %>
    			<% if ojumun.FMasterItemList(ix).FCurrstate = 0 then %>
    			<font color="blue">결제완료</font>
    			<% elseif ojumun.FMasterItemList(ix).FCurrstate = 2 then %>
    			<font color="#000000">주문통보</font>
    			<% elseif ojumun.FMasterItemList(ix).FCurrstate = 3 then %>
    			<font color="#CC9933">주문확인</font>
    			<% elseif ojumun.FMasterItemList(ix).FCurrstate = 7 then %>
    			<font color="#FF0000">출고완료</font>
    			<% end if %>
			<% end if %>
		</td>
	</tr>
	</form>
	<% next %>
<% end if %>
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="15" align="center">
		<% if ojumun.HasPreScroll then %>
			<a href="javascript:NextPage('<%= ojumun.StartScrollPage-1 %>')">[pre]</a>
		<% else %>
			[pre]
		<% end if %>
		<% for ix=0 + ojumun.StartScrollPage to ojumun.FScrollCount + ojumun.StartScrollPage - 1 %>
			<% if ix>ojumun.FTotalpage then Exit for %>
			<% if CStr(page)=CStr(ix) then %>
			<font color="red">[<%= ix %>]</font>
			<% else %>
			<a href="javascript:NextPage('<%= ix %>')">[<%= ix %>]</a>
			<% end if %>
		<% next %>

		<% if ojumun.HasNextScroll then %>
			<a href="javascript:NextPage('<%= ix %>')">[next]</a>
		<% else %>
			[next]
		<% end if %>
		</td>
	</tr>
</table>

<%
set ojumun = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->