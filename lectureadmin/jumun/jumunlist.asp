<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/designer/incSessionDesigner.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/designer/lib/designerbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lectureadmin/lib/classes/jumun/jumuncls.asp"-->
<%

dim searchtype
dim searchrect

dim yyyy1,yyyy2,mm1,mm2,dd1,dd2
dim nowdate,searchpredate,searchnextdate
dim orderserial,cknodate, isupchebeasong
dim datetype

nowdate = Left(CStr(now()),10)

searchtype = RequestCheckvar(request("searchtype"),2)
searchrect = requestCheckVar(request("searchrect"),32)

datetype   = RequestCheckvar(request("datetype"),16)

''if (datetype="") then datetype="ipkumil"
if (datetype="") then datetype="jumunil"        ''2009 주문일로 변경 : 주문접수건도 표시.

yyyy1   = RequestCheckvar(request("yyyy1"),4)
mm1     = RequestCheckvar(request("mm1"),2)
dd1     = RequestCheckvar(request("dd1"),2)
yyyy2   = RequestCheckvar(request("yyyy2"),4)
mm2     = RequestCheckvar(request("mm2"),2)
dd2     = RequestCheckvar(request("dd2"),2)
isupchebeasong = RequestCheckvar(request("isupchebeasong"),2)

if (yyyy1="") then
	yyyy1 = Left(nowdate,4)
	mm1   = Mid(nowdate,6,2)
	dd1   = Mid(nowdate,9,2)

	yyyy2 = yyyy1
	mm2   = mm1
	dd2   = dd1
end if

'날짜형태를 맞춤 (2008.05.26;허진원)
'searchpredate 수정 (2009.01.09;서동석)
searchpredate = Left(CStr(DateSerial(yyyy1 , mm1 , dd1)),10)
searchnextdate = Left(CStr(DateAdd("d",DateSerial(yyyy2 , mm2 , dd2),1)),10)

dim page
dim ojumun

page = RequestCheckvar(request("page"),10)
if (page="") then page=1

set ojumun = new CJumunMaster

'if cknodate="" and searchrect="" then
	ojumun.FRectRegStart = searchpredate
	ojumun.FRectRegEnd = searchnextdate
'end if

if searchtype="01" then
	ojumun.FRectOrderSerial = searchrect
elseif searchtype="02" then
	ojumun.FRectBuyname = searchrect
elseif searchtype="03" then
	ojumun.FRectReqName = searchrect
elseif searchtype="04" then
	ojumun.FRectUserID = searchrect
elseif searchtype="05" then
	ojumun.FRectIpkumName = searchrect
elseif searchtype="06" then
	ojumun.FRectSubTotalPrice = searchrect
elseif searchtype="11" then
	ojumun.FRectitemid = searchrect
end if

ojumun.FRectDesignerID = session("ssBctID")
ojumun.FPageSize = 50
ojumun.FCurrPage = page
ojumun.FRectDateType = datetype
ojumun.FRectIsUpcheBeasong = isupchebeasong
ojumun.SearchJumunListByDesigner

dim ix,iy
dim isalltenbeasong
isalltenbeasong = ojumun.IsAllTenBeasong


%>
<script language='javascript'>
function ViewOrderDetail(frm){
	//var popwin;
    //popwin = window.open('','upcheorderpop');
    frm.target = 'upcheorderpop';
    frm.action="popviewordermaster.asp"
	frm.submit();

}
function ShowOrderInfo(frm,orderserial){
	var props = "width=600, height=600, location=no, status=yes, resizable=no, scrollbars=yes";
	window.open("about:blank", "orderdetail", props);
    frm.target = "orderdetail";
    frm.orderserial.value = orderserial;
    frm.action="popviewordermaster.asp";
	frm.submit();
}
function ViewUserInfo(frm){

}

function ViewItem(itemid){
    var popwin = window.open("http://www.thefingers.co.kr/diyshop/shop_prd.asp?itemid=" + itemid,"sample");
    popwin.focus();
}

function NextPage(ipage){
	document.frm.page.value= ipage;
	document.frm.submit();
}

function checkdate(){
    var frm=document.frm;
	if ((frm.searchrect.value.length>0)&&(frm.searchtype.value=="")){
		alert("검색조건을 선택 하세요.");
		frm.searchtype.focus();
		return;
	}

    if ((frm.searchtype.value=="11")&&(!IsDigit(frm.searchrect.value))){
        alert("상품코드는 숫자만 가능합니다.");
		frm.searchrect.focus();
		return;
    }

    if((frm.yyyy2.value - frm.yyyy1.value) > 1){
	    alert("3개월 이내로 검색하셔야 합니다.");
		return;
	}
	else if(frm.yyyy1.value == frm.yyyy2.value){
	  if(((frm.mm2.value * 30) - (frm.dd2.value - 30))-((frm.mm1.value * 30) - (frm.dd1.value - 30)) > 90){
	    alert("3개월 이내로 검색하셔야 합니다.");
		return;
      }
	}
    else if(frm.yyyy1.value < frm.yyyy2.value){
	  if(((frm.mm2.value * 30) - (frm.dd2.value - 30)) + (((12-frm.mm1.value)*30) - (frm.dd1.value - 30)) > 90){
	    alert("3개월 이내로 검색하셔야 합니다.");
		return;
      }
	}



    frm.submit();

}
</script>

<!-- 검색 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frm" method="get" action="jumunlist.asp">
	<input type="hidden" name="page" value="1">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<tr align="center" bgcolor="#FFFFFF" >
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
		<td align="left">
			검색조건 :
			<select class="select" name="searchtype">
				<option value="">선택</option>
				<option value="01" <% if searchtype="01" then response.write "selected" %> >주문번호</option>
				<option value="02" <% if searchtype="02" then response.write "selected" %> >구매자</option>
				<option value="03" <% if searchtype="03" then response.write "selected" %> >수령인</option>
				<option value="04" <% if searchtype="04" then response.write "selected" %> >아이디</option>
				<!-- option value="05" <% if searchtype="05" then response.write "selected" %> >입금자</option -->
				<!-- option value="06" <% if searchtype="06" then response.write "selected" %> >결제금액</option -->
				<option value="11" <% if searchtype="11" then response.write "selected" %> >상품코드</option>
			</select>
			<input type="text" class="text" name="searchrect" value="<%= searchrect %>" size="11" maxlength="16">
			&nbsp;
			검색기간 : <% DrawDateBox yyyy1,yyyy2,mm1,mm2,dd1,dd2 %>
			<input type="radio" name="datetype" value="jumunil" <% if (datetype="jumunil") then response.write "checked" %> >주문일
			<input type="radio" name="datetype" value="ipkumil" <% if (datetype="ipkumil") then response.write "checked" %> >결제일
			<input type="radio" name="datetype" value="upbeasongdate" <% if (datetype="upbeasongdate") then response.write "checked" %> >출고일
			<!-- 상품별 출고일로 검색 텐배 업배 상관없이 -->
			<!--<input type="radio" name="datetype" value="tenbeasongdate" <% if (datetype="tenbeasongdate") then response.write "checked" %> >출고일(텐바이텐)-->
		</td>

		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button_s" value="검색" onClick="javascript:checkdate();">
		</td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF" >
		<td align="left">
	     	배송구분 :
			<select class="select" name="isupchebeasong">
	     	<option value="">전체</option>
	     	<!--
	     	<option value="N" <%= CHKIIF(isupchebeasong="N","selected","") %> >텐바이텐배송</option>
	     	-->
	     	<option value="Y" <%= CHKIIF(isupchebeasong="Y","selected","") %> >업체개별배송</option>
	     	</select>
		</td>
	</tr>
	</form>
</table>
<!-- 검색 끝 -->

<p>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="15">
		    <table width="100%" border="0" cellspacing="0" cellpadding="0" class="a" >
		    <tr>
		        <td>
    			검색결과 : <b><% =ojumun.FTotalCount %></b>
    			&nbsp;
    			페이지 : <b><%= page %> / <%= ojumun.FTotalpage %></b>
        		</td>
        		<td align="right"> 공급가계 : <strong><%= FormatNumber(ojumun.FTotalBuyCash,0) %></strong></td>
    		</tr>
    		</table >
		</td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td width="70">주문번호</td>
		<td width="50">구매자</td>
		<td width="50">수령인</td>
		<td width="50">상품코드</td>
		<td>상품명<font color="blue">[옵션명]</font></td>
		<td width="30">수량</td>
		<td width="40">판매가</td>

		<td width="40">공급가</td>
<!--	<td width="60">결제방법</td>	-->
<!--	<td width="60">텐바이텐<br>진행상태</td>	-->

		<td width="60">주문일</td>
		<td width="60">결제일</td>
		<td width="60">출고일</td>

		<td width="60">배송<br>구분</td>
		<td width="60">진행상태</td>
	</tr>
<% if ojumun.FresultCount<1 then %>
	<tr align="center" bgcolor="#FFFFFF">
		<td colspan="14">[검색결과가 없습니다.]</td>
	</tr>
<% else %>
	<% for ix=0 to ojumun.FresultCount-1 %>
	<form name="frmOnerder_<%= ix %>" method="post" >
	<input type="hidden" name="orderserial" value="<%= ojumun.FMasterItemList(ix).FOrderSerial %>">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<input type="hidden" name="sitename" value="<%= ojumun.FMasterItemList(ix).FSiteName %>">
	<input type="hidden" name="userid" value="<%= ojumun.FMasterItemList(ix).FUserID %>">
	<% if ojumun.FMasterItemList(ix).IsAvailJumun then %>
	<tr class="a" align="center" bgcolor="#FFFFFF">
	<% else %>
	<tr class="gray" align="center" bgcolor="#FFFFFF">
	<% end if %>
		<td><a href="#" onclick="ViewOrderDetail(frmOnerder_<%= ix %>)" class="zzz"><%= ojumun.FMasterItemList(ix).FOrderSerial %></a></td>
		<td><%= ojumun.FMasterItemList(ix).FBuyName %></td>
		<td><%= ojumun.FMasterItemList(ix).FReqName %></td>
		<td><%= ojumun.FMasterItemList(ix).FItemID %></td>
		<td align="left">
			<%= ojumun.FMasterItemList(ix).FItemName %>
			<% if (ojumun.FMasterItemList(ix).FItemOptionStr<>"") then %>
				<font color="blue">[<%= ojumun.FMasterItemList(ix).FItemOptionStr %>]</font>
			<% end if %>
		</td>
		<td>
			<% if CStr(ojumun.FMasterItemList(ix).FItemNo)<>"1" then %>
			<font color="red"><%= ojumun.FMasterItemList(ix).FItemNo %></font>
			<% else %>
			<%= ojumun.FMasterItemList(ix).FItemNo %>
			<% end if %>
		</td>
		<td align="right"><%= Formatnumber(ojumun.FMasterItemList(ix).Fitemcost,0) %></td>
		<td align="right"><%= Formatnumber(ojumun.FMasterItemList(ix).Fbuycash,0) %></td>
<!--
		<td>
			<% if ojumun.FMasterItemList(ix).Fjumundiv = "9" then %>
	        <font color="red">마이너스</font>
			<% else %>
			<%= ojumun.FMasterItemList(ix).JumunMethodName %>
			<% end if %>
		</td>
-->
<!--	<td><font color="<%= ojumun.FMasterItemList(ix).IpkumDivColor %>"><%= ojumun.FMasterItemList(ix).IpkumDivName %></font></td>	-->
		<td><acronym title="<%= ojumun.FMasterItemList(ix).FRegdate %>"><%= left(ojumun.FMasterItemList(ix).FRegdate,10) %></acronym></td>
		<td><acronym title="<%= ojumun.FMasterItemList(ix).FIpkumdate %>"><%= left(ojumun.FMasterItemList(ix).FIpkumdate,10) %></acronym></td>
		<td><acronym title="<%= ojumun.FMasterItemList(ix).FUpcheBaesongDate %>"><%= left(ojumun.FMasterItemList(ix).FUpcheBaesongDate,10) %></acronym></td>

		<td>
			<% if ojumun.FMasterItemList(ix).FIsUpcheBeasong="Y" then %>
			<font color="#22AA22">업체배송</font>
			<% else %>
			텐바이텐
			<% end if %>
		</td>

		<td>
			<% if ojumun.FMasterItemList(ix).FJumunDiv="9" then %>
				<font color="red">마이너스</font>
			<% else %>
				<font color="<%= ojumun.FMasterItemList(ix).UpCheDeliverStateColor %>"><%= ojumun.FMasterItemList(ix).NormalUpcheDeliverState %></font>
			<% end if %>
		</td>
	</tr>
	</form>
	<% next %>

	<tr height="25" bgcolor="FFFFFF">
		<td colspan="15" align="center">
			<% if ojumun.HasPreScroll then %>
				<a href="javascript:NextPage('<%= ojumun.StartScrollPage-1 %>')">[pre]</a>
			<% else %>
				[pre]
			<% end if %>
			<% for ix=0 + ojumun.StartScrollPage to ojumun.StartScrollPage + ojumun.FScrollCount - 1 %>
				<% if (ix > ojumun.FTotalpage) then Exit for %>
				<% if CStr(ix) = CStr(ojumun.FCurrPage) then %>
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

<% end if %>
</table>


<%
set ojumun = Nothing
%>
<!-- #include virtual="/designer/lib/designerbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->