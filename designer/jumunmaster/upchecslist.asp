<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/designer/incSessionDesigner.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/designer/lib/designerbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/classes/cscenter/cs_aslistcls.asp"-->
<%

dim currstate,research,page,orderserial
dim searchfield, searchstring, divcd
dim showAX12, excludeB006yn, excludeA004yn, excludeOLDCSyn
dim receiveyn
dim itemid

research    = requestCheckVar(request("research"),10)
currstate   = requestCheckVar(request("currstate"),10)
page        = requestCheckVar(request("page"),10)

searchfield = requestCheckVar(request("searchfield"),32)
searchstring = requestCheckVar(request("searchstring"),32)
divcd       = requestCheckVar(request("divcd"),10)

showAX12	= requestCheckVar(request("showAX12"),10)
receiveyn	= requestCheckVar(request("receiveyn"),10)
excludeB006yn	= requestCheckVar(request("excludeB006yn"),10)
excludeA004yn	= requestCheckVar(request("excludeA004yn"),10)
excludeOLDCSyn	= requestCheckVar(request("excludeOLDCSyn"),10)
itemid		= requestCheckVar(request("itemid"),64)

if page="" then page=1

if research="" then
	currstate="notfinish"
	excludeOLDCSyn = "Y"
end if

if searchstring="" then searchfield="" end if
if searchfield="" then searchstring="" end if

if (searchfield = "") or (searchstring = "") then
	itemid = ""
end if

dim ioneas,i
set ioneas = new CCSASList
ioneas.FPageSize = 20
ioneas.FCurrPage = page

if searchfield="01" then
	ioneas.FRectOrderserial = searchstring
elseif searchfield="02" then
	ioneas.FRectUserName = searchstring
elseif searchfield="03" then
	ioneas.FRectUserID = searchstring
end if


ioneas.FRectCurrstate  = currstate
ioneas.FRectDivcd = divcd
ioneas.FRectMakerID = session("ssBctID")
ioneas.FRectShowAX12 = showAX12
ioneas.FRectReceiveYN = receiveyn
ioneas.FRectExcludeB006YN = excludeB006yn
ioneas.FRectExcludeA004YN = excludeA004yn
ioneas.FRectExcludeOLDCSYN = excludeOLDCSyn
ioneas.FRectItemID = itemid

ioneas.GetCSASMasterListUpcheNew


Public Function GetAsDivCDString(divcd)
    if (divcd = "A000") or (divcd = "A100") or (divcd = "A001") or (divcd = "A002") then
    	GetAsDivCDString = "<font color=blue>출고</font>"
    elseif (divcd = "A004") then
    	GetAsDivCDString = "<font color=red>반품</font>"
    else
    	GetAsDivCDString = "기타"
    end if
end function

%>

<script language='javascript'>

function ShowOrderInfo(frm,orderserial){
    var props = "width=600, height=600, location=no, status=yes, resizable=no, scrollbars=yes";
	window.open("about:blank", "orderdetail", props);
    frm.target = "orderdetail";
    frm.orderserial.value = orderserial;
    frm.action="/designer/common/viewordermaster.asp";
	frm.submit();
}

function reSearchByOrderserial(frm, orderserial) {
	frm.searchfield.value = "01";
	frm.searchstring.value = orderserial;

	frm.submit();
}


function popOpenJupsuItems() {
    var props = "width=1000, height=600, location=no, status=yes, resizable=yes, scrollbars=yes";
	var w = window.open("upchecs_jupsu_item_list.asp", "popOpenJupsuItems", props);

	w.focus();
}

function NextPage(page){
    frm.page.value = page;
    frm.submit();
}

function SubmitSearch(frm) {
    if (frm.itemid.value != "") {
    	if ((frm.searchfield.value == "") || (frm.searchstring.value == "")) {
    		alert("기타 검색조건을 입력하셔야 상품코드 검색이 가능합니다.");
    		return;
    	} else if (frm.itemid.value*0 != 0) {
    		alert("상품코드는 숫자만 가능합니다.");
    		return;
    	}
    }
    frm.submit();
}
</script>

<!-- 검색 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frm" method="get" >
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<input type="hidden" name="research" value="T">
	<input type="hidden" name="page" value="">
	<tr align="center" bgcolor="<%= adminColor("topbar") %>" >
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
		<td align="left">
			상태:
			<select class="select" name="currstate">
	     	<option value='' selected>전체</option>
	     	<option value='notfinish' <% if currstate="notfinish" then response.write "selected" %>>미처리</option>
	     	<option value='B007' <% if currstate="B007" then response.write "selected" %>>처리완료</option>
	     	</select>
			&nbsp;
			기타검색:
			<select class="select" name="searchfield">
				<option value="">검색조건</option>
				<option value="01" <% if searchfield="01" then response.write "selected" %>>주문번호</option>
				<option value="02" <% if searchfield="02" then response.write "selected" %>>고객명</option>
				<option value="03" <% if searchfield="03" then response.write "selected" %>>고객ID</option>
			</select>
			<input type="text" class="text" name="searchstring" value="<%= searchstring %>" size="13" maxlength="11">
			&nbsp;
			상품코드:
			<input type="text" class="text" name="itemid" value="<%= itemid %>" size="10" maxlength="16">
			&nbsp;
			접수구분:
			<select class="select" name="divcd">
			    <option value="">전체</option>
			    <option value="A000" <% if divcd="A000" then response.write "selected" %>>맞교환출고</option>
			    <option value="A100" <% if divcd="A100" then response.write "selected" %>>상품변경 맞교환출고</option>
			    <option value="A001" <% if divcd="A001" then response.write "selected" %>>누락재발송</option>
			    <option value="A002" <% if divcd="A002" then response.write "selected" %>>서비스발송</option>
			    <option value="A004" <% if divcd="A004" then response.write "selected" %>>반품접수</option>
			    <option value="A006" <% if divcd="A006" then response.write "selected" %>>출고시유의사항</option>
			    <!--
			    <option value="A012" <% if divcd="A012" then response.write "selected" %>>맞교환반품</option>
			    <option value="A112" <% if divcd="A112" then response.write "selected" %>>상품변경 맞교환반품</option>
			    -->
			</select>
			&nbsp;
			맞교환회수상태:
			<select class="select" name="receiveyn">
			    <option value="">전체</option>
			    <option value="N" <% if receiveyn="N" then response.write "selected" %>>회수이전</option>
			    <option value="Y" <% if receiveyn="Y" then response.write "selected" %>>회수완료</option>
			</select>

			<!--
			<input type="checkbox" name="showAX12" value="Y" <% if (showAX12 = "Y") then %>checked<% end if %>> 맞교환반품 포함
			-->
		</td>

		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button_s" value="검색" onClick="SubmitSearch(frm)">
		</td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("topbar") %>" >
		<td align="left">
			<input type="checkbox" name="excludeB006yn" value="Y" <% if (excludeB006yn = "Y") then %>checked<% end if %>> 업체처리완료 건 제외
			<input type="checkbox" name="excludeA004yn" value="Y" <% if (excludeA004yn = "Y") then %>checked<% end if %>> 반품 건 제외
			<input type="checkbox" name="excludeOLDCSyn" value="Y" <% if (excludeOLDCSyn = "Y") then %>checked<% end if %>> 장기간(3개월) 미처리건 제외
		</td>
	</tr>
	</form>
</table>
<!-- 검색 끝 -->

<p>

<input type="button" class="button_s" value="CS접수상품조회" onClick="popOpenJupsuItems()">
* 원주문번호에 "+" 표시가 있는 경우, 교환주문에 대한 CS 입니다.

<p>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="15">
			검색결과 : <b><% = ioneas.FTotalCount %></b>
			&nbsp;
			페이지 : <b><%= page %> / <%= ioneas.FTotalPage %></b>
		</td>
	</tr>
    <tr align="center" bgcolor="<%= adminColor("tabletop") %>" height="25">
		<td width="40">구분</td>
		<td width="80">원주문번호</td>
		<td>고객명</td>
		<td>고객ID</td>
		<td>접수구분</td>
		<td>제목</td>
		<td>접수사유</td>
		<td>등록일</td>
		<td>처리완료일</td>
		<td>처리상태</td>
		<td>관련회수</td>
	</tr>
	<% for i=0 to ioneas.FresultCount-1 %>

		<%
		'// 고객 접수 강조
		ioneas.FItemList(i).FTitle = Replace(ioneas.FItemList(i).FTitle, "고객접수", "고객 직접 접수")
		ioneas.FItemList(i).FTitle = Replace(ioneas.FItemList(i).FTitle, "고객 직접 접수", "<font color=red>고객 등록</font>")
		%>

	<tr align="center" bgcolor="#FFFFFF" height="25">

		<td>
			<%= GetAsDivCDString(ioneas.FItemList(i).Fdivcd) %>
		</td>

		<td>
			<a href="upchecsdetail.asp?idx=<%= ioneas.FItemList(i).Fid %>&menupos=<%= menupos %>"><%= ioneas.FItemList(i).FOrgOrderSerial %></a>
    		<% if (ioneas.FItemList(i).Forderserial <> ioneas.FItemList(i).Forgorderserial) then %>
    			+
    		<% end if %>
		</td>

		<td><%= ioneas.FItemList(i).FCustomerName %></td>
		<td><%= printUserId(ioneas.FItemList(i).FUserID,2,"**") %></td>
		<td><%= ioneas.FItemList(i).FdivcdName %></td>
		<td align="left" ><a href="upchecsdetail.asp?idx=<%= ioneas.FItemList(i).Fid %>&menupos=<%= menupos %>"><%= ioneas.FItemList(i).FTitle %></a></td>
		<td align="left"><%= (ioneas.FItemList(i).Fgubun01Name) %>&gt;&gt;<%= (ioneas.FItemList(i).Fgubun02Name) %></td>
		<td><%= Left(CStr(ioneas.FItemList(i).Fregdate),10) %></td>
		<td>
			<% if ioneas.FItemList(i).Ffinishdate<>"" then %>
			<%= Left(CStr(ioneas.FItemList(i).Ffinishdate),10) %>
			<% else %>
    			<% if ((ioneas.FItemList(i).Fdivcd = "A000") or (ioneas.FItemList(i).Fdivcd = "A100")) and (Not IsNull(ioneas.FItemList(i).Freceivestate)) then %>
    			<input type="button" class="button" value="출고처리" onclick="location.href='upchecsdetail.asp?idx=<%= ioneas.FItemList(i).Fid %>&menupos=<%= menupos %>';">
    			<% else %>
    			<input type="button" class="button" value="완료처리" onclick="location.href='upchecsdetail.asp?idx=<%= ioneas.FItemList(i).Fid %>&menupos=<%= menupos %>';">
    			<% end if %>
			<% end if %>
		</td>
		<td>
			<%= CsState2Name(ioneas.FItemList(i).FCurrstate) %>
		</td>
		<td>
			<% if (ioneas.FItemList(i).Fdivcd = "A000") or (ioneas.FItemList(i).Fdivcd = "A100") then %>
				<!-- 맞교환출고, 상품변경 맞교환출고 -->
				<% if (Not IsNull(ioneas.FItemList(i).Freceivestate)) then %>
					<% if (ioneas.FItemList(i).Freceivestate < "B006") then %>
						<input type="button" class="button" value="회수처리" onclick="location.href='upchecsdetail.asp?idx=<%= ioneas.FItemList(i).Fid %>&menupos=<%= menupos %>&receiveonly=Y';">
					<% else %>
						<a href="javascript:location.href='upchecsdetail.asp?idx=<%= ioneas.FItemList(i).Fid %>&menupos=<%= menupos %>&receiveonly=Y'"><%= Left(CStr(ioneas.FItemList(i).Freceivefinishdate),10) %></a>
					<% end if %>
				<% end if %>
			<% end if %>
		</td>
	</tr>
	<% next %>
    <tr height="25" bgcolor="FFFFFF">
		<td colspan="15" align="center">
	        <% if ioneas.HasPreScroll then %>
				<a href="javascript:NextPage('<%= CStr(ioneas.StarScrollPage - 1) %>')">[prev]</a>
			<% else %>
				[prev]
			<% end if %>
			<% for i = ioneas.StarScrollPage to (ioneas.StarScrollPage + ioneas.FScrollCount - 1) %>
			  <% if (i > ioneas.FTotalPage) then Exit For %>
			  <% if CStr(i) = CStr(ioneas.FCurrPage) then %>
				 <font color="red"><b>[<%= i %>]</b></font>
			  <% else %>
				 <a href="javascript:NextPage('<%= i %>')" class="id_link">[<%= i %>]</a>
			  <% end if %>
			<% next %>
			<% if ioneas.HasNextScroll then %>
				<a href="javascript:NextPage('<%= CStr(ioneas.StarScrollPage + ioneas.FScrollCount) %>')">[next]</a>
			<% else %>
				[next]
			<% end if %>
	    </td>
	</tr>
</table>

<form name="frmshow" method="post">
<input type="hidden" name="orderserial" value="">

</form>


<%
set ioneas = Nothing
%>
<!-- #include virtual="/designer/lib/designerbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
