<%@ language=vbscript %>
<% option explicit %>
<%
'#######################################################
' Description : 판매내역[특정상품]
' History	:  서동석 생성
'              2022.09.19 한용민 수정(보안 취약부분 수정, 쿼리 클래스로 분리, 엑셀 다운로드 추가)
'#######################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/items/new_itemcls.asp"-->
<!-- #include virtual="/lib/classes/maechul/managementSupport/maechulCls.asp" -->
<%
dim itemid, itemoption, itemstate, sitename, inccancel, yyyy1,yyyy2,mm1,mm2,dd1,dd2, nowdate,oldlist
dim a1010,w1010,m1010,w10102,a10102,m10102, premonthdate, datetype, sortType, RowArr, oItemOrder
dim fromDate,toDate, page, RowCount, jumuncnt, totno, i, oitem, oitemoption
dim itemnosum, sellprice, realsellprice, upchejungsanprice
    page = RequestCheckVar(getNumeric(trim(request("page"))),10)
	nowdate         = Left(CStr(now()),10)
	premonthdate    = DateAdd("d",-14,nowdate)
	itemid = requestCheckvar(getNumeric(trim(request("itemid"))),10)
	itemoption = requestCheckvar(request("itemoption"),10)
	itemstate = request("itemstate")
	oldlist = request("oldlist")
	yyyy1   = requestCheckvar(getNumeric(trim(request("yyyy1"))),4)
	mm1     = requestCheckvar(getNumeric(trim(request("mm1"))),2)
	dd1     = requestCheckvar(getNumeric(trim(request("dd1"))),2)
	yyyy2   = requestCheckvar(getNumeric(trim(request("yyyy2"))),4)
	mm2     = requestCheckvar(getNumeric(trim(request("mm2"))),2)
	dd2     = requestCheckvar(getNumeric(trim(request("dd2"))),2)
	datetype = request("datetype")
	sitename = requestCheckvar(request("sitename"),32)
	inccancel = requestCheckvar(request("inccancel"),1)
	a1010 = requestCheckvar(request("a1010"),10)
	w1010 = requestCheckvar(request("w1010"),1)
	m1010 = requestCheckvar(request("m1010"),10)
	sortType = requestCheckvar(request("sortType"),2)

if sortType="" then sortType="od"
if (itemstate="5") then itemstate="6"
if (yyyy1="") then
	yyyy1 = Left(premonthdate,4)
	mm1   = Mid(premonthdate,6,2)
	dd1   = Mid(premonthdate,9,2)

	nowdate = Left(CStr(now()),10)
	yyyy2 = Left(nowdate,4)
	mm2   = Mid(nowdate,6,2)
	dd2   = Mid(nowdate,9,2)
else
	nowdate = Left(CStr(DateSerial(yyyy1 , mm1 , dd1)),10)
	yyyy1 = Left(nowdate,4)
	mm1   = Mid(nowdate,6,2)
	dd1   = Mid(nowdate,9,2)
end if
if (page="") then page=1
fromDate = CStr(DateSerial(yyyy1, mm1, dd1))
toDate = CStr(DateSerial(yyyy2, mm2, dd2+1))

if (datetype="") then datetype="reg"

if w1010 <> "" or m1010 <> "" or a1010 <> "" then
	if w1010="Y" then
		w10102=""
	else
		w10102="N"
	end if
	if m1010="" then
		m10102="N"
	else
		m10102=m1010
	end if
	if a1010="" then
		a10102="N"
	else
		a10102=a1010
	end if
end if

'상품코드 유효성 검사(2008.08.05;허진원)
if itemid<>"" then
	if Not(isNumeric(itemid)) then
		Response.Write "<script type='text/javascript'>alert('[" & itemid & "]은(는) 유효한 상품코드가 아닙니다.');history.back();</script>"
		dbget.close()	:	response.End
	end if
end if

set oItemOrder = new cManagementSupportMaechul_list
	oItemOrder.FCurrPage = page
	oItemOrder.FPageSize = 2000
	oItemOrder.FRectStartDate = fromDate
	oItemOrder.FRectEndDate   = toDate
	oItemOrder.frectdatetype=datetype
	oItemOrder.frectinccancel=inccancel
	oItemOrder.frectitemoption=itemoption
	oItemOrder.frectitemstate=itemstate
	oItemOrder.frectsitename=sitename
	oItemOrder.frectw10102=w10102
	oItemOrder.frectm10102=m10102
	oItemOrder.frecta10102=a10102

	if itemid<>"" and not(isnull(itemid)) then
		oItemOrder.GetOneItemOrderListNotPaging
	end if

if oItemOrder.FTotalCount>0 then
    RowArr=oItemOrder.fArrLIst
end if

RowCount = 0
jumuncnt = 0
if IsArray(RowArr) then
    RowCount = Ubound(RowArr,2)
    jumuncnt = RowCount + 1
end if

totno = 0

set oitem = new CItemInfo
oitem.FRectItemID = itemid

if itemid<>"" then
	oitem.GetOneItemInfo
end if

set oitemoption = new CItemOption
oitemoption.FRectItemID = itemid
if itemid<>"" then
	oitemoption.GetItemOptionInfo
end if

%>
<link rel="stylesheet" type="text/css" href="/css/adminPartnerDefault.css">
<link rel="stylesheet" type="text/css" href="/css/adminPartnerCommon.css">
<script type='text/javascript'>

function submitfrm(page){
	document.frm.target = "";
	document.frm.action = "";
    document.frm.page.value=page;
    document.frm.submit();
}

function PopItemSellEdit(iitemid){
	var popwin = window.open('/admin/lib/popitemsellinfo.asp?itemid=' + iitemid,'itemselledit','width=500,height=600,scrollbars=yes,resizable=yes')
	popwin.focus();
}

function popOrderDetailEdit(idx){
	var popwin = window.open('/common/orderdetailedit_UTF8.asp?idx=' + idx,'orderdetailedit','width=600,height=480,scrollbars=yes,resizable=yes');
	popwin.focus();
}

function chgSortType(st) {
	document.frm.sortType.value=st;
	document.frm.submit();
}

function downloadexcel(){
	document.frm.target = "view";
	document.frm.action = "/admin/ordermaster/oneitembuylist_excel.asp";
	document.frm.submit();
	document.frm.target = "";
	document.frm.action = "";
}

</script>
<div class="wrap">
	<div class="container">
		<div class="content scrl" style="top:0;">
			<!-- 검색 시작 -->
			<div class="searchWrap">
				<form name="frm" method="get" action="" style="margin:0px;">
				<input type="hidden" name="page" value="1">
				<input type="hidden" name="menupos" value="<%= menupos %>">
				<input type="hidden" name="sortType" value="<%=sortType%>">
				<div class="search rowSum1">
					<ul>
						<li>
							<label class="formTit" for="itemid">아이템ID :</label>
							<input type="text" class="formTxt" name="itemid" id="itemid" style="width:130px" value="<%= itemid %>" maxlength="16" placeholder="상품코드 검색" />
						</li>
						<% if oitemoption.FResultCount>0 then %>
						<li>
							<label class="formTit" for="itemoption">옵션선택 :</label>
							<select class="formSlt" name="itemoption" id="itemoption" title="상품옵션 선택">
								<option  value="">----
								<% for i=0 to oitemoption.FResultCount-1 %>
								<option value="<%= oitemoption.FITemList(i).FItemOption %>" <% if itemoption=oitemoption.FITemList(i).FItemOption then response.write "selected" %> ><%= oitemoption.FITemList(i).FOptionName %>
								<% next %>
							</select>
						</li>
						<% end if %>
						<li>
							<label class="formTit" for="datetype">검색기간 :</label>
							<select class="formSlt" name="datetype" id="datetype">
								<option value="reg" <%= chkIIF(datetype="reg","selected","") %> >주문일
								<option value="ipkum" <%= chkIIF(datetype="ipkum","selected","") %> >결제일
								<option value="beasong" <%= chkIIF(datetype="beasong","selected","") %> >출고일
							</select>
							<% DrawDateBox yyyy1,yyyy2,mm1,mm2,dd1,dd2 %>
							<span class="lMar10">
								<input type="checkbox" name="oldlist" id="oldlist" class="formCheck" <% if oldlist="on" then response.write "checked" %> />
								<label for="oldlist">6개월이전내역</label>
							</span>
						</li>
					</ul>
				</div>
				<dfn class="line"></dfn>
				<div class="search">
					<ul>
						<li>
							<label class="formTit" for="itemstate">주문상태 :</label>
							<select class="formSlt" name="itemstate" id="itemstate">
								<option value="availall" <% if itemstate="availall" then response.write "selected" %>>정상건 전체
								<option value="ipkumfinishall" <% if itemstate="ipkumfinishall" then response.write "selected" %>>결제완료이상
								<option value="2" <% if itemstate="2" then response.write "selected" %>>주문접수
								<option value="4" <% if itemstate="4" then response.write "selected" %>>결제완료
								<option value="6" <% if itemstate="6" then response.write "selected" %>>상품준비
								<option value="8" <% if itemstate="8" then response.write "selected" %>>출고완료
								<option value="9" <% if itemstate="9" then response.write "selected" %>>마이너스
							</select>
							(최대 2000건 까지만 검색됩니다.)
						</li>
						<li>
							<label class="formTit" for="sitename">사이트 :</label>
							<% Drawsitename "sitename",sitename %>
						</li>
						<li>
							<p class="formTit">추가 :</p>
							<span class="rMar10">
								<input type="checkbox" name="inccancel" id="inccancel" class="formCheck" value="Y" <%= CHKIIF(inccancel="Y", "checked", "") %> />
								<label for="inccancel">취소내역 포함</label>
							</span>
							<span class="rMar10">
								<input type="checkbox" name="w1010" id="w1010" class="formCheck" value="Y" <%= CHKIIF(w1010="Y", "checked", "") %> />
								<label for="w1010">10x10 Web</label>
							</span>
							<span class="rMar10">
								<input type="checkbox" name="m1010" id="m1010" class="formCheck" value="mobile" <%= CHKIIF(m1010="mobile", "checked", "") %> />
								<label for="m1010">10x10 Mobile</label>
							</span>
							<span>
								<input type="checkbox" name="a1010" id="a1010" class="formCheck" value="app_wish2" <%= CHKIIF(a1010="app_wish2", "checked", "") %> />
								<label for="a1010">10x10 App</label>
							</span>
						</li>
					</ul>
				</div>
				<input type="button" class="schBtn" onClick="submitfrm('1');" value="검색" />
				</form>
			</div>
			<!-- 검색 끝 -->

			<!-- 상품정보 시작 -->
			<div class="cont tMar10">
				<div style="padding:0 10px;">
				<% if oitem.FResultCount>0 then %>
				<table class="tbType1 listTb">
				<tbody>
					<tr>
						<td rowspan=<%= 5 + oitemoption.FResultCount -1 %> width="110" valign=top align=center><img src="<%= oitem.FOneItem.FListImage %>" width="100" height="100"></td>
						<th width="60">상품코드</th>
						<td>
							10 <b><%= CHKIIF(oitem.FOneItem.FItemID>=1000000,Format00(8,oitem.FOneItem.FItemID),Format00(6,oitem.FOneItem.FItemID)) %></b> <%= itemoption %>
							&nbsp;
							<!--
							<input type="button" value="수정" onclick="PopItemSellEdit('<%= itemid %>');">
							-->
						</td>
						<th width="60">전시여부</th>
						<td colspan=2><font color="<%= ynColor(oitem.FOneItem.FDispyn) %>"><%= oitem.FOneItem.FDispyn %></font></td>
					</tr>
					<tr>
						<th>브랜드ID</th>
						<td><%= oitem.FOneItem.FMakerid %></td>
						<th>판매여부</th>
						<td colspan=2><font color="<%= ynColor(oitem.FOneItem.FSellyn) %>"><%= oitem.FOneItem.FSellyn %></font></td>
					</tr>
					<tr>
						<th>상품명</th>
						<td><%= oitem.FOneItem.FItemName %></td>
						<th>사용여부</th>
						<td colspan=2><font color="<%= ynColor(oitem.FOneItem.FIsUsing) %>"><%= oitem.FOneItem.FIsUsing %></font></td>
					</tr>
					<tr>
						<th>판매가</th>
						<td>
							<%= FormatNumber(oitem.FOneItem.FSellcash,0) %> / <%= FormatNumber(oitem.FOneItem.FBuycash,0) %>
							&nbsp;&nbsp;
							<font color="<%= oitem.FOneItem.getMwDivColor %>"><%= oitem.FOneItem.getMwDivName %></font>
							<% if oitem.FOneItem.FSellcash<>0 then %>
							<%= CLng((1- oitem.FOneItem.FBuycash/oitem.FOneItem.FSellcash)*100) %> %
							<% end if %>
							&nbsp;&nbsp;
							<!-- 할인여부/쿠폰적용여부 -->
							<% if (oitem.FOneItem.FSailYn="Y") then %>
								<font color=red>
								<% if (oitem.FOneItem.Forgprice<>0) then %>
									<%= CLng((oitem.FOneItem.Forgprice-oitem.FOneItem.Fsellcash)/oitem.FOneItem.Forgprice*100) %> %
								<% end if %>
								할인
								</font>
							<% end if %>

							<% if (oitem.FOneItem.Fitemcouponyn="Y") then %>

								<font color=green><%= oitem.FOneItem.GetCouponDiscountStr %> 쿠폰
								(<%= FormatNumber(oitem.FOneItem.GetCouponAssignPrice,0) %>)</font>
							<% end if %>
						</td>
						<th>단종여부</th>
						<td colspan=2>
							<% if oitem.FOneItem.Fdanjongyn="Y" then %>
							<font color="#33CC33">단종</font>
							<% elseif oitem.FOneItem.Fdanjongyn="S" then %>
							<font color="#33CC33">일시품절</font>
							<% else %>
							생산중
							<% end if %>
						</td>
					</tr>

					<% if oitemoption.FResultCount>1 then %>
						<!-- 옵션이 있는경우 -->
						<% for i=0 to oitemoption.FResultCount -1 %>
							<% if oitemoption.FITemList(i).FOptIsUsing<>"Y" then %>
							<tr>
								<th><font color="#AAAAAA">옵션명 :</font></th>
								<td><font color="#AAAAAA"><%= oitemoption.FITemList(i).FOptionName %></font></td>
								<th><font color="#AAAAAA">한정여부 : </font></th>
								<td><font color="#AAAAAA"><font color="<%= ynColor(oitemoption.FITemList(i).Foptlimityn) %>"><%= oitemoption.FITemList(i).Foptlimityn %></font> (<%= oitemoption.FITemList(i).GetOptLimitEa %>)</font></td>
								<td>한정 비교재고 (<b><%= oitemoption.FITemList(i).GetLimitStockNo %></b>)</td>
							</tr>
							<% else %>

							<% if oitemoption.FITemList(i).Fitemoption=itemoption then %>
							<tr>
							<% else %>
							<tr>
							<% end if %>
								<th>옵션명</th>
								<td><%= oitemoption.FITemList(i).FOptionName %></td>
								<th>한정여부</th>
								<td><font color="<%= ynColor(oitemoption.FITemList(i).Foptlimityn) %>"><%= oitemoption.FITemList(i).Foptlimityn %></font> (<%= oitemoption.FITemList(i).GetOptLimitEa %>)</td>
								<td>한정 비교재고 (<b><%= oitemoption.FITemList(i).GetLimitStockNo %></b>)</td>
							</tr>
							<% end if %>
						<% next %>
					<% else %>
						<tr>
							<th>옵션명</th>
							<td>-</td>
							<th>한정여부</th>
							<td><font color="<%= ynColor(oitem.FOneItem.Flimityn) %>"><%= oitem.FOneItem.Flimityn %> (<%= oitem.FOneItem.GetLimitEa %>)</font></td>
							<td>한정 비교재고 (<b><%= oitem.FOneItem.GetLimitStockNo %></b>)</td>
						</tr>
					<% end if %>
				</tbody>
				</table>
				<% end if %>
				</div>
			</div>
			<!-- 상품정보 끝 -->

            <p class="pad10">
                * <font color="red">3PL 에서 출고된 수량</font> 표시안되어 있습니다.
            </p>

			<!-- 주문목록 시작 -->
			<div class="cont">
				<div class="pad10">
					<div class="panel1 rt pad10">
						<span><input type="button" onclick="downloadexcel();" value="엑셀다운로드" class="button"></span>
					</div>
					<div class="panel1 rt pad10">
						<span id="totDisp"></span>
					</div>

					<table class="tbType1 listTb">
					<thead>
						<tr>
							<th><%=pointUpDown("주문번호","o",left(sortType,1)="o",right(sortType,1)="d")%></th>
							<th>구분</th>
							<th>매입<br>구분</th>
							<th>과세<br>구분</th>
							<th>Site</th>
							<th><%=pointUpDown("rdSite","r",left(sortType,1)="r",right(sortType,1)="d")%></th>
							<th>주문상태</th>
							<th>상품상태</th>
							<th><%=pointUpDown("수량","c",left(sortType,1)="c",right(sortType,1)="d")%></th>
							<th>옵션명</th>
							<th>옵션코드</th>
							<th>회원ID</th>
							<th><%=pointUpDown("회원등급","l",left(sortType,1)="l",right(sortType,1)="d")%></th>
							<% if (FALSE) then %>
								<th>구매자</th>
							<% end if %>
							<th>수령인</th>
							<th>판매가</th>
							<% if (C_InspectorUser = False) then %>
								<th>실판매가<br>(쿠폰적용)</th>
							<% end if %>
							<th>업체정산액</th>
							<th>주문일</th>
							<th>출고일</th>
							<th>배송일</th>
							<th>정산일</th>
							<th>결제수단</th>
						</tr>
					</thead>
					<tbody>
				<%
					itemnosum = 0
					sellprice = 0
					realsellprice = 0
					upchejungsanprice = 0

					if IsArray(RowArr) then
						for i=0 to RowCount
						itemnosum = itemnosum + RowArr(2,i)
						sellprice = sellprice + RowArr(17,i)
						realsellprice = realsellprice + RowArr(18,i)
						upchejungsanprice = upchejungsanprice + RowArr(19,i)
				%>
						<tr align="center" bgcolor="<%= CHKIIF(RowArr(27,i)="N", "#FFFFFF", "#EEEEEE") %>">
							<td><%= RowArr(0,i) %></td>
							<td><%= getJumundivName(RowArr(15,i)) %></td>
							<td><a href="javascript:popOrderDetailEdit(<%=RowArr(20,i)%>)"><%= (RowArr(16,i)) %></a></td>
							<td><%= RowArr(24,i) %></td>
							<td><%= RowArr(12,i) %></td>
							<td><%= RowArr(22,i) %></td>
							<td><font color="<%= IpkumDivColor(RowArr(1,i)) %>"><%= IpkumDivName(RowArr(1,i)) %></font></td>
							<td><font color="<%= getCurrstateNameColor(RowArr(1,i),RowArr(11,i)) %>"><%= getCurrstateName(RowArr(1,i),RowArr(11,i)) %></font></td>
							<td><%= RowArr(2,i) %></td>
							<td><%= DdotFormat(RowArr(10,i),20) %></td>
							<td><%= RowArr(23,i) %></td>
							<td>
								<% if C_CriticInfoUserLV1 or C_CriticInfoUserLV2 or C_CriticInfoUserLV3 then %>
									<%= RowArr(14,i) %>
								<% else %>
									<%= printUserId(RowArr(14,i),2,"*") %>
								<% end if %>
							</td>
							<td>
								<font color="<%= getUserLevelColorByDate(RowArr(25,i), left(RowArr(21,i),10)) %>">
									<%= getUserLevelStrByDate(RowArr(25,i), left(RowArr(21,i),10)) %>
								</font>
							</td>
							<% if (FALSE) then %>
							<td><%= RowArr(3,i) %></td>
							<% end if %>
							<td><%= RowArr(7,i) %></td>
							<% if (C_InspectorUser = False) then %>
							<td><%= FormatNumber(RowArr(17,i),0) %></td>
							<% end if %>
							<td><%= FormatNumber(RowArr(18,i),0) %></td>
							<td><%= FormatNumber(RowArr(19,i),0) %></td>
							<td><%= RowArr(21,i) %></td>
							<td><%= RowArr(13,i) %></td>
							<td><%= RowArr(28,i) %></td>
							<td><%= RowArr(29,i) %></td>
							<td><%= GetaccountdivName(RowArr(26,i)) %></td>
						</tr>
					<%
								totno = totno + RowArr(2,i)
						if i mod 1000 = 0 then
							Response.Flush		' 버퍼리플래쉬
						end if
						next
					%>
						<tr align="center" bgcolor="#FFFFFF">
							<td colspan=8>총액</td>
							<td><%= FormatNumber(itemnosum,0) %></td>
							<td colspan=5>&nbsp;</td>

							<% if (C_InspectorUser = False) then %>
								<td><%= FormatNumber(sellprice,0) %></td>
							<% end if %>

							<td><%= FormatNumber(realsellprice,0) %></td>
							<td><%= FormatNumber(upchejungsanprice,0) %></td>
							<td colspan=5>&nbsp;</td>
						</tr>
					<% end if %>
						<tr height="26">
							<td align="right" colspan="22">총상품수 <%= totno %> 개 / 총주문건수 <%= jumuncnt %> 건</td>
						</tr>
						<tbody>
					</table>
				</div>
			</div>
			<% IF application("Svr_Info")="Dev" THEN %>
				<iframe id="view" name="view" src="" width="100%" height=300 frameborder="0" scrolling="no"></iframe>
			<% else %>
				<iframe id="view" name="view" src="" width=0 height=0 frameborder="0" scrolling="no"></iframe>
			<% end if %>
		</div>
	</div>
</div>

<%
function IpkumDivName(byval v )
	if v="0" then
		IpkumDivName="주문대기"
	elseif v="1" then
		IpkumDivName="주문실패"
	elseif v="2" then
		IpkumDivName="주문접수"
	elseif v="3" then
		IpkumDivName="주문접수"
	elseif v="4" then
		IpkumDivName="결제완료"
	elseif v="5" then
		IpkumDivName="주문통보"
	elseif v="6" then
		IpkumDivName="상품준비"
	elseif v="7" then
		IpkumDivName="일부출고"
	elseif v="8" then
		IpkumDivName="출고완료"
	elseif v="9" then
		IpkumDivName="마이너스"
	end if
end function

function getCurrstateName(byval v1, byval v)
    if (v=0) then
        if (v1>3) and (v1<8) then
           getCurrstateName = "결제완료"
        else
            getCurrstateName = IpkumDivName(v1)
        end if
    else
        if v=2 then
            getCurrstateName = "주문통보"
        elseif v=3 then
            getCurrstateName = "상품준비"
        elseif v=7 then
            getCurrstateName = "출고완료"
        else
            getCurrstateName = v
        end if
    end if
end function

function getCurrstateNameColor(byval v1, byval v)
    if (v=0) then
        if (v1>3) and (v1<8) then
            getCurrstateNameColor = IpkumDivColor(4)
        else
            getCurrstateNameColor = IpkumDivName(v1)
        end if
    else
        if v=2 then
            getCurrstateNameColor = IpkumDivColor(v)
        elseif v=3 then
            getCurrstateNameColor = IpkumDivColor(v)
        elseif v=7 then
            getCurrstateNameColor = IpkumDivColor(v)
        else
            getCurrstateNameColor = "#000000"
        end if
    end if
end function

function getJumundivName(byval ijumundiv)
    if (isNULL(ijumundiv)) then
        getJumundivName = ""
        Exit function
    end if

    if ijumundiv="1" then
		getJumundivName="출고"
	elseif ijumundiv="5" then
	    getJumundivName="출고"
    elseif ijumundiv="9" then
        getJumundivName="<font color='red'>반품</font>"
    elseif ijumundiv="6" then
        getJumundivName="<font color='blue'>교환</font>"
    else
        getJumundivName=ijumundiv
    end if
end function

Function pointUpDown(txt,tp,sw,ud)
	dim ret, st
	st = tp & chkIIF(sw and ud,"a","d")
	ret = "<div class=""sorting"" style=""" & chkIIF(sw,"font-weight:bold;","") & """ onClick=""chgSortType('" & st & "')"">"
	ret = ret & txt
	ret = ret & "<span class=""" & chkIIF(sw and ud,"sortWay","") & """></span>"
	ret = ret & "</div>"
	pointUpDown = ret
end function

set oitem = Nothing
set oitemoption = Nothing
%>
<script type='text/javascript'>

window.onload = function(e) {
	var totDisp = document.getElementById('totDisp');
	totDisp.innerHTML = '총상품수 <%= totno %> 개 / 총주문건수 <%= jumuncnt %> 건';
}

</script>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
