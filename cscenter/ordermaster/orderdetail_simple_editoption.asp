<%@ language=vbscript %>
<% option explicit %>
<%
session.codePage = 949
Response.CharSet = "EUC-KR"
%>
<%
'###########################################################
' Description : cs센터 옵션교환
' History : 이상구 생성
'			2023.09.05 한용민 수정(6개월이전 주문도 교환 가능하게 처리)
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbHelper.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/order/jumuncls.asp"-->
<!-- #include virtual="/lib/classes/order/new_ordercls.asp"-->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/classes/cscenter/cs_couponcls.asp" -->
<!-- #include virtual="/lib/classes/cscenter/sp_itemcouponcls.asp" -->
<!-- #include virtual="/cscenter/lib/csOrderFunction.asp" -->
<%
dim i, idx, orderserial, result, ojumunDetail, ojumun, IsOrderCanceled, IsChangeOrder
dim itemoption, optionname, optsellyn, optlimityn, optlimitno, optlimitsold, issameoptaddprice, isusing, optaddprice
Dim sqlStr, rsOption, k, optionText, itemStatus, sqlsub, changedindex
dim prevregno, contents_jupsu, title, divcd, oupchebeasongpay, upchebeasongpay, isupchebeasong, requiremakerid
	idx = requestCheckVar(getNumeric(request("idx")),10)

set ojumunDetail = new CJumunMaster
ojumunDetail.SearchOneJumunDetail idx

if (ojumunDetail.FResultCount < 1) then
	ojumunDetail.FRectOldJumun = "on"
	ojumunDetail.SearchOneJumunDetail idx
end if

orderserial = ojumunDetail.FJumunDetail.FOrderSerial

set ojumun = new COrderMaster

if (orderserial <> "") then
    ojumun.FRectOrderSerial = orderserial
    ojumun.QuickSearchOrderMaster
end if

if (ojumun.FResultCount<1) and (Len(orderserial)=11) and (IsNumeric(orderserial)) then
    ojumun.FRectOldOrder = "on"
    ojumun.QuickSearchOrderMaster
end if

if ojumun.FTotalCount < 1 then
	response.write "해당되는 주문건이 없습니다."
	dbget.close() : response.end
end if

IsOrderCanceled = (ojumun.FOneItem.Fcancelyn = "Y")
IsChangeOrder   = (ojumun.FOneItem.FjumunDiv="6")

If ojumunDetail.FJumunDetail.Fitemoption <> "0000" Then
	'* 옵션변경은 <font color=red>옵션가</font>가 동일한 옵션상품만 가능합니다.<br>
	'* 주문당시 옵션가격에 상관없이 현재 상품정보 상의 옵션가격으로 비교합니다.<br>
	'* 상품할인정보(판매가,매입가 등)는 주문당시 정보가 복사됩니다.<br>
	' 주문후 사용안함 처리가 되어도 표시

	sqlsub = "select top 1 optaddprice "
	sqlsub = sqlsub + "from [db_item].[dbo].tbl_item_option "
	sqlsub = sqlsub + "where 1 = 1 "
	sqlsub = sqlsub + "and itemid = " & CStr(ojumunDetail.FJumunDetail.Fitemid) & " "
	sqlsub = sqlsub + "and itemoption = '" & CStr(ojumunDetail.FJumunDetail.Fitemoption) & "' "

	sqlStr = " select "
	sqlStr = sqlStr + " v.itemoption "
	sqlStr = sqlStr + " , v.optionname "
	sqlStr = sqlStr + " , v.optsellyn "
	sqlStr = sqlStr + " , v.optlimityn "
	sqlStr = sqlStr + " , v.optlimitno "
	sqlStr = sqlStr + " , v.optlimitsold "
	sqlStr = sqlStr + " , case when v.optaddprice=IsNULL((" & sqlsub & "),0) " & " then 'T' else 'F' end "
	sqlStr = sqlStr + " , v.isusing "
	sqlStr = sqlStr + " , IsNull(P.regno, 0) as prevregno "
	sqlStr = sqlStr + " from [db_item].[dbo].tbl_item i "
	sqlStr = sqlStr + " left join [db_item].[dbo].tbl_item_option v "
	sqlStr = sqlStr + " on i.itemid=v.itemid "

	'이전 CS반품내역(접수+완료내역, 반품사유고려안함)
	sqlStr = sqlStr + "		LEFT JOIN (" + VbCrlf
	sqlStr = sqlStr + "		    select d.itemid, d.itemoption, sum(confirmitemno) as regno, max(a.id) asId " + VbCrlf
    sqlStr = sqlStr + "		    from" + VbCrlf
    sqlStr = sqlStr + "		    	[db_cs].[dbo].tbl_new_as_list a" + VbCrlf
    sqlStr = sqlStr + "		    	Join [db_cs].[dbo].tbl_new_as_detail d" + VbCrlf
    sqlStr = sqlStr + "		    on a.id=d.masterid" + VbCrlf
    sqlStr = sqlStr + "		    where a.orderserial='" + CStr(orderserial) + "'" + VbCrlf
    sqlStr = sqlStr + "		    and a.divcd in ('A004','A010', 'A111', 'A112')" + VbCrlf                ''반품 / 회수 / 상품변경 맞교환회수(텐바이텐배송) / 상품변경 맞교환반품(업체배송).
    sqlStr = sqlStr + "		    and a.deleteyn='N'" + VbCrlf
    'sqlStr = sqlStr + "		    	and a.currstate='B007'" + VbCrlf					'접수+완료 모두 계산
    sqlStr = sqlStr + "			group by d.itemid, d.itemoption" + VbCrlf
    sqlStr = sqlStr + " ) P " + VbCrlf
    sqlStr = sqlStr + "     ON i.itemid=P.itemid and v.itemoption=P.itemoption" + VbCrlf

	sqlStr = sqlStr + " WHERE 1=1 "
	sqlStr = sqlStr + " and i.itemid=" & ojumunDetail.FJumunDetail.Fitemid & ""
	sqlStr = sqlStr + " order by i.itemid desc, v.itemoption"

	rsget.Open sqlStr,dbget,1
	If Not rsget.EOF Then
		rsOption = rsget.getrows
	End If
	rsget.close()

	'response.write sqlStr
End If

prevregno = 0

'// 기본문구 설정
if Not IsNull(session("ssBctCname")) then
	contents_jupsu = "텐바이텐 고객센터 " + CStr(session("ssBctCname")) + " 입니다"
end if

if (ojumunDetail.FJumunDetail.FcurrState = "7") or IsChangeOrder then

	'==============================================================================
	'출고이후

	isupchebeasong = ojumunDetail.FJumunDetail.Fisupchebeasong

	if (isupchebeasong = "Y") then
		requiremakerid = ojumunDetail.FJumunDetail.Fmakerid
	end if

	divCd = "A100"	' 상품변경 맞교환출고
	title = "교환출고(옵션변경)"

	'// 옵션변경 맞교환의 경우 기존반품수량
	For i = 0 To UBound(rsOption,2)
		itemoption = rsOption(0,i)
		prevregno = rsOption(8,i)

		if (ojumunDetail.FJumunDetail.Fitemoption = itemoption) then
			Exit For
		end if
	Next

	set oupchebeasongpay = new COrderMaster
	upchebeasongpay = getDefaultBeasongPayByDate(Left(Now, 10))		' 배송비

	if (orderserial <> "") and (isupchebeasong = "Y") then
		oupchebeasongpay.FRectOrderSerial = orderserial
		oupchebeasongpay.getUpcheBeasongPayList

		for i = 0 to oupchebeasongpay.FResultCount - 1
			if (oupchebeasongpay.FItemList(i).Fmakerid = requiremakerid) then
				'// 업체배송이면 업체 기본배송비 가져오기
				upchebeasongpay = oupchebeasongpay.FItemList(i).Fdefaultdeliverpay
			end if
		next

		if (upchebeasongpay = 0) then
			'// XXXX 업체무료배송이면 텐텐배송비로 설정
			'기본배송비 설정 않되어 있으면 2500원(since 2012-06-18)
			upchebeasongpay = 2500
		end if
	end if

else
	'==============================================================================
	' 출고 이전

	divCd = "A900"	' 주문내역변경
	title = "상품옵션변경"
end if

%>
<script type="text/javascript" src="/cscenter/js/cscenter.js"></script>
<script type="text/javascript" SRC="/js/ajax.js"></script>
<script type="text/javascript" SRC="/cscenter/js/newcsas.js"></script>
<script type='text/javascript'>

// 사유구분(ajax) 를 사용하기 위해 필요
var IsPossibleModifyCSMaster = true;
var IsPossibleModifyItemList = true;

// ============================================================================
// 옵션별 수량 자동조절(마이너스 입력불가)
// ============================================================================
function CheckItemOptionNo(changedindex) {
    var frm = document.frm;
    var i;

	var orgitemno = parseInt(frm.orgitemoptionno.value);
	var regitemno = 0;

	if (frm.itemoptionno[changedindex].value*1 < 0) {
		alert('수량에 마이너스를 입력할 수 없습니다.');
		return;
	}

	if ((frm.itemoptionno[changedindex].value.length < 1) || (frm.itemoptionno[changedindex].value*0 != 0)) {
		alert('수량에 숫자를 입력하세요.');
		return;
	}

	for (i = 1; i < parseInt(frm.itemoptionno.length); i++) {
		regitemno = regitemno + parseInt(frm.itemoptionno[i].value);
	}

	if (regitemno > orgitemno) {
		alert('변경가능한 수량을 초과하였습니다.');
		frm.itemoptionno[changedindex].value = frm.itemoptionno[changedindex].value - (regitemno - orgitemno);
		regitemno = orgitemno;
	}

	frm.itemoptionno[0].value = orgitemno - regitemno;
}

function SetAddBeasongPay() {
    var frm = document.frm;

	if (!frm.isupchebeasong) {
		return;
	}

	if ((frm.gubun01.value == "C004") && (frm.gubun02.value == "CD01")) {
		// 단순변심
		frm.add_customeraddbeasongpay.value = frm.upchebeasongpay.value*2;
		frm.add_customeraddmethod.value = "1";
	} else {
		frm.add_customeraddbeasongpay.value = 0;
		frm.add_customeraddmethod.value = "";
	}
}

// ============================================================================
// 옵션변경 주문변경
// ============================================================================
function SaveItemOptionNo() {
    var frm = document.frm;

	var orgitemno = parseInt(frm.orgitemoptionno.value);
	var remainitemno = parseInt(frm.itemoptionno[0].value);

	var isupchebeasong = "<%= ojumunDetail.FJumunDetail.Fisupchebeasong %>";
	var itemstate = "<%= ojumunDetail.FJumunDetail.FcurrState %>";

	if (<%= LCase(IsChangeOrder) %> == true) {
		alert('교환주문은 상품변경 할 수 없습니다.');
		return;
	}

	if (orgitemno == remainitemno) {
		alert('변경할 수량이 0입니다.');
		return;
	}

	if (frm.gubun01.value == "") {
		alert("사유구분을 선택하세요.");
		return;
	}

	if (frm.title.value == "") {
		alert("제목을 입력하세요.");
		return;
	}

	if ((isupchebeasong == "Y") && (itemstate >= "3") && (itemstate < "7")) {
		// 업체배송, 상품준비 이후
		if (confirm('업체배송이면서 상품준비 이후입니다.\n\진행 하시겠습니까?') != true) {
			return;
		}
	}

	if (itemstate == "7") {
		// 상품출고 이후
		alert('상품출고 이후입니다. 등록할 수 없습니다.');
		return;
	}


	if (confirm('수정 하시겠습니까?')) {
		frm.mode.value="EditItemNoPart";
		frm.submit();
	}
}

// ============================================================================
// 옵션변경 맞교환
// ============================================================================
function SaveChangeItemOptionNo(){
    var frm = document.frm;

	<% 'if (ojumun.FRectOldOrder = "on") then %>
		//alert("6개월 이전주문 처리불가!!");
		//return;
	<% 'end if %>

	var orgitemno = parseInt(frm.orgitemoptionno.value);
	var remainitemno = parseInt(frm.itemoptionno[0].value);

	var itemstate = "<%= ojumunDetail.FJumunDetail.FcurrState %>";

	if (orgitemno < 1) {
		alert('원주문 수량이 없습니다.');
		return;
	}

	if (orgitemno == remainitemno) {
		alert('변경할 수량이 0입니다.');
		return;
	}

	if (frm.gubun01.value == "") {
		alert("사유구분을 선택하세요.");
		return;
	}

	if (frm.title.value == "") {
		alert("제목을 입력하세요.");
		return;
	}

	if ((itemstate < "7") && (<%= LCase(IsChangeOrder) %> != true)) {
		// 상품출고 이전
		alert('상품출고 이전 상품입니다. 교환출고(옵션변경)할 수 없습니다.');
		return;
	}

	if (confirm('교환(옵션변경) 접수하시겠습니까?')){
		frm.mode.value="ChangeEditItemNoPart";
		frm.submit();
	}
}

</script>

<form name="frm" method="post" action="/cscenter/ordermaster/orderdetail_simple_editoption_process.asp" style="margin:0px;">
<input type="hidden" name="mode" value="">
<input type="hidden" name="detailidx" value="<%= ojumunDetail.FJumunDetail.Fdetailidx %>">
<input type="hidden" name="itemid" value="<%= ojumunDetail.FJumunDetail.Fitemid %>">
<input type="hidden" name="orderserial" value="<%= ojumunDetail.FJumunDetail.FOrderSerial %>">
<input type="hidden" name="divcd" value="<%= divcd %>">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="15">
			<img src="/images/icon_star.gif" align="absbottom">&nbsp;<b>주문아이템정보 수정</b>
		</td>
	</tr>
	<tr bgcolor="#FFFFFF" height="25">
		<td width="100" bgcolor="<%= adminColor("tabletop") %>">IDX</td>
		<td><%= ojumunDetail.FJumunDetail.Fdetailidx %></td>
		<td width="110" rowspan="4"><img src="<%= ojumunDetail.FJumunDetail.FImageList %>"></td>
	</tr>
	<tr bgcolor="#FFFFFF" height="25">
		<td bgcolor="<%= adminColor("tabletop") %>">브랜드 ID</td>
		<td><%= ojumunDetail.FJumunDetail.Fmakerid %></td>
	</tr>
	<tr bgcolor="#FFFFFF" height="25">
		<td bgcolor="<%= adminColor("tabletop") %>">상품코드</td>
		<td><%= ojumunDetail.FJumunDetail.Fitemid %></td>
	</tr>
	<tr bgcolor="#FFFFFF" height="25">
		<td bgcolor="<%= adminColor("tabletop") %>">상품명</td>
		<td><%= ojumunDetail.FJumunDetail.Fitemname %></td>
	</tr>
	<tr bgcolor="#FFFFFF" height="25">
		<td bgcolor="<%= adminColor("tabletop") %>">취소상태</td>
		<td><%= ojumunDetail.FJumunDetail.Fcancelyn %></td>
		<td>

		</td>
	</tr>
</table>

<br>

<%

if ojumunDetail.FJumunDetail.Fitemoption = "0000" Then
	response.write "옵션이 없습니다. 등록할 수 없습니다."
	dbget.close() : response.end
end If

%>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr bgcolor="#FFFFFF">
		<td width="100" bgcolor="<%= adminColor("tabletop") %>">
			현재옵션
		</td>
		<td>
			<%= "[" & ojumunDetail.FJumunDetail.FitemOption & "] " & ojumunDetail.FJumunDetail.FitemOptionName %>
		</td>
		<input type=hidden name=orgitemoptionno value="<%= ojumunDetail.FJumunDetail.Fitemno - prevregno %>">
		<input type=hidden name=itemoptioncode value="<%= ojumunDetail.FJumunDetail.FitemOption %>">
		<td>
			<input type="text" class="text_ro" name="itemoptionno" value="<%= (ojumunDetail.FJumunDetail.Fitemno - prevregno) %>" size="3" maxlength="9" readonly> 개(<%= ojumunDetail.FJumunDetail.Fitemno %>개)

			<% if (prevregno <> 0) then %>
				<font color=red>(기존반품 : <%= prevregno %> 개)</font>
			<% end if %>
		</td>
	</tr>
	<%
	changedindex = 0
	%>
	<% For i = 0 To UBound(rsOption,2) %>
		<%
		itemoption 			= rsOption(0,i)
		optionname 			= rsOption(1,i)
		optsellyn 			= rsOption(2,i)
		optlimityn 			= rsOption(3,i)
		optlimitno 			= rsOption(4,i)
		optlimitsold 		= rsOption(5,i)
		issameoptaddprice 	= rsOption(6,i)
		isusing 			= rsOption(7,i)

		itemStatus = ""

		if (itemoption < ojumunDetail.FJumunDetail.Fitemoption) then
			changedindex = i + 1
		else
			changedindex = i
		end if

		%>
		<% if (itemoption = ojumunDetail.FJumunDetail.Fitemoption) then %>
			<!-- 현재옵션 스킵 -->
		<% else %>
			<%

			if (optsellyn = "N") then
				itemStatus = itemStatus + "<font color=red>판매않함</font>,"
			end if

			if (optlimityn = "Y") then
				if ((optlimitno - optlimitsold) < 1) then
					itemStatus = itemStatus + "<font color=red>한정:0</font>,"
				else
					itemStatus = itemStatus + "한정:" & ( optlimitno - optlimitsold ) & ","
				end if
			end if

			if (isusing = "N") then
				itemStatus = itemStatus + "<font color=red>사용안함</font>,"
			end if

			If itemStatus <> "" Then
				itemStatus = " ( " & Mid(itemStatus, 1, Len(itemStatus) - 1) & " )"
			End If

			optionText = "[" & itemoption & "] " & optionname

			%>
			<% if (issameoptaddprice = "F") then %>
			<tr bgcolor="#FFFFFF">
				<td width="100" bgcolor="<%= adminColor("tabletop") %>">
					변경불가옵션(<%= changedindex %>)
				</td>
				<td>
					<%=optionText%><font color=red>(옵션가 다름)</font>
				</td>
				<input type=hidden name=itemoptioncode value="<%= itemoption %>">
				<td>
					<input type="text" class="text_ro" name="itemoptionno" value="0" size="3" maxlength="9" onKeyUp="CheckItemOptionNo(<%= changedindex %>)" readonly> 개
					<%= itemStatus %>
				</td>
			</tr>
			<% else %>
			<tr bgcolor="#FFFFFF">
				<td width="100" bgcolor="<%= adminColor("tabletop") %>">
					변경가능옵션(<%= changedindex %>)
				</td>
				<td>
					<%=optionText%>
				</td>
				<input type=hidden name=itemoptioncode value="<%= itemoption %>">
				<td>
					<input type="text" class="text" name="itemoptionno" value="0" size="3" maxlength="9" onKeyUp="CheckItemOptionNo(<%= changedindex %>)"> 개
					<%= itemStatus %>
				</td>
			</tr>
			<% end if %>
		<% end If %>
	<% Next %>

	<tr bgcolor="#FFFFFF">
		<td width="100" height=30 bgcolor="<%= adminColor("tabletop") %>">
			사유구분
		</td>
		<td colspan=2>
                <input type="hidden" name="gubun01" value="">
                <input type="hidden" name="gubun02" value="">
                <input class="text_ro" type="text" name="gubun01name" value="" size="16" Readonly >
                &gt;
                <input class="text_ro" type="text" name="gubun02name" value="" size="16" Readonly >
                <input class="csbutton" type="button" value="선택" onClick="divCsAsGubunSelect(frm.gubun01.value, frm.gubun02.value, frm.gubun01.name, frm.gubun02.name, frm.gubun01name.name, frm.gubun02name.name,'frm','causepop');">
                <div id="causepop" style="position:absolute;"></div>

                <!-- 일부 사유 미리 표시 -->
                <%
                '참조쿼리
				'select top 100 m.comm_cd, m.comm_name, d.comm_cd, d.comm_name
				'from
				'	db_cs.dbo.tbl_cs_comm_code m
				'	left join db_cs.dbo.tbl_cs_comm_code d
				'	on
				'		m.comm_cd = d.comm_group
				'where
				'	1 = 1
				'	and m.comm_group = 'Z020'
				'	and m.comm_isdel <> 'Y'
				'	and d.comm_isdel <> 'Y'
				'order by m.comm_cd, d.comm_cd
                %>
                [<a href="javascript:selectGubun('C004','CD01','공통','단순변심','gubun01','gubun02','gubun01name','gubun02name','frm','causepop'); SetAddBeasongPay();">단순변심</a>]
                [<a href="javascript:selectGubun('C004','CD05','공통','품절','gubun01','gubun02','gubun01name','gubun02name','frm','causepop'); SetAddBeasongPay();">품절</a>]
                [<a href="javascript:selectGubun('C005','CE01','상품관련','상품불량','gubun01','gubun02','gubun01name','gubun02name','frm','causepop'); SetAddBeasongPay();">상품불량</a>]
                [<a href="javascript:selectGubun('C004','CD99','공통','기타','gubun01','gubun02','gubun01name','gubun02name','frm','causepop'); SetAddBeasongPay();">기타</a>]
            	&nbsp; &nbsp; &nbsp;
            	<div id="chkmodifyitemstockoutyn" style="display: inline;"><input type="checkbox" name="modifyitemstockoutyn" value="Y" checked> 품절정보 저장(업배상품)</div>
		</td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td width="100" height=30 bgcolor="<%= adminColor("tabletop") %>">
			접수제목
		</td>
		<td colspan=2>
                <input class='text' type="text" name="title" value="<%= title %>" size="56" maxlength="56">
		</td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td width="100" height=30 bgcolor="<%= adminColor("tabletop") %>">
			접수내용
		</td>
		<td colspan=2>
                <textarea class='textarea' name="contents_jupsu" cols="68" rows="6"><%= contents_jupsu %></textarea>
		</td>
	</tr>

	<% if (divcd = "A100") then %>
		<input type="hidden" name="isupchebeasong" value="<%= isupchebeasong %>">
		<input type="hidden" name="requiremakerid" value="<%= requiremakerid %>">
		<input type="hidden" name="upchebeasongpay" value="<%= upchebeasongpay %>">
		<tr bgcolor="#FFFFFF">
			<td width="100" height=30 bgcolor="<%= adminColor("tabletop") %>">
				배송구분
			</td>
			<td colspan=2>
		    	<% if (isupchebeasong = "Y") then %>
		    		<font color=red><%= requiremakerid %></font> (기본배송비 : <%= FormatNumber(upchebeasongpay, 0) %>원)
		    	<% else %>
		    		텐바이텐배송
		    	<% end if %>
			</td>
		</tr>
		<tr bgcolor="#FFFFFF">
			<td width="100" height=30 bgcolor="<%= adminColor("tabletop") %>">
				추가배송비
			</td>
			<td colspan=2>
		    	<input type="text" class="text" name="add_customeraddbeasongpay" value="0" size="20">
		    	&nbsp;
	    	    <select class="select" name="add_customeraddmethod" class="text">
		    	    <option value="">선택
		    	    <option value="1">박스동봉
		    	    <option value="2">택배비 고객부담
		    	    <option value="5">기타
	    	    </select>
			</td>
		</tr>
	<% end if %>

	<tr bgcolor="#FFFFFF" height=35>
		<td colspan="3" align="center">
			<% if Not IsOrderCanceled then %>
				<input type="button" class="button" value="옵션변경" onclick="javascript:SaveItemOptionNo()" <% if (ojumunDetail.FJumunDetail.FcurrState = "7") then %>disabled<% end if %>>
				<input type="button" class="button" value="옵션변경 맞교환" onclick="javascript:SaveChangeItemOptionNo()" <% if (ojumunDetail.FJumunDetail.FcurrState <> "7") and Not IsChangeOrder then %>disabled<% end if %>>
			<% else %>
				취소된 상품은 수량변경 불가
			<% end if %>
		</td>
	</tr>
</table>
</form>
<div>
* 옵션변경은 <font color=red>옵션가</font>가 동일한 옵션상품만 가능합니다.<br>
* 주문당시 옵션가격에 상관없이 현재 상품정보 상의 옵션가격으로 비교합니다.<br>
* 상품할인정보(판매가,매입가 등)는 주문당시 정보가 복사됩니다.<br>
* <font color=red>추가할 상품이 출고완료</font> 상태이면 주문변경 불가합니다.<br>
</div>
<%
set ojumun       = Nothing
set ojumunDetail = Nothing
%>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
