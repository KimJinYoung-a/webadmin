<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/classes/cscenter/oldmisendcls.asp"-->
<!-- #include virtual="/lib/classes/cscenter/cs_aslistcls.asp" -->
<%
dim orderserial
dim mode
Dim sqlStr, rowCount

orderserial = request("orderserial")

if (orderserial = "") then
    orderserial = "-"
end if


'// ============================================================================
'// CS 내역
'// ============================================================================
dim ocsaslist
set ocsaslist = New CCSASList
ocsaslist.FPageSize = 50
ocsaslist.FCurrPage = 1

ocsaslist.FRectOrderSerial = orderserial

ocsaslist.GetCSASMasterListByProcedure


'// ============================================================================
'// 제휴사 주문정보
'// ============================================================================

Dim ArrExtSiteOrderDetailList

sqlStr = " select "
sqlStr = sqlStr & " 	T.sellSite "
sqlStr = sqlStr & " 	, T.orderserial "
sqlStr = sqlStr & " 	, T.outmallorderserial "
sqlStr = sqlStr & " 	, T.OrgDetailKey  "
sqlStr = sqlStr & " 	, T.matchitemid "
sqlStr = sqlStr & " 	, T.matchitemoption "
sqlStr = sqlStr & " 	, T.orderitemname "
sqlStr = sqlStr & " 	, T.orderitemoptionname "
sqlStr = sqlStr & " 	, T.itemordercount "
sqlStr = sqlStr & " 	, d.itemno "
sqlStr = sqlStr & " 	, m.cancelyn "
sqlStr = sqlStr & " 	, d.cancelyn "
sqlStr = sqlStr & " 	, D.beasongdate "
sqlStr = sqlStr & " 	, IsNull(D.currstate, 0) as currstate "
sqlStr = sqlStr & " 	, IsNULL(T.sendState,0) as sendState "
sqlStr = sqlStr & " 	, T.matchState as matchState "				'// 15
sqlStr = sqlStr & " 	, T.outmallorderseq "
sqlStr = sqlStr & " from "
sqlStr = sqlStr & " 	db_temp.dbo.tbl_xSite_TMPOrder T "
sqlStr = sqlStr & "  	left Join db_order.dbo.tbl_order_detail D "
sqlStr = sqlStr & "  	on "
sqlStr = sqlStr & " 		T.orderserial=D.orderserial "
sqlStr = sqlStr & "  		and T.matchitemid=D.itemid "
sqlStr = sqlStr & "  		and T.matchitemoption=D.itemoption "
sqlStr = sqlStr & "  	left Join db_order.dbo.tbl_order_master M "
sqlStr = sqlStr & "  	on "
sqlStr = sqlStr & " 		D.orderserial=M.orderserial "
sqlStr = sqlStr & " where "
sqlStr = sqlStr & " 	1 = 1 "
sqlStr = sqlStr & " 	and T.orderserial = '" + CStr(orderserial) + "' "
sqlStr = sqlStr & " order by "
sqlStr = sqlStr & " 	T.orderserial, T.matchitemid, T.matchitemoption  "
''response.write sqlStr

rsget.Open sqlStr,dbget,1
if Not rsget.Eof then
	ArrExtSiteOrderDetailList = rsget.getRows
end if
rsget.Close


'// ============================================================================
dim omasterwithcs
set omasterwithcs = new COldMiSend
omasterwithcs.FRectOrderSerial = orderserial
omasterwithcs.GetOneOrderMasterWithCS

orderserial = omasterwithcs.FOneItem.FOrderSerial

dim omisendList
set omisendList = new COldMiSend
omisendList.FRectOrderSerial = orderserial
omisendList.GetMiSendOrderDetailList


dim validtenitemno
validtenitemno = 0

for i = 0 to omisendList.FResultCount - 1
	if (omisendList.FItemList(i).FDetailCancelYn <> "Y") then
		validtenitemno = validtenitemno + omisendList.FItemList(i).FItemNo
	end if
next

dim i
dim prevmatchitemid, prevmatchitemoption, IsCSDetail
%>
<script language='javascript'>
function SetCancelAllOrder() {
	var frm = document.frmact;
    if (confirm("삭제처리 하시겠습니까?") == true) {
		frm.mode.value = "cancelall";
        frm.submit();
    }
}

function SetCancelSelectedOrder(isforce) {
	var arrchk = "";
	var validextitemno = 0;
	var validtenitemno = <%= validtenitemno %>;
	for (var i = 0; ; i++) {
		var chk = document.getElementById("chk_" + i);
		var extitemno = document.getElementById("extitemno_" + i);
		if (chk == undefined) {
			break;
		}

		if (chk.checked == true) {
			arrchk = arrchk + "," + chk.value;
		} else {
			validextitemno = validextitemno + extitemno.value*1;
		}
	}

	if (arrchk == "") {
		alert("선택된 주문이 없습니다.");
		return;
	}

	if (isforce != true) {
		if (validextitemno != validtenitemno) {
			if (validextitemno > validtenitemno) {
				alert("제휴주문수량(" + validextitemno + ")과 텐바이텐 주문수량(" + validtenitemno + ")이 일치하지 않습니다.\n\n상품을 취소된 경우 제휴몰 주문내역을 삭제하세요");
			} else {
				alert("제휴주문수량(" + validextitemno + ")과 텐바이텐 주문수량(" + validtenitemno + ")이 일치하지 않습니다.\n\n상품을 변경한 경우 연결상품을 변경하세요.");
			}

			return;
		}
	}

	var frm = document.frmact;
	if (confirm("선택 제휴주문을 삭제하시겠습니까?") == true) {
		frm.mode.value = "cancelselected";
		frm.arrchk.value = arrchk;
		frm.submit();
	}
}

function ModifyMatchItem(outmallorderseq, extitemid, extitemoption, extitemno) {
	// =========================================================================
	// 수량 검증
	// =========================================================================
	var validextitemno = extitemno;
	var validtenitemno = 0;
	for (var i = 0; ; i++) {
		var chk = document.getElementById("chkten_" + i);
		var itemno = document.getElementById("tenitemno_" + i);
		if (chk == undefined) {
			break;
		}

		if (chk.checked) {
			validtenitemno = validtenitemno + itemno.value*1;
		}
	}

	if (validextitemno != validtenitemno) {
		if (validextitemno > validtenitemno) {
			alert("제휴주문수량(" + validextitemno + ")과 텐바이텐 주문수량(" + validtenitemno + ")이 일치하지 않습니다.\n\n상품을 취소된 경우 제휴몰 주문내역을 삭제하세요");
		} else {
			alert("제휴주문수량(" + validextitemno + ")과 텐바이텐 주문수량(" + validtenitemno + ")이 일치하지 않습니다.\n\n상품을 변경한 경우 연결상품을 변경하세요.");
		}

		return;
	}

	// =========================================================================
	// 선택상품 확인
	// =========================================================================
	var selecteditemid = 0;
	var selecteditemoption = "";
	var selecteditemno = 0;
	var selecteditemcount = 0;

	for (var i = 0; ; i++) {
		var chkten = document.getElementById("chkten_" + i);
		var tenitemid = document.getElementById("tenitemid_" + i);
		var tenitemoption = document.getElementById("tenitemoption_" + i);
		var tenitemno = document.getElementById("tenitemno_" + i);

		if (chkten == undefined) {
			break;
		}

		if (chkten.checked == true) {
			if (selecteditemcount < 1) {
				selecteditemid = tenitemid.value;
				selecteditemoption = tenitemoption.value;
			}
			selecteditemno = selecteditemno + tenitemno.value*1;
			selecteditemcount = selecteditemcount + 1;
		}
	}

	if (selecteditemcount < 1) {
		alert("텐바이텐 주문내역에서 변경된 상품을 선택하세요.");
		return;
	}

	/*
	// 두가지 이상의 상품으로 나눠서 주문내역을 변경한경우
	// 변경된 상품 모두를 체크해서 수량을 비교하고, 수량이 맞으면 첫번째 상품의 송장번호 복사
	if (selecteditemcount > 1) {
		alert("변경할 상품은 하나만 선택 가능합니다.");
		return;
	}
	*/

	if (selecteditemno*1 != extitemno*1) {
		if (confirm("주의!!\n\n수량이 다릅니다.\n강제로 송장입력 하시겠습니까?") != true) {
			return;
		}
		// alert("변경불가 - 수량이 다릅니다." + selecteditemno);
		// return;
	}

	// =========================================================================
	// 선택 상품이 이미 제휴몰 주문 상품에 있는지 확인
	// =========================================================================
	for (var i = 0; ; i++) {
		var chk = document.getElementById("chk_" + i);
		var itemid = document.getElementById("extitemid_" + i);
		var itemoption = document.getElementById("extitemoption_" + i);
		if (chk == undefined) {
			break;
		}

		if ((itemid.value*1 == selecteditemid*1) && (itemoption.value*1 == selecteditemoption*1)) {
			alert("중복에러 : 변경할 상품이 이미 제휴몰 주문내역에 있습니다.");
			return;
		}
	}

	var frm = document.frmact;
	if (confirm("선택 상품으로 연결상품을 변경하시겠습니까?") == true) {
		frm.mode.value = "modifymatchitem";
		frm.chk.value = outmallorderseq;
		frm.itemid.value = selecteditemid;
		frm.itemoption.value = selecteditemoption;
		frm.submit();
	}
}

function ModifyMatchItemNo(outmallorderseq, extitemid, extitemoption, extitemno) {
	// =========================================================================
	var validextitemno = 0;
	var validtenitemno = <%= validtenitemno %>;
	for (var i = 0; ; i++) {
		var chk = document.getElementById("chk_" + i);
		var itemno = document.getElementById("extitemno_" + i);
		if (chk == undefined) {
			break;
		}

		validextitemno = validextitemno + itemno.value*1;
	}

	if (validextitemno != validtenitemno) {
		if (validextitemno > validtenitemno) {
			alert("제휴주문수량(" + validextitemno + ")과 텐바이텐 주문수량(" + validtenitemno + ")이 일치하지 않습니다.\n\n상품을 취소된 경우 제휴몰 주문내역을 삭제하세요");
		} else {
			alert("제휴주문수량(" + validextitemno + ")과 텐바이텐 주문수량(" + validtenitemno + ")이 일치하지 않습니다.\n\n상품을 변경한 경우 연결상품을 변경하세요.");
		}

		return;
	}

	// =========================================================================
	var changeditemid = 0;
	var changeditemoption = "";
	for (var i = 0; ; i++) {
		var chk = document.getElementById("chk_" + i);
		var extchangeitemid = document.getElementById("extitemid_" + i);
		var extchangeitemoption = document.getElementById("extitemoption_" + i);
		var extchangeitemno = document.getElementById("extitemno_" + i);
		var extchangetenitemno = document.getElementById("exttenitemno_" + i);

		if (chk == undefined) {
			break;
		}

		if ((extchangeitemno.value*1 + extitemno*1) == extchangetenitemno.value*1) {
			changeditemid = extchangeitemid.value;
			changeditemoption = extchangeitemoption.value;
			break;
		}
	}

	if (changeditemid == 0) {
		alert("변경불가 : 시스템팀 문의");
		return;
	}

	var frm = document.frmact;
	if (confirm("수량이 늘어난 상품으로 변경하시겠습니까?") == true) {
		frm.mode.value = "modifymatchitemno";
		frm.chk.value = outmallorderseq;
		frm.itemid.value = changeditemid;
		frm.itemoption.value = changeditemoption;
		frm.submit();
	}
}

</script>
<style type="text/css">
<!--
td { font-size:9pt; font-family:Verdana;}

.button {
	font-family: "굴림", "돋움";
	font-size: 10px;
	background-color: #E4E4E4;
	border: 1px solid #000000;
	color: #000000;
	height: 20px;
}
//-->
</style>

<!-- 검색 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name=frmsearch>
	<tr align="center" bgcolor="#FFFFFF" >
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
		<td align="left">
			주문번호 : <input type="text" class="text" name="orderserial" value="<%= orderserial %>" size=13 >
        	<% if omasterwithcs.FOneItem.FCancelyn<>"N" then %>
			<b><font color="#CC3333">[취소주문]</font></b>
			<script language='javascript'>alert('취소된 거래 입니다.');</script>
			<% else %>
			[정상주문]
			<% end if %>
		</td>
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button_s" value="검색" onClick="SearchThis();">
		</td>
	</tr>
	</form>
</table>

<p>

<br><b>[CS처리내역]</b>
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td height="25" width="80">idx</td>
		<td width="120">구분</td>
		<td width="100">원주문번호</td>
		<td width="100">Site</td>
		<td>제목</td>
		<td width="35">상태</td>
		<td width="80">등록일</td>
		<td width="80">처리일</td>
		<td width="220">관련송장</td>
		<td width="30">삭제</td>
	</tr>
<% for i = 0 to (ocsaslist.FResultCount - 1) %>
    <% if (ocsaslist.FItemList(i).Fdeleteyn <> "N") then %>
    <tr bgcolor="#EEEEEE" style="color:gray" align="center">
    <% else %>
    <tr bgcolor="#FFFFFF" align="center">
    <% end if %>
        <td height="25" nowrap><%= ocsaslist.FItemList(i).Fid %></td>
        <td nowrap align="left"><acronym title="<%= ocsaslist.FItemList(i).GetAsDivCDName %>"><font color="<%= ocsaslist.FItemList(i).GetAsDivCDColor %>"><%= ocsaslist.FItemList(i).GetAsDivCDName %></font></acronym></td>
        <td nowrap>
			<%= ocsaslist.FItemList(i).Forgorderserial %>
			<% if (ocsaslist.FItemList(i).Forderserial <> ocsaslist.FItemList(i).Forgorderserial) then %>
				+
			<% end if %>
        </td>
        <td nowrap><%= ocsaslist.FItemList(i).FExtsitename %></td>
        <td nowrap align="left"><acronym title="<%= ocsaslist.FItemList(i).Ftitle %>"><%= ocsaslist.FItemList(i).Ftitle %></acronym></td>
        <td nowrap><font color="<%= ocsaslist.FItemList(i).GetCurrstateColor %>"><%= ocsaslist.FItemList(i).GetCurrstateName %></font></td>
        <td nowrap><acronym title="<%= ocsaslist.FItemList(i).Fregdate %>"><%= Left(ocsaslist.FItemList(i).Fregdate,10) %></acronym></td>
        <td nowrap><acronym title="<%= ocsaslist.FItemList(i).Ffinishdate %>"><%= Left(ocsaslist.FItemList(i).Ffinishdate,10) %></acronym></td>
		<td nowrap>
			<% Call drawSelectBoxDeliverCompany ("songjangdiv", ocsaslist.FItemList(i).Fsongjangdiv) %>
			<%= ocsaslist.FItemList(i).Fsongjangno %>
		</td>
        <td nowrap>
			<% if ocsaslist.FItemList(i).Fdeleteyn="Y" then %>
			<font color="red">삭제</font>
			<% elseif ocsaslist.FItemList(i).Fdeleteyn="C" then %>
			<font color="red"><strong>취소</strong></font>
			<% end if %>
        </td>
    </tr>
<% next %>

</table>

<p>

<br><b>[텐바이텐주문내역]</b>
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
	<form name="frmtensite">
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td width="20"></td>
		<td width="50">상품코드</td>
		<td width="50">이미지</td>
		<td>상품명<font color="blue">[옵션]</font></td>
		<td width="40">주문<br>수량</td>
		<td width="35">취소<br>삭제</td>
		<td width="60">진행상태</td>
		<td width="80">미출고사유</td>
		<td width="120">송장번호</td>
	</tr>
	<% for i=0 to omisendList.FResultCount -1 %>
	<% if omisendList.FItemList(i).FItemLackNo<>0 then %>
	<tr align="center" bgcolor="<%= adminColor("pink") %>">
	<% else %>
	<tr align="center" bgcolor="FFFFFF">
	<% end if %>
		<td><input type="checkbox" name="chkten" id="chkten_<%= i %>" value="" <% if omisendList.FItemList(i).FDetailCancelYn = "Y" then %>disabled<% end if %> ></td>
		<input type="hidden" name="tenitemid" id="tenitemid_<%= i %>" value="<%= omisendList.FItemList(i).FItemID %>">
		<input type="hidden" name="tenitemoption" id="tenitemoption_<%= i %>" value="<%= omisendList.FItemList(i).FItemOption %>">
		<input type="hidden" name="tenitemno" id="tenitemno_<%= i %>" value="<%= omisendList.FItemList(i).FItemNo %>">
		<td>
			<% if omisendList.FItemList(i).FIsUpcheBeasong="Y" then %>
			<font color="red"><%= omisendList.FItemList(i).FItemID %></font>
			<% else %>
			<%= omisendList.FItemList(i).FItemID %>
			<% end if %>
		</td>
		<td><img src="<%= omisendList.FItemList(i).FSmallImage %>" width="50" height="50"></td>
		<td align="left">
			<%= omisendList.FItemList(i).FItemName %>
			<% if omisendList.FItemList(i).FItemOptionName<>"" then %>
			<br>
			<font color="blue">[<%= omisendList.FItemList(i).FItemOptionName %>]</font>
			<% end if %>
		</td>
		<td><%= omisendList.FItemList(i).FItemNo %></td>
		<td>
		    <%= fnColor(omisendList.FItemList(i).FDetailCancelYn,"cancelyn") %>
		</td>
		<td>
		    <font color="<%= omisendList.FItemList(i).getUpCheDeliverStateColor %>"><%= omisendList.FItemList(i).getUpCheDeliverStateName %></font>
		</td>
		<td>
				<font color="<%= omisendList.FItemList(i).getMiSendCodeColor %>"><%= omisendList.FItemList(i).getMiSendCodeName %></font>
		</td>
		<td>
			<% if (omisendList.FItemList(i).FSongjangno <> "") then %>
				<%= omisendList.FItemList(i).FSongjangdiv %>
				<%= omisendList.FItemList(i).FSongjangno %>
			<% end if %>
		</td>
	</tr>
	<% next %>
	</form>
</table>

<p>

<br><b>[제휴몰주문내역]</b>
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
	<form name="frmextsite">
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	    <td width="20"></td>
		<td width="30">주문<br>구분</td>
		<td width="70">제휴사</td>
    	<td>제휴주문번호</td>
      	<td>제휴<br>상품코드</td>
      	<td width="50">상품코드</td>
      	<td width="50">옵션코드</td>
      	<td align="left">상품명<br><font color="blue">[옵션명]</font></td>
        <td width="30">제휴<br>주문<br>수량</td>
		<td width="30">텐텐<br>주문<br>수량</td>
		<td width="30">제휴<br>취소<br>상태</td>
		<td width="30">배송<br>상태</td>
		<td width="30">송장<br>전송</td>
      	<td>비고</td>
    </tr>
<% if (IsArray(ArrExtSiteOrderDetailList)) THEN %>
<%
rowCount = UBound(ArrExtSiteOrderDetailList,2)

for i=0 to rowCount
	'// 상품코드, 옵션코드 모두 동일하면 제휴몰 CS건이다.
	IsCSDetail = (prevmatchitemid = ArrExtSiteOrderDetailList(4,i)) and (prevmatchitemoption = ArrExtSiteOrderDetailList(5,i))
%>
<tr bgcolor="#FFFFFF" align="center">
	<td><input type="checkbox" name="chk" id="chk_<%= i %>" value="<%= ArrExtSiteOrderDetailList(16,i) %>" <% if ArrExtSiteOrderDetailList(14,i) = "1" or ArrExtSiteOrderDetailList(13,i) = "7" and Not IsCSDetail then %>disabled<% end if %> ></td>
	<input type="hidden" name="extitemid" id="extitemid_<%= i %>" value="<%= ArrExtSiteOrderDetailList(4,i) %>">
	<input type="hidden" name="extitemoption" id="extitemoption_<%= i %>" value="<%= ArrExtSiteOrderDetailList(5,i) %>">
	<input type="hidden" name="extitemno" id="extitemno_<%= i %>" value="<%= ArrExtSiteOrderDetailList(8,i) %>">
	<input type="hidden" name="exttenitemno" id="exttenitemno_<%= i %>" value="<%= ArrExtSiteOrderDetailList(9,i) %>">
    <td height="45">
		<% if IsCSDetail then %>
			<font color=red>CS</font>
		<% else %>
			정상
		<% end if %>
	</td>
	<td><%= ArrExtSiteOrderDetailList(0,i) %></td>
    <td><%= ArrExtSiteOrderDetailList(2,i) %></td>
    <td><%= ArrExtSiteOrderDetailList(3,i) %></td>
    <td><%= ArrExtSiteOrderDetailList(4,i) %></td>
	<td><%= ArrExtSiteOrderDetailList(5,i) %></td>
    <td align="left">
		<%= ArrExtSiteOrderDetailList(6,i) %>
		<% if (ArrExtSiteOrderDetailList(5,i) <> "0000") then %>
			<br><font color="blue">[<%= ArrExtSiteOrderDetailList(7,i) %>]</font>
		<% end if %>
	</td>
    <td align="center"><%= ArrExtSiteOrderDetailList(8,i) %></td>
    <td align="center">
		<% if (ArrExtSiteOrderDetailList(8,i) <> ArrExtSiteOrderDetailList(9,i)) then %><font color="red"><% end if %>
		<% if (ArrExtSiteOrderDetailList(11,i) <> "Y") then %>
			<%= ArrExtSiteOrderDetailList(9,i) %>
		<% end if %>
	</td>
	<td>
		<% if ArrExtSiteOrderDetailList(15,i) = "D" then %>
		취소
		<% end if %>
	</td>
	<td align="center">
		<% if ArrExtSiteOrderDetailList(13,i) = "7" then %>
			출고<br>완료
		<% end if %>
	</td>
	<td align="center">
		<% if ArrExtSiteOrderDetailList(14,i) = "1" then %>
			Y
		<% end if %>
	</td>
    <td align="center">
		<input type="button" class="csbutton" value="상품변경" onClick="ModifyMatchItem(<%= ArrExtSiteOrderDetailList(16,i) %>, <%= ArrExtSiteOrderDetailList(4,i) %>, '<%= ArrExtSiteOrderDetailList(5,i) %>', <%= ArrExtSiteOrderDetailList(8,i) %>)" <% if ArrExtSiteOrderDetailList(14,i) = "1" then %>disabled<% end if %> >
		<input type="button" class="csbutton" value="수량추가" onClick="ModifyMatchItemNo(<%= ArrExtSiteOrderDetailList(16,i) %>, <%= ArrExtSiteOrderDetailList(4,i) %>, '<%= ArrExtSiteOrderDetailList(5,i) %>', <%= ArrExtSiteOrderDetailList(8,i) %>)" <% if ArrExtSiteOrderDetailList(14,i) = "1" then %>disabled<% end if %> >
	</td>
</tr>
<%
prevmatchitemid = ArrExtSiteOrderDetailList(4,i)
prevmatchitemoption = ArrExtSiteOrderDetailList(5,i)
%>
<% next %>
<% ELSE %>
<tr>
    <td colspan="11" align="center">[검색 결과가 없습니다.]</td>
</tr>
<% end if %>
</form>
</table>

<p>

<input type="button" class="csbutton" value="선택 제휴몰 주문삭제" onClick="SetCancelSelectedOrder(false)">
&nbsp;
<input type="button" class="csbutton" value="[강제삭제] 선택 제휴몰 주문삭제" onClick="SetCancelSelectedOrder(true)">
&nbsp;
<input type="button" class="csbutton" value="제휴몰주문 전체삭제" onClick="SetCancelAllOrder()" <%if (omasterwithcs.FOneItem.FCancelyn <> "Y") then %>disabled<% end if %> >

<p>

<form name="frmact" method="post" action="etcSiteOrderProc.asp">
<input type="hidden" name="mode" value="">
<input type="hidden" name="orderserial" value="<%= orderserial %>">
<input type="hidden" name="arrchk" value="">
<input type="hidden" name="chk" value="">
<input type="hidden" name="itemid" value="">
<input type="hidden" name="itemoption" value="">
</form>
<!-- 표 하단바 끝-->


<%
set omasterwithcs = Nothing
set omisendList = Nothing
%>

<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
