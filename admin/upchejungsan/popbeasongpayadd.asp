<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->

<%
dim gubun, orderserial
dim yyyy1, mm1

gubun       = "witakchulgo"
orderserial = request("orderserial")
yyyy1       = request("yyyy1")
mm1         = request("mm1")

dim sqlStr
dim jungsanDataExists
dim orderRows, jungsanRows
sqlStr = " select distinct m.buyname, m.reqname, d.makerid "
sqlStr = sqlStr + " from [db_order].[dbo].tbl_order_master m "
sqlStr = sqlStr + " left join [db_order].[dbo].tbl_order_detail d on m.orderserial=d.orderserial"
sqlStr = sqlStr + " where m.orderserial='" + CStr(orderserial) + "'"

if (orderserial<>"") then
    rsget.Open sqlStr,dbget,1
    If Not rsget.Eof then
        orderRows = rsget.getRows()
    end if
    rsget.Close
end if

sqlStr = " select top 10 m.yyyymm, m.designerid, d.gubuncd, d.mastercode, d.itemid, d.itemname, d.itemno, d.sellcash, d.suplycash "
sqlStr = sqlStr + " from [db_jungsan].[dbo].tbl_designer_jungsan_master m,"
sqlStr = sqlStr + " [db_jungsan].[dbo].tbl_designer_jungsan_detail d"
sqlStr = sqlStr + " where m.id=d.masteridx"
sqlStr = sqlStr + " and d.itemid=0"
sqlStr = sqlStr + " and d.gubuncd='witakchulgo'"
sqlStr = sqlStr + " and d.mastercode='" + CStr(orderserial) + "'"

if (orderserial<>"") then
    rsget.Open sqlStr,dbget,1
    If Not rsget.Eof then
        jungsanRows = rsget.getRows()
    end if
    rsget.Close
end if

dim i
%>

<script language='javascript'>
function searchOrder(frm){
    if (frm.orderserial.value.length!=11){
        alert('주문번호 11 자리를 입력하세요.');
        frm.orderserial.focus();
        return;
    }
    frm.method="get";
    frm.action="/admin/upchejungsan/popbeasongpayadd.asp";
    frm.submit();
}

function adddata(frm){
    if (frm.orderserial.value.length!=11){
        alert('주문번호를 입력하세요.');
		frm.orderserial.focus();
		return;
    }
    
    if (frm.makerid.value.length<1){
        alert('브랜드ID를 선택하세요.');
		frm.makerid.focus();
		return;
    }
    
	if (frm.itemname.value.length<1){
		alert('내용을 입력하세요.');
		frm.itemname.focus();
		return;
	}

	if (frm.itemno.value.length<1){
		alert('갯수를 입력하세요.');
		frm.itemno.focus();
		return;
	}

	if (!IsDigit(frm.itemno.value)){
		alert('갯수는 숫자만 가능합니다.');
		frm.itemno.focus();
		return;
	}

	if (frm.sellcash.value.length<1){
		alert('판매가를 입력하세요.');
		frm.sellcash.focus();
		return;
	}

	if (!IsDigit(frm.sellcash.value)){
		alert('판매가는 숫자만 가능합니다.');
		frm.sellcash.focus();
		return;
	}

	if (frm.suplycash.value.length<1){
		alert('매입가를 입력하세요.');
		frm.suplycash.focus();
		return;
	}

	if (!IsDigit(frm.suplycash.value)){
		alert('매입가는 숫자만 가능합니다.');
		frm.suplycash.focus();
		return;
	}

	var ret = confirm('저장 하시겠습니까?');
	if (ret){
	    
		frm.submit();
	}
}
</script>


<!-- 표 상단바 시작-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
   	<tr height="10" valign="bottom">
        <td width="10" align="right"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
        <td background="/images/tbl_blue_round_02.gif"></td>
        <td background="/images/tbl_blue_round_02.gif"></td>
        <td width="10" align="left" ><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
	</tr>
	<tr height="25" valign="top">
        <td background="/images/tbl_blue_round_04.gif"></td>
        <td>
        	<img src="/images/icon_star.gif" align="absbottom">
			<font color="red"><strong>기타내역추가</strong></font>
        </td>
        <td align="right">배송비차액
        </td>
        <td background="/images/tbl_blue_round_05.gif"></td>
	</tr>
</table>
<!-- 표 상단바 끝-->
<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frmadd" method="post" action="dodesignerjungsan.asp">
    <input type="hidden" name="mode" value="etcbeasongpayadd">
    <input type="hidden" name="gubun" value="<%= gubun %>">
    <input type="hidden" name="yyyy1" value="<%= yyyy1 %>">
    <input type="hidden" name="mm1" value="<%= mm1 %>">
    <tr align="center" bgcolor="<%= adminColor("tabletop") %>">
        <td width="80">정산년월</td>
        <td width="100">주문번호</td>
        <td width="100">브랜드ID</td>
        <td width="80">구매자</td>
        <td width="80">수령인</td>
		<td>내용</td>
		<td width="40">수량</td>
		<td width="80">판매가</td>
		<td width="80">공급가</td>
    </tr>
    <tr bgcolor="#FFFFFF">
        <td><%= yyyy1 %>-<%= mm1 %></td>
        <td><input type="text" name="orderserial" value="<%= orderserial %>" size="12" maxlength="11"><input type="button" value="검색" onClick="searchOrder(frmadd);" onFocus="this.blur"></td>
		<td>
		    <select name="makerid">
            <% if IsArray(orderRows) then %>
            <% for i=0 to UBound(orderRows,2) %>
            <option value="<%= orderRows(2,i) %>"><%= orderRows(2,i) %>
            <% next %>
            <% end if %>
            </select>
        </td>
		<td>
		    <% if IsArray(orderRows) then %>
		    <input type="text" name="buyname" value="<%= db2html(orderRows(0,0)) %>" size="8">
		    <% else %>
		    <input type="text" name="buyname" value="" size="8">
		    <% end if %>
		</td>
		<td>
		    <% if IsArray(orderRows) then %>
		    <input type="text" name="reqname" value="<%= db2html(orderRows(1,0)) %>" size="8">
		    <% else %>
		    <input type="text" name="reqname" value="" size="8">
		    <% end if %>
		</td>
		<td><input type="text" name="itemname" value="" size="40"></td>
		<td><input type="text" name="itemno" value="1" size="3"></td>
		<td><input type="text" name="sellcash" value="" size="8"></td>
		<td><input type="text" name="suplycash" value="" size="8"></td>
    </tr>
</table>

<!-- 표 하단바 시작-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
    <tr valign="bottom" height="25">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td valign="bottom" align="center"><input type="button" value="내역 추가" onclick="adddata(frmadd)"></td>
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
    </tr>
    <tr valign="top" height="10">
        <td width="10" align="right"><img src="/images/tbl_blue_round_07.gif" width="10" height="10"></td>
        <td background="/images/tbl_blue_round_08.gif"></td>
        <td width="10" align="left"><img src="/images/tbl_blue_round_09.gif" width="10" height="10"></td>
    </tr>
    </form>
</table>
<br>
<% if IsArray(jungsanRows) then %>
<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<tr bgcolor="#FFFFFF">
    <td colspan="9">등록된 내역 검토</td>
</tr>
<tr bgcolor="#DDDDFF">
    <td>정산월</td>
    <td>브랜드</td>
    <td>구분</td>
    <td>주문번호</td>
    <td>상품코드</td>
    <td>상품명</td>
    <td>갯수</td>
    <td>판매가</td>
    <td>정산반영액</td>
</tr>
<% for i=0 to UBound(jungsanRows,2) %>
<tr bgcolor="#FFFFFF">
    <td><%= jungsanRows(0,i) %></td>
    <td><%= jungsanRows(1,i) %></td>
    <td><%= jungsanRows(2,i) %></td>
    <td><%= jungsanRows(3,i) %></td>
    <td><%= jungsanRows(4,i) %></td>
    <td><%= jungsanRows(5,i) %></td>
    <td><%= jungsanRows(6,i) %></td>
    <td><%= jungsanRows(7,i) %></td>
    <td><%= jungsanRows(8,i) %></td>
</tr>
<% next %>
</table>

<% end if %>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->