<%@ language=vbscript %>
<%
option explicit
Response.Expires = -1
%>
<%
'###########################################################
' Description : 오프라인 배송
' Hieditor : 2011.02.25 한용민 생성
'###########################################################
%>
<!-- #include virtual="/common/incSessionAdminOrUpche.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/common/lib/commonbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/offshop/upche/upchebeasong_cls.asp" -->

<%
dim ojumun ,i
set ojumun = new cupchebeasong_list
	ojumun.FRectDesignerID = session("ssBctId")
	ojumun.fDesignerDateBaljuList()

%>

<script language='javascript'>

//전체선택
function switchCheckBox(comp){
    var frm = comp.form;

	if(frm.detailidx.length>1){
		for(i=0;i<frm.detailidx.length;i++){
			frm.detailidx[i].checked = comp.checked;
			AnCheckClick(frm.detailidx[i]);
		}
	}else{
		frm.detailidx.checked = comp.checked;
		AnCheckClick(frm.detailidx);
	}
}

//선택 주문 확인
function CheckNBaljusu(){
	var frm = document.frmbalju;
	var pass = false;

    if(frm.detailidx.length>1){
    	for (var i=0;i<frm.detailidx.length;i++){
    	    pass = (pass||frm.detailidx[i].checked);
    	}
    }else{
        pass = frm.detailidx.checked;
    }

	if (!pass) {
		alert("선택 주문이 없습니다.");
		return;
	}

	var ret = confirm("선택 주문을 확인 하시겠습니까?");

	if (ret){
 		frm.action="/common/offshop/beasong/upche_selectbaljulist.asp";
		frm.submit();

	}
}

</script>

<!-- 검색 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get">
<input type="hidden" name="menupos" value="<%= menupos %>">
<tr align="center" bgcolor="<%= adminColor("topbar") %>" >
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
	<td align="left" bgcolor="#FFFFFF">
		<input type="radio" name="" value="" checked >배송요청 리스트
		<!-- <input type="radio" name="" value="">요청이전 주문리스트(주문접수 포함) -->
	</td>
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="검색" onClick="javascript:document.frm.submit();">
	</td>
</tr>
</form>
</table>

<br>
※ 배송요청선택 후, [선택요청확인]을 클릭하시면, 발주서 출력이 가능합니다.
<br>발주서 재출력을 원하실 경우, [오프샵미배송리스트]를 이용하시기 바랍니다.
<br>(요청확인을 하셔야 배송정보 확인이 가능합니다.)

<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
<tr align="center" bgcolor="<%= adminColor("topbar") %>" >
	<td align="left">
		<input type="button" class="button" value="선택요청확인" onclick="CheckNBaljusu()">
	</td>
</tr>
</table>
<!-- 액션 끝 -->

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frmbalju" method="post">
<tr bgcolor="FFFFFF">
	<td height="25" colspan="15">
		검색결과 : <b><% = ojumun.FTotalCount %></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td><input type="checkbox" name="chkAll" onClick="switchCheckBox(this);"></td>
	<td>IDX</td>
	<td>매장주문번호</td>
	<td>수령인</td>
	<td>상품코드</td>
	<td>상품명<font color="blue">&nbsp;[옵션]</font></td>
	<td>공급가</td>
	<td>판매가</td>
	<td>수량</td>
	<td>배송요청일</td>
	<td>출고기준일<!--주문통보일--></td>
	<td>경과일</td>
</tr>
<% if ojumun.ftotalcount > 0 then %>
<% for i=0 to ojumun.ftotalcount-1 %>
<tr align="center" class="a" bgcolor="#FFFFFF">
	<td>
	    <!-- detail Index -->
		<input type="checkbox" name="detailidx"  onClick="AnCheckClick(this);" value="<% =ojumun.fitemlist(i).fdetailidx %>">
	</td>
	<td><%= ojumun.fitemlist(i).fdetailidx %></td>
	<td><%= ojumun.fitemlist(i).forderno %></td>
	<td><%= ojumun.fitemlist(i).FReqname %></td>
	<td><%= ojumun.fitemlist(i).fitemgubun %>-<%= CHKIIF(ojumun.fitemlist(i).FitemID>=1000000,Format00(8,ojumun.fitemlist(i).FitemID),Format00(6,ojumun.fitemlist(i).FitemID)) %>-<%= ojumun.fitemlist(i).fitemoption %></td>
	<td align="left">
		<%= ojumun.fitemlist(i).FItemname %>
		<% if (ojumun.fitemlist(i).fitemoptionname<>"") then %>
		<font color="blue">[<%= ojumun.fitemlist(i).fitemoptionname %>]</font>
		<% end if %>
	</td>
	<td><%= FormatNumber(ojumun.fitemlist(i).fsuplyprice,0) %></td>
	<td><%= FormatNumber(ojumun.fitemlist(i).fsellprice,0) %></td>
	<td><%= ojumun.fitemlist(i).FItemno %></td>
	<td><acronym title="<%= ojumun.fitemlist(i).Fregdate %>"><%= left(ojumun.fitemlist(i).Fregdate,10) %></acronym></td>
	<td><acronym title="<%= ojumun.fitemlist(i).Fbaljudate %>"><%= left(ojumun.fitemlist(i).Fbaljudate,10) %></acronym></td>
    <td>
        <% if IsNULL(ojumun.fitemlist(i).Fbaljudate) then %>
        	D+0
        <% elseif datediff("d",(left(ojumun.fitemlist(i).Fbaljudate,10)) , (left(now,10)) )>2 then %>
        	<font color="red"><b>D+<%= datediff("d",(left(ojumun.fitemlist(i).Fbaljudate,10)) , (left(now,10)) ) %></b></font>
        <% else %>
        	D+<%= datediff("d",(left(ojumun.fitemlist(i).Fbaljudate,10)) , (left(now,10)) ) %>
        <% end if %>
    </td>
</tr>

<% next %>

<% else %>
	<tr bgcolor="#FFFFFF">
		<td colspan="15" align="center">[검색결과가 없습니다.]</td>
	</tr>
	<% end if %>


    </form>
</table>


<%
set ojumun = Nothing
%>

<!-- #include virtual="/common/lib/commonbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->