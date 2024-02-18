<%@ language=vbscript %>
<%
option explicit
Response.Expires = -1
%>
<%
'###########################################################
' Description : 오프라인 배송
' Hieditor : 2011.02.26 한용민 생성
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
dim searchType, searchValue, MisendReason ,ojumun ,i,iy
	searchType      = request("searchType")
	searchValue     = request("searchValue")
	MisendReason    = request("MisendReason")

set ojumun = new cupchebeasong_list
	ojumun.FRectSearchType  = SearchType
	ojumun.FRectSearchValue = SearchValue

	if (MisendReason="") then
	    ojumun.FRectMisendReason = "AA"
	else
	    ojumun.FRectMisendReason = MisendReason
	end if

	ojumun.FRectDesignerID = session("ssBctID")
	ojumun.fDesignerDateBaljuinputlist()
%>

<script language='javascript'>

function chksubmit(){
    var frm = document.frm;

    if ((frm.searchType.value.length>0)&&(frm.searchValue.value.length<1)){
        alert('검색 내용을 입력하세요.');
        frm.searchValue.focus();
        return;
    }

    if ((frm.searchType.value=="orderno")||(frm.searchType.value=="itemid")){
        if (!IsDigit(frm.searchValue.value)){
            alert('검색 내용은 숫자만 가능합니다.');
            frm.searchValue.focus();
            return;
        }
    }

    frm.submit();
}


function ShowOrderInfo(masteridx){
	var ShowOrderInfo = window.open('/common/offshop/beasong/upche_viewordermaster.asp?masteridx='+masteridx,'ShowOrderInfo','width=800,height=768,scrollbars=yes,resizable=yes');
	ShowOrderInfo.focus();
}

function switchCheckBox(comp){
    var frm = comp.form;

	if(frm.chkidx.length>1){
		for(i=0;i<frm.chkidx.length;i++){
			frm.chkidx[i].checked = comp.checked;
			AnCheckClick(frm.chkidx[i]);
		}
	}else{
		frm.chkidx.checked = comp.checked;
		AnCheckClick(frm.chkidx);
	}
}

function BaljuReprint(){
    var frm = document.frmbalju;
	var pass = false;

    if(!frm.chkidx.length){
    	pass = frm.chkidx.checked;
    }else{
        for (var i=0;i<frm.chkidx.length;i++){
    	    pass = (pass||frm.chkidx[i].checked);
    	}
    }

	if (!pass) {
		alert("재출력할 내역을 선택하세요.");
		return;
	}else{
	    var popwin = window.open("about:blank","PopBaljuList","width=800,scrollbars=yes,resizable");
	    frm.target = "PopBaljuList";
	    frm.isall.value = "";
 		frm.action = "/common/offshop/beasong/upche_reselectbaljulist.asp";
		frm.submit();
	}
}

function BaljuReprintAll(){
    var frm = document.frmbalju;

    if (confirm('미출고 내역 전체 발주서를 재출력 하시겠습니까?')){
        var popwin = window.open("about:blank","PopBaljuList","width=800,scrollbars=yes,resizable");
	    frm.target = "PopBaljuList";
	    frm.isall.value = "on";
 		frm.action = "/common/offshop/beasong/upche_reselectbaljulist.asp";
		frm.submit();
    }
}

function popMisendInput(detailidx){
    var popwin = window.open('/common/offshop/beasong/upche_popMisendInput.asp?detailidx=' + detailidx,'popMisendInput','width=600,height=768,scrollbars=yes,resizable=yes');
    popwin.focus();
}

</script>

<!-- 검색 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get" action="" onsubmit="chksubmit(); return false">
<input type="hidden" name="page" value="1">
<input type="hidden" name="menupos" value="<%= menupos %>">
<tr align="center" bgcolor="<%= adminColor("topbar") %>" >
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
	<td align="left" bgcolor="#FFFFFF">
		<select class="select" name="searchType" >
			<option value="">검색조건</option>
			<option value="orderno" <%= ChkIIF(searchType="orderno","selected","") %> >주문번호</option>
			<option value="itemid" <%= ChkIIF(searchType="itemid","selected","") %> >상품코드</option>
			<option value="buyname" <%= ChkIIF(searchType="buyname","selected","") %> >구매자</option>
			<option value="reqname" <%= ChkIIF(searchType="reqname","selected","") %> >수령인</option>
		</select>
		<input type="text" class="text" name="searchValue" value="<%= searchValue %>" size="20" maxlength="20">
		&nbsp;
		<!--사유입력여부 :-->
		<!--<select class="select" name="MisendReason">-->
		<!--	<option value="" >전체</option>-->
		<!--	<option value="NN" <%'= ChkIIF(MisendReason="NN","selected","") %> >사유미입력</option>-->
		<!--	<option value="03" <%'= ChkIIF(MisendReason="03","selected","") %> >출고지연</option>-->
			<!--<option value="05" <%'= ChkIIF(MisendReason="05","selected","") %> >품절출고불가</option>-->
			<!--<option value="02" <%'= ChkIIF(MisendReason="02","selected","") %> >주문제작</option>-->
		<!--</select>-->
		<br>
	</td>
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="검색" onClick="javascript:chksubmit();">
	</td>
</tr>
</form>
</table>

<br>
<!--
※ 출고지연의 경우 고객에게 SMS 및 안내메일 발송<br>
품절출고불가의 경우, 고객센터에서 처리
-->
<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
<tr align="center">
	<td align="left">
    	<input type="button" class="button" value="선택내역 발주서 재출력" onclick="javascript:BaljuReprint()">
		&nbsp;
    	<input type="button" class="button" value="미출고전체 발주서 재출력" onclick="javascript:BaljuReprintAll()">
    </td>
    <td align="right">
	</td>
</tr>
</table>
<!-- 액션 끝 -->

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frmbalju" method="post" >
<input type="hidden" name="mode" value="">
<input type="hidden" name="isall" value="">
<input type="hidden" name="ArrChkVal" value="">
<tr bgcolor="FFFFFF">
	<td height="25" colspan="15">
		검색결과 : <b><% = ojumun.FResultCount %></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td><input type="checkbox" name="chkAll" onClick="switchCheckBox(this);"></td>
	<td>일렬번호</td>
	<td>주문번호</td>
	<td>수령인</td>
	<td>상품코드</td>
	<td>상품명<font color="blue">&nbsp;[옵션]</font></td>
	<td>공급가</td>
	<td>판매가</td>
	<td>수량</td>
	<td>출고기준일<!-- 주문통보일 --></td>
	<td>주문확인일</td>
	<td>경과일</td>
	<!--<td>미출고사유</td>
	<td>출고예정일</td>
	<td>미출고사유<br>입력</td>-->
</tr>
<% if ojumun.FResultCount > 0 then %>
<% for i=0 to ojumun.FresultCount-1 %>
<input type="hidden" name="detailidx" value="<%= ojumun.FItemList(i).fdetailidx %>">
<tr align="center" class="a" bgcolor="#FFFFFF">
	<td><input type="checkbox" name="chkidx" value="<%= ojumun.FItemList(i).Fdetailidx %>" onClick="AnCheckClick(this);"></td>
	<td><%= ojumun.FItemList(i).fdetailidx %></td>
	<td height="25">
		<a href="javascript:ShowOrderInfo('<%= ojumun.FItemList(i).fmasteridx %>')">
		<%= ojumun.FItemList(i).Forderno %></a>
	</td>
	<td><%= ojumun.FItemList(i).FReqname %></td>
	<td><%= ojumun.fitemlist(i).fitemgubun %>-<%= FormatCode(ojumun.fitemlist(i).FitemID) %>-<%= ojumun.fitemlist(i).fitemoption %></td>
	<td align="left">
		<%= ojumun.FItemList(i).FItemname %>
		<% if (ojumun.FItemList(i).fitemoptionname<>"") then %>
		<font color="blue">[<%= ojumun.FItemList(i).fitemoptionname %>]</font>
		<% end if %>
	</td>
	<td><%= FormatNumber(ojumun.fitemlist(i).fsuplyprice,0) %></td>
	<td><%= FormatNumber(ojumun.fitemlist(i).fsellprice,0) %></td>
	<td><%= ojumun.FItemList(i).Fitemno %></td>
	<td><acronym title="<%= ojumun.FItemList(i).Fbaljudate %>"><%= left(ojumun.FItemList(i).Fbaljudate,10) %></acronym></td>
	<td><acronym title="<%= ojumun.FItemList(i).Fupcheconfirmdate %>"><%= left(ojumun.FItemList(i).Fupcheconfirmdate,10) %></acronym></td>
	<td>
	    <% if IsNULL(ojumun.FItemList(i).Fbaljudate) then %>
        D+0
        <% elseif datediff("d",(left(ojumun.FItemList(i).Fbaljudate,10)) , (left(now,10)) )>2 then %>
        <font color="red"><b>D+<%= datediff("d",(left(ojumun.FItemList(i).Fbaljudate,10)) , (left(now,10)) ) %></b></font>
        <% else %>
        D+<%= datediff("d",(left(ojumun.FItemList(i).Fbaljudate,10)) , (left(now,10)) ) %>
        <% end if %>
    </td>
	<!--<td><%'= ojumun.FItemList(i).getMisendText %></td>
	<td><%'= ojumun.FItemList(i).FMisendIpgodate %></td>
    <td>
        <%' if (ojumun.FItemList(i).isMisendAlreadyInputed) then %>
        <a href="javascript:popMisendInput('<%= ojumun.FItemList(i).Fdetailidx %>');">상세보기</a>
        <%' else %>
        <a href="javascript:popMisendInput('<%= ojumun.FItemList(i).Fdetailidx %>');"><font color="#AAAAAA">입력</font></a>
        <%' end if %>
    </td>-->
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