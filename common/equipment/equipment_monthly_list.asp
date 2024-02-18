<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description : 월별 장비자산관리
' History : 				 이상구 생성
'			2016년 04월 26일 한용민 수정
'####################################################
%>
<!-- #include virtual="/common/incSessionAdminOrShop.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/common/lib/commonbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/common/equipment/equipment_cls.asp"-->
<!-- #include virtual="/lib/classes/admin/tenbyten/TenByTenDepartmentCls.asp"-->
<%
dim page, research, yyyy, mm, tmpDate, accountGubun, BIZSECTION_CD, BIZSECTION_NM, showsum, i, prev_account_gubun
dim tot_buy_cost, tot_prev_remain_value, tot_buy_cost_this_month, tot_month_down_value, tot_month_out_value, tot_month_remain_value
dim tot_buy_cost_sum, tot_prev_remain_value_sum, tot_buy_cost_this_month_sum, tot_month_down_value_sum, tot_month_out_value_sum
dim tot_month_remain_value_sum
	page = requestcheckvar(request("page"),10)
	research = requestcheckvar(Request("research"),2)
	yyyy = requestcheckvar(request("yyyy1"),4)
	mm = requestcheckvar(request("mm1"),2)
	accountGubun = requestcheckvar(request("accountGubun"),5)
	showsum = requestcheckvar(request("showsum"),5)
	BIZSECTION_CD = requestcheckvar(Request("BIZSECTION_CD"),15)
	BIZSECTION_NM = requestcheckvar(Request("BIZSECTION_NM"),55)

if (yyyy = "") then
	tmpDate = DateAdd("m", -1, Now())
	yyyy = Year(tmpDate)
	mm = Month(tmpDate)
	if (mm < 10) then
		mm = "0" & mm
	end if
end if
if page="" then page=1
if (research = "") then
	''onlyusing = "Y"
end if

dim oequip
set oequip = new CEquipment
	oequip.FPageSize = 50
	oequip.FCurrPage = page
	oequip.FRectYYYYMM = yyyy & "-" & mm
	oequip.FRectAccountGubun = accountGubun
	oequip.FRectBIZSECTION_CD = BIZSECTION_CD
	oequip.getEquipmentMonthlyList

dim oequipsum
set oequipsum = new CEquipment
	oequipsum.FPageSize = 50
	oequipsum.FCurrPage = 1
	oequipsum.FRectYYYYMM = yyyy & "-" & mm
	oequipsum.FRectAccountGubun = accountGubun
	oequipsum.FRectBIZSECTION_CD = BIZSECTION_CD

	if (showsum = "Y") then
		oequipsum.getEquipmentMonthlySUM
	end if

%>
<script type="text/javascript">

function NextPage(page){
	frm.page.value= page;
	frm.submit();
}

function jsMakeMonthlyData() {
	var frm = document.frmAct;

	<% if oequip.FTotalCount > 0 then %>
		<% if C_ADMIN_AUTH then %>
			alert("[관리자권한]기존에 작성된 자산은 빼고 추가된 자산만 추가 작성 합니다.\n계속 하실려면 다음 얼럿에서 확인을 누르세요.");
		<% else %>
			alert("이미 작성되었습니다.(재작성불가)\n다시 작성하시려면 관리자 문의 하세요.");
			return;
		<% end if %>
	<% end if %>

	if (confirm("작성(<%= yyyy & "-" & mm %>) 하시겠습니까?") == true) {
		frm.mode.value = "makemonthlydata";
		frm.yyyymm.value = "<%= yyyy & "-" & mm %>";
		frm.submit();
	}
}

function jsGetPart(){
	var winP = window.open('/admin/linkedERP/Biz/popGetBizOne.asp','popGetBizOne','width=600, height=500, resizable=yes, scrollbars=yes');
	winP.focus();
}

//자금관리부서 등록
function jsSetPart(selUP, sPNM){
	document.frm.BIZSECTION_CD.value = selUP;
	document.frm.BIZSECTION_NM.value = sPNM;
}

function jsClearPart() {
	document.frm.BIZSECTION_CD.value = "";
	document.frm.BIZSECTION_NM.value = "";
}

function fnSearch(accountGubun, BIZSECTION_CD, BIZSECTION_NM){
	frm.accountGubun.value = accountGubun;
	frm.BIZSECTION_CD.value = BIZSECTION_CD;
	frm.BIZSECTION_NM.value = BIZSECTION_NM;

	frm.submit();
}

function edit_equipmentreg_monthly(yyyymm, idx){
	var edit_equipmentreg_monthly = window.open('/common/equipment/pop_equipmentreg_monthly.asp?yyyymm='+yyyymm+'&idx='+idx,'edit_equipmentreg_monthly','width=800,height=400,scrollbars=yes,resizable=yes');
	edit_equipmentreg_monthly.focus();
}

</script>

<!-- 검색 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get" action="">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="research" value="on">
<input type="hidden" name="page" value="1">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="3" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
	<td align="left">
		연월 : <% Call DrawYMBox(yyyy, mm) %>
		&nbsp;&nbsp;
		자산구분 : <% drawEquipmentAccountCode "accountGubun" ,accountGubun, "" %>
		&nbsp;&nbsp;
		손익부서 :
		<input type="text" name="BIZSECTION_CD" value="<%= BIZSECTION_CD %>" size="15"  class="text_ro"> <input type="text" name="BIZSECTION_NM" value="<%= BIZSECTION_NM %>" class="text_ro" size="15">
		<input type="button" class="button" value="X" onClick="jsClearPart()">
		<a href="javascript:jsGetPart();"> <img src="/images/icon_search.jpg" border="0"></a>
	</td>
	<td rowspan="3" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="검색" onClick="NextPage('');">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
		<input type="checkbox" name="showsum" value="Y" <% if (showsum = "Y") then %>checked<% end if %> > 통계보기
	</td>
</tr>
</form>
</table>
<!-- 검색 끝 -->

<br>

<% if oequipsum.FResultCount > 0 then %>
	<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td width="60">연월</td>
		<td width="80">자산구분</td>
		<td width="100">손익부서</td>
		<td width="100">구매금액</td>
		<td width="100">전월<br>잔존가치</td>
		<td width="100">당월<br>구매액</td>
		<td width="100">당월<br>감가상각</td>
		<td width="100">당월<br>폐기액</td>
		<td width="100">월말<br>잔존가치</td>
		<td></td>
	</tr>
	<% for i=0 to oequipsum.FResultCount - 1 %>
	<%
	if (i = 0) then
		prev_account_gubun = oequipsum.FItemList(i).Faccount_gubun
	elseif (prev_account_gubun <> oequipsum.FItemList(i).Faccount_gubun) then
		prev_account_gubun = oequipsum.FItemList(i).Faccount_gubun
		%>
		<tr align="center" bgcolor="#FFFFFF" onMouseOver= this.style.background='f1f1f1'; onMouseOut=this.style.background='#ffffff'; height="25">
			<td colspan="3">소계</td>
			<td align="right"><%= FormatNumber(tot_buy_cost, 0) %></td>
			<td align="right"><%= FormatNumber(tot_prev_remain_value, 0) %></td>
			<td align="right"><%= FormatNumber(tot_buy_cost_this_month, 0) %></td>
			<td align="right"><%= FormatNumber(tot_month_down_value, 0) %></td>
			<td align="right"><%= FormatNumber(tot_month_out_value, 0) %></td>
			<td align="right"><%= FormatNumber(tot_month_remain_value, 0) %></td>
			<td></td>
		</tr>
		<%
		tot_buy_cost = 0
		tot_prev_remain_value = 0
		tot_buy_cost_this_month = 0
		tot_month_down_value = 0
		tot_month_out_value = 0
		tot_month_remain_value = 0
	end if
	
	tot_buy_cost = tot_buy_cost + oequipsum.FItemList(i).Ftot_buy_cost
	tot_prev_remain_value = tot_prev_remain_value + oequipsum.FItemList(i).Ftot_prev_remain_value
	tot_buy_cost_this_month = tot_buy_cost_this_month + oequipsum.FItemList(i).Ftot_buy_cost_this_month
	tot_month_down_value = tot_month_down_value + oequipsum.FItemList(i).Ftot_month_down_value
	tot_month_out_value = tot_month_out_value + oequipsum.FItemList(i).Ftot_month_out_value
	tot_month_remain_value = tot_month_remain_value + oequipsum.FItemList(i).Ftot_month_remain_value
	
	tot_buy_cost_sum = tot_buy_cost_sum + oequipsum.FItemList(i).Ftot_buy_cost
	tot_prev_remain_value_sum = tot_prev_remain_value_sum + oequipsum.FItemList(i).Ftot_prev_remain_value
	tot_buy_cost_this_month_sum = tot_buy_cost_this_month_sum + oequipsum.FItemList(i).Ftot_buy_cost_this_month
	tot_month_down_value_sum = tot_month_down_value_sum + oequipsum.FItemList(i).Ftot_month_down_value
	tot_month_out_value_sum = tot_month_out_value_sum + oequipsum.FItemList(i).Ftot_month_out_value
	tot_month_remain_value_sum = tot_month_remain_value_sum + oequipsum.FItemList(i).Ftot_month_remain_value
	%>
	<tr align="center" bgcolor="#FFFFFF" onMouseOver= this.style.background='f1f1f1'; onMouseOut=this.style.background='#ffffff'; height="25">
		<td><%= oequipsum.FItemList(i).Fyyyymm %></td>
		<td><a href="javascript:fnSearch('<%= oequipsum.FItemList(i).Faccount_gubun %>', '<%= oequipsum.FItemList(i).FBIZSECTION_CD %>', '<%= oequipsum.FItemList(i).FBIZSECTION_NM %>')"><%= oequipsum.FItemList(i).GetAccountGubunName %></a></td>
		<td><a href="javascript:fnSearch('<%= oequipsum.FItemList(i).Faccount_gubun %>', '<%= oequipsum.FItemList(i).FBIZSECTION_CD %>', '<%= oequipsum.FItemList(i).FBIZSECTION_NM %>')"><%= oequipsum.FItemList(i).FBIZSECTION_NM %></a></td>
		<td align="right"><%= FormatNumber(oequipsum.FItemList(i).Ftot_buy_cost, 0) %></td>
		<td align="right"><%= FormatNumber(oequipsum.FItemList(i).Ftot_prev_remain_value, 0) %></td>
		<td align="right"><%= FormatNumber(oequipsum.FItemList(i).Ftot_buy_cost_this_month, 0) %></td>
		<td align="right"><%= FormatNumber(oequipsum.FItemList(i).Ftot_month_down_value, 0) %></td>
		<td align="right"><%= FormatNumber(oequipsum.FItemList(i).Ftot_month_out_value, 0) %></td>
		<td align="right"><%= FormatNumber(oequipsum.FItemList(i).Ftot_month_remain_value, 0) %></td>
		<td></td>
	</tr>
	<% next %>

	<tr align="center" bgcolor="#FFFFFF" onMouseOver= this.style.background='f1f1f1'; onMouseOut=this.style.background='#ffffff'; height="25">
		<td colspan="3">소계</td>	
		<td align="right"><%= FormatNumber(tot_buy_cost, 0) %></td>
		<td align="right"><%= FormatNumber(tot_prev_remain_value, 0) %></td>
		<td align="right"><%= FormatNumber(tot_buy_cost_this_month, 0) %></td>
		<td align="right"><%= FormatNumber(tot_month_down_value, 0) %></td>
		<td align="right"><%= FormatNumber(tot_month_out_value, 0) %></td>
		<td align="right"><%= FormatNumber(tot_month_remain_value, 0) %></td>
		<td></td>
	</tr>
	<tr align="center" bgcolor="#EEEEEE" onMouseOver= this.style.background='f1f1f1'; onMouseOut=this.style.background='#ffffff'; height="25">
		<td colspan="3">합계</td>
		<td align="right"><%= FormatNumber(tot_buy_cost_sum, 0) %></td>
		<td align="right"><%= FormatNumber(tot_prev_remain_value_sum, 0) %></td>
		<td align="right"><%= FormatNumber(tot_buy_cost_this_month_sum, 0) %></td>
		<td align="right"><%= FormatNumber(tot_month_down_value_sum, 0) %></td>
		<td align="right"><%= FormatNumber(tot_month_out_value_sum, 0) %></td>
		<td align="right"><%= FormatNumber(tot_month_remain_value_sum, 0) %></td>
		<td></td>
	</tr>
	</table>
	<br>
<% end if %>

<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a">
<tr>
	<td align="left">
		<input type="button" class="button" onclick="pageexcelsheet11();" value="엑셀출력">
	</td>
	<td align="right">
		<%
		if C_ADMIN_AUTH then		'// 관리자만
		%>
			<input type="button" class="button" onclick="jsMakeMonthlyData();" value="작성">
		<% end  if %>
	</td>
</tr>
</table>
<!-- 액션 끝 -->

<!-- 리스트 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="17">
		검색결과 : <b><%= oequip.FTotalCount %></b>
		&nbsp;
		페이지 : <b><%= page %>/ <%= oequip.FTotalPage %></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width="60">연월</td>
	<td width="150">장비코드</td>
	<td width="80">자산구분</td>
	<td width="100">손익부서</td>
	<td width="80">구매일자</td>
	<td width="70">구매금액</td>
	<td width="70">누적월</td>
	<td width="70">전월<br>잔존가치</td>
	<td width="70">당월<br>구매액</td>
	<td width="70">당월<br>감가상각</td>
	<td width="70">당월<br>폐기액</td>
	<td width="70">월말<br>잔존가치</td>
	<td width="70">상태</td>
	<td width="80">등록일</td>
	<td>비고</td>
</tr>
<% if oequip.FResultCount > 0 then %>
<% for i=0 to oequip.FResultCount - 1 %>
<form name=frm_<%= i %> method="post">
<input type="hidden" name="idx" value="<%= oequip.FItemList(i).Fidx %>">
<tr align="center" bgcolor="#FFFFFF" onMouseOver= this.style.background='f1f1f1'; onMouseOut=this.style.background='#ffffff'; height="25">
	<td><%= oequip.FItemList(i).Fyyyymm %></td>
	<td><%= oequip.FItemList(i).Fequip_code %></td>
	<td><%= oequip.FItemList(i).GetAccountGubunName %></td>
	<td><%= oequip.FItemList(i).FBIZSECTION_NM %></td>
	<td><%= oequip.FItemList(i).Fbuy_date %></td>
	<td align="right"><%= FormatNumber(oequip.FItemList(i).Fbuy_cost, 0) %></td>
	<td align="right"><%= oequip.FItemList(i).GetAccMonthCount %></td>

	<td align="right"><%= FormatNumber(oequip.FItemList(i).Fprev_remain_value, 0) %></td>
	<td align="right"><%= FormatNumber(oequip.FItemList(i).GetBuyThisMonth, 0) %></td>
	<td align="right"><%= FormatNumber(oequip.FItemList(i).Fmonth_down_value, 0) %></td>
	<td align="right"><%= FormatNumber(oequip.FItemList(i).GetDiscardThisMonth, 0) %></td>
	<td align="right"><%= FormatNumber(oequip.FItemList(i).GetRemainThisMonth, 0) %></td>
	<td><%= oequip.FItemList(i).Fstate_name %></td>
	<td><%= Left(oequip.FItemList(i).Fregdate, 10) %></td>
	<td>
		<input type="button" onclick="edit_equipmentreg_monthly('<%= oequip.FItemList(i).Fyyyymm %>','<%= oequip.FItemList(i).Fidx %>');" value="수정" class="button">
	</td>
</tr>
</form>
<% next %>

<!--
<tr bgcolor="<%= adminColor("tabletop") %>">
	<td align="center" colspan=1>총계</td>
	<td align="right"></td>
</tr>
-->

<tr height="25" bgcolor="FFFFFF">
	<td colspan="17" align="center">
    	<% if oequip.HasPreScroll then %>
		<a href="javascript:NextPage('<%= oequip.StartScrollPage-1 %>')">[pre]</a>
		<% else %>
			[pre]
		<% end if %>

		<% for i=0 + oequip.StartScrollPage to oequip.FScrollCount + oequip.StartScrollPage - 1 %>
			<% if i>oequip.FTotalpage then Exit for %>
			<% if CStr(page)=CStr(i) then %>
			<font color="red">[<%= i %>]</font>
			<% else %>
			<a href="javascript:NextPage('<%= i %>')">[<%= i %>]</a>
			<% end if %>
		<% next %>

		<% if oequip.HasNextScroll then %>
			<a href="javascript:NextPage('<%= i %>')">[next]</a>
		<% else %>
			[next]
		<% end if %>
	</td>
</tr>
<% else %>
	<tr bgcolor="#FFFFFF">
		<td colspan="15" align="center" class="page_link">[검색결과가 없습니다.]</td>
	</tr>
<% end if %>

</table>

<form name="frmAct" method="post" action="do_equipment.asp">
	<input type="hidden" name="mode" value="">
	<input type="hidden" name="yyyymm" value="">
</form>

<%
set oequip = Nothing
set oequipsum = Nothing
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/common/lib/commonbodytail.asp"-->
