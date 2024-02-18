<%@ language=vbscript %>
<% option explicit %>

<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
%>
<%
'####################################################
' Description : 세금계산서
' History : 서동석 생성
'			2017.04.12 한용민 수정(보안관련처리)
'####################################################
%>
<!-- #include virtual="/designer/incSessionDesigner.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/designer/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/offshop/fran_chulgojungsancls.asp"-->
<!-- #include virtual="/lib/classes/offshop/offshopchargecls.asp"-->
<!-- #include virtual="/lib/classes/partners/partnerusercls.asp"-->

<%
dim idx
idx = requestCheckVar(request("idx"),10)

dim obj
set obj = new CFranjungsan
obj.FRectidx = idx
obj.getOneFranJungsan


if obj.FResultCount<1 then
	response.write "<script type='text/javascript'>alert('정산정보가 없습니다.');</script>"
	response.write "<script type='text/javascript'>window.close()</script>"
	dbget.close()	:	response.End
end if

if (obj.FoneItem.FStateCd>"0") and (obj.FoneItem.FStateCd<"4") then
	stypename = "세금계산서"
else
	response.write "<script type='text/javascript'>alert('세금계산서 혹은 계산서만 발행 가능합니다. - 이미 발행 하였거나 발행할 정보가 없습니다.');</script>"
	response.write "<script type='text/javascript'>window.close()</script>"
	dbget.close()	:	response.End
end if

dim objShop, ogroup
dim stypename

set objShop = new COffShopChargeUser
objShop.FRectShopID = obj.FoneItem.Fshopid
objShop.GetOffShopList

set ogroup = new CPartnerGroup
ogroup.FRectGroupid = objShop.FItemList(0).Fgroupid
ogroup.GetOneGroupInfo

'rw objShop.FItemList(0).Fgroupid

dim jungsan_hpall, jungsan_hp1,jungsan_hp2,jungsan_hp3
jungsan_hpall = ogroup.FOneItem.Fjungsan_hp
If Not IsNull(jungsan_hpall) Then 
	jungsan_hpall = split(jungsan_hpall,"-")
	if UBound(jungsan_hpall) >= 2 then
		jungsan_hp1 = jungsan_hpall(0)
		jungsan_hp2 = jungsan_hpall(1)
		jungsan_hp3 = jungsan_hpall(2)
	end if
End If 

If IsNull(ogroup.FOneItem.Fcompany_no) Then 
	ogroup.FOneItem.Fcompany_no = ""
End If 

Dim totalCost, supplyCost, vatCost

totalCost	= CDbl(obj.FoneItem.Ftotalsum)
supplyCost	= Round(totalCost / 1.1)
vatCost		= totalCost - supplyCost

%>

<script type='text/javascript'>

function ActTaxReg(frm){
//alert('점검중입니다');
//return;
	if (frm.biz_no.value.length!=10){
		alert('사업자 등록 번호가 올바르지 않거나 등록되어 있지 않습니다. - 가맹점정보 수정후 사용하세요.');
		return;
	}

	if (frm.corp_nm.value.length<1){
		alert('사업자 명이 등록되어 있지 않습니다. - 가맹점정보 수정후 사용하세요.');
		return;
	}

	if (frm.ceo_nm.value.length<1){
		alert('대표자 명이 등록되어 있지 않습니다. - 가맹점정보 수정후 사용하세요.');
		return;
	}

	if (frm.biz_status.value.length<1){
		alert('업태가 등록되어 있지 않습니다. - 가맹점정보 수정후 사용하세요.');
		return;
	}

	if (frm.biz_type.value.length<1){
		alert('업종이 등록되어 있지 않습니다. - 가맹점정보 수정후 사용하세요.');
		return;
	}

	if (frm.addr.value.length<1){
		alert('사업장 주소가 등록되어 있지 않습니다. - 가맹점정보 수정후 사용하세요.');
		return;
	}

	if (frm.dam_nm.value.length<1){
		alert('담당자 성명이 등록되어 있지 않습니다. - 가맹점정보 수정후 사용하세요.');
		return;
	}

	if (frm.email.value.length<1){
		alert('담당자 이메일이 등록되어 있지 않습니다. - 가맹점정보 수정후 사용하세요.');
		return;
	}

	if (frm.write_date.value.length<1){
		alert('계산서 발행일 입력 후 사용하세요.');
		return;
	}


	if (confirm('<%= stypename %> 를 발행 하시겠습니까?')){
		frm.submit();
	}
}

</script>

<!-- 표 상단바 시작-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
   	<tr height="10" valign="bottom">
        <td width="10" align="right"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
        <td background="/images/tbl_blue_round_02.gif"></td>
        <td width="10" align="left" ><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
	</tr>
	<tr height="25" valign="top">
        <td background="/images/tbl_blue_round_04.gif"></td>
        <td>
        	<img src="/images/icon_star.gif" width="16" height="16" align="absbottom">
        	<strong>전자 <%= stypename %> 발행</strong>
        </td>
        <td background="/images/tbl_blue_round_05.gif"></td>
	</tr>
</table>
<!-- 표 상단바 끝-->

<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frm" method="post" action="pop_offshop_TaxReg_Proc.asp">
	<input type=hidden name=jungsanid value="<%=obj.FoneItem.FIdx%>">
	<input type=hidden name=jungsanname value="<%=obj.FoneItem.Ftitle%>">
	<input type=hidden name=jungsangubun value="OFFSHOP">
	<input type=hidden name=makerid value="<%=obj.FoneItem.Fshopid%>">
	
	<input type=hidden name=biz_no value="<%= replace(replace(ogroup.FOneItem.Fcompany_no,"-","")," ","") %>" >
	<input type=hidden name=corp_nm value="<%= ogroup.FOneItem.FCompany_name %>">
	<input type=hidden name=ceo_nm value="<%= ogroup.FOneItem.Fceoname %>">
	<input type=hidden name=biz_status value="<%= ogroup.FOneItem.Fcompany_uptae %>">
	<input type=hidden name=biz_type value="<%= ogroup.FOneItem.Fcompany_upjong %>">
	
	
	<input type=hidden name=addr value="<%= ogroup.FOneItem.Fcompany_address %> <%= ogroup.FOneItem.Fcompany_address2 %>">
	<input type=hidden name=dam_nm value="<%= ogroup.FOneItem.Fjungsan_name %>">
	<input type=hidden name=email value="<%= ogroup.FOneItem.Fjungsan_email %>">
	<input type=hidden name=hp_no1 value="<%= jungsan_hp1 %>">
	<input type=hidden name=hp_no2 value="<%= jungsan_hp2 %>">
	<input type=hidden name=hp_no3 value="<%= jungsan_hp3 %>">
	
	<input type=hidden name=sb_type value="01"> <!-- 매출 01 매입 02 -->
	<input type=hidden name=tax_type value="01"> <!-- 세금계산서 01 -->
	<input type=hidden name=bill_type value="18"> <!-- 영수 01 청구 18 -->
	<input type=hidden name=pc_gbn value="C"> <!-- 개인 P 기업 C -->
	
	<input type=hidden name=item_count value="1">
	<input type=hidden name=item_nm value="<%=obj.FoneItem.Ftitle%>">
	<input type=hidden name=item_qty value="1">
	<input type=hidden name=item_price value="<%=supplyCost%>">
	<input type=hidden name=item_amt value="<%=supplyCost%>">
	<input type=hidden name=item_vat value="<%=vatCost%>">
	<input type=hidden name=item_remark value="">
	
	<input type=hidden name=credit_amt value="<%=totalCost%>">

	<!-- DEV 1000394, REAL 244730, ON 261744 -->
<!-- 
	<input type=hidden name=cur_u_user_no value="261744"> 
	<input type=hidden name=cur_dam_nm value="이문재">
	<input type=hidden name=cur_email value="moon@10x10.co.kr">
	<input type=hidden name=cur_hp_no1 value="000">
	<input type=hidden name=cur_hp_no2 value="000">
	<input type=hidden name=cur_hp_no3 value="0000">
 -->

	<input type=hidden name=cur_u_user_no value="261748">
	<input type=hidden name=cur_dam_nm value="신희영">
	<input type=hidden name=cur_email value="shyoung@10x10.co.kr">
	<input type=hidden name=cur_hp_no1 value="010">
	<input type=hidden name=cur_hp_no2 value="4260">
	<input type=hidden name=cur_hp_no3 value="0622">

    <tr align="center" bgcolor="#FFFFFF">
		<td colspan="2">
		* 2005년 3월분 정산분(발행일 3월 31일)부터는 전자 <%= stypename %> 발행을 사용하셔야 합니다.
		</td>
	</tr>
    <tr bgcolor="<%= adminColor("tabletop") %>">
   		<td height="20" colspan="2">
	   		<img src="/images/icon_arrow_down.gif" width="16" height="16" align="absbottom">
	   		<strong>전자 <%= stypename %> 발행방법</strong>
	   		&nbsp;&nbsp;&nbsp;&nbsp;
	   		<a href="http://www.neoport.net" target="_blank"><font color="blue">>>네오포트 회원가입하기</font></a>
   		</td>
 	</tr>
	<tr bgcolor="#FFFFFF">
		<td colspan="2">
			<img src="/images/icon_num01.gif" width="16" height="16" align="absbottom">
			<font color="red"><b>www.neoport.net에 회원가입(회원가입무료)</b></font><br>
				&nbsp;&nbsp;1.네오포트에 기업회원으로 무료가입하기기 바랍니다.(사업자번호 정확히 입력)<br>
				&nbsp;&nbsp;2.인증서는 구매하실 필요 없습니다.<br>
			<img src="/images/icon_num02.gif" width="16" height="16" align="absbottom">
			<font color="red"><b>이용료 결제(건수 충전)</b></font><br>
				&nbsp;&nbsp;1.네오포트에 로그인 후, 이용료를 구매하세요.(건수 충전)<br>
				&nbsp;&nbsp;2.메인화면 오른쪽 보이는 "서비스/제품구매"로 들어가시면 됩니다.<br>
				&nbsp;&nbsp;3.건당 이용료는 200원이며, 원하시는 건수를 미리 충전하시면 됩니다.<br>
			<img src="/images/icon_num03.gif" width="16" height="16" align="absbottom">
			<font color="red"><b>전자(세금)계산서 발행</b></font><br>
				&nbsp;&nbsp;1.1번과 2번을 완료하시면, 전자(세금)계산서 발행이 가능합니다.<br>
				&nbsp;&nbsp;2.발행은 꼭 텐바이텐 어드민에서 해주셔야 자동처리가 됩니다.
		</td>
	</tr>
    <tr bgcolor="<%= adminColor("tabletop") %>">
   		<td colspan="2" height="20" valign="middle">
	   		<img src="/images/icon_arrow_down.gif" width="16" height="16" align="absbottom">
	   		<strong>등록된 사업자정보 확인</strong>
   		</td>
 	</tr>
	<tr bgcolor="#FFFFFF">
		<td bgcolor="#F3F3FF" width="30%">사업자명</td>
		<td><%= ogroup.FOneItem.FCompany_name %></td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td bgcolor="#F3F3FF">대표자명</td>
		<td><%= ogroup.FOneItem.Fceoname %></td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td bgcolor="#F3F3FF">사업자번호</td>
		<td><%= ogroup.FOneItem.Fcompany_no %></td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td bgcolor="#F3F3FF">과세구분</td>
		<td><%= ogroup.FOneItem.Fjungsan_gubun %></td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td bgcolor="#F3F3FF">사업장소재지</td>
		<td><%= ogroup.FOneItem.Fcompany_address %>&nbsp;<%= ogroup.FOneItem.Fcompany_address2 %></td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td bgcolor="#F3F3FF">업태</td>
		<td><%= ogroup.FOneItem.Fcompany_uptae %></td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td bgcolor="#F3F3FF">업종</td>
		<td><%= ogroup.FOneItem.Fcompany_upjong %></td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td bgcolor="#FFFFFF">계산서발행일</td>
	<% if (obj.FoneItem.FStateCd>"0") and (obj.FoneItem.FStateCd<"4") then %>
		<td><input type=text name=write_date value="<%=Left(obj.FoneItem.Ftaxdate,10)%>" size="10" maxlength=10 readonly ><a href="javascript:calendarOpen(frm.write_date);"><img src="/images/calicon.gif" border=0 align=absmiddle></a></td>
	<% else %>
		<td><input type=text name=write_date value="<%=Left(obj.FoneItem.Ftaxdate,10)%>" size="10" maxlength=10 readonly style="border:0"></td>
	<% end if %>

	</tr>
	<tr bgcolor="#FFFFFF">
		<td bgcolor="#FFFFFF">발행금액</td>
		<td><b><%=formatNumber(totalCost,0)%></b> (공급가 : <%=FormatNumber(supplyCost,0) %> 부가세: <%=FormatNumber(vatCost,0) %>)</td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td colspan="2">
			&nbsp;&nbsp;<b>* 매월 12일 까지 발행시 : 정상발행</b><br>
			&nbsp;&nbsp;<b>* 매월 13일 이후 발행시 : 이월발행(입금처리도 이월(15일)됩니다.)</b>
		</td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td bgcolor="#F3F3FF">정산담당자명</td>
		<td><%= ogroup.FOneItem.Fjungsan_name %></td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td bgcolor="#F3F3FF">정산담당자E-mail</td>
		<td><%= ogroup.FOneItem.Fjungsan_email %></td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td bgcolor="#F3F3FF">정산담당자 핸드폰번호</td>
		<td><%= ogroup.FOneItem.Fjungsan_hp %></td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td colspan="2">
			&nbsp;&nbsp;* 가맹점정보를 확인하시고, 미입력된 정보는 어드민 업체정보수정에서 수정후 진행하시기 바랍니다.<br>
			&nbsp;&nbsp;* 정산담당자의 정보를 입력하시면, 세금계산서의 발행상황을 E-mail과 문자서비스로 알려드립니다.
		</td>
	</tr>
</table>

<!-- 표 하단바 시작-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
    <tr valign="bottom" height="25">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td valign="bottom" align="center">
        	<input type=button value="전자 <%= stypename %> 발행" onClick="ActTaxReg(frm)">
        </td>
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
    </tr>
    <tr valign="top" height="10">
        <td width="10" align="right"><img src="/images/tbl_blue_round_07.gif" width="10" height="10"></td>
        <td background="/images/tbl_blue_round_08.gif"></td>
        <td width="10" align="left"><img src="/images/tbl_blue_round_09.gif" width="10" height="10"></td>
    </tr>
</form>
</table>
<!-- 표 하단바 끝-->



<%
set obj = Nothing
set objShop = Nothing
set ogroup = Nothing
%>

<script language=javascript>
function SvcErrMsg(){
    //alert('이번달 계산서 마감은 4월 14일(월) 까지 연장합니다. ');
}
window.onload = SvcErrMsg;
</script>
<!-- #include virtual="/designer/lib/designerbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
