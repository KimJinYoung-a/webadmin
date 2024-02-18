<%@ language=vbscript %>
<% option explicit %>
<%
'#######################################################
'	2009년 01월 19일 한용민 수정
'#######################################################
%>
<!-- #include virtual="/common/incSessionBctId.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/common/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/bscclass/equipmentcls.asp"-->

<%
dim idx
	idx = request("idx")

if idx="" then idx=0
dim oequip
set oequip = new CEquipment
	oequip.FRectIdx = idx
	oequip.getOneEquipment
%>

<script language="javascript">

//사용가능한ip선택시작
function checkip(frm)
{
	if (document.frmreg.checkipform.value!="")
	{
		document.frmreg.detail_ip.value = ""
		document.frmreg.detail_ip.value = document.frmreg.checkipform.value;
	}	
}

//저장
function regEquip(frm){
	//필수입력체크
	if (frm.equip_gubun.value.length<1){
		alert('장비구분을 선택하세요.');
		frm.equip_gubun.focus();
		return;
	}

	if (frm.part_code.value.length<1){
		alert('사용구분을 선택하세요.');
		frm.part_code.focus();
		return;
	}


	if (confirm('저장 하시겠습니까?')){
		frm.submit();
	}
}

//옵션 상세 선택부분시작
function selectChange(comp){
	if (comp.name=="equip_gubun"){
		//옵션1
		if ((comp.value=="PC")||(comp.value=="NB")||(comp.value=="MO")||(comp.value=="SV")||(comp.value=="FS")){
			div_detail_quality1.style.display="inline";
		}else{
			div_detail_quality1.style.display="none";
		}

		if ((comp.name=="equip_gubun")&&(comp.value=="MO")){
			detail_quality1_name.innerText = "모니터사양 :";
			detail_quality1_etc.innerText = "(LCD 17, CRT 19)";
		}else{
			detail_quality1_name.innerText = "CPU :";
			detail_quality1_etc.innerText = "(P2.8, C2.4, AMD 1800 ..)";
		}

		//옵션2
		if ((comp.value=="PC")||(comp.value=="NB")||(comp.value=="SV")||(comp.value=="FS")){
			div_detail_quality2.style.display="inline";
		}else{
			div_detail_quality2.style.display="none";
		}
		if ((comp.value=="SC")||(comp.value=="PR")||(comp.value=="CX")){
			div_detail_quality3.style.display="inline";
		}else{
			div_detail_quality3.style.display="none";
		}
		if (comp.value=="NE"){
			div_detail_quality4.style.display="inline";
		}else{
			div_detail_quality4.style.display="none";
		}
		if (comp.value=="UP"){
			div_detail_quality5.style.display="inline";
		}else{
			div_detail_quality5.style.display="none";
		}
		
	<!--}else if (comp.name=="part_code"){
		if ((comp.value=="10")&&(frmreg.usinguserid.value.length<1)){
			frmreg.usinguserid.value = frmreg.curruserid.value;
		}else{
			//frmreg.usinguserid.value = "";
		}-->
	}
}	

</script>

<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a">
<tr>
	<td align="left">
		※ 장비 자산 리스트 추가 </strong> / 되도록 자세히 입력해 주세요. 	
	</td>
	<td align="right">		
	</td>
</tr>
</form>
</table>
<!-- 액션 끝 -->

<!--하단테이블시작-->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="#BABABA">
<form name="frmreg" method="post" action="do_equipment.asp">
<input type="hidden" name="idx" value="<%= oequip.FOneItem.Fidx %>">
<input type="hidden" name="curruserid" value="<%= session("ssBctId") %>">
<input type="hidden" name="currusername" value="<%= session("ssBctCname") %>">

<tr bgcolor="#FFFFFF">
	<td width="100" bgcolor="F4F4F4">장비코드</td>
	<td colspan="2"><%= oequip.FOneItem.getEquipCode %></td>
</tr>
<tr bgcolor="#FFFFFF">
	<td width="100" bgcolor="F4F4F4">장비구분</td>
	<td >
		<% DrawEquipMentGubun "10","equip_gubun",oequip.FOneItem.Fequip_gubun ," onchange='selectChange(frmreg.equip_gubun)'" %>
	</td>
	
</tr>
<tr bgcolor="#FFFFFF">
	<td width="100" bgcolor="F4F4F4">사용구분</td>
	<td >
		<% DrawEquipMentGubun "20","part_code",oequip.FOneItem.Fpart_code ,"" %>
	</td>
	
</tr>
<tr bgcolor="#FFFFFF">
	<td width="100" bgcolor="F4F4F4">제품명</td>
	<td colspan="2">
		<input type="text" name="equip_name" value="<%= oequip.FOneItem.Fequip_name %>" size="60" maxlength="60">
		(ex : 삼보 드림시스 74SC)
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td width="100" bgcolor="F4F4F4">시리얼번호</td>
	<td colspan="2">
		<input type="text" name="model_name" value="<%= oequip.FOneItem.Fmodel_name %>" size="60" maxlength="60">
		(ex : PN17AS)
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td width="100" bgcolor="F4F4F4">제조사</td>
	<td colspan="2">
		<input type="text" name="manufacture_company" value="<%= oequip.FOneItem.Fmanufacture_company %>" size="60" maxlength="60">
		(ex : 삼성전자, LG전자)
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td width="100" bgcolor="F4F4F4">구매처</td>
	<td colspan="2">
		<input type="text" name="buy_company_name" value="<%= oequip.FOneItem.Fbuy_company_name %>" size="60" maxlength="60">
		(ex : 삼성몰, 인터파크, DELL코리아)
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td width="100" bgcolor="F4F4F4">구매일</td>
	<td colspan="2">
		<input type="text" name="buy_date" value="<%= oequip.FOneItem.Fbuy_date %>" size="10" maxlength="10" readonly>
		<a href="javascript:calendarOpen3(frmreg.buy_date,'구매일',frmreg.buy_date.value)"><img src="/images/calicon.gif" width="21" border="0" align="middle"></a>
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td width="100" bgcolor="F4F4F4">구매가격</td>
	<td colspan="2">
		<input type="text" name="buy_sum" value="<%= oequip.FOneItem.Fbuy_sum %>" size="10" maxlength="9">
		(부가세 포함가)
		<!-- <input type="checkbox" name="" value="">부가세포함 -->
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td width="100" bgcolor="F4F4F4">상세사양</td>
	<td colspan="2">
		
	<!-- 장비구분에 따라 뿌려줌-->	
		<div id="div_detail_quality1" style="display:none">
		<table border="0" cellspacing="0" cellpadding="0" class="a">
		<tr>
			<td width="80"><span id="detail_quality1_name">CPU :</span></td>
			<td>
				<input type="text" name="detail_quality1" value="<%= oequip.FOneItem.Fdetail_quality1 %>" size="50" maxlength="50">
				<span id="detail_quality1_etc">(ex: P2.8, C2.4, AMD 1800)</span>
			</td>
		</tr>
		</table>
		</div>

		<div id="div_detail_quality2" style="display:none">
		<table border="0" cellspacing="0" cellpadding="0" class="a">
		<tr>
			<td width="80">Memory :</td>
			<td>
				<input type="text" name="detail_quality2" value="<%= oequip.FOneItem.Fdetail_quality2 %>" size="50" maxlength="50">
				(ex: 512M, 1G)
			</td>
		</tr>
		</table>
		</div>
		
		<div id="div_detail_quality3" style="display:none">
		<table border="0" cellspacing="0" cellpadding="0" class="a">
		<tr>
			<td width="80">해상도 :</td>
			<td>
				<input type="text" name="detail_quality3" value="<%= oequip.FOneItem.Fdetail_quality2 %>" size="50" maxlength="50">
				(ex: 600DPI, 1200DPI)
			</td>
		</tr>
		</table>
		</div>
		
		<div id="div_detail_quality4" style="display:none">
		<table border="0" cellspacing="0" cellpadding="0" class="a">
		<tr>
			<td width="80">종류 :</td>
			<td>
				<input type="text" name="detail_quality4" value="<%= oequip.FOneItem.Fdetail_quality2 %>" size="50" maxlength="50">
				(ex: 15포트허브, 5포트IP공유기)
			</td>
		</tr>
		</table>
		</div>
		
		<div id="div_detail_quality5" style="display:none">
		<table border="0" cellspacing="0" cellpadding="0" class="a">
		<tr>
			<td width="80">용도 :</td>
			<td>
				<input type="text" name="detail_quality5" value="<%= oequip.FOneItem.Fdetail_quality2 %>" size="50" maxlength="50">
				(ex: 컴퓨터부품, 책)
			</td>
		</tr>
		</table>
		</div>

		<table border="0" cellspacing="0" cellpadding="0" class="a">
		<tr>
			<td width="80">사양 :</td>
			<td>
				<textarea cols="60" rows="4" name="detail_qualityetc"><%= oequip.FOneItem.Fdetail_qualityetc %></textarea>
			</td>
		</tr>
		<!--<tr>
			<td width="80">IP :</td>
			<td>
				<input type="text" name="detail_ip" value="<%= oequip.FOneItem.Fdetail_ip %>" size="16" maxlength="16">		
				<%' DrawipGubun "equip_gubun" %>			
			</td>
		</tr>-->
		<tr>
			<td></td>
			<td>
				<%' DrawipGubun2 "equip_gubun" %>							
			</td>
		</tr>
		</table>

	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td width="100" bgcolor="F4F4F4">사용자 ID</td>
	<td colspan="2">
		<% drawpartneruser "usinguserid", oequip.FOneItem.Fusinguserid ,"" %>
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td width="100" bgcolor="F4F4F4">기타사항<br>제품위치</td>
	<td colspan="2">
		<textarea cols="80" rows="5" name="etc_str"><%= oequip.FOneItem.Fetc_str %></textarea><br>
		<font size="2">(ex : 3층 사장님자리 왼쪽 모니터)</font>
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td colspan="3" align="center"><input type="button" value="저장" onclick="regEquip(frmreg);" class="button"></td>
</tr>
</form>
</table>

<%
set oequip = Nothing
%>

<script>
	selectChange(frmreg.equip_gubun);
	selectChange(frmreg.part_code);
</script>

<!-- #include virtual="/common/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->