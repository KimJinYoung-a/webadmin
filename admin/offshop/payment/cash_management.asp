<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  오프라인 현금매출정산관리
' History : 2013.10.24 한용민 생성
'####################################################
%>
<!-- #include virtual="/common/incSessionAdminOrShop.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/common/lib/commonbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/offshop/payment/payment_cls.asp"-->
<%
dim page,shopid,yyyy1,mm1,dd1,yyyy2,mm2,dd2, fromDate,toDate
dim datefg , i, ToTcashsum, intLoop, isedityn, inc3pl
	shopid = requestCheckVar(request("shopid"),32)
	page = requestCheckVar(request("page"),10)
	if page="" then page=1
	yyyy1 = requestCheckVar(request("yyyy1"),4)
	mm1 = requestCheckVar(request("mm1"),2)
	dd1 = requestCheckVar(request("dd1"),2)
	yyyy2 = requestCheckVar(request("yyyy2"),4)
	mm2 = requestCheckVar(request("mm2"),2)
	dd2 = requestCheckVar(request("dd2"),2)
	datefg = requestCheckVar(request("datefg"),10)
    inc3pl = requestCheckVar(request("inc3pl"),32)

if datefg = "" then datefg = "maechul"

if (yyyy1="") then
	fromDate = DateSerial(Cstr(Year(now())), Cstr(Month(now())), Cstr(day(now()))-14)
else
	fromDate = DateSerial(yyyy1, mm1, dd1)
end if

if (yyyy2="") then yyyy2 = Cstr(Year(now()))
if (mm2="") then mm2 = Cstr(Month(now()))
if (dd2="") then dd2 = Cstr(day(now()))

toDate = DateSerial(yyyy2, mm2, dd2+1)

yyyy1 = left(fromDate,4)
mm1 = Mid(fromDate,6,2)
dd1 = Mid(fromDate,9,2)

'/매장
if (C_IS_SHOP) then

	'//직영점일때
	if C_IS_OWN_SHOP then

		'/어드민권한 점장 미만
		'if getlevel_sn("",session("ssBctId")) > 6 then
			shopid = C_STREETSHOPID		'"streetshop011"
		'end if
	else
		shopid = C_STREETSHOPID
	end if
else
	'/업체
	if (C_IS_Maker_Upche) then

	else
		if (Not C_ADMIN_USER) then
		    shopid = "X"                ''다른매장조회 막음.
		else
		end if
	end if
end if

dim opayment, opaymentetc
set opayment = new Cpayment
	opayment.FRectShopID = shopid
	opayment.FRectStartDay = fromDate
	opayment.FRectEndDay = toDate
	opayment.frectdatefg = datefg
	opayment.FRectInc3pl = inc3pl
	opayment.FPageSize = 500
	opayment.FCurrPage = 1

	if shopid<>"" then
		opayment.Getcash_management
	else
		response.write "<script language='javascript'>"
		response.write "alert('매장을 선택하신 후 검색하세요.');"
		response.write "</script>"
	end if

ToTcashsum = 0

'/오프라인 매장직원 한번만 쿼리해서 통채로 내역 다 가져옴
dim arrListpartpeopleshop
	arrListpartpeopleshop = getpartpeopleshoparray("Y")

%>

<script type="text/javascript">

function cash_edit(ix){
	var yyyymmdd = eval("frmarr.yyyymmdd_"+ix);
	var posid = eval("frmarr.posid_"+ix);
	var cnt100000won = eval("frmarr.cnt100000won_"+ix);
	var cnt50000won = eval("frmarr.cnt50000won_"+ix);
	var cnt10000won = eval("frmarr.cnt10000won_"+ix);
	var cnt5000won = eval("frmarr.cnt5000won_"+ix);
	var cnt1000won = eval("frmarr.cnt1000won_"+ix);
	var cnt500won = eval("frmarr.cnt500won_"+ix);
	var cnt100won = eval("frmarr.cnt100won_"+ix);
	var cnt50won = eval("frmarr.cnt50won_"+ix);
	var cnt10won = eval("frmarr.cnt10won_"+ix);
	var vaultcash = eval("frmarr.vaultcash_"+ix);
	var jungsanadminid = eval("frmarr.jungsanadminid_"+ix);
	var depositadminid = eval("frmarr.depositadminid_"+ix);
	var etctype = eval("frmarr.etctype_"+ix);
	var etcwon = eval("frmarr.etcwon_"+ix);
	var bigo = eval("frmarr.bigo_"+ix);

	if (!IsDouble(cnt100000won.value)){
		alert('숫자만 입력 가능합니다.'); cnt100000won.focus(); return;
	}
	if (!IsDouble(cnt50000won.value)){
		alert('숫자만 입력 가능합니다.'); cnt50000won.focus(); return;
	}
	if (!IsDouble(cnt10000won.value)){
		alert('숫자만 입력 가능합니다.'); cnt10000won.focus(); return;
	}
	if (!IsDouble(cnt5000won.value)){
		alert('숫자만 입력 가능합니다.'); cnt5000won.focus(); return;
	}
	if (!IsDouble(cnt1000won.value)){
		alert('숫자만 입력 가능합니다.'); cnt1000won.focus(); return;
	}
	if (!IsDouble(cnt500won.value)){
		alert('숫자만 입력 가능합니다.'); cnt500won.focus(); return;
	}
	if (!IsDouble(cnt100won.value)){
		alert('숫자만 입력 가능합니다.'); cnt100won.focus(); return;
	}
	if (!IsDouble(cnt50won.value)){
		alert('숫자만 입력 가능합니다.'); cnt50won.focus(); return;
	}
	if (!IsDouble(cnt10won.value)){
		alert('숫자만 입력 가능합니다.'); cnt10won.focus(); return;
	}
	if (!IsDouble(vaultcash.value)){
		alert('매장시재금은 숫자만 입력 가능합니다.'); vaultcash.focus(); return;
	}
	if(jungsanadminid.value==''){
		alert('정산담당자를 선택 해주세요.'); jungsanadminid.focus(); return;
	}
	if(depositadminid.value==''){
		alert('입금당담자를 선택 해주세요.'); depositadminid.focus(); return;
	}

	for (var i=0; i<etctype.length; i++){
		if (etctype[i].checked){
			if (!IsDouble(etcwon[i].value)){
				alert('숫자만 입력 가능합니다.'); etcwon[i].focus(); return;
			}

			frmarr.etctypearr.value = frmarr.etctypearr.value + etctype[i].value + ","
			frmarr.etcwonarr.value = frmarr.etcwonarr.value + etcwon[i].value + ","
		}
	}

	frmarr.bigoarr.value = bigo.value;
	frmarr.yyyymmddarr.value = yyyymmdd.value;
	frmarr.posidarr.value = posid.value;
	frmarr.cnt100000wonarr.value = cnt100000won.value;
	frmarr.cnt50000wonarr.value = cnt50000won.value;
	frmarr.cnt10000wonarr.value = cnt10000won.value;
	frmarr.cnt5000wonarr.value = cnt5000won.value;
	frmarr.cnt1000wonarr.value = cnt1000won.value;
	frmarr.cnt500wonarr.value = cnt500won.value;
	frmarr.cnt100wonarr.value = cnt100won.value;
	frmarr.cnt50wonarr.value = cnt50won.value;
	frmarr.cnt10wonarr.value = cnt10won.value;
	frmarr.vaultcasharr.value = vaultcash.value;
	frmarr.jungsanadminidarr.value = jungsanadminid.value;
	frmarr.depositadminidarr.value = depositadminid.value;
	frmarr.submit();
}

//현금계산금액
function reautocashsum(ix){
	var tmpetcwon = 0;

	var cnt100000won = eval("frmarr.cnt100000won_"+ix);
	var cnt50000won = eval("frmarr.cnt50000won_"+ix);
	var cnt10000won = eval("frmarr.cnt10000won_"+ix);
	var cnt5000won = eval("frmarr.cnt5000won_"+ix);
	var cnt1000won = eval("frmarr.cnt1000won_"+ix);
	var cnt500won = eval("frmarr.cnt500won_"+ix);
	var cnt100won = eval("frmarr.cnt100won_"+ix);
	var cnt50won = eval("frmarr.cnt50won_"+ix);
	var cnt10won = eval("frmarr.cnt10won_"+ix);
	var vaultcash = eval("frmarr.vaultcash_"+ix);
	var autocashsum = eval("frmarr.autocashsum_"+ix);
	var etcwon = eval("frmarr.etcwon_"+ix);
	var etctype = eval("frmarr.etctype_"+ix);

	for (var i=0; i<etcwon.length; i++){
		<% if shopid="streetshop011" then %>
			//대학로 매장일 경우 해피머니 상품권은 안더함
			if (etctype[i].value!='GC_1'){
				tmpetcwon = tmpetcwon + parseInt(etcwon[i].value)
			}
		<% else %>
			tmpetcwon = tmpetcwon + parseInt(etcwon[i].value)
		<% end if %>
	}

	autocashsum.value = (cnt100000won.value*100000) + (cnt50000won.value*50000) + (cnt10000won.value*10000) + (cnt5000won.value*5000) + (cnt1000won.value*1000) + (cnt500won.value*500) + (cnt100won.value*100) + (cnt50won.value*50) + (cnt10won.value*10) + tmpetcwon

	//매출차액
	reautocashsumdifference(ix)

	//시재금차액
	reautovaultcash(ix)
}

//매출차액
function reautocashsumdifference(ix){
	var autocashsum = eval("frmarr.autocashsum_"+ix);
	var cashsum = eval("frmarr.cashsum_"+ix);
	var autocashsumdifference = eval("frmarr.autocashsumdifference_"+ix);

	autocashsumdifference.value = autocashsum.value - cashsum.value
}

//시재금차액
function reautovaultcash(ix){
	var autocashsum = eval("frmarr.autocashsum_"+ix);
	var cashsum = eval("frmarr.cashsum_"+ix);
	var vaultcash = eval("frmarr.vaultcash_"+ix);
	var autovaultcash = eval("frmarr.autovaultcash_"+ix);

	autovaultcash.value = (autocashsum.value - cashsum.value) - vaultcash.value
}

function divbigodisp(ix,sw){
	var divbigo = document.getElementById("divbigo_"+ix);

	if (sw=="ON"){
		divbigo.style.visibility = "visible";
	} else {
		divbigo.style.visibility = 'hidden';
	}
}

function divetcdisp(ix,sw){
	var divetc = document.getElementById("divetc_"+ix);

	if (sw=="ON"){
		divetc.style.visibility = "visible";
	} else {
		divetc.style.visibility = 'hidden';
	}
}

</script>

<!-- 표 상단바 시작-->
<table width="100%" align="center" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>" class="A">
<form name="frm" method="get" action="">
<input type="hidden" name="menupos" value="<%= menupos %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
	<td align="left">
		<table border="0" width="100%" cellpadding="3" cellspacing="0" class="a">
		<tr>
			<td>
				* 기간 :
				<% drawmaechuldatefg "datefg" ,datefg ,""%>
				<% DrawDateBox yyyy1,yyyy2,mm1,mm2,dd1,dd2 %>
				&nbsp;&nbsp;
				<%
				'직영/가맹점
				if (C_IS_SHOP) then
				%>
					<% if not C_IS_OWN_SHOP and shopid <> "" then %>
						* 매장 : <%=shopid%><input type="hidden" name="shopid" value="<%= shopid %>">
					<% else %>
						* 매장 : <% Call NewDrawSelectBoxDesignerwithNameAndUserDIV("shopid",shopid, "21") %>
					<% end if %>
				<% else %>
					* 매장 : <% Call NewDrawSelectBoxDesignerwithNameAndUserDIV("shopid",shopid, "21") %>
				<% end if %>
	            &nbsp;&nbsp;
	            <b>* 매출처구분</b>
	            <% Call draw3plMeachulComboBox("inc3pl",inc3pl) %>
			</td>
		</tr>
		</table>
    </td>
		<td  width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="검색" onClick="frm.submit();">
	</td>
</tr>
</form>
</table>
<!-- 표 상단바 끝-->

<Br>

<!-- 표 중간바 시작-->
<table width="100%" align="center" cellpadding="1" cellspacing="1" class="a">
<tr valign="bottom">
    <td align="left">
    	<font color="red">※ 매장시재금은 한번입력후, 수정하실수 없습니다.</font>(오프라인관리자만가능)
    </td>
    <td align="right">
    </td>
</tr>
</table>
<!-- 표 중간바 끝-->

<form name="frmarr" method="post" action="/admin/offshop/payment/cash_management_process.asp" style="margin:0px;">
<input type="hidden" name="mode" value="cash_edit">
<input type="hidden" name="menupos" value="<%=menupos%>">
<input type="hidden" name="shopid" value="<%=shopid%>">
<input type="hidden" name="bigoarr">
<input type="hidden" name="posidarr">
<input type="hidden" name="yyyymmddarr">
<input type="hidden" name="cnt100000wonarr">
<input type="hidden" name="cnt50000wonarr">
<input type="hidden" name="cnt10000wonarr">
<input type="hidden" name="cnt5000wonarr">
<input type="hidden" name="cnt1000wonarr">
<input type="hidden" name="cnt500wonarr">
<input type="hidden" name="cnt100wonarr">
<input type="hidden" name="cnt50wonarr">
<input type="hidden" name="cnt10wonarr">
<input type="hidden" name="vaultcasharr">
<input type="hidden" name="jungsanadminidarr">
<input type="hidden" name="depositadminidarr">
<input type="hidden" name="etctypearr">
<input type="hidden" name="etcwonarr">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="25">
		검색결과 : <b><%= opayment.FResultCount %></b> ※ 최대 500건 까지 검색 됩니다.
	</td>
</tr>
<tr bgcolor="<%= adminColor("tabletop") %>" align="center">
	<td rowspan=2>기간</td>
	<td rowspan=2>요일</td>
	<td rowspan=2>포스<br>ID</td>
	<td colspan=9>권종</td>
	<td rowspan=2>기타<Br>상품권</td>
	<td rowspan=2>현금<br>계산금액<br>J</td>
	<td rowspan=2>현금<br>매출액<br>K</td>
	<td rowspan=2>매출차액<br>J-K</td>
	<td rowspan=2>매장시재금<br>L</td>
	<td rowspan=2>시재금<br>차액<br>(J-K)-L</td>
	<td rowspan=2>코맨트</td>
	<td rowspan=2>정산<br>담당</td>
	<td rowspan=2>입금<br>담당</td>
	<td rowspan=2>비고</td>
</tr>
<tr bgcolor="<%= adminColor("tabletop") %>" align="center">
	<td>10만원</td>
	<td>5만원</td>
	<td>1만원</td>
	<td>5천원</td>
	<td>천원</td>
	<td>5백원</td>
	<td>백원</td>
	<td>50원</td>
	<td>10원</td>
</tr>
<%
dim tmpyyyymmdd, tmpposid, tmpcnt100000won, tmpcnt50000won, tmpcnt10000won, tmpcnt5000won, tmpcnt1000won
dim tmpcnt500won, tmpcnt100won, tmpcnt50won, tmpcnt10won, tmpcashsum, tmpvaultcash, tmpetcwon, tmpautocashsum
dim tmpautocashsumdifference, tmpautovaultcash
	tmpyyyymmdd=""
	tmpposid=""
	tmpcnt100000won=0
	tmpcnt50000won=0
	tmpcnt10000won=0
	tmpcnt5000won=0
	tmpcnt1000won=0
	tmpcnt500won=0
	tmpcnt100won=0
	tmpcnt50won=0
	tmpcnt10won=0
	tmpcashsum=0
	tmpvaultcash=0
	tmpetcwon=0
	tmpautocashsum=0
	tmpautocashsumdifference=0
	tmpautovaultcash=0

if opayment.FResultCount > 0 then

for i=0 to opayment.FResultCount -1

'//시재금이 따로 입력이 안되어 있는경우, 일정하게 정해진 값 입력
if opayment.FItemList(i).fvaultcash="" or opayment.FItemList(i).fvaultcash="0" then
	if opayment.FItemList(i).Fcashsum>0 then
		if opayment.FItemList(i).fshopid="streetshop012" then
			opayment.FItemList(i).fvaultcash = "200000"
		elseif opayment.FItemList(i).fshopid="streetshop014" then
			if opayment.FItemList(i).fposid="11" then
				opayment.FItemList(i).fvaultcash = "50000"
			else
				opayment.FItemList(i).fvaultcash = "100000"
			end if
		else
			opayment.FItemList(i).fvaultcash = "100000"
		end if
	end if
end if

isedityn = "N"
if isnull(opayment.FItemList(i).fidx) or opayment.FItemList(i).fidx="" or C_ADMIN_AUTH or C_OFF_AUTH then
	isedityn = "Y"
end if
%>
<% if tmpyyyymmdd<>opayment.FItemList(i).fyyyymmdd and tmpyyyymmdd<>"" then %>
	<tr bgcolor="f1f1f1" align="center">
		<td colspan=3>
			<%= tmpyyyymmdd %> 합계
		</td>
		<td><%= FormatNumber(tmpcnt100000won,0) %></td>
		<td><%= FormatNumber(tmpcnt50000won,0) %></td>
		<td><%= FormatNumber(tmpcnt10000won,0) %></td>
		<td><%= FormatNumber(tmpcnt5000won,0) %></td>
		<td><%= FormatNumber(tmpcnt1000won,0) %></td>
		<td><%= FormatNumber(tmpcnt500won,0) %></td>
		<td><%= FormatNumber(tmpcnt100won,0) %></td>
		<td><%= FormatNumber(tmpcnt50won,0) %></td>
		<td><%= FormatNumber(tmpcnt10won,0) %></td>
		<td align="right"><%= FormatNumber(tmpetcwon,0) %></td>
		<td align="right"><%= FormatNumber(tmpautocashsum,0) %></td>
		<td align="right"><%= FormatNumber(tmpcashsum,0) %></td>
		<td align="right"><%= FormatNumber(tmpautocashsumdifference,0) %></td>
		<td align="right"><%= FormatNumber(tmpvaultcash,0) %></td>
		<td align="right"><%= FormatNumber(tmpautovaultcash,0) %></td>
		<td colspan=4></td>
	</tr>
	<%
	tmpcnt100000won=0
	tmpcnt50000won=0
	tmpcnt10000won=0
	tmpcnt5000won=0
	tmpcnt1000won=0
	tmpcnt500won=0
	tmpcnt100won=0
	tmpcnt50won=0
	tmpcnt10won=0
	tmpcashsum=0
	tmpvaultcash=0
	tmpetcwon=0
	tmpautocashsum=0
	tmpautocashsumdifference=0
	tmpautovaultcash=0
	%>
<% end if %>

<tr bgcolor="#FFFFFF" onmouseover=this.style.background="#f1f1f1"; onmouseout=this.style.background='#ffffff'; align="center" height=40>
	<input type="hidden" name="yyyymmdd_<%=i%>" value="<%= opayment.FItemList(i).fyyyymmdd %>">
	<%
	'/셀병합
	if instr(opayment.fposidarr,opayment.FItemList(i).fyyyymmdd)="0" then
	%>
		<td width=70>
			<%= getweekendcolor(opayment.FItemList(i).fyyyymmdd) %>
		</td>
		<td width=30>
			<%= getweekend(opayment.FItemList(i).fyyyymmdd) %>
		</td>
	<% else %>
		<% if opayment.FItemList(i).fposidcnt="1" then %>
			<td width=70 rowspan="<%= mid(opayment.fposidarr, instr(opayment.fposidarr,opayment.FItemList(i).fyyyymmdd)+11,1) %>">
				<%= getweekendcolor(opayment.FItemList(i).fyyyymmdd) %>
			</td>
			<td width=30 rowspan="<%= mid(opayment.fposidarr, instr(opayment.fposidarr,opayment.FItemList(i).fyyyymmdd)+11,1) %>">
				<%= getweekend(opayment.FItemList(i).fyyyymmdd) %>
			</td>
		<% end if %>
	<% end if %>

	<td width=30>
		<input type="hidden" name="posid_<%=i%>" value="<%= opayment.FItemList(i).fposid %>">
		<%= opayment.FItemList(i).fposid %>
	</td>
	<td width=60>
		<input type="text" name="cnt100000won_<%=i%>" size=5 maxlength=5 value="<%= opayment.FItemList(i).fcnt100000won %>" onkeyup="reautocashsum('<%=i%>')" style="text-align:right;">
	</td>
	<td width=60>
		<input type="text" name="cnt50000won_<%=i%>" size=5 maxlength=5 value="<%= opayment.FItemList(i).fcnt50000won %>" onkeyup="reautocashsum('<%=i%>')" style="text-align:right;">
	</td>
	<td width=60>
		<input type="text" name="cnt10000won_<%=i%>" size=5 maxlength=5 value="<%= opayment.FItemList(i).fcnt10000won %>" onkeyup="reautocashsum('<%=i%>')" style="text-align:right;">
	</td>
	<td width=60>
		<input type="text" name="cnt5000won_<%=i%>" size=5 maxlength=5 value="<%= opayment.FItemList(i).fcnt5000won %>" onkeyup="reautocashsum('<%=i%>')" style="text-align:right;">
	</td>
	<td width=60>
		<input type="text" name="cnt1000won_<%=i%>" size=5 maxlength=5 value="<%= opayment.FItemList(i).fcnt1000won %>" onkeyup="reautocashsum('<%=i%>')" style="text-align:right;">
	</td>
	<td width=60>
		<input type="text" name="cnt500won_<%=i%>" size=5 maxlength=5 value="<%= opayment.FItemList(i).fcnt500won %>" onkeyup="reautocashsum('<%=i%>')" style="text-align:right;">
	</td>
	<td width=60>
		<input type="text" name="cnt100won_<%=i%>" size=5 maxlength=5 value="<%= opayment.FItemList(i).fcnt100won %>" onkeyup="reautocashsum('<%=i%>')" style="text-align:right;">
	</td>
	<td width=60>
		<input type="text" name="cnt50won_<%=i%>" size=5 maxlength=5 value="<%= opayment.FItemList(i).fcnt50won %>" onkeyup="reautocashsum('<%=i%>')" style="text-align:right;">
	</td>
	<td width=60>
		<input type="text" name="cnt10won_<%=i%>" size=5 maxlength=5 value="<%= opayment.FItemList(i).fcnt10won %>" onkeyup="reautocashsum('<%=i%>')" style="text-align:right;">
	</td>
	<td width=40>
		<%
		set opaymentetc = new Cpayment
			opaymentetc.frectmasteridx = opayment.FItemList(i).fidx
			opaymentetc.FPageSize = 50
			opaymentetc.FCurrPage = 1
			opaymentetc.Getcash_management_etc
		%>
		<div id="divetc_<%=i%>" name="divetc_<%=i%>" style='position:absolute; width:250px; margin-top:15px; margin-left:0px;visibility:hidden; background-color:white; border-width:1px; border-style:solid;'>
		<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
		<tr bgcolor="#FFFFFF" align="center">
			<td colspan=2 align="right">
				<input type="button" onclick="divetcdisp(<%=i%>,'OFF')" value="닫기" class="button"> ※ 리스트의 저장버튼을 누르세요.
			</td>
		</tr>
		<%
		If opaymentetc.FResultCount > 0 THEN

		For intLoop = 0 To opaymentetc.FResultCount -1

		if shopid="streetshop011" then
			'//대학로 매장일 경우 해피머니 상품권은 안더함
			if opaymentetc.FItemList(intLoop).fcodeid<>"GC_1" then
				tmpetcwon = tmpetcwon + opaymentetc.FItemList(intLoop).fetcwon
			end if
		else
			tmpetcwon = tmpetcwon + opaymentetc.FItemList(intLoop).fetcwon
		end if
		%>
		<tr bgcolor="#FFFFFF" align="center">
			<td align="left">
				<input type="checkbox" name="etctype_<%=i%>" value="<%= opaymentetc.FItemList(intLoop).fcodeid %>" <% if opaymentetc.FItemList(intLoop).fdetailidx<>"" then response.write " checked" %>><%= opaymentetc.FItemList(intLoop).fcodename %>
			</td>
			<td align="right">
				<input type="text" name="etcwon_<%=i%>" size=8 maxlength=8 value="<%= opaymentetc.FItemList(intLoop).fetcwon %>" onkeyup="reautocashsum('<%=i%>')" style="text-align:right;">
			</td>
		</tr>
		<% Next %>
		<% End IF %>
		</table>
		</div>
		<a href="javascript:divetcdisp(<%=i%>,'ON')" onfocus="this.blur()"><img src="/images/icon_search.jpg" border=0></a>
		<%
		set opaymentetc = nothing
		%>
	</td>
	<td width=90>
		<input type="text" name="autocashsum_<%=i%>" size=10 maxlength=10 readonly class="text_ro" style="text-align:right;">
	</td>
    <td align="right" width=80 bgcolor="#E6B9B8">
    	<input type="hidden" name="cashsum_<%=i%>" value="<%= opayment.FItemList(i).Fcashsum %>">
    	<%= FormatNumber(opayment.FItemList(i).Fcashsum,0) %>
    </td>
	<td width=90>
		<input type="text" name="autocashsumdifference_<%=i%>" size=10 maxlength=10 readonly class="text_ro" style="text-align:right;">
	</td>
	<td width=90>
		<input type="text" name="vaultcash_<%=i%>" size=10 maxlength=10 value="<%= opayment.FItemList(i).fvaultcash %>" <% if isedityn="N" then response.write " readonly" %> onkeyup="reautovaultcash('<%=i%>')" style="text-align:right;">
	</td>
	<td width=90>
		<input type="text" name="autovaultcash_<%=i%>" size=10 maxlength=10 readonly class="text_ro" style="text-align:right;">
	</td>
	<td width=40>
		<div id="divbigo_<%=i%>" name="divbigo_<%=i%>" style='position:absolute; width:250px; margin-top:15px; margin-left:0px;visibility:hidden; background-color:white; border-width:1px; border-style:solid;'>
		<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
		<tr bgcolor="#FFFFFF" align="center">
			<td align="right">
				<input type="button" onclick="divbigodisp(<%=i%>,'OFF')" value="닫기" class="button"> ※ 리스트의 저장버튼을 누르세요.
			</td>
		</tr>
		<tr bgcolor="#FFFFFF" align="center">
			<td>
				<textarea name="bigo_<%=i%>" cols=30 rows=15><%= opayment.FItemList(i).fbigo %></textarea>
			</td>
		</tr>
		</table>
		</div>
		<a href="javascript:divbigodisp(<%=i%>,'ON')" onfocus="this.blur()"><img src="/images/icon_search.jpg" border=0></a>
	</td>
	<td width=80>
		<select name="jungsanadminid_<%=i%>">
			<option value="" <% if opayment.FItemList(i).fjungsanadminid = "" then response.write " selected" %>>선택</option>
			<%
			If isArray(arrListpartpeopleshop) THEN

			For intLoop = 0 To UBound(arrListpartpeopleshop,2)
			%>
			<option value="<%=arrListpartpeopleshop(0,intLoop)%>" <%if arrListpartpeopleshop(0,intLoop) = opayment.FItemList(i).fjungsanadminid then %> selected<%end if%>><%=arrListpartpeopleshop(1,intLoop)%></option>
			<%
			Next

			End IF
			%>
		</select>
	</td>
	<td width=80>
		<select name="depositadminid_<%=i%>">
			<option value="" <% if opayment.FItemList(i).fdepositadminid = "" then response.write " selected" %>>선택</option>
			<%
			If isArray(arrListpartpeopleshop) THEN

			For intLoop = 0 To UBound(arrListpartpeopleshop,2)
			%>
			<option value="<%=arrListpartpeopleshop(0,intLoop)%>" <%if arrListpartpeopleshop(0,intLoop) = opayment.FItemList(i).fdepositadminid then %> selected<%end if%>><%=arrListpartpeopleshop(1,intLoop)%></option>
			<%
			Next

			End IF
			%>
		</select>
	</td>
	<td>
		<% if shopid<>"" then %>
			<input type="button" onclick="cash_edit('<%=i%>')" value="저장" class="button">
		<% end if %>
	</td>
</tr>
<%
ToTcashsum = ToTcashsum + opayment.FItemList(i).Fcashsum
tmpyyyymmdd=opayment.FItemList(i).fyyyymmdd
tmpposid=opayment.FItemList(i).fposid
tmpcnt100000won=tmpcnt100000won + opayment.FItemList(i).fcnt100000won
tmpcnt50000won=tmpcnt50000won + opayment.FItemList(i).fcnt50000won
tmpcnt10000won=tmpcnt10000won + opayment.FItemList(i).fcnt10000won
tmpcnt5000won=tmpcnt5000won + opayment.FItemList(i).fcnt5000won
tmpcnt1000won=tmpcnt1000won + opayment.FItemList(i).fcnt1000won
tmpcnt500won=tmpcnt500won + opayment.FItemList(i).fcnt500won
tmpcnt100won=tmpcnt100won + opayment.FItemList(i).fcnt100won
tmpcnt50won=tmpcnt50won + opayment.FItemList(i).fcnt50won
tmpcnt10won=tmpcnt10won + opayment.FItemList(i).fcnt10won
tmpcashsum=tmpcashsum + opayment.FItemList(i).fcashsum
tmpvaultcash=tmpvaultcash + opayment.FItemList(i).fvaultcash
tmpautocashsum=(tmpcnt100000won*100000)+(tmpcnt50000won*50000)+(tmpcnt10000won*10000)+(tmpcnt5000won*5000)+(tmpcnt1000won*1000)+(tmpcnt500won*500)+(tmpcnt100won*100)+(tmpcnt50won*50)+(tmpcnt10won*10)+tmpetcwon
tmpautocashsumdifference=(tmpautocashsum-tmpcashsum)
tmpautovaultcash=(tmpautocashsum-tmpcashsum)-tmpvaultcash
%>
<script type="text/javascript">
	reautocashsum('<%=i%>')
</script>

<% next %>

<tr bgcolor="f1f1f1" align="center">
	<td colspan=3>
		<%= tmpyyyymmdd %> 합계
	</td>
	<td><%= FormatNumber(tmpcnt100000won,0) %></td>
	<td><%= FormatNumber(tmpcnt50000won,0) %></td>
	<td><%= FormatNumber(tmpcnt10000won,0) %></td>
	<td><%= FormatNumber(tmpcnt5000won,0) %></td>
	<td><%= FormatNumber(tmpcnt1000won,0) %></td>
	<td><%= FormatNumber(tmpcnt500won,0) %></td>
	<td><%= FormatNumber(tmpcnt100won,0) %></td>
	<td><%= FormatNumber(tmpcnt50won,0) %></td>
	<td><%= FormatNumber(tmpcnt10won,0) %></td>
	<td align="right"><%= FormatNumber(tmpetcwon,0) %></td>
	<td align="right"><%= FormatNumber(tmpautocashsum,0) %></td>
	<td align="right"><%= FormatNumber(tmpcashsum,0) %></td>
	<td align="right"><%= FormatNumber(tmpautocashsumdifference,0) %></td>
	<td align="right"><%= FormatNumber(tmpvaultcash,0) %></td>
	<td align="right"><%= FormatNumber(tmpautovaultcash,0) %></td>
	<td colspan=4></td>
</tr>

<tr bgcolor="<%= adminColor("tabletop") %>" align="center">
    <td colspan=14>총합계</td>
    <td align="right"><%= FormatNumber(ToTcashsum,0) %></td>
    <td colspan=7></td>
</tr>

<% else %>
<tr align="center" bgcolor="#FFFFFF">
	<td colspan="25">등록된 내용이 없습니다.</td>
</tr>
<% end if %>

</table>

</form>

<%
set opayment= Nothing
%>
<!-- #include virtual="/common/lib/commonbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
