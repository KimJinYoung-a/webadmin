<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  핑거스
' History : 2010.05.12 한용민 수정
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/academy/lib/academy_function.asp"-->
<!-- #include virtual="/academy/lib/classes/fingers_lecturecls.asp"-->

<%
dim lec_idx , i , lecOption
lec_idx = RequestCheckvar(request("lec_idx"),10)
lecOption = RequestCheckvar(request("lecOption"),4)

dim olecture
set olecture = new CLecture
	olecture.FRectIdx = lec_idx
	'olecture.FRectLecOpt = lecOption

	if lec_idx<>"" then
		olecture.GetOneLecture
	end if
%>
<script language='javascript'>

	// 숫자 자리수 표시
	function plusComma(num){
		if (num < 0) { num *= -1; var minus = true}
		else var minus = false

		var dotPos = (num+"").split(".")
		var dotU = dotPos[0]
		var dotD = dotPos[1]
		var commaFlag = dotU.length%3

		if(commaFlag) {
			var out = dotU.substring(0, commaFlag)
			if (dotU.length > 3) out += ","
		}
		else var out = ""

		for (var i=commaFlag; i < dotU.length; i+=3) {
			out += dotU.substring(i, i+3)
			if( i < dotU.length-3) out += ","
		}

		if(minus) out = "-" + out
		if(dotD) return out + "." + dotD
		else return out
	}
    
    function chgEntryDetail(comp){
        var cnt = comp.value;
		var ttlsumid = document.all["htmlttlsum"];
		
		var matinclude_yn = eval("frmlec.matinclude_yn");
		var mat_cost = eval("frmlec.mat_cost");
		var mat_buying_cost = eval("frmlec.mat_buying_cost");

		var lec_cost = eval("frmlec.lec_cost");
		var buying_cost = eval("frmlec.buying_cost");

		var itemno = eval("frmlec.itemea").value;

		var sellprice = eval("frmlec.sellprice").value;
		var ttlsumvalue = itemno*1*sellprice;
		var soldoutflagform = eval("frmlec.soldoutflag");
		var itemsubtotalsumfrm = eval("frmlec.itemsubtotalsum");

		itemsubtotalsumfrm.value = ttlsumvalue;
        
        if (matinclude_yn.value == "C") {
			ttlsumid.innerHTML = "<b>" + plusComma(ttlsumvalue) + "</b><br>선납";
		} else {
			ttlsumid.innerHTML = "<b>" + plusComma(ttlsumvalue) + "</b><br>현장";
		}
    }
    
	// 수강생 6숫자에 따른 테이블 표시
	function ShowEntryDetail(comp){

		var cnt = comp.value;
		var ttlsumid = document.all["htmlttlsum"];

		var matinclude_yn = eval("frmlec.matinclude_yn");
		var mat_cost = eval("frmlec.mat_cost");
		var mat_buying_cost = eval("frmlec.mat_buying_cost");

		var lec_cost = eval("frmlec.lec_cost");
		var buying_cost = eval("frmlec.buying_cost");

		var itemno = eval("frmlec.itemea").value;

		var sellprice = eval("frmlec.sellprice").value;
		var ttlsumvalue = itemno*1*sellprice;
		var soldoutflagform = eval("frmlec.soldoutflag");
		var itemsubtotalsumfrm = eval("frmlec.itemsubtotalsum");

		itemsubtotalsumfrm.value = ttlsumvalue;

		for (i=0;i<cnt;i++){
			document.all["entry"+(i)].style.display="";
		}

		for (i=3;i>=cnt;i--){
			document.all["entry"+(i)].style.display="none";
		}

		if (matinclude_yn.value == "C") {
			ttlsumid.innerHTML = "<b>" + plusComma(ttlsumvalue) + "</b><br>선납";
		} else {
			ttlsumid.innerHTML = "<b>" + plusComma(ttlsumvalue) + "</b><br>현장";
		}

		//RecalcuSubTotal();

		//포커싱이동
		//eval("baguniFrm.buy_name").focus();
	}


	// 폼 검사 및 전송
	function SaveItem()
	{
		var frm = document.frmlec;

		if(!frm.lecOption.value)
		{
			alert("강좌시간을 선택해주십시오.");
			frm.lecOption.focus();
			return;
		} else if(frm.lecOption.options[frm.lecOption.selectedIndex].id=="S") {
			alert("마감된 강좌는 수강신청을 할 수 없습니다.");
			frm.lecOption.focus();
			return;
		}

		if(!frm.buy_name.value)
		{
			alert("주문자의 이름을 입력해주십시오.");
			frm.buy_name.focus();
			return;
		}

		if(!(frm.buy_phone1.value&&frm.buy_phone2.value&&frm.buy_phone3.value))
		{
			alert("주문자의 전화번호를 입력해주십시오.");
			frm.buy_phone1.focus();
			return;
		}

		if(!(frm.buy_hp1.value&&frm.buy_hp2.value&&frm.buy_hp3.value))
		{
			alert("주문자의 휴대폰번호를 입력해주십시오.");
			frm.buy_hp1.focus();
			return;
		}

		if(!frm.buy_email.value)
		{
			alert("주문자의 이메일을 입력해주십시오.");
			frm.buy_email.focus();
			return;
		}
        <% If Not olecture.FOneItem.isWeClass Then %>
		for(i=1;i<frm.itemea.value;i++)
		{
			if(!frm['entryname' + i].value)
			{
				alert("수강생#" + (i+1) + "의 이름을 입력해주십시오.");
				frm['entryname' + i].focus();
				return;
			}

			if(!(frm['entry' + i + '_hp1'].value&&frm['entry' + i + '_hp2'].value&&frm['entry' + i + '_hp3'].value))
			{
				alert("수강생#" + (i+1) + "의 연락처를 입력해주십시오.");
				frm['entry' + i + '_hp1'].focus();
				return;
			}
		}
        <% end if %>
        
        if ((frm.paymethod.value=="7")&&(frm.lecOption.value!="0000")){
            alert('단체강좌 주문 접수로 등록시 강좌시간=상시(기본값)으로 선택하세요.');
            return;
        }
        
        <% If olecture.FOneItem.isWeClass Then %>
        if (frm.wantstudyName.value.length<1){
    		alert('주문자 업체(동호회)명을 입력하시기 바랍니다.');
    		frm.wantstudyName.focus();
    		return;
    	}
	
    	if (frm.wantstudyPlace.value.length<1){
    		alert('강의장소 입력하시기 바랍니다.');
    		frm.wantstudyPlace.focus();
    		return;
    	}
    	if (!(frm.wantstudyWho[0].checked) && !(frm.wantstudyWho[1].checked) && !(frm.wantstudyWho[2].checked) && !(frm.wantstudyWho[3].checked)){
    		alert('강의대상을 선택하시기 바랍니다.');
    		return;
    	}
    	<% End If %>
	
		// 전송
		if (confirm('강좌 수기 등록 하시겠습니까?')){
		    frm.submit();
		}
	}
	

	// 아이디 찾기
	function popSrcId()
	{
		window.open("popSearchId.asp", "popId", "width=418,height=300,scrollbars=yes")
	}

</script>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frmlec" method="POST" action="DoPopLecReg.asp">
<input type="hidden" name="lec_idx" value="<%=lec_idx%>">
<input type="hidden" name="lec_title" value="<%=olecture.FOneItem.Flec_title%>">

<input type="hidden" name="lec_cost" value="<%=olecture.FOneItem.Flec_cost%>">
<input type="hidden" name="buying_cost" value="<%=olecture.FOneItem.Fbuying_cost%>">
<input type="hidden" name="matinclude_yn" value="<%=olecture.FOneItem.Fmatinclude_yn%>">
<input type="hidden" name="mat_cost" value="<%=olecture.FOneItem.Fmat_cost%>">
<input type="hidden" name="mat_buying_cost" value="<%=olecture.FOneItem.Fmat_buying_cost%>">

<% if (olecture.FOneItem.Fmatinclude_yn = "C") then %>
	<input type="hidden" name="sellprice" value="<%= (olecture.FOneItem.Flec_cost + olecture.FOneItem.Fmat_cost) %>">
	<input type="hidden" name="buycash" value="<%= (olecture.FOneItem.Fbuying_cost + olecture.FOneItem.Fmat_buying_cost) %>">
	<input type="hidden" name="itemsubtotalsum" value="<%= (olecture.FOneItem.Flec_cost + olecture.FOneItem.Fmat_cost) %>">
<% else %>
	<input type="hidden" name="sellprice" value="<%= (olecture.FOneItem.Flec_cost) %>">
	<input type="hidden" name="buycash" value="<%=olecture.FOneItem.Fbuying_cost%>">
	<input type="hidden" name="itemsubtotalsum" value="<%=olecture.FOneItem.Flec_cost%>">
<% end if %>

<input type="hidden" name="mileage" value="<%=olecture.FOneItem.Fmileage%>">
<input type="hidden" name="makerId" value="<%=olecture.FOneItem.Flecturer_id%>">
<input type="hidden" name="sitename" value="academy">
<input type="hidden" name="buy_level" value="0">
<input type="hidden" name="weclassyn" value="<%= CHKIIF(olecture.FOneItem.isWeClass,"Y","N") %>">
<tr bgcolor="ffffff">
	<td valign="top" colspan=5>강좌 : <b><%= lec_idx & " / " & olecture.FOneItem.Flec_title%></b>
	<% if (olecture.FOneItem.isWeClass) then %>
	<b><font color=red>[단체강좌]</font></b>
	<% end if %>
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td colspan="2">
		<table width="100%" border="0" cellpadding="2" cellspacing="1" class="a">
		<tr align="center" bgcolor="#F7F7F7">
			<td>수강료<br>재료비</td>
			<td>마일리지</td>
			<td>신청인원</td>
			<td>총금액<br>재료비</td>
		</tr>
		<tr align="center" bgcolor="#FFFFFF">
			<td><%= FormatNumber(olecture.FOneItem.Flec_cost,0) %><br><%= FormatNumber(olecture.FOneItem.Fmat_cost,0) %></td>
			<td><%= olecture.FOneItem.Fmileage %> (point)</td>
			<td>
			    <% IF (olecture.FOneItem.isWeClass) THEN %>
			    <input type="text" name="itemea" value="1" size=3 maxlength=3 onChange="chgEntryDetail(this)">
			    <% ELSE %>
				<select name="itemea" onChange="ShowEntryDetail(this)">
					<option value="1">1 명</option>
					<option value="2">2 명</option>
					<option value="3">3 명</option>
					<option value="4">4 명</option>
				</select>
				<% end if %>
			</td>
			<td id="htmlttlsum">
				<% if (olecture.FOneItem.Fmatinclude_yn = "C") then %>
				<b><%= FormatNumber((olecture.FOneItem.Flec_cost + olecture.FOneItem.Fmat_cost),0) %></b><br>
				선납
				<% else %>
				<b><%= FormatNumber(olecture.FOneItem.Flec_cost,0) %></b><br>
				현장
				<% end if %>
			</td>
		</tr>
		</table>
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td width=120 bgcolor="#DDDDFF" align="center">강좌시간 선택</td>
	<td><%= getLecOptionBoxHTML(lec_idx,"lecOption","") %></td>
</tr>
<%	For i=0 to 3 %>
<tr id="entry<%=i%>" <% if ((i>0) ) then Response.Write "style='display:none;'"%> bgcolor="#FFFFFF">
	<td width=120 bgcolor="#DDDDFF" align="center">
		수강생#<%=i+1%>
		<% if i=0 then Response.Write "<br>(주문자)" %>
	</td>
	<td>
		<% if i=0 then %>
		<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="a">
		<tr>
			<td width="60" bgcolor="#F8F8F8">아이디</td>
			<td>
				<input type="text" name="buy_userid" value="" size="12" maxlength="32" class="input" readonly>
				<img src="/images/icon_search.gif" onClick="popSrcId()" style="cursor:pointer" align="absmiddle">
			</td>
		</tr>
		<tr>
			<td width="60" bgcolor="#F8F8F8"><font color="orange">* </font>성 명</td>
			<td><input type="text" name="buy_name" value="" size="10" maxlength="16" class="input"></td>
		</tr>
		<tr>
			<td bgcolor="#F8F8F8"><font color="orange">* </font>전화번호</td>
			<td>
				<input name="buy_phone1" type="text" size="4" maxlength="4" value="" maxlength="4" class="input">
				-
				<input name="buy_phone2" type="text" size="4" maxlength="4" value="" maxlength="4" class="input">
				-
				<input name="buy_phone3" type="text" size="4" maxlength="4" value="" maxlength="4" class="input">
			</td>
		</tr>
		<tr>
			<td bgcolor="#F8F8F8"><font color="orange">* </font>휴대폰</td>
			<td>
				<input name="buy_hp1" type="text" size="4" maxlength="4" value="" maxlength="4" class="input">
				-
				<input name="buy_hp2" type="text" size="4" maxlength="4" value="" maxlength="4" class="input">
				-
				<input name="buy_hp3" type="text" size="4" maxlength="4" value="" maxlength="4" class="input">
			</td>
		</tr>
		<tr>
			<td bgcolor="#F8F8F8"><font color="orange">* </font>이메일</td>
			<td><input name="buy_email" type="text" size="26" value="" maxlength="90" class="input"></td>
		</tr>
		<tr>
			<td bgcolor="#F8F8F8"><font color="orange">* </font>마일리지 적립여부</td>
			<td><input type="radio" name="mileagegubun" value="ON" checked>적립 <input type="radio" name="mileagegubun" value="OFF">적립안함
			</td>
		</tr>
		</table>
		<% else %>
		<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="a">
		<tr>
			<td width="60" bgcolor="#F8F8F8"><font color="orange">* </font>성 명</td>
			<td><input type="text" name="entryname<%=i%>" value="" size="8" maxlength="16" class="input"></td>
		</tr>
		<tr>
			<td bgcolor="#F8F8F8"><font color="orange">* </font>연락처</td>
			<td>
				<input name="entry<%=i%>_hp1" type="text" size="4" maxlength="4" value="" maxlength="4" class="input">
				-
				<input name="entry<%=i%>_hp2" type="text" size="4" maxlength="4" value="" maxlength="4" class="input">
				-
				<input name="entry<%=i%>_hp3" type="text" size="4" maxlength="4" value="" maxlength="4" class="input">
			</td>
		</tr>
		</table>
		<% end if %>
	</td>
</tr>
<%	next %>
<% if olecture.FOneItem.isWeClass THEN %>
<tr  bgcolor="#FFFFFF">
    <td bgcolor="#DDDDFF" align="center">단체정보</td>
    <td >
        <table>
        <tbody>
		<tr>
			<th><span>업체(동호회)명</span></th>
			<td>
				<span><input type="text" name="wantstudyName" class="txtBasic tblInput" style="width:200px;" maxlength="100" value="" /></span>
				<span class="lPad0">(업체, 동호회 혹은 대표자 명을 입력해주세요.)</span>
			</td>
		</tr>
		<tr>
			<th><span>강의 희망일</span></th>
			<td>
				<span>
					<select name="wantstudyYear" class="select tblInput" >
						<option value="2012">2012</option>
						<option value="2013">2013</option>
						<option value="2014">2014</option>
						<option value="2015">2015</option>
						<option value="2016">2016</option>
						<option value="2017">2017</option>
						<option value="2018">2018</option>
						<option value="2019">2019</option>
						<option value="2020">2020</option>
					</select> 년
					<select name="wantstudyMonth" class="select tblInput" >
						<% For i=1 To 12 %>
						<option value="<%=i%>"><%=i%></option>
						<% Next %>
					</select> 월
					<select name="wantstudyDay" class="select tblInput" >
						<% For i=1 To 31 %>
						<option value="<%=i%>"><%=i%></option>
						<% Next %>
					</select> 일
				</span>
				<span class="lPad0">
					<select name="wantstudyAmPm" class="select tblInput" >
						<option value="오전">오전</option>
						<option value="오후">오후</option>
					</select>
					<select name="wantstudyHour" class="select tblInput" >
						<% For i=1 To 12 %>
						<option value="<%=i%>"><%=i%></option>
						<% Next %>
					</select> 시
					<select name="wantstudyMin" class="select tblInput" >
						<% For i=0 To 50 step 10 %>
						<option value="<%=i%>"><%=i%></option>
						<% Next %>
					</select> 분
				</span>
			</td>
		</tr>
		<tr>
			<th><span>강의장소</span></th>
			<td><span><input type="text" name="wantstudyPlace" class="txtBasic tblInput" style="width:500px;" maxlength="100" value="" /></span></td>
		</tr>
		<tr>
			<th><span>강의대상</span></th>
			<td>
				<span><input name="wantstudyWho" type="radio" class="radio" value="1" /> 기업</span>
				<span><input name="wantstudyWho" type="radio" class="radio" value="2" /> 동호회</span>
				<span><input name="wantstudyWho" type="radio" class="radio" value="3" /> 학생</span>
				<span><input name="wantstudyWho" type="radio" class="radio" value="0" /> 기타</span>
			</td>
		</tr>
		</tbody>
        </table>
    </td>
</tr>
<% end if %>
<tr bgcolor="#FFFFFF">
	<td bgcolor="#DDDDFF" align="center">결제방법</td>
	<td>
	    <% if (olecture.FOneItem.isWeClass) then %>
	    <select name="paymethod">
	    <option value="7">주문접수
	    <option value="900">수기입력(결제완료)
	    </select>
	    <% else %>
		수기입력<br>
		<font color="gray">※ 참고 : 입력 완료시 결제 상태는 [결제완료] 입니다.</font>
		<input type="hidden" name="paymethod" value="900">
		<% end if %>
	</td>
</tr>

<tr bgcolor="ffffff">
	<td valign="top" colspan=5 align="center">
		<img src="/images/icon_save.gif" onClick="SaveItem()" style="cursor:pointer" align="absbottom"> &nbsp;
		<img src="/images/icon_cancel.gif" onClick="self.close()" style="cursor:pointer" align="absbottom">
	</td>
</tr>
</form>
</table>

<%
	set olecture = Nothing
%>

<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->