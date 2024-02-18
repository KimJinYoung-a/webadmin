<%@ language=vbscript %>
<% option explicit %>
<%
'###############################################
' PageName : main_manager.asp
' Discription : 사이트 메인 관리
' History : 2008.04.11 허진원 : 실서버에서 이전
'			2009.04.19 한용민 2009에 맞게 수정
'           2009.12.21 허진원 : 일자별 플래시 예약 기능 추가
'			2012.02.08 허진원 : 미니달력 교체
'           2013.09.28 허진원 : 2013리뉴얼 - 추가선택 필드 추가
'           2015.04.07 원승현 : 2015리뉴얼 - 추가선택 필드 추가
'           2018-01-15 이종화 : 구분 PC배너 관리 추가
'###############################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/classes/sitemasterclass/main_enjoyContentsManageCls.asp" -->
<!-- #include virtual="/admin/eventmanage/common/event_function.asp"-->
<%
Dim isusing, fixtype, validdate, prevDate
Dim idx, poscode, reload, gubun, edid
Dim culturecode
	idx = request("idx")
	poscode = request("poscode")
	reload = request("reload")
	gubun = request("gubun")

	isusing = request("isusing")
	fixtype = request("fixtype")
	validdate= request("validdate")
	prevDate = request("prevDate")

	culturecode = request("eC")

	if idx="" then idx=0

	Response.write culturecode

	if reload="on" then
	    response.write "<script>opener.location.reload(); window.close();</script>"
	    dbget.close()	:	response.End
	end if

	dim oMainContents
		set oMainContents = new CMainEnjoyContents
		oMainContents.FRectIdx = idx
		oMainContents.GetOneGatherEventMainContents

	If gubun = "" Then
		gubun = "index"
	End If

%>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script language='javascript' src="/js/jsCal/js/jscal2.js"></script>
<script language='javascript' src="/js/jsCal/js/lang/ko.js"></script>
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/jscal2.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/border-radius.css" />
<script language='javascript'>

	function SaveMainContents(frm){
	    if (frm.MainCopy1.value==""){
	        alert('메인카피를 입력 하세요.');
	        frm.MainCopy1.focus();
	        return;
	    }

		if(!maxLengthCheck("MainCopy1","메인카피",60))
		{
			frm.MainCopy1.focus();
			return;
		}

		if(!maxLengthCheck("MainCopy2","메인카피",60))
		{
			frm.MainCopy2.focus();
			return;
		}

	    if (frm.Evt_Code1.value==""){
	        alert('이벤트 번호를 입력 하세요.');
	        frm.Evt_Code1.focus();
	        return;
	    }

		if (frm.Evt_Title1.value==""){
	        alert('이벤트 메인카피를 입력 하세요.');
	        frm.Evt_Title1.focus();
	        return;
	    }

		if(!maxLengthCheck("Evt_Title1","이벤트 메인카피",40))
		{
			frm.Evt_Title1.focus();
			return;
		}

		if (frm.Evt_Subcopy1.value==""){
	        alert('이벤트 서브카피를 입력 하세요.');
	        frm.Evt_Subcopy1.focus();
	        return;
	    }

		if(!maxLengthCheck("Evt_Subcopy1","이벤트 서브카피",56))
		{
			frm.Evt_Subcopy1.focus();
			return;
		}

		if(!maxLengthCheck("Evt_Title2","이벤트 메인카피",40))
		{
			frm.Evt_Title2.focus();
			return;
		}
		if(!maxLengthCheck("Evt_Subcopy2","이벤트 서브카피",56))
		{
			frm.Evt_Subcopy2.focus();
			return;
		}

		if(!maxLengthCheck("Evt_Title3","이벤트 메인카피",40))
		{
			frm.Evt_Title3.focus();
			return;
		}
		if(!maxLengthCheck("Evt_Subcopy3","이벤트 서브카피",56))
		{
			frm.Evt_Subcopy3.focus();
			return;
		}

	    if (frm.startdate.value.length!=10){
	        alert('시작일을 입력  하세요.');
	        return;
	    }

	    if (frm.enddate.value.length!=10){
	        alert('종료일을 입력  하세요.');
	        return;
	    }

	    if (frm.startdate.value>frm.enddate.value){
	        alert('종료일이 시작일보다 빠르면 안됩니다.');
	        return;
	    }

	    if (confirm('저장 하시겠습니까?')){
	        frm.submit();
	    }
	}

	function ChangeLinktype(comp){
	    if (comp.value=="M"){
	       document.all.link_M.style.display = "";
	       document.all.link_L.style.display = "none";
	    }else{
	       document.all.link_M.style.display = "none";
	       document.all.link_L.style.display = "";
	    }
	}

	//function getOnLoad(){
	//    ChangeLinktype(frmcontents.linktype.value);
	//}

	//window.onload = getOnLoad;

	function ChangeGubun(comp){
	    location.href = "?gubun=<%=gubun%>&poscode=" + comp.value;
	    // nothing;
	}


	function ChangeGroupGubun(comp){
	    location.href = "?gubun=" + comp.value;
	    // nothing;
	}

	function cultureloadpop(){
		winLast = window.open('pop_culturelist.asp','pLast','width=1200,height=600, scrollbars=yes')
		winLast.focus();
	}

	//색상코드 선택
	function selColorChip(bg,cd) {
		var i;
		document.frmcontents.BGColor.value= bg;
		for(i=1;i<=11;i++) {
			document.all("cline"+i).bgColor='#DDDDDD';
		}
		if(!cd) document.all("cline0").bgColor='#DD3300';
		else document.all("cline"+cd).bgColor='#DD3300';
	}

	//-- jsLastEvent : 지난 이벤트 불러오기 --//
	function jsLastEvent(num){
	  winLast = window.open('pop_event_lastlist.asp?num='+num,'pLast','width=800,height=600, scrollbars=yes')
	  winLast.focus();
	}

	/**
	 * 바이트 문자 입력가능 문자수 체크
	 * 
	 * @param id : tag id 
	 * @param title : tag title
	 * @param maxLength : 최대 입력가능 수 (byte)
	 * @returns {Boolean}
	 */
	function maxLengthCheck(id, title, maxLength){
		 var obj = $("#"+id);
		 if(maxLength == null) {
			 maxLength = obj.attr("maxLength") != null ? obj.attr("maxLength") : 1000;
		 }
		 
		 if(Number(byteCheck(obj)) > Number(maxLength)) {
			 alert(title + "이(가) 입력가능문자수를 초과하였습니다.\n(영문, 숫자, 일반 특수문자 : " + maxLength + " / 한글, 한자, 기타 특수문자 : " + parseInt(maxLength/2, 10) + ").");
			 obj.focus();
			 return false;
		 } else {
			 return true;
		}
	}

	/**
	 * 바이트수 반환  
	 * 
	 * @param el : tag jquery object
	 * @returns {Number}
	 */
	function byteCheck(el){
		var codeByte = 0;
		for (var idx = 0; idx < el.val().length; idx++) {
			var oneChar = escape(el.val().charAt(idx));
			if ( oneChar.length == 1 ) {
				codeByte ++;
			} else if (oneChar.indexOf("%u") != -1) {
				codeByte += 2;
			} else if (oneChar.indexOf("%") != -1) {
				codeByte ++;
			}
		}
		return codeByte;
	}

</script>

<table width="100%" cellpadding="2" cellspacing="1" class="a" bgcolor="#3d3d3d">
<form name="frmcontents" method="post" action="doMainGatherEventReg.asp" onsubmit="return false;">
<tr bgcolor="#FFFFFF">
    <td width="100" bgcolor="#DDDDFF">Idx</td>
    <td>
        <% if oMainContents.FOneItem.Fidx<>"" then %>
        <%= oMainContents.FOneItem.Fidx %>
        <input type="hidden" name="idx" value="<%= oMainContents.FOneItem.Fidx %>">
        <% else %>

        <% end if %>
    </td>
</tr>
<tr bgcolor="#FFFFFF">
    <td width="100" bgcolor="#DDDDFF">메인카피</td>
    <td height="60">
		<input type="text" name="MainCopy1" id="MainCopy1" value="<%=oMainContents.FOneItem.FMainCopy1%>" size="80"><p>
		<input type="text" name="MainCopy2" id="MainCopy2" value="<%=oMainContents.FOneItem.FMainCopy2%>" size="80">
    </td>
</tr>
<tr bgcolor="#FFFFFF">
    <td width="100" bgcolor="#DDDDFF">기획전1</td>
    <td>
		<table>
		<tr bgcolor="#FFFFFF" height="30">
			<td>이벤트 코드 : </td>
			<td><input type="text" name="Evt_Code1" value="<%=oMainContents.FOneItem.FEvt_Code1%>"> <a href="javascript:jsLastEvent(1);">불러오기</a></td>
		</tr>
		<tr bgcolor="#FFFFFF" height="30">
			<td>메인카피 : </td>
			<td><input type="text" name="Evt_Title1" id="Evt_Title1" value="<%=oMainContents.FOneItem.FEvt_Title1%>" size="50" maxlength="30"></td>
		</tr>
		<tr bgcolor="#FFFFFF" height="30">
			<td>할인율 : </td>
			<td><input type="text" name="Evt_Discount1" value="<%=oMainContents.FOneItem.FEvt_Discount1%>"></td>
		</tr>
		<tr bgcolor="#FFFFFF" height="30">
			<td>쿠폰 할인율 : </td>
			<td><input type="text" name="Evt_Coupon1" value="<%=oMainContents.FOneItem.FEvt_Coupon1%>"></td>
		</tr>
		<tr bgcolor="#FFFFFF" height="30">
			<td>서브카피 : </td>
			<td><input type="text" name="Evt_Subcopy1" id="Evt_Subcopy1" value="<%=oMainContents.FOneItem.FEvt_Subcopy1%>" size="70" maxlength="20"></td>
		</tr>
		</table>
    </td>
</tr>
<tr bgcolor="#FFFFFF">
    <td width="100" bgcolor="#DDDDFF">기획전2</td>
    <td>
		<table>
		<tr bgcolor="#FFFFFF" height="30">
			<td>이벤트 코드 : </td>
			<td><input type="text" name="Evt_Code2" value="<%=oMainContents.FOneItem.FEvt_Code2%>"> <a href="javascript:jsLastEvent(2);">불러오기</a></td>
		</tr>
		<tr bgcolor="#FFFFFF" height="30">
			<td>메인카피 : </td>
			<td><input type="text" name="Evt_Title2" id="Evt_Title2" value="<%=oMainContents.FOneItem.FEvt_Title2%>" size="50" maxlength="30"></td>
		</tr>
		<tr bgcolor="#FFFFFF" height="30">
			<td>할인율 : </td>
			<td><input type="text" name="Evt_Discount2" value="<%=oMainContents.FOneItem.FEvt_Discount2%>"></td>
		</tr>
		<tr bgcolor="#FFFFFF" height="30">
			<td>쿠폰 할인율 : </td>
			<td><input type="text" name="Evt_Coupon2" value="<%=oMainContents.FOneItem.FEvt_Coupon2%>"></td>
		</tr>
		<tr bgcolor="#FFFFFF" height="30">
			<td>서브카피 : </td>
			<td><input type="text" name="Evt_Subcopy2" id="Evt_Subcopy2" value="<%=oMainContents.FOneItem.FEvt_Subcopy2%>" size="70" maxlength="20"></td>
		</tr>
		</table>
    </td>
</tr>
<tr bgcolor="#FFFFFF">
    <td width="100" bgcolor="#DDDDFF">기획전3</td>
    <td>
		<table>
		<tr bgcolor="#FFFFFF" height="30">
			<td>이벤트 코드 : </td>
			<td><input type="text" name="Evt_Code3" value="<%=oMainContents.FOneItem.FEvt_Code3%>"> <a href="javascript:jsLastEvent(3);">불러오기</a></td>
		</tr>
		<tr bgcolor="#FFFFFF" height="30">
			<td>메인카피 : </td>
			<td><input type="text" name="Evt_Title3" id="Evt_Title3" value="<%=oMainContents.FOneItem.FEvt_Title3%>" size="50" maxlength="30"></td>
		</tr>
		<tr bgcolor="#FFFFFF" height="30">
			<td>할인율 : </td>
			<td><input type="text" name="Evt_Discount3" value="<%=oMainContents.FOneItem.FEvt_Discount3%>"></td>
		</tr>
		<tr bgcolor="#FFFFFF" height="30">
			<td>쿠폰 할인율 : </td>
			<td><input type="text" name="Evt_Coupon3" value="<%=oMainContents.FOneItem.FEvt_Coupon3%>"></td>
		</tr>
		<tr bgcolor="#FFFFFF" height="30">
			<td>서브카피 : </td>
			<td><input type="text" name="Evt_Subcopy3" id="Evt_Subcopy3" value="<%=oMainContents.FOneItem.FEvt_Subcopy3%>" size="70" maxlength="20"></td>
		</tr>
		</table>
    </td>
</tr>
<tr bgcolor="#FFFFFF">
  <td width="100" bgcolor="#DDDDFF">시작일</td>
  <td>
  	<input type="text" name="StartDate" id="startdate" value="<%=oMainContents.FOneItem.FStartDate%>">
	<img src="http://webadmin.10x10.co.kr/images/calicon.gif" id="startdate_trigger" border="0" style="cursor:pointer;" align="absbottom" />
	<script type="text/javascript">
	var CAL_Start = new Calendar({
		inputField : "startdate",
		trigger    : "startdate_trigger",
		onSelect: function() {
			var date = Calendar.intToDate(this.selection.get());
			CAL_End.args.min = date;
			CAL_End.redraw();
			this.hide();
		},
		bottomBar: true,
		dateFormat: "%Y-%m-%d"
	});
	</script>
  </td>
</tr>
<tr bgcolor="#FFFFFF">
  <td width="100" bgcolor="#DDDDFF">종료일</td>
  <td>
  	<input type="text" name="EndDate" id="enddate" value="<%=oMainContents.FOneItem.FEndDate%>">
	<img src="http://webadmin.10x10.co.kr/images/calicon.gif" id="enddate_trigger" border="0" style="cursor:pointer;" align="absbottom" />
	<script type="text/javascript">
	var CAL_End = new Calendar({
		inputField : "enddate",
		trigger    : "enddate_trigger",
		onSelect: function() {
			var date = Calendar.intToDate(this.selection.get());
			CAL_Start.args.max = date;
			CAL_Start.redraw();
			this.hide();
		},
		bottomBar: true,
		dateFormat: "%Y-%m-%d"
	});
	</script>
  </td>
</tr>
<tr bgcolor="#FFFFFF">
  <td width="100" bgcolor="#DDDDFF">우선순위</td>
  <td>
  	<input type="text" name="DispOrder" value="<%=oMainContents.FOneItem.FDispOrder%>">
  </td>
</tr>
<tr bgcolor="#FFFFFF">
  <td width="100" bgcolor="#DDDDFF">사용여부</td>
  <td>
  	<input type="radio" name="Isusing" value="Y"<% If oMainContents.FOneItem.FIsusing="Y" Or oMainContents.FOneItem.FIsusing="" Then Response.write " checked" %>> 사용함
	<input type="radio" name="Isusing" value="N"<% If oMainContents.FOneItem.FIsusing="N" Then Response.write " checked" %>> 사용안함
  </td>
</tr>
<% If oMainContents.FOneItem.FRegUser<>"" Then %>
<tr bgcolor="#FFFFFF">
  <td width="100" bgcolor="#DDDDFF">작업자</td>
  <td>
  	작업자 : <%=oMainContents.FOneItem.FRegUser %><br>
	최종작업자 : <%=oMainContents.FOneItem.FLastUser %>
  </td>
</tr>
<% End If %>
<tr bgcolor="#FFFFFF">
    <td colspan="2" align="center"><input type="button" value=" 저 장 " onClick="SaveMainContents(frmcontents);"></td>
</tr>
</form>
</table>
<%
set oMainContents = Nothing
%>

<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
