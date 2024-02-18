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
		oMainContents.GetOneEnjoyMainContents

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
	    if (frm.BGColor.value==""){
	        alert('배경색을 먼저 선택 하세요.');
	        frm.BGColor.focus();
	        return;
	    }

	    if (frm.Evt_Code.value==""){
	        alert('이벤트 번호를 입력 하세요.');
	        frm.Evt_Code.focus();
	        return;
	    }

		if (frm.Evt_Title.value==""){
	        alert('이벤트 메인카피를 입력 하세요.');
	        frm.Evt_Title.focus();
	        return;
	    }
		
		if(!maxLengthCheck("Evt_Title","메인카피",48))
		{
			frm.Evt_Title.focus();
			return;
		}

		if (frm.Evt_Subcopy.value==""){
	        alert('이벤트 서브카피를 입력 하세요.');
	        frm.Evt_Subcopy.focus();
	        return;
	    }

		if(!maxLengthCheck("Evt_Subcopy","서브카피",80))
		{
			frm.Evt_Title.focus();
			return;
		}

		if (frm.Item1.value==""){
	        alert('상품 1번을 입력 하세요.');
	        frm.Item1.focus();
	        return;
	    }

		if (frm.Item2.value==""){
	        alert('상품 2번을 입력 하세요.');
	        frm.Item2.focus();
	        return;
	    }

		if (frm.Item3.value==""){
	        alert('상품 3번을 입력 하세요.');
	        frm.Item3.focus();
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
	function jsLastEvent(){
	  winLast = window.open('pop_event_lastlist.asp','pLast','width=800,height=600, scrollbars=yes')
	  winLast.focus();
	}

	function fnviewimg(num){
		var itemid = $("#Item"+num).val();
		$.ajax({
			type: "POST",
			url: "/admin/sitemaster/lib/item_image_view_act.asp",
			data: "Itemid="+itemid,
			cache: false,
			success: function(message) {
				$("#img"+num).empty().html("<img src='"+message+"' width='100' height='100' border='0'>");
			},
			error: function(err) {
				alert(err.responseText);
			}
		});
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
<form name="frmcontents" method="post" action="doMainEnjoyContentsReg.asp" onsubmit="return false;">
<tr bgcolor="#FFFFFF">
    <td width="150" bgcolor="#DDDDFF">Idx</td>
    <td>
        <% if oMainContents.FOneItem.Fidx<>"" then %>
        <%= oMainContents.FOneItem.Fidx %>
        <input type="hidden" name="idx" value="<%= oMainContents.FOneItem.Fidx %>">
        <% else %>

        <% end if %>
    </td>
</tr>
<tr bgcolor="#FFFFFF">
    <td width="150" bgcolor="#DDDDFF">배경색</td>
    <td>
		<table id="colorselect1">
			<tr>
				<td><table id='cline11' border='0' cellpadding='0' cellspacing='1' bgcolor='<% If oMainContents.FOneItem.FBGColor="#FBD65C" Or oMainContents.FOneItem.FBGColor="" Then %>#DD3300<% Else %>#dddddd<% End If %>'><tr><td bgcolor='#FFFFFF'><table border='0' cellpadding='0' cellspacing='2'><tr><td bgcolor='#FBD65C' style="font-size:8px"><a href='javascript:selColorChip("#FBD65C",11)' onfocus='this.blur()'><img src='http://testwebadmin.10x10.co.kr/images/space.gif' alt='주황' width='12' height='12' border='0'></a></td></tr></table></td></tr></table></td>
				<td><table id='cline1' border='0' cellpadding='0' cellspacing='1' bgcolor='<% If oMainContents.FOneItem.FBGColor="#FFB137" Then %>#DD3300<% Else %>#dddddd<% End If %>'><tr><td bgcolor='#FFFFFF'><table border='0' cellpadding='0' cellspacing='2'><tr><td bgcolor='#FFB137' style="font-size:8px"><a href='javascript:selColorChip("#FFB137",1)' onfocus='this.blur()'><img src='http://testwebadmin.10x10.co.kr/images/space.gif' alt='주황' width='12' height='12' border='0'></a></td></tr></table></td></tr></table></td>
				<td><table id='cline2' border='0' cellpadding='0' cellspacing='1' bgcolor='<% If oMainContents.FOneItem.FBGColor="#DCBBEC" Then %>#DD3300<% Else %>#dddddd<% End If %>'><tr><td bgcolor='#FFFFFF'><table border='0' cellpadding='0' cellspacing='2'><tr><td bgcolor='#DCBBEC' style="font-size:8px"><a href='javascript:selColorChip("#DCBBEC",2)' onfocus='this.blur()'><img src='http://testwebadmin.10x10.co.kr/images/space.gif' alt='주황' width='12' height='12' border='0'></a></td></tr></table></td></tr></table></td>
				<td><table id='cline3' border='0' cellpadding='0' cellspacing='1' bgcolor='<% If oMainContents.FOneItem.FBGColor="#C1BEFE" Then %>#DD3300<% Else %>#dddddd<% End If %>'><tr><td bgcolor='#FFFFFF'><table border='0' cellpadding='0' cellspacing='2'><tr><td bgcolor='#C1BEFE' style="font-size:8px"><a href='javascript:selColorChip("#C1BEFE",3)' onfocus='this.blur()'><img src='http://testwebadmin.10x10.co.kr/images/space.gif' alt='주황' width='12' height='12' border='0'></a></td></tr></table></td></tr></table></td>
				<td><table id='cline4' border='0' cellpadding='0' cellspacing='1' bgcolor='<% If oMainContents.FOneItem.FBGColor="#B9DAFA" Then %>#DD3300<% Else %>#dddddd<% End If %>'><tr><td bgcolor='#FFFFFF'><table border='0' cellpadding='0' cellspacing='2'><tr><td bgcolor='#B9DAFA' style="font-size:8px"><a href='javascript:selColorChip("#B9DAFA",4)' onfocus='this.blur()'><img src='http://testwebadmin.10x10.co.kr/images/space.gif' alt='주황' width='12' height='12' border='0'></a></td></tr></table></td></tr></table></td>
				<td><table id='cline5' border='0' cellpadding='0' cellspacing='1' bgcolor='<% If oMainContents.FOneItem.FBGColor="#AAE9DB" Then %>#DD3300<% Else %>#dddddd<% End If %>'><tr><td bgcolor='#FFFFFF'><table border='0' cellpadding='0' cellspacing='2'><tr><td bgcolor='#AAE9DB' style="font-size:8px"><a href='javascript:selColorChip("#AAE9DB",5)' onfocus='this.blur()'><img src='http://testwebadmin.10x10.co.kr/images/space.gif' alt='주황' width='12' height='12' border='0'></a></td></tr></table></td></tr></table></td>
				<td><table id='cline6' border='0' cellpadding='0' cellspacing='1' bgcolor='<% If oMainContents.FOneItem.FBGColor="#CBF09C" Then %>#DD3300<% Else %>#dddddd<% End If %>'><tr><td bgcolor='#FFFFFF'><table border='0' cellpadding='0' cellspacing='2'><tr><td bgcolor='#CBF09C' style="font-size:8px"><a href='javascript:selColorChip("#CBF09C",6)' onfocus='this.blur()'><img src='http://testwebadmin.10x10.co.kr/images/space.gif' alt='주황' width='12' height='12' border='0'></a></td></tr></table></td></tr></table></td>
				<td><table id='cline7' border='0' cellpadding='0' cellspacing='1' bgcolor='<% If oMainContents.FOneItem.FBGColor="#DFDFDF" Then %>#DD3300<% Else %>#dddddd<% End If %>'><tr><td bgcolor='#FFFFFF'><table border='0' cellpadding='0' cellspacing='2'><tr><td bgcolor='#DFDFDF' style="font-size:8px"><a href='javascript:selColorChip("#DFDFDF",7)' onfocus='this.blur()'><img src='http://testwebadmin.10x10.co.kr/images/space.gif' alt='주황' width='12' height='12' border='0'></a></td></tr></table></td></tr></table></td>
				<td><table id='cline8' border='0' cellpadding='0' cellspacing='1' bgcolor='<% If oMainContents.FOneItem.FBGColor="#C0C0C0" Then %>#DD3300<% Else %>#dddddd<% End If %>'><tr><td bgcolor='#FFFFFF'><table border='0' cellpadding='0' cellspacing='2'><tr><td bgcolor='#C0C0C0' style="font-size:8px"><a href='javascript:selColorChip("#C0C0C0",8)' onfocus='this.blur()'><img src='http://testwebadmin.10x10.co.kr/images/space.gif' alt='주황' width='12' height='12' border='0'></a></td></tr></table></td></tr></table></td>
				<td>직접입력<input type="text" name="BGColor" value="<%=oMainContents.FOneItem.FBGColor%>"></td>
			</tr>
		</table>
    </td>
</tr>
<tr bgcolor="#FFFFFF">
    <td width="150" bgcolor="#DDDDFF">타입</td>
    <td>
    	<input type="radio" name="Evt_Type" value="1"<% If oMainContents.FOneItem.FEvt_Type="1" Then Response.write " checked" %> onClick="jsLastEvent()"> 이벤트 불러오기
		<input type="radio" name="Evt_Type" value="2"<% If oMainContents.FOneItem.FEvt_Type="2" Then Response.write " checked" %>> 직접입력
    </td>
</tr>
<tr bgcolor="#FFFFFF">
    <td width="150" bgcolor="#DDDDFF">이벤트</td>
    <td>
    	이벤트 코드 : <input type="text" name="Evt_Code" value="<%=oMainContents.FOneItem.FEvt_Code%>"> <a href="">미리보기</a>
    </td>
</tr>
<tr bgcolor="#FFFFFF">
    <td width="150" bgcolor="#DDDDFF">메인카피</td>
    <td>
		<input type="text" name="Evt_Title" id="Evt_Title" value="<%=oMainContents.FOneItem.FEvt_Title%>" size="50"><br>
		할인율 : <input type="text" name="Evt_Discount" value="<%=oMainContents.FOneItem.FEvt_Discount%>">
    </td>
</tr>
<tr bgcolor="#FFFFFF">
    <td width="150" bgcolor="#DDDDFF">서브카피</td>
    <td>
    	<input type="text" name="Evt_Subcopy" id="Evt_Subcopy" value="<%=oMainContents.FOneItem.FEvt_Subcopy%>" size="80">
    </td>
</tr>
<tr bgcolor="#FFFFFF">
  <td width="150" bgcolor="#DDDDFF">상품1</td>
  <td>
  	<input type="text" name="Item1" id="Item1" value="<%=oMainContents.FOneItem.FItem1%>"> <a href="javascript:fnviewimg(1);">미리보기</a><br>
	<div id="img1"><img src="<% = GetItemImageLoad(oMainContents.FOneItem.FItem1) %>" width="100" height="100" border="0"></div>
  </td>
</tr>
<tr bgcolor="#FFFFFF">
  <td width="150" bgcolor="#DDDDFF">상품2</td>
  <td>
  	<input type="text" name="Item2" id="Item2" value="<%=oMainContents.FOneItem.FItem2%>"> <a href="javascript:fnviewimg(2);">미리보기</a><br>
	<div id="img2"><img src="<% = GetItemImageLoad(oMainContents.FOneItem.FItem2) %>" width="100" height="100" border="0"></div>
  </td>
</tr>
<tr bgcolor="#FFFFFF">
  <td width="150" bgcolor="#DDDDFF">상품3</td>
  <td>
  	<input type="text" name="Item3" id="Item3" value="<%=oMainContents.FOneItem.FItem3%>"> <a href="javascript:fnviewimg(3);">미리보기</a><br>
	<div id="img3"><img src="<% = GetItemImageLoad(oMainContents.FOneItem.FItem3) %>" width="100" height="100" border="0"></div>
  </td>
</tr>
<tr bgcolor="#FFFFFF">
  <td width="150" bgcolor="#DDDDFF">시작일</td>
  <td>
  	<input type="text" name="StartDate" id="startdate" value="<%=chkiif(idx=0,prevDate,oMainContents.FOneItem.FStartDate)%>">
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
  <td width="150" bgcolor="#DDDDFF">종료일</td>
  <td>
  	<input type="text" name="EndDate" id="enddate" value="<%=chkiif(idx=0,prevDate,oMainContents.FOneItem.FEndDate)%>">
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
  <td width="150" bgcolor="#DDDDFF">우선순위</td>
  <td>
  	<input type="text" name="DispOrder" value="<%=oMainContents.FOneItem.FDispOrder%>">
  </td>
</tr>
<tr bgcolor="#FFFFFF">
  <td width="150" bgcolor="#DDDDFF">사용여부</td>
  <td>
  	<input type="radio" name="Isusing" value="Y"<% If oMainContents.FOneItem.FIsusing="Y" Or oMainContents.FOneItem.FIsusing="" Then Response.write " checked" %>> 사용함
	<input type="radio" name="Isusing" value="N"<% If oMainContents.FOneItem.FIsusing="N" Then Response.write " checked" %>> 사용안함
  </td>
</tr>
<% If oMainContents.FOneItem.FRegUser<>"" Then %>
<tr bgcolor="#FFFFFF">
  <td width="150" bgcolor="#DDDDFF">작업자</td>
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
