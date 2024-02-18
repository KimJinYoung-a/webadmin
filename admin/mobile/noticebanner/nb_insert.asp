<%@ language=vbscript %>
<% option explicit %>
<%
'###############################################
' PageName : nb_insert.asp
' Discription : 모바일 사이트 알림배너
' History : 2013.04.01 이종화
'###############################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/classes/mobile/main_noticebanner.asp" -->
<%
Dim sdate , edate , idx , mode
Dim writer
Dim temputarr
Dim tempSdate , tempEdate , tempShour , tempEhour

	idx = request("idx")
	writer  = session("ssBctCname")

If idx = "" Then 
	mode = "add" 
	idx = 0
Else 
	mode = "modify" 
End If 

dim oNoticeBannerOne
set oNoticeBannerOne = new CMainbanner
oNoticeBannerOne.FRectIdx = idx
oNoticeBannerOne.GetOneContents()

	Function checksel(val) ''셀렉트박스 체크
		temputarr = Split(oNoticeBannerOne.FOneItem.FutArr,",")
		Dim ii  
			For ii = 0 To UBound(temputarr)
				If val = temputarr(ii) Then
					checksel = "checked"
				Exit for
				End If 
			next
	End Function

	''//날짜 시간 분해
	sdate = oNoticeBannerOne.FOneItem.Fstartday
	edate = oNoticeBannerOne.FOneItem.Fendday

	If Not(sdate="" or isNull(sdate)) then
		tempSdate = Left(sdate,10)
		tempShour = Num2Str(hour(sdate),2,"0","R") & ":" & Num2Str(minute(sdate),2,"0","R") & ":" & Num2Str(second(sdate),2,"0","R")
	else
		tempSdate = date
		tempShour = "00:00:00"
	end if

	If Not(edate="" or isNull(edate)) then
		tempEdate = Left(edate,10)
		tempEhour = Num2Str(hour(edate),2,"0","R") & ":" & Num2Str(minute(edate),2,"0","R") & ":" & Num2Str(second(edate),2,"0","R")
	else
		tempEdate = date
		tempEhour = "23:59:59"
	end if
%>
<link href="/js/jqueryui/css/jquery-ui.css" rel="stylesheet">
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript" src="/js/jqueryui/jquery-ui-1.10.2.custom.min.js"></script>
<script type='text/javascript'>
	function jsSubmit(){
		var frm = document.frm;
		if (frm.utArr.value ==""){
			checkval();
		}

		if (frm.utArr.value ==""){
			alert('등급 구분을 먼저 선택 하세요.');
			return;
		}
		
		if (frm.startday.value.length!=10){
			alert('노출기간 시작일을 입력 하세요.');
			frm.startday.focus();
			return;
		}
		
		if (frm.endday.value.length!=10){
			alert('노출기간 종료일을 입력 하세요.');
			frm.endday.focus();
			return;
		}

		if (frm.title.value == "" ){
			alert('제목을 입력 하세요');
			frm.title.focus();
			return;
		}

		if (frm.sorting.value == "" ){
			alert('우선순위를 입력 하세요');
			frm.sorting.focus();
			return;
		}

		if (frm.text.value == "" ){
			alert('텍스트 내용을 입력 하세요');
			frm.text.focus();
			return;
		}

		if (confirm('저장 하시겠습니까?')){
			//frm.target = "_blank";
			frm.action = "/admin/mobile/noticebanner/nb_proc.asp";
			frm.submit();
		}
	}

	//제목 글자수 제한
	function textCounter(field, countfield, maxlimit) {
		if (field.value.length > maxlimit){ 
			field.blur();
			field.value = field.value.substring(0, maxlimit);
			//alert("15자 이내에서 적어주세요");
			field.focus();
		}
		else {
			countfield.value = maxlimit - field.value.length;
			document.getElementById("counttxt").innerHTML = maxlimit - countfield.value;
		}
	}
	// 등급구분 전체선택
	function allcheck(){
		var adddot = ",";
		var frm = document.frm;
		var length = frm.usertype.length-1;
		frm.utnArr.value  = "";
		frm.utArr.value = "";
		if ( frm.usertype[0].checked == true ){
			for ( i = 1 ; i <= length ; i++ )
			{
				if ( i ==  length)
				{
					adddot = "";
				}
				frm.usertype[i].checked = true;
				frm.utnArr.value = frm.utnArr.value+ $("input[name='usertype']").eq(i).attr("value2") + adddot;
				frm.utArr.value =  frm.utArr.value + $("input[name='usertype']").eq(i).attr("value") + adddot;
			}
		}else{
			for ( i = 1 ; i <= length ; i++ )
			{
				frm.usertype[i].checked = false;
				frm.utnArr.value = "";
				frm.utArr.value = "";
			}
		}
	}
	// 등급구분 개별 선택
	function checkval(){
		var adddot = ",";
		var frm = document.frm;
		var length = frm.usertype.length-1;
		frm.usertype[0].checked = false;
		frm.utnArr.value  = "";
		frm.utArr.value = "";
		for ( i = 1 ; i <= length ; i++ )
		{
			if ( frm.usertype[i].checked )
			{
				if ( i ==  length)
				{
					adddot = "";
				}
				frm.utnArr.value = frm.utnArr.value+ $("input[name='usertype']").eq(i).attr("value2") + adddot;
				frm.utArr.value =  frm.utArr.value + $("input[name='usertype']").eq(i).attr("value") + adddot;
			}
		}
	}

	$(function(){
		//달력대화창 설정
		var arrDayMin = ["일","월","화","수","목","금","토"];
		var arrMonth = ["1월","2월","3월","4월","5월","6월","7월","8월","9월","10월","11월","12월"];
	    $("#sDt").datepicker({
			dateFormat: "yy-mm-dd",
			prevText: '이전달', nextText: '다음달', yearSuffix: '년',
			dayNamesMin: arrDayMin,
			monthNames: arrMonth,
			showMonthAfterYear: true,
	    	numberOfMonths: 2,
	    	showCurrentAtPos: 1,
	      	showOn: "button",
	      	maxDate: "<%=tempEdate%>",
	    	onClose: function( selectedDate ) {
	    		$( "#eDt" ).datepicker( "option", "minDate", selectedDate );
	    	}
	    });
	    $("#eDt").datepicker({
			dateFormat: "yy-mm-dd",
			prevText: '이전달', nextText: '다음달', yearSuffix: '년',
			dayNamesMin: arrDayMin,
			monthNames: arrMonth,
			showMonthAfterYear: true,
	    	numberOfMonths: 2,
	      	showOn: "button",
	      	minDate: "<%=tempSdate%>",
	    	onClose: function( selectedDate ) {
	    		$( "#sDt" ).datepicker( "option", "maxDate", selectedDate );
	    	}
	    });
	});
</script>

<table width="800" cellpadding="2" cellspacing="1" class="a" bgcolor="#3d3d3d">
<form name="frm" method="post">
<input type="hidden" name="utnArr" value="">
<input type="hidden" name="utArr" value="">
<input type="hidden" name="iidx" value="<%=idx%>">
<input type="hidden" name="mode" value="<%=mode%>">
<tr bgcolor="#FFFFFF">
	<td bgcolor="#FFF999" align="center">등급 구분</td>
	<td>
		<input type="checkbox" name="usertype" value="9" onclick="allcheck();" value2="전체" <% If mode="modify" then%><%=checksel("9")%><%End If %>/> 전체
		<input type="checkbox" name="usertype" value="5" onclick="checkval();" value2="오렌지" <% If mode="modify" then%><%=checksel("5")%><%End If %>/> 오렌지
		<input type="checkbox" name="usertype" value="0" onclick="checkval();" value2="옐로우" <% If mode="modify" then%><%=checksel("0")%><%End If %>/> 옐로우
		<input type="checkbox" name="usertype" value="1" onclick="checkval();" value2="그린" <% If mode="modify" then%><%=checksel("1")%><%End If %>/> 그린
		<input type="checkbox" name="usertype" value="2" onclick="checkval();" value2="블루" <% If mode="modify" then%><%=checksel("2")%><%End If %>/> 블루
		<input type="checkbox" name="usertype" value="3" onclick="checkval();" value2="VIP실버" <% If mode="modify" then%><%=checksel("3")%><%End If %>/> VIP실버
		<input type="checkbox" name="usertype" value="4" onclick="checkval();" value2="VIP골드" <% If mode="modify" then%><%=checksel("4")%><%End If %>/> VIP골드
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="#FFF999" align="center">노출 기간</td>
	<td>
		<input type="text" id="sDt" name="startday" size="10" value="<%=tempSdate%>" />
		<input type="text" name="sthh" size="8" value="<%=tempShour%>" /> ~
		<input type="text" id="eDt" name="endday" size="10" value="<%=tempEdate%>" />
		<input type="text" name="edhh" size="8" value="<%=tempEhour%>" />
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="#FFF999" align="center">제목</td>
	<td><input type="text" class="text"  name="title" size="50" maxlength="40" onKeyDown="textCounter(this.form.title,this.form.remLen,20);" onKeyUp="textCounter(this.form.title,this.form.remLen,20);" value="<%=oNoticeBannerOne.FOneItem.Ftitle%>"/><input type="hidden" name="remLen" value="20"/>&nbsp;<span id="counttxt"/>0</span>자 / 20자</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="#FFF999" align="center">우선순위</td>
	<td><div style="float:left;"><input type="text" class="text" name="sorting" size="5" maxlength="3" value="<%=oNoticeBannerOne.FOneItem.Fsorting%>"/></div> <div style="float:right;margin-top:5px;margin-right:10px;">※노출방법 : 99(최상단)~1(최하단)</div></td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="#FFF999" align="center">텍스트</td>
	<td>
		<table width="80%" cellpadding="2" cellspacing="1" class="a" bgcolor="#3d3d3d">
			<tr bgcolor="#FFFFFF">
				<td bgcolor="#FFFFFF" width="15%" align="center">내용</td>
				<td><input type="text" class="text" name="text" size="50" maxlength="22" value="<%=oNoticeBannerOne.FOneItem.Ftext%>"/></td>
			</tr>
			<tr bgcolor="#FFFFFF">
				<td colspan="2" bgcolor="#87ceeb" align="left">최대 22자</td>
			</tr>
			<tr bgcolor="#FFFFFF">
				<td bgcolor="#FFFFFF" width="15%"  align="center">연결 카피</td>
				<td><input type="text" class="text" name="infourl" size="50" value="<%=oNoticeBannerOne.FOneItem.Ftextcopy%>" /></td>
			</tr>
			<tr bgcolor="#FFFFFF">
				<td colspan="2" bgcolor="#87ceeb" align="left">예) 확인하러 가기> ( > 자동입력 생략 )</td>
			</tr>
			<tr bgcolor="#FFFFFF">
				<td bgcolor="#FFFFFF" width="15%"  align="center">연결URL</td>
				<td><input type="text" class="text" name="texturl" size="50" value="<%=oNoticeBannerOne.FOneItem.Ftexturl%>" /></td>
			</tr>
			<tr bgcolor="#FFFFFF">
				<td colspan="2" bgcolor="#87ceeb" align="left">예) /event/eventm.asp?eventid=6264 (연결이 없을경우 비워두기)</td>
			</tr>
		</table>
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="#FFF999" align="center">사용여부</td>
	<td><div style="float:left;"><input type="radio" name="isusing" value="1" <%=chkiif(oNoticeBannerOne.FOneItem.Fisusing = "1","checked","")%>/>사용함 &nbsp;&nbsp;&nbsp; <input type="radio" name="isusing" value="0" checked <%=chkiif(oNoticeBannerOne.FOneItem.Fisusing = "0","checked","")%>/>사용안함</div> <div style="float:right;margin-top:5px;margin-right:10px;">※ 사용함 : 메인노출 / 사용안함 : 메인노출 금지</div></td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="#FFF999" align="center" height="25">작업자</td>
	<td><div style="float:left"><strong><%=chkiif(mode="add",writer,oNoticeBannerOne.FOneItem.Fwriter)%></strong></div><div style="float:right;margin-right:10px;">※ 작업자는 파일을 업로드 또는 컨텐츠 등록자가 기록됩니다.</div></td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="#FFF999" align="center" height="25">최종 수정자</td>
	<td><div style="float:left"><strong><%=writer%></strong></div><div style="float:right;margin-right:10px;">※ 최종 수정자는 마지막 업데이트 등록자가 기록됩니다.</div></td>
</tr>
<tr bgcolor="#FFFFFF" align="center">
    <td colspan="2"><input type="button" value=" 취 소 " onClick="history.back(-1)"/><input type="button" value=" 저 장 " onClick="jsSubmit();"/></td>
</tr>
</form>
</table>
<%
set oNoticeBannerOne = Nothing
%>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->