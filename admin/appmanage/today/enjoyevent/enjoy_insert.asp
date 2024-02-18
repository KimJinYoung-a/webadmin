<%@ language=vbscript %>
<% option explicit %>
<%
'###############################################
' PageName : enjoy_insert.asp
' Discription : 모바일 enjoybanner
' History : 2013.12.14 이종화
'###############################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/classes/mobile/enjoyeventCls.asp" -->
<%
Dim idx , subImage1 , subImage2 , subImage3 , subImage4 , isusing , mode
Dim srcSDT , srcEDT 
Dim mainStartDate, mainEndDate
Dim sDt, sTm, eDt, eTm
Dim img1alt , img2alt , img3alt , img4alt , img1url , img2url , img3url , img4url
Dim img1text , img2text , img3text , img4text
Dim img1sale ,img2sale , img3sale , img4sale
Dim prevDate , ordertext
Dim img1stdate , img2stdate , img3stdate , img4stdate 
Dim img1eddate , img2eddate , img3eddate , img4eddate
Dim img1sc , img2sc , img3sc , img4sc

	idx = requestCheckvar(request("idx"),16)
	srcSDT = request("sDt")
	srcEDT = request("eDt")
	prevDate = request("prevDate")

If idx = "" Then 
	mode = "add" 
Else 
	mode = "modify" 
End If 

If idx <> "" then
	dim oEnjoyeventOne
	set oEnjoyeventOne = new CMainbanner
	oEnjoyeventOne.FRectIdx = idx
	oEnjoyeventOne.GetOneContents()

	img1alt				=	oEnjoyeventOne.FOneItem.Fimg1alt
	img2alt				=	oEnjoyeventOne.FOneItem.Fimg2alt
	img3alt				=	oEnjoyeventOne.FOneItem.Fimg3alt
	img4alt				=	oEnjoyeventOne.FOneItem.Fimg4alt

	img1url				=	oEnjoyeventOne.FOneItem.Fimg1url
	img2url				=	oEnjoyeventOne.FOneItem.Fimg2url
	img3url				=	oEnjoyeventOne.FOneItem.Fimg3url
	img4url				=	oEnjoyeventOne.FOneItem.Fimg4url

	subImage1			=	oEnjoyeventOne.FOneItem.Fimg1
	subImage2			=	oEnjoyeventOne.FOneItem.Fimg2
	subImage3			=	oEnjoyeventOne.FOneItem.Fimg3
	subImage4			=	oEnjoyeventOne.FOneItem.Fimg4

	img1text			=	oEnjoyeventOne.FOneItem.Fimg1text
	img2text			=	oEnjoyeventOne.FOneItem.Fimg2text
	img3text			=	oEnjoyeventOne.FOneItem.Fimg3text
	img4text			=	oEnjoyeventOne.FOneItem.Fimg4text

	img1sale			=	oEnjoyeventOne.FOneItem.Fimg1sale
	img2sale			=	oEnjoyeventOne.FOneItem.Fimg2sale
	img3sale			=	oEnjoyeventOne.FOneItem.Fimg3sale
	img4sale			=	oEnjoyeventOne.FOneItem.Fimg4sale

	img1stdate			=	oEnjoyeventOne.FOneItem.Fimg1stdate
	img2stdate			=	oEnjoyeventOne.FOneItem.Fimg2stdate
	img3stdate			=	oEnjoyeventOne.FOneItem.Fimg3stdate
	img4stdate			=	oEnjoyeventOne.FOneItem.Fimg4stdate

	img1eddate			=	oEnjoyeventOne.FOneItem.Fimg1eddate
	img2eddate			=	oEnjoyeventOne.FOneItem.Fimg2eddate
	img3eddate			=	oEnjoyeventOne.FOneItem.Fimg3eddate
	img4eddate			=	oEnjoyeventOne.FOneItem.Fimg4eddate

	img1sc				=	oEnjoyeventOne.FOneItem.Fimg1sc
	img2sc				=	oEnjoyeventOne.FOneItem.Fimg2sc
	img3sc				=	oEnjoyeventOne.FOneItem.Fimg3sc
	img4sc				=	oEnjoyeventOne.FOneItem.Fimg4sc

	mainStartDate		=	oEnjoyeventOne.FOneItem.Fstartdate
	mainEndDate			=	oEnjoyeventOne.FOneItem.Fenddate 
	isusing				=	oEnjoyeventOne.FOneItem.Fisusing
	ordertext			=	oEnjoyeventOne.FOneItem.Fordertext

	set oEnjoyeventOne = Nothing
End If 

if Not(mainStartDate="" or isNull(mainStartDate)) then
	sDt = left(mainStartDate,10)
	sTm = Num2Str(hour(mainStartDate),2,"0","R") &":"& Num2Str(minute(mainStartDate),2,"0","R") &":"& Num2Str(second(mainStartDate),2,"0","R")
else
	if srcSDT<>"" then
		sDt = left(srcSDT,10)
	else
		sDt = date
	end if
	sTm = "00:00:00"
end if

if Not(mainEndDate="" or isNull(mainEndDate)) then
	eDt = left(mainEndDate,10)
	eTm = Num2Str(hour(mainEndDate),2,"0","R") &":"& Num2Str(minute(mainEndDate),2,"0","R") &":"& Num2Str(second(mainEndDate),2,"0","R")
else
	if srcEDT<>"" then
		eDt = left(srcEDT,10)
	else
		eDt = date
	end if
	eTm = "23:59:59"
end if
%>
<link href="/js/jqueryui/css/jquery-ui.css" rel="stylesheet">
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript" src="/js/jqueryui/jquery-ui-1.10.2.custom.min.js"></script>
<script type='text/javascript'>
	function jsSubmit(){
		var frm = document.frm;

		if (confirm('저장 하시겠습니까?')){
			//frm.target = "blank";
			frm.submit();
		}
	}
	function jsgolist(){
		self.location.href="/admin/appmanage/today/enjoyevent/";
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
      	<% if Idx<>"" then %>maxDate: "<%=eDt%>",<% end if %>
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
      	<% if Idx<>"" then %>minDate: "<%=sDt%>",<% end if %>
    	onClose: function( selectedDate ) {
    		$( "#sDt" ).datepicker( "option", "maxDate", selectedDate );
    	}
    });

});

//-- jsPopCal : 달력 팝업 --//
function jsPopCal(sName){
	var winCal;
	winCal = window.open('/lib/common_cal.asp?DN='+sName,'pCal','width=250, height=200');
	winCal.focus();
}

function putLinkText(key,gubun) {
	var frm = document.frm;
	var urllink
	if (gubun == "1" )
	{
		urllink = frm.img1url;
	}else if (gubun == "2" ){
		urllink = frm.img2url;
	}else if (gubun == "3" ){
		urllink = frm.img3url;
	}else{
		urllink = frm.img4url;
	}
	switch(key) {
		case 'event':
			urllink.value='/event/eventmain.asp?eventid=이벤트번호';
			break;
		case 'itemid':
			urllink.value='/category/category_itemprd.asp?itemid=상품코드';
			break;
	}
}
</script>
<table width="1000" cellpadding="2" cellspacing="1" class="a" bgcolor="#3d3d3d">
<form name="frm" method="post" action="<%=uploadUrl%>/linkweb/mobile/enjoyevent_proc.asp" enctype="multipart/form-data" style="margin:0px;">
<input type="hidden" name="mode" value="<%=mode%>">
<input type="hidden" name="adminid" value="<%=session("ssBctId")%>">
<input type="hidden" name="idx" value="<%=idx%>">
<input type="hidden" name="prevDate" value="<%=prevDate%>">
<tr bgcolor="#FFFFFF">
    <td bgcolor="#FFF999" align="center" width="15%">노출기간</td>
    <td colspan="3">
		<input type="text" id="sDt" name="StartDate" size="10" value="<%=sDt%>" />
		<input type="text" name="sTm" size="8" value="<%=sTm%>" /> ~
		<input type="text" id="eDt" name="EndDate" size="10" value="<%=eDt%>" />
		<input type="text" name="eTm" size="8" value="<%=eTm%>" />
    </td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="#FFF999" height="30" colspan="4" style="text-align:center">1번 이벤트</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="#FFF999" align="center" width="15%">1번이벤트</td>
	<td width="45%">
		<input type="file" name="subImage1" class="file" title="이벤트 #1" require="N" style="width:80%;" />
		<% if subImage1<>"" then %>
		<br>
		<img src="<%= subImage1 %>" width="100" /><br><%= subImage1 %>
		<% end if %>		
	</td>
	<td bgcolor="#FFF999" align="center" width="10%">1번이벤트<br/>이미지 alt</td>
	<td width="20%"><input type="text" name="img1alt" value="<%=img1alt%>" size="20" maxlength="20"/></td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="#FFF999" align="center">1번이벤트 제목</td>
	<td width="45%"><input type="text" name="img1text" size="50" value="<%=img1text%>"/></td>
	<td bgcolor="#FFF999" align="center" width="10%">1번이벤트 할인</td>
	<td width="20%">할인 : <input type="radio" name="img1sc" value="1" <%=chkiif(img1sc = 1,"checked","")%>/> 쿠폰 : <input type="radio" name="img1sc" value="2" <%=chkiif(img1sc = 2,"checked","")%>/> <input type="text" name="img1sale" size="10" value="<%=img1sale%>" maxlength="10"/>&nbsp;</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="#FFF999" align="center">1번이벤트 시작일 - 종료일</td>
	<td colspan="3">
		<input type="text" name="img1stdate" size="10" value="<%=img1stdate%>" onClick="jsPopCal('img1stdate');"/>
		-
		<input type="text" name="img1eddate" size="10" value="<%=img1eddate%>" onClick="jsPopCal('img1eddate');"/>
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="#FFF999" align="center">1번이벤트 URL</td>
	<td colspan="3"><input type="text" name="img1url" size="80" value="<%=img1url%>"/>
	<br/><br/>ex)
		<font color="#707070">
		- <span style="cursor:pointer" onClick="putLinkText('event','1')">이벤트 링크 : /event/eventmain.asp?eventid=<font color="darkred">이벤트코드</font></span><br>
		- <span style="cursor:pointer" onClick="putLinkText('itemid','1')">상품코드 링크 : /category/category_itemprd.asp?itemid=<font color="darkred">상품코드 (O)</font></span><br>
		</font>
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="#FFF999" height="30" colspan="4" style="text-align:center">2번 이벤트</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="#FFF999" align="center" width="15%">2번이벤트</td>
	<td width="45%">
		<input type="file" name="subImage2" class="file" title="이벤트 #1" require="N" style="width:80%;" />
		<% if subImage2<>"" then %>
		<br>
		<img src="<%= subImage2 %>" width="100" /><br><%= subImage2 %>
		<% end if %>		
	</td>
	<td bgcolor="#FFF999" align="center" width="10%">2번이벤트<br/>이미지 alt</td>
	<td width="20%"><input type="text" name="img2alt" value="<%=img2alt%>" size="20" maxlength="20"/></td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="#FFF999" align="center">2번이벤트 제목</td>
	<td width="45%"><input type="text" name="img2text" size="50" value="<%=img2text%>"/></td>
	<td bgcolor="#FFF999" align="center" width="10%">2번이벤트 할인</td>
	<td width="20%">할인 : <input type="radio" name="img2sc" value="1" <%=chkiif(img2sc = 1,"checked","")%>/> 쿠폰 : <input type="radio" name="img2sc" value="2" <%=chkiif(img2sc = 2,"checked","")%>/> <input type="text" name="img2sale" size="10" value="<%=img2sale%>" maxlength="10"/>&nbsp;</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="#FFF999" align="center">2번이벤트 시작일 - 종료일</td>
	<td colspan="3">
		<input type="text" name="img2stdate" size="10" value="<%=img2stdate%>" onClick="jsPopCal('img2stdate');"/>
		-
		<input type="text" name="img2eddate" size="10" value="<%=img2eddate%>" onClick="jsPopCal('img2eddate');"/>
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="#FFF999" align="center">2번이벤트 URL</td>
	<td colspan="3"><input type="text" name="img2url" size="80" value="<%=img2url%>"/>
	<br/><br/>ex)
		<font color="#707070">
		- <span style="cursor:pointer" onClick="putLinkText('event','2')">이벤트 링크 : /event/eventmain.asp?eventid=<font color="darkred">이벤트코드</font></span><br>
		- <span style="cursor:pointer" onClick="putLinkText('itemid','2')">상품코드 링크 : /category/category_itemprd.asp?itemid=<font color="darkred">상품코드 (O)</font></span><br>
		</font>
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="#FFF999" height="30" colspan="4" style="text-align:center">3번 이벤트</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="#FFF999" align="center">3번이벤트</td>
	<td>
		<input type="file" name="subImage3" class="file" title="이벤트 #1" require="N" style="width:80%;" />
		<% if subImage3<>"" then %>
		<br>
		<img src="<%= subImage3 %>" width="100" /><br><%= subImage3 %>
		<% end if %>		
	</td>
	<td bgcolor="#FFF999" align="center">3번이벤트<br/>이미지 alt</td>
	<td><input type="text" name="img3alt" value="<%=img3alt%>" size="20" maxlength="20"/></td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="#FFF999" align="center">3번이벤트 제목</td>
	<td width="45%"><input type="text" name="img3text" size="50" value="<%=img3text%>"/></td>
	<td bgcolor="#FFF999" align="center" width="10%">3번이벤트 할인</td>
	<td width="20%">할인 : <input type="radio" name="img3sc" value="1" <%=chkiif(img3sc = 1,"checked","")%>/> 쿠폰 : <input type="radio" name="img3sc" value="2" <%=chkiif(img3sc = 2,"checked","")%>/> <input type="text" name="img3sale" size="10" value="<%=img3sale%>" maxlength="10"/>&nbsp;</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="#FFF999" align="center">3번이벤트 시작일 - 종료일</td>
	<td colspan="3">
		<input type="text" name="img3stdate" size="10" value="<%=img3stdate%>" onClick="jsPopCal('img3stdate');"/>
		-
		<input type="text" name="img3eddate" size="10" value="<%=img3eddate%>" onClick="jsPopCal('img3eddate');"/>
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="#FFF999" align="center">3번이벤트 URL</td>
	<td colspan="3"><input type="text" name="img3url" size="80" value="<%=img3url%>"/>
	<br/><br/>ex)
		<font color="#707070">
		- <span style="cursor:pointer" onClick="putLinkText('event','3')">이벤트 링크 : /event/eventmain.asp?eventid=<font color="darkred">이벤트코드</font></span><br>
		- <span style="cursor:pointer" onClick="putLinkText('itemid','3')">상품코드 링크 : /category/category_itemprd.asp?itemid=<font color="darkred">상품코드 (O)</font></span><br>
		</font>
	</td>
</tr>
<!--<tr bgcolor="#FFFFFF">
	<td bgcolor="#FFF999" height="30" colspan="4" style="text-align:center">4번 이벤트</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="#FFF999" align="center">4번이벤트</td>
	<td>
		<input type="file" name="subImage4" class="file" title="이벤트 #1" require="N" style="width:80%;" />
		<% if subImage4<>"" then %>
		<br>
		<img src="<%= subImage4 %>" width="100" /><br><%= subImage4 %>
		<% end if %>		
	</td>
	<td bgcolor="#FFF999" align="center">4번이벤트<br/>이미지 alt</td>
	<td><input type="text" name="img4alt" value="<%=img4alt%>" size="20" maxlength="20"/></td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="#FFF999" align="center">4번이벤트 제목</td>
	<td width="45%"><input type="text" name="img4text" size="50" value="<%=img4text%>"/></td>
	<td bgcolor="#FFF999" align="center" width="10%">4번이벤트 할인</td>
	<td width="20%">할인 : <input type="radio" name="img4sc" value="1"/> 쿠폰 : <input type="radio" name="img4sc" value="2"/> <input type="text" name="img4sale" size="10" value="<%=img4sale%>" maxlength="2"/>&nbsp;</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="#FFF999" align="center">4번이벤트 시작일 - 종료일</td>
	<td colspan="3">
		<input type="text" name="img4stdate" size="10" value="<%=img4stdate%>" onClick="jsPopCal('img4stdate');"/>
		-
		<input type="text" name="img4eddate" size="10" value="<%=img4eddate%>" onClick="jsPopCal('img4eddate');"/>
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="#FFF999" align="center">4번이벤트 URL</td>
	<td colspan="3"><input type="text" name="img4url" size="80" value="<%=img4url%>"/>
	<br/><br/>ex)
		<font color="#707070">
		- <span style="cursor:pointer" onClick="putLinkText('event','4')">이벤트 링크 : /event/eventmain.asp?eventid=<font color="darkred">이벤트코드</font></span><br>
		- <span style="cursor:pointer" onClick="putLinkText('itemid','4')">상품코드 링크 : /category/category_itemprd.asp?itemid=<font color="darkred">상품코드 (O)</font></span><br>
		</font>
	</td>
</tr> -->
<tr bgcolor="#FFFFFF">
	<td bgcolor="#FFF999" align="center">사용여부</td>
	<td colspan="3"><div style="float:left;"><input type="radio" name="isusing" value="Y" <%=chkiif(isusing = "Y","checked","")%> checked />사용함 &nbsp;&nbsp;&nbsp; <input type="radio" name="isusing" value="N"  <%=chkiif(isusing = "N","checked","")%>/>사용안함</div> <div style="float:right;margin-top:5px;margin-right:10px;"></div></td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="#FFF999" align="center">작업자 지시사항</td>
	<td colspan="3"><textarea name="ordertext" cols="80" rows="8"/><%=ordertext%></textarea></td>
</tr>
<tr bgcolor="#FFFFFF" align="center">
    <td colspan="4"><input type="button" value=" 취 소 " onClick="jsgolist();"/><input type="button" value=" 저 장 " onClick="jsSubmit();"/></td>
</tr>
</form>
</table>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->