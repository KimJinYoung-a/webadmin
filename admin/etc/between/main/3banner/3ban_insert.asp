<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbCTopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/admin/etc/between/mainCls.asp"-->
<%
Dim idx, imgurl, mode, sortno, isusing, imglink
Dim mainStartDate, mainEndDate, gender
Dim sDt, eDt

idx = requestCheckvar(request("idx"),16)

If idx = "" Then
	mode = "I"
Else
	mode = "U"
End If

If idx <> "" then
	Dim o3ban
	SET o3ban = new cMain
		o3ban.FRectIdx = idx
		o3ban.GetOne3Banner()

		imgurl			= o3ban.FItemList(0).FImgurl
		sortno			= o3ban.FItemList(0).FSortno
		mainStartDate	= Left(o3ban.FItemList(0).FStartdate, 10)
		mainEndDate		= Left(o3ban.FItemList(0).FEnddate, 10)
		isusing			= o3ban.FItemList(0).FIsusing
		imgurl			= o3ban.FItemList(0).FImgurl
		imglink			= o3ban.FItemList(0).FImglink
		gender			= o3ban.FItemList(0).FGender
	SET o3ban = Nothing
End If
%>
<link href="/js/jqueryui/css/jquery-ui.css" rel="stylesheet">
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript" src="/js/jqueryui/jquery-ui-1.10.2.custom.min.js"></script>
<script type='text/javascript'>
function jsSubmit(){
	var frm = document.frm;
	if (confirm('저장 하시겠습니까?')){
		frm.submit();
	}
}

function jsgolist(){
	self.location.href="/admin/etc/between/main/3banner/index.asp";
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

function putLinkText(key,gubun) {
	var frm = document.frm;
	var urllink
	if (gubun == "3" ){
		urllink = frm.imglink;
	}
	switch(key) {
		case 'search':
			urllink.value='/apps/appCom/between/project/?pjt_code=코드';
			break;
	}
}
function jsSetImg(sFolder, sImg, sName, sSpan){
	document.domain ="10x10.co.kr";
	var winImg;
	winImg = window.open('/admin/etc/between/main/3banner/pop_3Banner_uploadimg.asp?sF='+sFolder+'&sImg='+sImg+'&sName='+sName+'&sSpan='+sSpan,'popImg','width=370,height=150');
	winImg.focus();
}
function jsDelImg(sName, sSpan){
	if(confirm("이미지를 삭제하시겠습니까?\n\n삭제 후 저장버튼을 눌러야 처리완료됩니다.")){
	   eval("document.all."+sName).value = "";
	   eval("document.all."+sSpan).style.display = "none";
	}
}
</script>
<table width="900" cellpadding="2" cellspacing="1" class="a" bgcolor="#3d3d3d">
<form name="frm" method="post" action="3ban_process.asp" style="margin:0px;">
<input type="hidden" name="mode" value="<%=mode%>">
<input type="hidden" name="idx" value="<%=idx%>">
<input type="hidden" name="ban" value="<%=imgurl%>">
<input type="hidden" name="menupos" value="<%=menupos%>">
<tr bgcolor="#FFFFFF">
    <td bgcolor="#FFF999" align="center" width="15%">성별</td>
    <td colspan="3">
    	<select name="gender" class="select">
    		<option value="M" <%= Chkiif(gender="M", "selected", "") %> >남자</option>
    		<option value="F" <%= Chkiif(gender="F", "selected", "") %> >여자</option>
    	</select>
    </td>
</tr>
<tr bgcolor="#FFFFFF">
    <td bgcolor="#FFF999" align="center" width="15%">노출기간</td>
    <td colspan="3">
		<input type="text" id="sDt" name="startDate" size="10" value="<%=mainStartDate%>" readonly />
		<input type="text" name="sTm" size="8" value="00:00:00" disabled /> ~
		<input type="text" id="eDt" name="endDate" size="10" value="<%=mainEndDate%>" readonly />
		<input type="text" name="eTm" size="8" value="23:59:59" disabled />
    </td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="#FFF999" align="center" width="15%">이미지</td>
	<td width="45%">
	<input type="button" name="btnBan" value="이미지 등록" onClick="jsSetImg('<%=idx%>','<%= imgurl %>','ban','spanban')" class="button">
		<div id="spanban" style="padding: 5 5 5 5">
		<% If imgurl <> "" Then %>
			<img src="<%=imgurl%>" border="0">
			<a href="javascript:jsDelImg('ban','spanban');"><img src="/images/icon_delete2.gif" border="0"></a>
		<% End If %>
		</div>
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="#FFF999" align="center">이미지 Link</td>
	<td colspan="3"><input type="text" name="imglink" size="80" value="<%=imglink%>"/>
	<br/><br/>ex)
		<font color="#707070">
		- <span style="cursor:pointer" onClick="putLinkText('search','3')">검색결과 링크 : /apps/appCom/between/project/?pjt_code=<font color="darkred">코드</font></span><br>
		</font>
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="#FFF999" align="center">사용여부</td>
	<td colspan="3"><div style="float:left;"><input type="radio" name="isusing" value="Y" <%=chkiif(isusing = "Y","checked","")%> checked />사용함 &nbsp;&nbsp;&nbsp; <input type="radio" name="isusing" value="N"  <%=chkiif(isusing = "N","checked","")%>/>사용안함</div> <div style="float:right;margin-top:5px;margin-right:10px;"></div></td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="#FFF999" align="center">정렬번호</td>
	<td colspan="3"><input type="text" name="sortno" value="<%=chkiif(sortno="","0",sortno)%>" size="2"/></td>
</tr>
<tr bgcolor="#FFFFFF" align="center">
    <td colspan="4"><input type="button" value=" 취 소 " onClick="jsgolist();"/><input type="button" value=" 저 장 " onClick="jsSubmit();"/></td>
</tr>
</form>
</table>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbCTclose.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->