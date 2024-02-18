<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/sitemasterclass/videoInfoCls.asp"-->
<%
'###############################################
' PageName : videoWrite.asp
' Discription : 동영상 관리 등록/수정
' History : 2009.09.29 허진원 생성
'###############################################

dim videoSn, mode, i
mode=request("mode")
videoSn=request("videoSn")

dim fmainitem
set fmainitem = New Cvideo
fmainitem.FCurrPage = 1
fmainitem.FPageSize=1
fmainitem.FRectVSN=videoSn
fmainitem.GetVideoList
%>
<link href="/js/jqueryui/css/jquery-ui.css" rel="stylesheet">
<script type="text/javascript" src="/js/jquery-1.7.2.min.js"></script>
<script type="text/javascript" src="/js/jqueryui/jquery-ui-1.10.2.custom.min.js"></script>
<script language="javascript">
<!--
$(document).ready(function(){
    $('#videoDiv').change(function(){
        if($('#videoDiv').val() == "mov"){
			$("#mlink").show();
			$("#mdate").show();
			$("#msize").hide();
		}
		else{
			$("#mlink").hide();
			$("#mdate").hide();
			$("#msize").show();
		}
    });

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
      	<% if videoSn<>"" then %>maxDate: "<%=fmainitem.FItemList(0).FendDate%>",<% end if %>
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
      	<% if videoSn<>"" then %>maxDate: "<%=fmainitem.FItemList(0).FstartDate%>",<% end if %>
    	onClose: function( selectedDate ) {
    		$( "#sDt" ).datepicker( "option", "maxDate", selectedDate );
    	}
    });
});

function editcont(){
    //오픈된후 설명만 수정할 경우;;
    var frm=document.inputfrm;
    
    if (confirm('수정 하시겠습니까?')){
        frm.sale_code.value="";
        frm.submit();
    }
    
}

function subcheck(){
	var frm=document.inputfrm;

	if(!frm.videoDiv.value) {
		alert("동영상 구분을 선택해주세요!");
		frm.videoDiv.focus();
		return;
	}

	if(!frm.videoTitle.value) {
		alert("동영상 제목을 입력해주세요!");
		frm.videoTitle.focus();
		return;
	}

	if(!frm.videoFile.value&&!frm.videoSn.value) {
		alert("FLV형태의 동영상 파일을 선택해주세요!");
		frm.videoFile.focus();
		return;
	}
	if($('#videoDiv').val()!="mov"){
		if(!frm.videoWidth.value||frm.videoWidth.value=='0') {
			alert("동영상 너비를 입력해주세요!");
			frm.videoWidth.focus();
			return;
		}

		if(!frm.videoHeight.value||frm.videoHeight.value=='0') {
			alert("동영상 높이를 입력해주세요!");
			frm.videoHeight.focus();
			return;
		}
	}
	else{
		if(frm.linkgubun.value=="") {
			alert("링크 구분을 선택해주세요!");
			frm.linkgubun.focus();
			return;
		}
		if(frm.linkinfo.value=="") {
			alert("링크 고유번호를 입력해주세요!");
			frm.linkinfo.focus();
			return;
		}
		if(frm.startDate.value=="") {
			alert("시작일을 선택해주세요!");
			frm.startDate.focus();
			return;
		}
		if(frm.endDate.value=="") {
			alert("종료일을 선택해주세요!");
			frm.endDate.focus();
			return;
		}
	}

	frm.submit();
}

function delitems(){
	var frm = document.inputfrm;
	if (confirm('본 동영상을 삭제하시겠습니까?')) {
		frm.mode.value="delete";
		frm.submit();
	}
}
//-->
</script>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="inputfrm" method="post" action="<%= uploadImgUrl %>/linkweb/sitemaster/doVideoProcess.asp" enctype="MULTIPART/FORM-DATA">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="mode" value="<% =mode %>">
<tr height="30">
	<td colspan="2" bgcolor="#FFFFFF">
		<img src="/images/icon_star.gif" align="absmiddle">
		<font color="red"><b>동영상 등록/수정</b></font>
	</td>
</tr>
<% if mode="add" then %>
<input type="hidden" name="videoSn" value="">
<tr>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">동영상 구분</td>
	<td bgcolor="#FFFFFF"><%=drawVDivSelect("videoDiv","")%></td>
</tr>
<tr>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">제목</td>
	<td bgcolor="#FFFFFF">
		<input type="text" class="text" name="videoTitle" value="" size="60">
	</td>
</tr>
<tr>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">동영상</td>
	<td bgcolor="#FFFFFF">
		<input type="file" class="text" name="videoFile" value="" size="40"> (※ MP4/FLV/MP3파일, 최대 20MB 이하)
	</td>
</tr>
<tr>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">미리보기 이미지</td>
	<td bgcolor="#FFFFFF">
		<input type="file" class="text" name="videoThumb" value="" size="40"> (※ JPG,GIF 이미지, 최대 300KB 이하)
	</td>
</tr>
<tr id="msize">
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">동영상 크기</td>
	<td bgcolor="#FFFFFF">
		가로 <input type="text" class="text" name="videoWidth" value="0" size="3" style="text-align:right">px × 세로 <input type="text" class="text" name="videoHeight" value="0" size="3" style="text-align:right">px 
		<br>※ 0 입력시 너비에 맞춤
		<br>※ 다이어리 : 450 × 280 / 기타 구분 : 필요에따라 지정
	</td>
</tr>
<tr id="mdate" style="display:none">
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">노출기간</td>
	<td bgcolor="#FFFFFF">
		<input type="text" id="sDt" name="startDate" size="10" />
		~
		<input type="text" id="eDt" name="endDate" size="10" />
	</td>
</tr>
<tr id="mlink" style="display:none">
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">링크 정보</td>
	<td bgcolor="#FFFFFF">
		구분
		<select name="linkgubun" class="select">
			<option value="">링크 구분</option>
			<option value="1">상품</option>
			<option value="2">이벤트</option>
			<option value="3">히치하이커</option>
			<option value="4">브랜드</option>
			<option value="5">다꾸TV</option>
		</select>
		&nbsp;&nbsp;링크 고유번호 <input type="text" class="text" name="linkinfo" size="15">
		<br><br>고유번호 예시)
		<br> 상품 : http://www.10x10.co.kr/shopping/category_prd.asp?itemid=<font color="red">2687010</font>
		<br>이벤트 : http://www.10x10.co.kr/event/eventmain.asp?eventid=<font color="red">102198</font>
		<br>히치하이커 : 고유번호 불필요
		<br>브랜드 : http://www.10x10.co.kr/street/street_brand.asp?makerid=<font color="red">tenten10000</font>
		<br>다꾸TV : http://www.10x10.co.kr/diarystory2020/daccutv_detail.asp?cidx=<font color="red">39</font>
	</td>
</tr>
<% elseif mode="edit" then %>
<tr>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">번호</td>
	<td bgcolor="#FFFFFF">
		<b><%=fmainitem.FItemList(0).FvideoSn%></b>
		<input type="hidden" name="videoSn" value="<%=fmainitem.FItemList(0).FvideoSn%>">
	</td>
</tr>
<tr>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">동영상 구분</td>
	<td bgcolor="#FFFFFF"><%=drawVDivSelect("videoDiv",fmainitem.FItemList(0).FvideoDiv)%></td>
</tr>
<tr>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">제목</td>
	<td bgcolor="#FFFFFF">
		<input type="text" class="text" name="videoTitle" value="<%=fmainitem.FItemList(0).FvideoTitle%>" size="60">
	</td>
</tr>
<tr>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">동영상</td>
	<td bgcolor="#FFFFFF">
		<input type="file" class="text" name="videoFile" value="" size="40"> (※ FLV파일)
		<%
			if Not(fmainitem.FItemList(0).FvideoFile="" or isNull(fmainitem.FItemList(0).FvideoFile)) then
				Response.Write "<br>(현재:" & fmainitem.FItemList(0).FvideoFile & ")"
			end if
		%>
	</td>
</tr>
<tr>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">미리보기 이미지</td>
	<td bgcolor="#FFFFFF">
		<input type="file" class="text" name="videoThumb" value="" size="40"> (※ JPG,GIF 이미지)
		<%
			if Not(fmainitem.FItemList(0).FvideoThumb="" or isNull(fmainitem.FItemList(0).FvideoThumb)) then
				Response.Write "<br>(현재:" & fmainitem.FItemList(0).FvideoThumb & ")"
			end if
		%>
	</td>
</tr>
<tr id="mwidth"<% if fmainitem.FItemList(0).FvideoDiv="mov" then %> style="display:none"<% end if %>>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">동영상 크기</td>
	<td bgcolor="#FFFFFF">
		가로 <input type="text" class="text" name="videoWidth" value="<%=fmainitem.FItemList(0).FvideoWidth%>" size="3" style="text-align:right">px ×
		세로 <input type="text" class="text" name="videoHeight" value="<%=fmainitem.FItemList(0).FvideoHeight%>" size="3" style="text-align:right">px 
		<br>※ 0 입력시 너비에 맞춤
	</td>
</tr>
<tr id="mdate"<% if fmainitem.FItemList(0).FvideoDiv<>"mov" then %> style="display:none"<% end if %>>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">노출기간</td>
	<td bgcolor="#FFFFFF">
		<input type="text" id="sDt" name="startDate" size="10" value="<%=fmainitem.FItemList(0).FstartDate%>" />
		~
		<input type="text" id="eDt" name="endDate" size="10" value="<%=fmainitem.FItemList(0).FendDate%>" />
	</td>
</tr>
<tr id="mlink"<% if fmainitem.FItemList(0).FvideoDiv<>"mov" then %> style="display:none"<% end if %>>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">링크 정보</td>
	<td bgcolor="#FFFFFF">
		구분
		<select name="linkgubun" class="select">
			<option value="">링크 구분</option>
			<option value="1"<% if fmainitem.FItemList(0).Flinkgubun="1" then response.write " selected" %>>상품</option>
			<option value="2"<% if fmainitem.FItemList(0).Flinkgubun="2" then response.write " selected" %>>이벤트</option>
			<option value="3"<% if fmainitem.FItemList(0).Flinkgubun="3" then response.write " selected" %>>히치하이커</option>
			<option value="4"<% if fmainitem.FItemList(0).Flinkgubun="4" then response.write " selected" %>>브랜드</option>
			<option value="5"<% if fmainitem.FItemList(0).Flinkgubun="5" then response.write " selected" %>>다꾸TV</option>
		</select>
		&nbsp;&nbsp;링크 고유번호 <input type="text" class="text" name="linkinfo" value="<%=fmainitem.FItemList(0).Flinkinfo %>" size="15">
		<br><br>고유번호 예시)
		<br> 상품 : http://www.10x10.co.kr/shopping/category_prd.asp?itemid=<font color="red">2687010</font>
		<br>이벤트 : http://www.10x10.co.kr/event/eventmain.asp?eventid=<font color="red">102198</font>
		<br>히치하이커 : 고유번호 불필요
		<br>브랜드 : http://www.10x10.co.kr/street/street_brand.asp?makerid=<font color="red">tenten10000</font>
		<br>다꾸TV : http://www.10x10.co.kr/diarystory2020/daccutv_detail.asp?cidx=<font color="red">39</font>
	</td>
</tr>
<% end if %>
<tr bgcolor="#FFFFFF" >
	<td colspan="2" align="center">
		<input type="button" value=" 저장 " class="button" onclick="subcheck();"> &nbsp;&nbsp;
		<% if mode="edit" then %><input type="button" value=" 삭제 " class="button" onclick="delitems();"> &nbsp;&nbsp;<% end if %>
		<input type="button" value=" 취소 " class="button" onclick="history.back();">
	</td>
</tr>
</form>
</table>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
