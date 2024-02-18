<%@ language=vbscript %>
<% option explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db3open.asp" -->
<!-- #include virtual="/lib/classes/admin/menucls.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/search/search_manageCls.asp"-->
<%
dim menupos, imenuposStr, imenuposnotice, imenuposhelp
menupos = request("menupos")
if menupos ="" then menupos=1

imenuposStr = fnGetMenuPos(menupos, imenuposnotice, imenuposhelp)
'// 즐겨찾기
dim IsMenuFavoriteAdded
IsMenuFavoriteAdded = fnGetMenuFavoriteAdded(session("ssBctID"), menupos)


Dim i, cMain, vIdx, vBgGubun, vBgColor, vBgImg, vMaskingImg, vViewGubun, vSDate, vEDate, vUseYN, vShhmmss, vEhhmmss
Dim vTextInfoUse, vTextInfo1, vTextInfo1url, vTextInfo2, vTextInfo2url, vMemo, vRegdate, vLastUserName, vLastdate, vRegUserName
vIdx = requestCheckVar(Request("idx"),15)

If vIdx <> "" Then
	Set cMain = New CSearchMng
	cMain.FRectIdx = vIdx
	cMain.sbMainManageDetail

	vBgGubun = cMain.FOneItem.Fbggubun
	vBgColor = cMain.FOneItem.Fbgcolor
	vBgImg = cMain.FOneItem.Fbgimg
	vMaskingImg = cMain.FOneItem.Fmaskingimg
	vViewGubun = cMain.FOneItem.Fviewgubun
	vSDate = cMain.FOneItem.Fsdate
	vEDate = cMain.FOneItem.Fedate
	vShhmmss = cMain.FOneItem.Fshhmmss
	vEhhmmss = cMain.FOneItem.Fehhmmss
	vUseYN = cMain.FOneItem.Fuseyn
	vTextInfoUse = cMain.FOneItem.Ftextinfouse
	vTextInfo1 = cMain.FOneItem.Ftextinfo1
	vTextInfo1url = cMain.FOneItem.Ftextinfo1url
	vTextInfo2 = cMain.FOneItem.Ftextinfo2
	vTextInfo2url = cMain.FOneItem.Ftextinfo2url
	vMemo = cMain.FOneItem.Fmemo
	vRegUserName = cMain.FOneItem.Fregusername
	vRegdate = cMain.FOneItem.Fregdate
	vLastUserName = cMain.FOneItem.Flastusername
	vLastdate = cMain.FOneItem.Flastdate

	Set cMain = Nothing
Else
	vViewGubun = "period"
	vUseYN = "y"
	vBgGubun = "c"
	vTextInfoUse = "0"
	vShhmmss = "10:00:00"
	vEhhmmss = "09:59:59"
End If

Dim vText1view, vText2view
If vTextInfoUse = "0" Then
	vText1view = "none"
	vText2view = "none"
ElseIf vTextInfoUse = "1" Then
	vText1view = "''"
	vText2view = "none"
ElseIf vTextInfoUse = "2" Then
	vText1view = "''"
	vText2view = "''"
End If
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" lang="ko" xml:lang="ko">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<script language="JavaScript" src="/js/xl.js"></script>
<script language="JavaScript" src="/js/common.js"></script>
<script language="JavaScript" src="/js/report.js"></script>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script language="JavaScript" src="/js/calendar.js"></script>
<script language='javascript' src="/js/jsCal/js/jscal2.js"></script>
<script language='javascript' src="/js/jsCal/js/lang/ko.js"></script>
<script language="JavaScript" src="/cscenter/js/cscenter.js"></script>
<link rel="stylesheet" type="text/css" href="/css/adminDefault.css" />
<link rel="stylesheet" type="text/css" href="/css/adminCommon.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/jscal2.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/border-radius.css" />
<!--[if IE]>
	<link rel="stylesheet" type="text/css" href="/css/adminPartnerIe.css" />
<![endif]-->
<style type="text/css">
.colorbtn {border-width:2px; border-style:solid; border-color:Red;}
</style>
<script language='javascript'>
document.domain = "10x10.co.kr";

function jsMainManageSave(){
	if($("#oneClrUse").is(":checked") == true){
		if($("#bgcolor").val() == ""){
			alert("배경 설정에 단색사용인 경우 컬러를 선택 또는 직접 입력해주세요.");
			return;
		}
	}
	if($("#imgUse").is(":checked") == true){
		if($("#mbgimgurlm").val() == ""){
			alert("배경 설정에 이미지 사용인 경우 이미지를 등록해주세요.");
			return;
		}
	}
	if($("#maskingimgurlm").val() == ""){
		alert("마스킹이미지를 등록해주세요.");
		return;
	}
	if($("#sdate").val() == "" || $("#edate").val() == ""){
		alert("시작일, 종료일을 입력해주세요.");
		return;
	}
	if($("#textinfouse").val() == "1"){
		if($("#textinfo1").val() == "" || $("#textinfo1url").val() == ""){
			alert("검색화면 텍스트 정보를 1개 사용 할 경우\n텍스트 1 과 텍스트 1 URL 모두 입력해 주세요.");
			return;
		}
		
		if(!jsURLchkeck($("#textinfo1url").val())){
			return;
		}
	}else if($("#textinfouse").val() == "2"){
		if($("#textinfo1").val() == "" || $("#textinfo1url").val() == ""){
			alert("검색화면 텍스트 정보를 2개 사용 할 경우\n텍스트 1, 2 와 텍스트 1, 2 URL 모두 입력해 주세요.");
			return;
		}
		if($("#textinfo2").val() == "" || $("#textinfo2url").val() == ""){
			alert("검색화면 텍스트 정보를 2개 사용 할 경우\n텍스트 1, 2 와 텍스트 1, 2 URL 모두 입력해 주세요.");
			return;
		}
		
		if(!jsURLchkeck($("#textinfo1url").val())){
			return;
		}
		
		if(!jsURLchkeck($("#textinfo2url").val())){
			return;
		}
	}

	frm1.submit();
}

function jsImgView(sImgUrl){
 var wImgView;
 wImgView = window.open('/admin/sitemaster/play/lib/pop_event_detailImg.asp?sUrl='+sImgUrl,'pImg','width=100,height=100');
 wImgView.focus();
}

function jsUploadImg(a,b){
	document.domain ="10x10.co.kr";
	var popupl;
	popupl = window.open('/admin/search/pop_uploadimg.asp?folder=main&span='+b+'&sname='+a+'','popupl','width=370,height=150');
	popupl.focus();
}

function jsDelImg(sName, sSpan){
	if(confirm("이미지를 삭제하시겠습니까?")){
	   eval("document.all."+sName).value = "";
	   eval("document.all."+sSpan).style.display = "none";
	}
}

//링크값선택
function showDrop(g){
	$("#selectLink"+g+"").show();
}

function linkcopy(g){
	var val = $("#text"+g+"url").val();
	$("#selectLink"+g+"").css("display","none");
}

//선택입력
function populateTextBox(v,g){
	var val = v;
	$("#text"+g+"url").val(val);
	$("#selectLink"+g+"").css("display","none");
}

function jsViewGubunClear(){
	$("#sdate").val("");
	$("#edate").val("");
}

function jsBgGubun(g){
	if(g == "c"){
		$("#bgcolorselect").show();
		$("#mbgimg").hide();
	}else{
		$("#bgcolorselect").hide();
		$("#mbgimg").show();
	}
}

function jsBGColor(a,v,btn,bgc){
	$("#"+a+" > span > button").removeClass("colorbtn");
	$("#"+btn+"").addClass("colorbtn");
	$("#"+v+"").val(bgc);
}

function jsTextInfo(v){
	if($("#textinfouse").val() == "0"){
		$("#text11").hide();
		$("#text12").hide();
		$("#text21").hide();
		$("#text22").hide();
	}else if($("#textinfouse").val() == "1"){
		$("#text11").show();
		$("#text12").show();
		$("#text21").hide();
		$("#text22").hide();
	}else if($("#textinfouse").val() == "2"){
		$("#text11").show();
		$("#text12").show();
		$("#text21").show();
		$("#text22").show();
	}
}

function jsURLchkeck(u){
	if(u.indexOf("/category/category_list.asp") > -1){
		if(u.indexOf("=") != u.lastIndexOf("=")){
			alert("카테고리리스트는 /category/catogory_list.asp?disp=카테고리코드 까지만 입력하세요.");
			return false;
		}
	}else if(u.indexOf("/search/search_item.asp") > -1){
		if(u.indexOf("=") != u.lastIndexOf("=")){
			alert("검색어는 /search/search_item.asp?rect=검색어 까지만 입력하세요.");
			return false;
		}
	}else if(u.indexOf("/street/") > -1){
		if(u.indexOf("=") != u.lastIndexOf("=")){
			alert("브랜드는 /street/street_brand.asp?makerid=브랜드아이디 까지만 입력하세요.");
			return false;
		}
	}else if(u.indexOf("/event/") > -1){
		if(u.indexOf("=") != u.lastIndexOf("=")){
			alert("이벤트는 /event/eventmain.asp?eventid=이벤트코드 까지만 입력하세요.");
			return false;
		}
	}
	return true;
}
</script>
</head>
<body>
<div id="calendarPopup" style="position: absolute; visibility: hidden; z-index: 2;"></div>
<div class="contSectFix scrl">
	<div class="cont">
	<form name="frm1" id="frm1" action="searchMainManageproc.asp" method="post" style="margin:0px;">
	<input type="hidden" name="idx" value="<%=vIdx%>">
		<div class="searchWrap inputWrap">
			<h3>- 모바일 검색 화면 기본 정보</h3>
			<table class="writeTb tMar10">
				<colgroup>
					<col width="16%" /><col width="" />
				</colgroup>
				<tbody>
				<tr>
					<th><div>배경 설정 *</div></th>
					<td>
						<p>
							<span class="rMar10"><input type="radio" id="oneClrUse" name="bggubun" value="c" <%=CHKIIF(vBgGubun="c","checked","")%> onClick="jsBgGubun('c');" /> <label for="oneClrUse">단색 사용</label></span>
							<span class="rMar10"><input type="radio" id="imgUse" name="bggubun" value="i" <%=CHKIIF(vBgGubun="i","checked","")%> onClick="jsBgGubun('i');" /> <label for="imgUse">이미지 사용</label></span>
						</p>
						<p class="tPad10" id="bgcolorselect" style="display:<%=CHKIIF(vBgGubun="c","block","none")%>">
							<span class="rMar10"><button type="button" class="colorChip <%=CHKIIF(vBgColor="BAD3E0","colorbtn","")%>" id="color1" onClick="jsBGColor('bgcolorselect','bgcolor','color1','BAD3E0');" style="background-color:#BAD3E0"></button></span>
							<span class="rMar10"><button type="button" class="colorChip <%=CHKIIF(vBgColor="84adc2","colorbtn","")%>" id="color2" onClick="jsBGColor('bgcolorselect','bgcolor','color2','84adc2');" style="background-color:#84adc2"></button></span>
							<span class="rMar10"><button type="button" class="colorChip <%=CHKIIF(vBgColor="9c7c6b","colorbtn","")%>" id="color3" onClick="jsBGColor('bgcolorselect','bgcolor','color3','9c7c6b');" style="background-color:#9c7c6b"></button></span>
							<span class="rMar10"><button type="button" class="colorChip <%=CHKIIF(vBgColor="7a88b8","colorbtn","")%>" id="color4" onClick="jsBGColor('bgcolorselect','bgcolor','color4','7a88b8');" style="background-color:#7a88b8"></button></span>
							<span class="rMar10"><button type="button" class="colorChip <%=CHKIIF(vBgColor="eff7fd","colorbtn","")%>" id="color5" onClick="jsBGColor('bgcolorselect','bgcolor','color5','eff7fd');" style="background-color:#eff7fd"></button></span>
							<span class="rMar10"><button type="button" class="colorChip <%=CHKIIF(vBgColor="fceef2","colorbtn","")%>" id="color6" onClick="jsBGColor('bgcolorselect','bgcolor','color6','fceef2');" style="background-color:#fceef2"></button></span>
							<span class="rMar10"><button type="button" class="colorChip <%=CHKIIF(vBgColor="e9f4ed","colorbtn","")%>" id="color7" onClick="jsBGColor('bgcolorselect','bgcolor','color7','e9f4ed');" style="background-color:#e9f4ed"></button></span>
							<span class="rMar10"><button type="button" class="colorChip <%=CHKIIF(vBgColor="fbf3e7","colorbtn","")%>" id="color8" onClick="jsBGColor('bgcolorselect','bgcolor','color8','fbf3e7');" style="background-color:#fbf3e7"></button></span>
							<span>#<input type="text" class="formTxt vTop" id="bgcolor" name="bgcolor" value="<%=vBgColor%>" style="width:10%" maxlength="6" /></span>
						</p>
						<p class="tPad10" id="mbgimg" style="display:<%=CHKIIF(vBgGubun="i","block","none")%>">
							<input type="button" value="이미지업로드" onClick="jsUploadImg('mbgimgurlm','mbgimgurlmspan');" /><br /><br />
							<span id="mbgimgurlmspan" style="padding:5px 5px 5px 0;"><%
								If vBgImg <> "" Then
									Response.Write "<img src='" & vBgImg & "' height='100' style='cursor:pointer;' onclick=jsImgView('" & vBgImg & "');>"
									Response.Write "<a href=javascript:jsDelImg('mbgimgurlm','mbgimgurlmspan');><img src='/images/icon_delete2.gif' border='0'></a>"
								End If
							%></span>
							<input type="hidden" id="mbgimgurlm" name="mbgimgurlm" value="<%=vBgImg%>">
							<br /><span class="tPad10 fs11 cBl3">* 2Mb 이하의(1024x200사이즈) png, jpg, gif등의 이미지파일을 선택해주세요.</span>
						</p>
					</td>
				</tr>
				<tr>
					<th><div>마스킹이미지 *</div></th>
					<td>
						<p class="tPad10">
							<input type="button" value="이미지업로드" onClick="jsUploadImg('maskingimgurlm','maskingimgurlmspan');" /><br /><br />
							<span id="maskingimgurlmspan" style="padding:5px 5px 5px 0;"><%
								If vMaskingImg <> "" Then
									Response.Write "<img src='" & vMaskingImg & "' height='100' style='cursor:pointer;' onclick=jsImgView('" & vMaskingImg & "');>"
									Response.Write "<a href=javascript:jsDelImg('maskingimgurlm','maskingimgurlmspan');><img src='/images/icon_delete2.gif' border='0'></a>"
								End If
							%></span>
							<input type="hidden" id="maskingimgurlm" name="maskingimgurlm" value="<%=vMaskingImg%>">
							<br /><span class="tPad10 fs11 cBl3">* 2Mb 이하의(1024x200사이즈) png, jpg, gif등의 이미지파일을 선택해주세요.</span>
						</p>
					</td>
				</tr>
				<tr>
					<th><div>노출 기간 *</div></th>
					<td>
						<span><input type="hidden" id="termSet" name="viewgubun" value="<%=vViewGubun%>" /></span>
						<span>
							<input type="text" class="formTxt" id="sdate" name="sdate" value="<%=vSDate%>" style="width:100px" placeholder="시작일" maxlength="10" readonly />
							<img src="/images/admin_calendar.png" id="sdate_trigger" alt="달력으로 검색" />
							<script language="javascript">
								var CAL_Start = new Calendar({
									inputField : "sdate", trigger    : "sdate_trigger",
									onSelect: function() {
										var date = Calendar.intToDate(this.selection.get());
										CAL_End.args.min = date;
										CAL_End.redraw();
										this.hide();
									}, bottomBar: true, dateFormat: "%Y-%m-%d"
								});
							</script>
							<input type="text" class="formTxt" id="shhmmss" name="shhmmss" value="<%=vShhmmss%>" style="width:60px" maxlength="8" readonly />
							~
							<input type="text" class="formTxt" id="edate" name="edate" value="<%=vEDate%>" style="width:100px" placeholder="종료일" maxlength="10" readonly />
							<img src="/images/admin_calendar.png" id="edate_trigger" alt="달력으로 검색" />
							<script language="javascript">
								var CAL_End = new Calendar({
									inputField : "edate", trigger    : "edate_trigger",
									onSelect: function() {
										var date = Calendar.intToDate(this.selection.get());
										CAL_Start.args.max = date;
										CAL_Start.redraw();
										this.hide();
									}, bottomBar: true, dateFormat: "%Y-%m-%d"
								});
							</script>
							<input type="text" class="formTxt" id="ehhmmss" name="ehhmmss" value="<%=vEhhmmss%>" style="width:60px" maxlength="8" readonly />
						</span>
					</td>
				</tr>
				<tr>
					<th><div>사용 여부 *</div></th>
					<td>
						<span class="rMar10"><input type="radio" id="useyny" name="useyn" value="y" <%=CHKIIF(vUseYN="y","checked","")%> /> <label for="useyny">사용함</label></span>
						<span class="rMar10"><input type="radio" id="useynn" name="useyn" value="n" <%=CHKIIF(vUseYN="n","checked","")%> /> <label for="useynn">사용안함</label></span>
					</td>
				</tr>
				<tr>
					<th><div>비고</div></th>
					<td><textarea class="formTxtA" rows="6" style="width:99%;" id="memo" name="memo"><%=vMemo%></textarea></td>
				</tr>
				</tbody>
			</table>
		</div>
		<div class="pad20">
			<h3>- 검색화면 텍스트 정보</h3>
			<table class="tbType1 writeTb tMar10">
				<colgroup>
					<col width="16%" /><col width="" />
				</colgroup>
				<tbody>
				<tr>
					<th><div>사용여부 *</div></th>
					<td>
						<select class="formSlt" title="옵션 선택" id="textinfouse" name="textinfouse" onChange="jsTextInfo(this.value);">
							<option value="0" <%=CHKIIF(vTextInfoUse="0","selected","")%>>사용안함</option>
							<option value="1" <%=CHKIIF(vTextInfoUse="1","selected","")%>>1개 사용</option>
							<option value="2" <%=CHKIIF(vTextInfoUse="2","selected","")%>>2개 사용</option>
						</select>
					</td>
				</tr>
				<tr id="text11" style="display:<%=vText1view%>;">
					<th><div>텍스트 1</div></th>
					<td><input type="text" class="formTxt" id="textinfo1" name="textinfo1" value="<%=vTextInfo1%>" placeholder="텍스트 정보를 10자 이내로 입력해주세요" style="width:99%" maxlength="10" /></td>
				</tr>
				<tr id="text12" style="display:<%=vText1view%>;">
					<th><div>텍스트 1 URL</div></th>
					<td>
						<div class="selectLink">
							<input type="text" class="formTxt" value="<%=CHKIIF(vTextInfo1url="","",vTextInfo1url)%>" placeholder="링크값 입력(선택)" onclick="showDrop('info1');" id="textinfo1url" name="textinfo1url" onkeyup="linkcopy('info1');" maxlength="200" />
							<ul style="display:none;" id="selectLinkinfo1">
								<li onclick="populateTextBox('<%=CHKIIF(vTextInfo1url="","",vTextInfo1url)%>','info1');">선택안함</li>
								<li onclick="populateTextBox('/category/category_itemPrd.asp?itemid=상품코드','info1');">/category/category_itemPrd.asp?itemid=상품코드</li>
								<li onclick="populateTextBox('/category/category_list.asp?disp=카테고리','info1');">/category/category_list.asp?disp=카테고리</li>
								<li onclick="populateTextBox('/street/street_brand.asp?makerid=브랜드아이디','info1');">/street/street_brand.asp?makerid=브랜드아이디</li>
								<li onclick="populateTextBox('/search/search_item.asp?rect=검색어','info1');">/search/search_item.asp?rect=검색어</li>
								<li onclick="populateTextBox('/event/eventmain.asp?eventid=이벤트코드','info1');">/event/eventmain.asp?eventid=이벤트코드</li>
								<li onclick="populateTextBox('/culturestation/culturestation_event.asp?evt_code=컬처스테이션이벤트코드','info1');">/culturestation/culturestation_event.asp?evt_code=컬처스테이션이벤트코드</li>
								<li onclick="populateTextBox('/gift/gifttalk/','info1');">기프트</li>
								<li onclick="populateTextBox('/wish/index.asp','info1');">위시</li>
							</ul>
						</div>
					</td>
				</tr>
				<tr id="text21" style="display:<%=vText2view%>;">
					<th><div>텍스트 2</div></th>
					<td><input type="text" class="formTxt" id="textinfo2" name="textinfo2" value="<%=vTextInfo2%>" placeholder="텍스트 정보를 10자 이내로 입력해주세요" style="width:99%" maxlength="110" /></td>
				</tr>
				<tr id="text22" style="display:<%=vText2view%>;">
					<th><div>텍스트 2 URL</div></th>
					<td>
						<div class="selectLink">
							<input type="text" class="formTxt" value="<%=CHKIIF(vTextInfo2url="","",vTextInfo2url)%>" placeholder="링크값 입력(선택)" onclick="showDrop('info2');" id="textinfo2url" name="textinfo2url" onkeyup="linkcopy('info2');" maxlength="200" />
							<ul style="display:none;" id="selectLinkinfo2">
								<li onclick="populateTextBox('<%=CHKIIF(vTextInfo2url="","",vTextInfo2url)%>','info2');">선택안함</li>
								<li onclick="populateTextBox('/category/category_itemPrd.asp?itemid=상품코드','info2');">/category/category_itemPrd.asp?itemid=상품코드</li>
								<li onclick="populateTextBox('/category/category_list.asp?disp=카테고리','info2');">/category/category_list.asp?disp=카테고리</li>
								<li onclick="populateTextBox('/street/street_brand.asp?makerid=브랜드아이디','info2');">/street/street_brand.asp?makerid=브랜드아이디</li>
								<li onclick="populateTextBox('/search/search_item.asp?rect=검색어','info2');">/search/search_item.asp?rect=검색어</li>
								<li onclick="populateTextBox('/event/eventmain.asp?eventid=이벤트코드','info2');">/event/eventmain.asp?eventid=이벤트코드</li>
								<li onclick="populateTextBox('/culturestation/culturestation_event.asp?evt_code=컬처스테이션이벤트코드','info2');">/culturestation/culturestation_event.asp?evt_code=컬처스테이션이벤트코드</li>
								<li onclick="populateTextBox('/gift/gifttalk/','info2');">기프트</li>
								<li onclick="populateTextBox('/wish/index.asp','info2');">위시</li>
							</ul>
						</div>
					</td>
				</tr>
				</tbody>
			</table>
		</div>
		<br /><br /><br /><br /><br /><br /><br />
		<div class="pad20">
			<div class="ct">
				<input type="button" value="저장" onclick="jsMainManageSave();" class="cRd1" style="width:100px; height:30px;" />
				<input type="button" value="취소" onclick="window.close();" style="width:100px; height:30px;" />
			</div>
		</div>
	</form>
	</div>
</div>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/db3close.asp" -->