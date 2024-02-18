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
<!-- #include virtual="/lib/classes/sitemaster/brand_banner_manageCls.asp"-->
<%
dim menupos, imenuposStr, imenuposnotice, imenuposhelp
menupos = request("menupos")
if menupos ="" then menupos=1

Dim i, cQuick, vIdx, vName, vURL_PC, vURL_M, vRegUserName
Dim vSDate, vEDate, vRegdate, vLastUserName, vLastdate, vUseYN
Dim vQImgPC, vQImgM, vBtnColor, vShhmmss, vEhhmmss
vIdx = requestCheckVar(Request("idx"),15)

If vIdx <> "" Then
	Set cQuick = New CSearchMng
	cQuick.FRectIdx = vIdx
	cQuick.sbQuickLinkDetail
	vName = cQuick.FOneItem.Fquickname
	vURL_PC = cQuick.FOneItem.Furl_pc
	vURL_M = cQuick.FOneItem.Furl_m
	vSDate = cQuick.FOneItem.Fsdate
	vEDate = cQuick.FOneItem.Fedate
	vShhmmss = cQuick.FOneItem.Fshhmmss
	vEhhmmss = cQuick.FOneItem.Fehhmmss
	vRegUserName = cQuick.FOneItem.Fregusername
	vRegdate = cQuick.FOneItem.Fregdate
	vLastUserName = cQuick.FOneItem.Flastusername
	vLastdate = cQuick.FOneItem.Flastdate
	vUseYN = cQuick.FOneItem.Fuseyn
	vQImgPC = cQuick.FOneItem.Fqimgpc
	vQImgM = cQuick.FOneItem.Fqimgm
	Set cQuick = Nothing
Else
	vUseYN = "y"
	vShhmmss = "10:00:00"
	vEhhmmss = "09:59:59"
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

function goTypeChange(a){
	location.href = "quickLinkManage.asp?idx=<%=vIdx%>&type="+a+"";
}

function jsQuickLinkSave(){

	if($("#quickname").val() == ""){
		alert("퀵링크 명을 입력하세요.");
		return;
	}
	if($("#url_pc").val() == ""){
		alert("URL PC를 입력하세요.");
		return;
	}
	if($("#url_m").val() == ""){
		alert("URL M를 입력하세요.");
		return;
	}
	if($("#sdate").val() == "" || $("#edate").val() == ""){
		alert("시작일, 종료일을 입력해주세요.");
		return;
	}
	frm1.submit();
}

//브랜드 ID 검색 팝업창
function jsSearchBrandID1(frmName,compName,compName2){
    var compVal = "";
    try{
        compVal = eval("document.all." + frmName + "." + compName).value;
    }catch(e){
        compVal = "";
    }

    var popwin = window.open("popBrandSearch_search.asp?isjsdomain=o&frmName=" + frmName + "&compName=" + compName + "&compName2=" + compName2 + "&socname_kr=" + compVal,"popBrandSearch","width=800 height=400 scrollbars=yes resizable=yes");

	popwin.focus();
}

function jsImgView(sImgUrl){
	var wImgView;
	wImgView = window.open('/admin/sitemaster/play/lib/pop_event_detailImg.asp?sUrl='+sImgUrl,'pImg','width=100,height=100');
	wImgView.focus();
}

function jsUploadImg(a,b){
	document.domain ="10x10.co.kr";
	var popupl;
	popupl = window.open('/admin/search/pop_uploadimg.asp?folder=quick&span='+b+'&sname='+a+'','popupl','width=370,height=150');
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
	var val = $("#url_"+g+"").val();
	$("#selectLink"+g+"").css("display","none");
}
function showDropL(g){
	$("#selectbtnLink"+g+"").show();
}
function linkcopyL(g){
	var val = $("#url_"+g+"").val();
	$("#selectbtnLink"+g+"").css("display","none");
}
//선택입력
function populateTextBox(v,g){
	var val = v;
	$("#url_"+g+"").val(val);
	$("#selectLink"+g+"").css("display","none");
}
function populateTextBoxL(v,g){
	var val = v;
	$("#btnlink"+g+"").val(val);
	$("#selectbtnLink"+g+"").css("display","none");
}

function jsViewGubunClear(){
	$("#sdate").val("");
	$("#edate").val("");
}

function jsBgGubun(g){
	if(g == "c"){
		$("#bgcolorselect").show();
		$("#qbgimg").hide();
	}else{
		$("#bgcolorselect").hide();
		$("#qbgimg").show();
	}
}

function jsBGColor(a,v,btn,bgc){
	$("#"+a+" > span > button").removeClass("colorbtn");
	$("#"+btn+"").addClass("colorbtn");
	$("#"+v+"").val(bgc);
}

function jsQimgGubun(g){
	if(g == "y"){
		$("#qimg").show();
	}else{
		$("#qimg").hide();
	}
}
</script>
</head>
<body>
<div id="calendarPopup" style="position: absolute; visibility: hidden; z-index: 2;"></div>
<div class="contSectFix scrl">
	<div class="cont">
	<form name="frm1" id="frm1" action="quickLinkproc.asp" method="post" style="margin:0px;">
	
	<input type="hidden" name="idx" value="<%=vIdx%>">
		<div class="searchWrap inputWrap">
			<h3>- 퀵링크 기본 정보</h3>
			<table class="writeTb tMar10">
				<colgroup>
					<col width="16%" /><col width="" />
				</colgroup>
				<tbody>
				<tr>
					<th><div>퀵링크 명 *</div></th>
					<td><input type="text" class="formTxt" value="<%=vName%>" id="quickname" name="quickname" placeholder="퀵링크명을 입력해주세요." style="width:45%" maxlength="20" />
					</td>
				</tr>
				<tr>
					<th><div>URL PC *</div></th>
					<td>
						<div class="selectLink">
							<input type="text" class="formTxt" value="<%=CHKIIF(vURL_PC="","",vURL_PC)%>" placeholder="링크값 입력(선택)" onclick="showDrop('pc');" id="url_pc" name="url_pc" onkeyup="linkcopy('pc');" maxlength="200" />
							<ul style="display:none;" id="selectLinkpc">
								<li onclick="populateTextBox('<%=CHKIIF(vURL_PC="","",vURL_PC)%>','pc');">선택안함</li>
								<li onclick="populateTextBox('/category/category_prd.asp?itemid=상품코드','pc');">/category/category_prd.asp?itemid=상품코드</li>
								<li onclick="populateTextBox('/shopping/category_list.asp?disp=카테고리','pc');">/shopping/category_list.asp?disp=카테고리</li>
								<li onclick="populateTextBox('/street/street_brand.asp?makerid=브랜드아이디','pc');">/street/street_brand.asp?makerid=브랜드아이디</li>
								<li onclick="populateTextBox('/event/eventmain.asp?eventid=이벤트코드','pc');">/event/eventmain.asp?eventid=이벤트코드</li>
								<li onclick="populateTextBox('/culturestation/culturestation_event.asp?evt_code=컬처스테이션이벤트코드','pc');">/culturestation/culturestation_event.asp?evt_code=컬처스테이션이벤트코드</li>
								<li onclick="populateTextBox('/gift/gifttalk/','pc');">기프트</li>
								<li onclick="populateTextBox('/wish/index.asp','pc');">위시</li>
							</ul>
						</div>
					</td>
				</tr>
				<tr>
					<th><div>URL M *</div></th>
					<td>
						<div class="selectLink">
							<input type="text" class="formTxt" value="<%=CHKIIF(vURL_M="","",vURL_M)%>" placeholder="링크값 입력(선택)" onclick="showDrop('m');" id="url_m" name="url_m" onkeyup="linkcopy('m');" maxlength="200" />
							<ul style="display:none;" id="selectLinkm">
								<li onclick="populateTextBox('<%=CHKIIF(vURL_M="","",vURL_M)%>','m');">선택안함</li>
								<li onclick="populateTextBox('/category/category_itemPrd.asp?itemid=상품코드','m');">/category/category_itemPrd.asp?itemid=상품코드</li>
								<li onclick="populateTextBox('/category/category_list.asp?disp=카테고리','m');">/category/category_list.asp?disp=카테고리</li>
								<li onclick="populateTextBox('/street/street_brand.asp?makerid=브랜드아이디','m');">/street/street_brand.asp?makerid=브랜드아이디</li>
								<li onclick="populateTextBox('/event/eventmain.asp?eventid=이벤트코드','m');">/event/eventmain.asp?eventid=이벤트코드</li>
								<li onclick="populateTextBox('/culturestation/culturestation_event.asp?evt_code=컬처스테이션이벤트코드','m');">/culturestation/culturestation_event.asp?evt_code=컬처스테이션이벤트코드</li>
								<li onclick="populateTextBox('/gift/gifttalk/','m');">기프트</li>
								<li onclick="populateTextBox('/wish/index.asp','m');">위시</li>
							</ul>
						</div>
					</td>
				</tr>
				<tr>
					<th><div>노출 기간 *</div></th>
					<td>
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
					<th><div>퀵링크 이미지 *</div></th>
					<td>
						<p class="tPad10" id="qimg">
							<input type="button" value="PC 업로드" onClick="jsUploadImg('qimgurlpc','qimgurlpcspan');" /><br /><br />
							<span id="qimgurlpcspan" style="padding:5px 5px 5px 0;"><%
								If vQImgPC <> "" Then
									Response.Write "<img src='" & vQImgPC & "' height='100' style='cursor:pointer;' onclick=jsImgView('" & vQImgPC & "');>"
									Response.Write "<a href=javascript:jsDelImg('qimgurlpc','qimgurlpcspan');><img src='/images/icon_delete2.gif' border='0'></a>"
								End If
							%></span>
							<input type="hidden" id="qimgurlpc" name="qimgurlpc" value="<%=vQImgPC%>">
							<br /><br />
							<input type="button" value="Mobile 업로드" onClick="jsUploadImg('qimgurlm','qimgurlmspan');" /><br /><br />
							<span id="qimgurlmspan" style="padding:5px 5px 5px 0;"><%
								If vQImgM <> "" Then
									Response.Write "<img src='" & vQImgM & "' height='100' style='cursor:pointer;' onclick=jsImgView('" & vQImgM & "');>"
									Response.Write "<a href=javascript:jsDelImg('qimgurlm','qimgurlmspan');><img src='/images/icon_delete2.gif' border='0'></a>"
								End If
							%></span>
							<input type="hidden" id="qimgurlm" name="qimgurlm" value="<%=vQImgM%>">
							<br /><span class="tPad10 fs11 cBl3">* 2Mb 이하의(1024x200사이즈) png, jpg, gif등의 이미지파일을 선택해주세요.</span>
						</p>
					</td>
				</tr>
				</tbody>
			</table>
		</div>
		<div class="pad20">
			<div class="ct">
				<input type="button" value="저장" onclick="jsQuickLinkSave();" class="cRd1" style="width:100px; height:30px;" />
				<input type="button" value="취소" onclick="window.close();" style="width:100px; height:30px;" />
			</div>
		</div>
	</form>
	</div>
</div>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/db3close.asp" -->