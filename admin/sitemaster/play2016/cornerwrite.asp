<%@ language=vbscript %>
<% option explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/play/play2016Cls.asp" -->
<%
	Dim i, p, l, tmp, cPl, vPlayImgList, vPlayAzitList, vVolNum, vTitleStyle
	Dim vMIdx, vDidx, vCate, vOpenDate, vState, vTitle, vSubCopy, vWorkText, vJikListImg, vJungListImg, vPartWDID, vPartMKID, vPartPBID, vMoBGColor, vPCBGColor
	Dim vIsTagView, vTagSDate, vTagEDate, vTagAnnounceDate, vKeyword, vSearchListImg
	Dim vCate1VideoURL, vCate1ImageURL, vCate1Type, vCate1Directer, vCate1PCLinkBanImg, vCate1PCLinkBanURL, vCate1MoLinkBanImg, vCate1MoLinkBanURL
	Dim vCate1CommTitle, vCate1Comment1, vCate1Comment2, vCate1Comment3, vCate1precomm1, vCate1precomm2, vCate1precomm3, vCate1VideoOrigin, vCate1RewardCopy
	Dim vCate21ImageURL(5), vCate21ImageIdx(5), vCate21Item, vCate22ImageURL(5), vCate22ImageIdx(5), vCate22Item
	Dim vCate3PCImageURL, vCate3MoImageURL, vCate3Icon, vCate3Notice, vCate3Ptitle(4), vCate3Pjuso(4), vCate3Plink(4), vCate3PImg(20), vCate3PCopy(20)
	Dim vCate3EntryCont, vCate3EntrySDate, vCate3EntryEDate, vCate3AnnounDate, vCate3EntryMethod
	Dim vCate41PCIsExec, vCate41PCExecFile, vCate41MoIsExec, vCate41MoExecFile, vCate41PCContent, vCate41MoContent
	Dim vCate42Img(3), vCate42EntrySDate, vCate42EntryEDate, vCate42AnnounDate, vCate42WinnerTxt, vCate42WinnerValue, vCate42Value(10), vCate42Item
	Dim vCate42WinList, vCate42Notice, vCate42EntryCopy, vCate42PCLinkBanImg, vCate42PCLinkBanURL, vCate42MoLinkBanImg, vCate42MoLinkBanURL
	Dim vCate43PCImageURL, vCate43MoImageURL, vCate43QRImageURL, vCate43PCDownList, vCate43MoDownList, vCate43PCDown(3), vCate43PCLink(3), vCate43MoDown(3), vCate43MoLink(3)
	Dim vCate5PCTopImageURL, vCate5MoTopImageURL, vCate5Directer, vCate5Img(5), vCate5Copy(5), vCate5PCLinkBanImg, vCate5PCLinkBanURL, vCate5MoLinkBanImg, vCate5MoLinkBanURL
	Dim vCate6VideoURL, vCate6BannSub, vCate6BannTitle, vCate6BannBtnTitle, vCate6BannBtnLink, vCate6Img(4), vCate6Copy(4)
	'// 2017.06.01 원승현 azit comma 스타일 추가
	Dim vCate31PCTopImageURL, vCate31MoTopImageURL, vCate31Directer, vCate31Img(5), vCate31Copy(5), vCate31PCLinkBanImg, vCate31PCLinkBanURL, vCate31MoLinkBanImg, vCate31MoLinkBanURL
	Dim vCate42Badgetag

	vVolNum = requestCheckVar(Request("volnum"),10)
	vMIdx = requestCheckVar(Request("midx"),10)
	vDidx = requestCheckVar(Request("didx"),10)
	vCate = requestCheckVar(Request("cate"),10)
	vCate41PCIsExec = True
	vCate41MoIsExec = True
	vCate3EntryMethod = "c"
	
	If vDidx <> "" Then
		Set cPl = New CPlay
		cPl.FRectDIdx = vDidx
		cPl.FRectCate = vCate
		cPl.sbPlayCornerDetail

		vOpenDate			= cPl.FOneItem.Fstartdate
		vState				= cPl.FOneItem.Fstate
		vTitle				= cPl.FOneItem.Ftitle
		vTitleStyle		= cPl.FOneItem.Ftitlestyle
		vSubCopy			= cPl.FOneItem.Fsubcopy
		vMoBGColor		= cPl.FOneItem.Fmobgcolor
		vWorkText			= cPl.FOneItem.Fworktext
		vPartWDID			= cPl.FOneItem.FpartWDID
		vPartMKID			= cPl.FOneItem.FpartMKID
		vPartPBID			= cPl.FOneItem.FpartPBID
		vIsTagView		= cPl.FOneItem.Fistagview
		vTagSDate			= cPl.FOneItem.Ftagsdate
		vTagEDate			= cPl.FOneItem.Ftagedate
		vTagAnnounceDate	= cPl.FOneItem.Ftagannouncedate
		vKeyword			= cPl.FOneItem.Fkeyword
		vPlayImgList		= cPl.FPlayImageList
		
		'### fnPlayImageSelect(array, cate, 이미지구분, 가져올항목)
		vJikListImg		= fnPlayImageSelect(vPlayImgList,vCate,"1","i")
		vJungListImg		= fnPlayImageSelect(vPlayImgList,vCate,"11","i")
		vSearchListImg	= fnPlayImageSelect(vPlayImgList,vCate,"28","i")

		If vCate = "1" Then '### playlist
			vCate1VideoURL		= cPl.FOneItem.FCate1VideoURL
			vCate1Type				= cPl.FOneItem.FCate1Type
			vCate1Directer		= cPl.FOneItem.FCate1Directer
			vCate1CommTitle		= cPl.FOneItem.FCate1CommTitle
			vCate1Comment1		= cPl.FOneItem.FCate1Comment1
			vCate1Comment2		= cPl.FOneItem.FCate1Comment2
			vCate1Comment3		= cPl.FOneItem.FCate1Comment3
			vCate1precomm1		= cPl.FOneItem.FCate1precomm1
			vCate1precomm2		= cPl.FOneItem.FCate1precomm2
			vCate1precomm3		= cPl.FOneItem.FCate1precomm3
			vCate1VideoOrigin		= cPl.FOneItem.FCate1VideoOrigin
			vCate1RewardCopy		= cPl.FOneItem.FCate1RewardCopy
			
			'### fnPlayImageSelect(array, cate, 이미지구분, 가져올항목)
			vCate1ImageURL			= fnPlayImageSelect(vPlayImgList,vCate,"2","i")
			vCate1PCLinkBanImg		= fnPlayImageSelect(vPlayImgList,vCate,"3","i")
			vCate1PCLinkBanURL		= fnPlayImageSelect(vPlayImgList,vCate,"3","l")
			vCate1MoLinkBanImg		= fnPlayImageSelect(vPlayImgList,vCate,"18","i")
			vCate1MoLinkBanURL		= fnPlayImageSelect(vPlayImgList,vCate,"18","l")
			
		ElseIf vCate = "21" Then '### inspiration design
			vCate21Item			= fnPlayItemList(vDidx)
			For l=1 To 5
				vCate21ImageURL(l) = fnPlayImageSelectSortNo(vPlayImgList,vCate,"4","i","0",l)
				vCate21ImageIdx(l) = fnPlayImageSelectSortNo(vPlayImgList,vCate,"4","x","0",l)
			Next
			
		ElseIf vCate = "22" Then '### inspiration style
			vCate22Item			= fnPlayItemList(vDidx)
			For l=1 To 5
				vCate22ImageURL(l) = fnPlayImageSelectSortNo(vPlayImgList,vCate,"5","i","0",l)
				vCate22ImageIdx(l) = fnPlayImageSelectSortNo(vPlayImgList,vCate,"5","x","0",l)
			Next
			
		ElseIf vCate = "3" Then '### azit
			tmp = 0
			vCate3Icon				= cPl.FOneItem.FCate3iconimg
			vCate3PCImageURL		= fnPlayImageSelect(vPlayImgList,vCate,"19","i")
			vCate3MoImageURL		= fnPlayImageSelect(vPlayImgList,vCate,"6","i")
			vCate3EntryCont		= cPl.FOneItem.FCate3EntryCont
			vCate3EntrySDate		= cPl.FOneItem.FCate3EntrySDate
			vCate3EntryEDate		= cPl.FOneItem.FCate3EntryEDate
			vCate3AnnounDate		= cPl.FOneItem.FCate3AnnounDate
			vCate3Notice			= cPl.FOneItem.FCate3Notice
			vCate3EntryMethod		= cPl.FOneItem.FCate3EntryMethod
			vPlayAzitList			= cPl.FPlayAzipList
			
			For p=1 To 4
				vCate3Ptitle(p)	= fnPlayAzitSelect(vPlayAzitList,p,"1")
				vCate3Pjuso(p)	= fnPlayAzitSelect(vPlayAzitList,p,"2")
				vCate3Plink(p)	= fnPlayAzitSelect(vPlayAzitList,p,"3")
				
				For l=1 To 5
					tmp = tmp + 1
					vCate3PImg(tmp)	= fnPlayImageSelectSortNo(vPlayImgList,vCate,"7","i",p,l)
					vCate3PCopy(tmp)	= fnPlayImageSelectSortNo(vPlayImgList,vCate,"7","c",p,l)
				Next
			Next

		'// 2017.06.01 원승현 azit comma 스타일 추가
		ElseIf vCate = "31" Then
			tmp = 0
			vCate31PCTopImageURL		= fnPlayImageSelect(vPlayImgList,vCate,"23","i")
			vCate31MoTopImageURL		= fnPlayImageSelect(vPlayImgList,vCate,"24","i")
			vCate31Directer			= cPl.FOneItem.FCate31Directer
			vCate31PCLinkBanImg		= fnPlayImageSelect(vPlayImgList,vCate,"26","i")
			vCate31PCLinkBanURL		= fnPlayImageSelect(vPlayImgList,vCate,"26","l")
			vCate31MoLinkBanImg		= fnPlayImageSelect(vPlayImgList,vCate,"27","i")
			vCate31MoLinkBanURL		= fnPlayImageSelect(vPlayImgList,vCate,"27","l")
			
			For l=1 To 5
				tmp = tmp + 1
				vCate31Img(tmp)	= fnPlayImageSelectSortNo(vPlayImgList,vCate,"25","i","0",l)
				vCate31Copy(tmp)	= fnPlayImageSelectSortNo(vPlayImgList,vCate,"25","c","0",l)
			Next
			
		ElseIf vCate = "41" Then '### THING thing
			vCate41PCIsExec	= cPl.FOneItem.FCate41PCIsExec
			vCate41PCExecFile	= cPl.FOneItem.FCate41PCExecFile
			vCate41MoIsExec	= cPl.FOneItem.FCate41MoIsExec
			vCate41MoExecFile	= cPl.FOneItem.FCate41MoExecFile
			vCate41PCContent	= cPl.FOneItem.FCate41PCContent
			vCate41MoContent	= cPl.FOneItem.FCate41MoContent
			
		ElseIf vCate = "42" Then '### THING thingthing
			tmp = 0
			vCate42Badgetag			= cPl.FOneItem.FCate42Badgetag
			vCate42EntrySDate		= cPl.FOneItem.FCate42EntrySDate
			vCate42EntryEDate		= cPl.FOneItem.FCate42EntryEDate
			vCate42AnnounDate		= cPl.FOneItem.FCate42AnnounDate
			vCate42WinnerTxt		= cPl.FOneItem.FCate42WinnerTxt
			vCate42WinnerValue		= cPl.FOneItem.FCate42WinnerValue
			vCate42Notice			= cPl.FOneItem.FCate42Notice
			vCate42EntryCopy		= cPl.FOneItem.FCate42Entrycopy
			vCate42WinList		= cPl.FPlayThingThingWinlist
			vCate42PCLinkBanImg	= fnPlayImageSelect(vPlayImgList,vCate,"21","i")
			vCate42PCLinkBanURL	= fnPlayImageSelect(vPlayImgList,vCate,"21","l")
			vCate42MoLinkBanImg	= fnPlayImageSelect(vPlayImgList,vCate,"22","i")
			vCate42MoLinkBanURL	= fnPlayImageSelect(vPlayImgList,vCate,"22","l")
			vCate42Item			= fnPlayItemList(vDidx)
			
			For l=1 To 3
				tmp = tmp + 1
				vCate42Img(tmp)	= fnPlayImageSelectSortNo(vPlayImgList,vCate,"8","i","0",l)
			Next
			
			IF isArray(vCate42WinList) THEN
				For l =0 To UBound(vCate42WinList,2)
					vCate42Value(l+1) = vCate42WinList(0,l)
				Next
			End If
			
		ElseIf vCate = "43" Then '### THING 배경화면
			tmp = 0
			vCate43PCImageURL		= fnPlayImageSelect(vPlayImgList,vCate,"20","i")
			vCate43MoImageURL		= fnPlayImageSelect(vPlayImgList,vCate,"9","i")
			vCate43QRImageURL		= fnPlayImageSelect(vPlayImgList,vCate,"10","i")
			vCate43PCDownList		= cPl.FPlayThingPCDownList
			vCate43MoDownList		= cPl.FPlayThingMoDownList

			IF isArray(vCate43PCDownList) THEN
				For l =0 To UBound(vCate43PCDownList,2)
					vCate43PCDown(l+1) = fnPlayDownloadSelect(vCate43PCDownList,"0",l)
					vCate43PCLink(l+1) = fnPlayDownloadSelect(vCate43PCDownList,"1",l)
				Next
			End If
			
			IF isArray(vCate43MoDownList) THEN
				For l =0 To UBound(vCate43MoDownList,2)
					vCate43MoDown(l+1) = fnPlayDownloadSelect(vCate43MoDownList,"0",l)
					vCate43MoLink(l+1) = fnPlayDownloadSelect(vCate43MoDownList,"1",l)
				Next
			End If
			
		ElseIf vCate = "5" Then '### COMMA
			tmp = 0
			vCate5PCTopImageURL		= fnPlayImageSelect(vPlayImgList,vCate,"12","i")
			vCate5MoTopImageURL		= fnPlayImageSelect(vPlayImgList,vCate,"13","i")
			vCate5Directer			= cPl.FOneItem.FCate5Directer
			vCate5PCLinkBanImg		= fnPlayImageSelect(vPlayImgList,vCate,"15","i")
			vCate5PCLinkBanURL		= fnPlayImageSelect(vPlayImgList,vCate,"15","l")
			vCate5MoLinkBanImg		= fnPlayImageSelect(vPlayImgList,vCate,"16","i")
			vCate5MoLinkBanURL		= fnPlayImageSelect(vPlayImgList,vCate,"16","l")
			
			For l=1 To 5
				tmp = tmp + 1
				vCate5Img(tmp)	= fnPlayImageSelectSortNo(vPlayImgList,vCate,"14","i","0",l)
				vCate5Copy(tmp)	= fnPlayImageSelectSortNo(vPlayImgList,vCate,"14","c","0",l)
			Next
			
		ElseIf vCate = "6" Then '### HOWHOW
			tmp = 0
			vCate6VideoURL		= cPl.FOneItem.FCate6VideoURL
			vCate6BannSub			= cPl.FOneItem.FCate6BannSub
			vCate6BannTitle		= cPl.FOneItem.FCate6BannTitle
			vCate6BannBtnTitle	= cPl.FOneItem.FCate6BannBtnTitle
			vCate6BannBtnLink		= cPl.FOneItem.FCate6BannBtnLink
			
			For l=1 To 4
				tmp = tmp + 1
				vCate6Img(tmp)	= fnPlayImageSelectSortNo(vPlayImgList,vCate,"17","i","0",l)
				vCate6Copy(tmp)	= fnPlayImageSelectSortNo(vPlayImgList,vCate,"17","c","0",l)
			Next

		End If
		SET cPl = Nothing
	End If
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" lang="ko" xml:lang="ko">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<style type="text/css">
body {font:9pt/135% "dotum";color:#000000}
.tbType1 {width:100%;}
.tbType1 th, .tbType1 td {color:#444;}
.tbType1 th {background-color:#eaeaea;}
.tbType1 th a, .tbType1 td a {color:#444;}
.tbType1 th a:hover, .tbType1 td a:hover {text-decoration:underline;}


.writeTb {border-top:2px solid #b9b9b9; border-bottom:2px solid #b9b9b9;}
.writeTb th, .writeTb td {border-bottom:1px solid #c9c9c9; vertical-align:middle;}
.writeTb th {font-weight:bold; text-align:center;}
.writeTb th div {padding:9px 10px 7px 10px; vertical-align:middle;}


.btn2 {display:inline-block; font-size:11px !important; letter-spacing:-0.025em; line-height:110%; border-left:1px solid #f0f0f0; border-top:1px solid #f0f0f0; border-right:1px solid #cdcdcd; border-bottom:1px solid #cdcdcd; background-color:#f2f2f2; background-image:-webkit-linear-gradient(#fff, #e1e1e1); background-image:-moz-linear-gradient(#fff, #e1e1e1); background-image:-ms-linear-gradient(#fff, #e1e1e1); background-image:linear-gradient(#fff, #e1e1e1); text-align:center; cursor:pointer;}
.cBk1, .cBk1 a {color:#000 !important;}
.ftLt {float:left;}
</style>
<link rel="stylesheet" href="/css/scm.css" type="text/css" />
<script language='javascript' src="/js/jsCal/js/jscal2.js"></script>
<script language='javascript' src="/js/jsCal/js/lang/ko.js"></script>
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/jscal2.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/border-radius.css" />
<link href="/js/jqueryui/css/jquery-ui.css" rel="stylesheet">
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript" src="/js/jqueryui/jquery-ui-1.10.2.custom.min.js"></script>
<script language="JavaScript" src="/js/common.js"></script>
<Script>
document.domain = "10x10.co.kr";

function goSavePlay(){
	if(frm1.opendate.value == ""){
		alert("오픈일을 입력하세요.");
		return;
	}
	if(frm1.state.value == ""){
		frm1.state.focus();
		alert("상태를 선택하세요.");
		return;
	}
	if(frm1.title.value == ""){
		alert("메인 타이틀 을 입력하세요.");
		frm1.title.focus();
		return;
	}
	if(frm1.keyword.value != ""){
		if(GetByteLength(frm1.keyword.value) > 250){
			alert("검색키워드는 250자 이내로 작성해주세요");
			frm1.keyword.focus();
			return false;
		}
	}

	frm1.submit();
}

function goCateOnChange(a){
	location.href = "cornerwrite.asp?midx=<%=vMIdx%>&didx=<%=vDidx%>&cate="+a+"";
}

function jsPopCal(sName){
	var winCal;

	winCal = window.open('/lib/common_cal.asp?in_domain=o&DN='+sName,'pCal','width=250, height=200');
	winCal.focus();
}

function jsImgView(sImgUrl){
 var wImgView;
 wImgView = window.open('/admin/sitemaster/play/lib/pop_event_detailImg.asp?sUrl='+sImgUrl,'pImg','width=100,height=100');
 wImgView.focus();
}

function jsUploadImg(a,b,c){
	document.domain ="10x10.co.kr";
	var popupl;
	popupl = window.open('/admin/sitemaster/play2016/pop_uploadimg.asp?folder='+a+'&span='+b+'&imggubun='+c+'','popupl','width=370,height=150');
	popupl.focus();
}

function jsDelImg(sName, sSpan){
	if(confirm("이미지를 삭제하시겠습니까?")){
	   eval("document.all."+sName).value = "";
	   eval("document.all."+sSpan).style.display = "none";
	}
}

function jsDelImgDB(i){
	if(confirm("이미지를 삭제하시겠습니까?\n\n즉시 데이터가 삭제됩니다.")){
	   frm3.idx.value = i;
	   frm3.submit();
	}
}

function jsSpanShowHide(s,n){
	$("#"+s+"").show();
	$("#"+n+"").hide();
}

function jsDelAzit(g){
	if(confirm("컨텐츠(장소"+g+")를 삭제하시겠습니까?") == true) {
		if(confirm("삭제하면 복구가 불가합니다.\n그래도 삭제하시겠습니까?") == true) {
			$("#groupnum").val(g);
			frm2.submit();
		}else{
			return false;
		}
	}else{
		return false;
	}
}

function putLinkText(device,key,valuename) {
	var frm = document.frm1;
	switch(key) {
		case 'event':
			$("#"+valuename+"").val("/event/eventmain.asp?eventid=이벤트번호");
			break;
		case 'itemid':
			if(device == "pc"){
				$("#"+valuename+"").val("/shopping/category_prd.asp?itemid=상품코드");
			}else{
				$("#"+valuename+"").val("/category/category_itemprd.asp?itemid=상품코드");
			}
			break;
		case 'brand':
			if(device == "pc"){
				$("#"+valuename+"").val("/street/street_brand_sub06.asp?makerid=브랜드아이디");
			}else{
				$("#"+valuename+"").val("/street/street_brand.asp?makerid=브랜드아이디");
			}
			break;
	}
}

function jsTagView(){
	if($("#istagview").is(":checked")){
		$("#tagspan").show();
	}else{
		$("#tagspan").hide();
	}
	
}
</Script>
</head>
<body TOPMARGIN="0">
<form name="frm3" action="imagedelproc.asp" method="post" style="margin:0px;">
<input type="hidden" name="idx" value="">
</form>
<form name="frm2" action="azitdelproc.asp" method="post" style="margin:0px;">
<input type="hidden" name="groupnum" id="groupnum" value="">
<input type="hidden" name="didx" value="<%=vDidx%>">
</form>
<form name="frm1" action="cornerproc.asp" method="post" style="margin:0px;">
<input type="hidden" name="action" value="<%=CHKIIF(vDidx="","insert","update")%>">
<input type="hidden" name="midx" value="<%=vMIdx%>">
<input type="hidden" name="didx" value="<%=vDidx%>">
<div style="padding:10px 0 5px 0;"><font size="3" color="#000000"><strong>│기본정보</strong></font></div>
<table class="tbType1 writeTb">
	<tbody>
		<tr>
			<th width="15%">Vol.</th>
			<td height="30"><%=Format00(3,vVolNum)%><% If vDidx <> "" Then %> (코너글번호 <%=vDidx%>)<% End If %></td>
		</tr>
		<tr>
			<th width="15%">코 너</th>
			<td height="30" style="padding-left:5px;">
				<% If vDidx <> "" Then
					Response.Write "<input type=hidden name=cate value=" & vCate & ">" & fnPlayCateName(vCate)
				Else %>
					<select name="cate" class="formSlt" onChange="goCateOnChange(this.value);">
						<option value=""> - 선택 - </option>
						<option value="41" <%=CHKIIF(vCate="41","selected","")%>>THING. > thing</option>
						<option value="42" <%=CHKIIF(vCate="42","selected","")%>>THING. > thingthing</option>
						<option value="43" <%=CHKIIF(vCate="43","selected","")%>>THING. > 배경화면</option>
						<option value="3" <%=CHKIIF(vCate="3","selected","")%>>TALK > AZIT&</option>
						<option value="31" <%=CHKIIF(vCate="31","selected","")%>>TALK > AZIT&Comma</option>
						<option value="1" <%=CHKIIF(vCate="1","selected","")%>>TALK > PLAYLIST♬</option>
						<option value="21" <%=CHKIIF(vCate="21","selected","")%>>!NSPIRATION > DESIGN</option>
						<option value="22" <%=CHKIIF(vCate="22","selected","")%>>!NSPIRATION > STYLE</option>
						<option value="5" <%=CHKIIF(vCate="5","selected","")%>>!NSPIRATION > COMMA,</option>
						<option value="6" <%=CHKIIF(vCate="6","selected","")%>>!NSPIRATION > HOWHOW?</option>
					</select>
				<% End If %>
			</td>
		</tr>
		<% If vCate <> "" Then %>
			<tr>
				<th width="15%">오픈일</th>
				<td height="30" style="padding-left:5px;">
					<input type="text" name="opendate" value="<%=vOpenDate%>" onClick="jsPopCal('opendate');" style="cursor:pointer;" size="10" maxlength="10" readonly>
				</td>
			</tr>
			<tr>
				<th width="15%">상 태</th>
				<td height="30" style="padding-left:5px;">
					<select name="state" class="formSlt">
						<%=fnStateSelectBox("select",vState)%>
					</select>
				</td>
			</tr>
			<tr>
				<th width="15%">담당자</th>
				<td height="30" style="padding-left:5px;">
					<select name="partmkid" >
						<option value="">선택</option>
						<option value="shaeiou" <%=CHKIIF(vPartMKID="shaeiou","selected","")%>>김시화</option>
						<option value="ascreem" <%=CHKIIF(vPartMKID="ascreem","selected","")%>>남찬</option>
						<option value="sss162000" <%=CHKIIF(vPartMKID="sss162000","selected","")%>>손아름</option>
						<option value="madebyash" <%=CHKIIF(vPartMKID="madebyash","selected","")%>>안서연</option>
						<option value="heejong1013" <%=CHKIIF(vPartMKID="heejong1013","selected","")%>>최희종</option>
						<option value="ppono2" <%=CHKIIF(vPartMKID="ppono2","selected","")%>>한유민</option>
						<option value="torymilk" <%=CHKIIF(vPartMKID="torymilk","selected","")%>>이서영</option>
						<option value="spinel93" <%=CHKIIF(vPartMKID="spinel93","selected","")%>>이수진</option>
					</select>
				</td>
			</tr>
			<tr>
				<th width="15%">작업자</th>
				<td height="30" style="padding-left:5px;">
					WD:<% sbGetpartid "partwdid",vPartWDID,"","12" %>
					&nbsp;&nbsp;&nbsp;
					퍼블리셔:
					<select name="partpbid">
						<option value="">선택</option>
						<option value="happyngirl" <%=CHKIIF(vPartPBID="happyngirl","selected","")%>>최선미</option>
						<option value="kyungae13" <%=CHKIIF(vPartPBID="kyungae13","selected","")%>>조경애</option>
						<option value="jinyeonmi" <%=CHKIIF(vPartPBID="jinyeonmi","selected","")%>>진연미</option>
						<option value="jj999a" <%=CHKIIF(vPartPBID="jj999a","selected","")%>>김송이</option>
					</select>
				</td>
			</tr>
			<tr>
				<th width="15%">검색키워드<br />(250자 이내)<br />(쉼표로 구분)</th>
				<td height="30" style="padding-left:5px;">
					<textarea name="keyword" rows="3" cols="80"><%=vKeyword%></textarea>
				</td>
			</tr>
		<% End If %>
	</tbody>
</table>
<% If vCate <> "" AND vCate <> "2" AND vCate <> "4" Then %>
<div style="padding:20px 0 5px 0;"><font size="3" color="#000000"><strong>│컨텐츠 등록</strong></font></div>
<table class="tbType1 writeTb">
	<tbody>
		<tr>
			<th width="15%">메인 타이틀<br>(텍스트만)</th>
			<td height="40" style="padding-left:5px;"><input type="text" name="title" value="<%=vTitle%>" size="60" maxlength="146"> * " 따옴표 금지요.</td>
		</tr>
		<tr>
			<th width="15%">메인 타이틀<br>(스타일용)</th>
			<td height="40" style="padding-left:5px;"><input type="text" name="titlestyle" value="<%=vTitleStyle%>" size="60" maxlength="196"> * " 따옴표 금지요.</td>
		</tr>
		<tr>
			<th width="15%">서브카피</th>
			<td height="30" style="padding:5px 5px 5px 5px;"><textarea name="subcopy" rows="3" cols="80"><%=vSubCopy%></textarea></td>
		</tr>
		<tr>
			<th width="15%">리스트 이미지<br />(<font color="blue">직사각형</font>)<br />823x540<br /><font size="1" color="silver">(이미지구분 : 1)</font></th>
			<td height="70" style="padding:5px 5px 5px 5px;">
				<input type="button" value="이미지등록" onClick="jsUploadImg('jiklistimg','jiklistimgspan','1');" /><br /><br />
				<span id="jiklistimgspan" style="padding:5px 5px 5px 0;"><%
					If vJikListImg <> "" Then
						Response.Write "<img src='" & vJikListImg & "' height='100' style='cursor:pointer;' onclick=jsImgView('" & vJikListImg & "');>"
						Response.Write "<a href=javascript:jsDelImg('jiklistimg','jiklistimgspan');><img src='/images/icon_delete2.gif' border='0'></a>"
					End If
				%></span>
				<input type="hidden" name="jiklistimg" value="<%=vJikListImg%>">
			</td>
		</tr>
		<tr>
			<th width="15%">리스트 이미지<br />(<font color="blue">정사각형</font>)<br />320x320<br /><font size="1" color="silver">(이미지구분 : 11)</font></th>
			<td height="70" style="padding:5px 5px 5px 5px;">
				<input type="button" value="이미지등록" onClick="jsUploadImg('junglistimg','junglistimgspan','11');" /><br /><br />
				<span id="junglistimgspan" style="padding:5px 5px 5px 0;"><%
					If vJungListImg <> "" Then
						Response.Write "<img src='" & vJungListImg & "' height='100' style='cursor:pointer;' onclick=jsImgView('" & vJungListImg & "');>"
						Response.Write "<a href=javascript:jsDelImg('junglistimg','junglistimgspan');><img src='/images/icon_delete2.gif' border='0'></a>"
					End If
				%></span>
				<input type="hidden" name="junglistimg" value="<%=vJungListImg%>">
			</td>
		</tr>
		<tr>
			<th width="15%">검색리스트<br />배너이미지<br />750x140<br /><font size="1" color="silver">(이미지구분 : 28)</font></th>
			<td height="70" style="padding:5px 5px 5px 5px;">
				<input type="button" value="이미지등록" onClick="jsUploadImg('searchlistimg','searchlistimgspan','28');" /><br /><br />
				<span id="searchlistimgspan" style="padding:5px 5px 5px 0;"><%
					If vSearchListImg <> "" Then
						Response.Write "<img src='" & vSearchListImg & "' height='100' style='cursor:pointer;' onclick=jsImgView('" & vSearchListImg & "');>"
						Response.Write "<a href=javascript:jsDelImg('searchlistimg','searchlistimgspan');><img src='/images/icon_delete2.gif' border='0'></a>"
					End If
				%></span>
				<input type="hidden" name="searchlistimg" value="<%=vSearchListImg%>">
			</td>
		</tr>
		<tr>
			<th width="15%">작업 전달 사항</th>
			<td height="30" style="padding:5px 5px 5px 5px;">
				<textarea name="worktext" rows="10" cols="80"><%=vWorkText%></textarea>
			</td>
		</tr>
		<%
		If vCate = "1" Then	'### PLAYLIST
			If vCate1Type = "" Then
				vCate1Type = "1"
			End If
		%>
			<tr>
				<th width="15%">컬러코드</th>
				<td height="30" style="padding-left:5px;">
					# <input type="text" name="mo_bgcolor" value="<%=vMoBGColor%>" size="10" maxlength="6"> * #을 제외한 6자 숫자로만 입력하세요.
				</td>
			</tr>
			<tr>
				<th width="15%">컨텐츠(제작자)</th>
				<td height="30" style="padding-left:5px;"><input type="text" name="cate1directer" value="<%=vCate1Directer%>" size="30" maxlength="146"></td>
			</tr>
			<tr>
				<th width="15%">컨텐츠(이미지)</th>
				<td height="50" style="padding-left:5px;">
					<input type="radio" name="cate1type" value="1" onClick="jsSpanShowHide('cate1span1','cate1span2');" <%=CHKIIF(vCate1Type="1","checked","")%>> 동영상&nbsp;&nbsp;&nbsp;
					<input type="radio" name="cate1type" value="2" onClick="jsSpanShowHide('cate1span2','cate1span1');" <%=CHKIIF(vCate1Type="2","checked","")%>> 이미지<font size="1" color="silver">(이미지구분 : 2)</font><br />
					<span id="cate1span1" style="display:none;">
						<input type="text" name="cate1videourl" value="<%=vCate1VideoURL%>" size="50" maxlength="190">  * http:// 를 포함한 주소를 입력하세요.<br />
						출처 : <input type="text" name="cate1videoorigin" value="<%=vCate1VideoOrigin%>" size="50" maxlength="190"> * 필수값 아님.
					</span>
					<span id="cate1span2" style="display:none;">
						<input type="button" value="이미지등록" onClick="jsUploadImg('cate1imageurl','cate1imgspan','2');" /><br /><br />
						<span id="cate1imgspan" style="padding:5px 5px 5px 0;"><%
							If vCate1ImageURL <> "" Then
								Response.Write "<img src='" & vCate1ImageURL & "' height='100' style='cursor:pointer;' onclick=jsImgView('" & vCate1ImageURL & "');>"
								Response.Write "<a href=javascript:jsDelImg('cate1imageurl','cate1imgspan');><img src='/images/icon_delete2.gif' border='0'></a>"
							End If
						%></span>
						<input type="hidden" name="cate1imageurl" value="<%=vCate1ImageURL%>">
					</span>
					<script>
						<% If vCate1Type = "1" Then %>
							jsSpanShowHide("cate1span1","cate1span2");
						<% Else %>
							jsSpanShowHide("cate1span2","cate1span1");
						<% End If %>
					</script>
				</td>
			</tr>
			<tr>
				<th width="15%">참여(코멘트)</th>
				<td height="130" style="padding-left:5px;">
					주제 : <input type="text" name="cate1commtitle" value="<%=vCate1CommTitle%>" size="50" maxlength="148"><br /><br />
					입력란 : 
					<input type="radio" name="cate1commcnt" value="2" onClick="jsSpanShowHide('tempspan','cate1commspan');$('#cate1comment3').val('');" <%=CHKIIF(vCate1Comment3="","checked","")%>> 2개&nbsp;&nbsp;&nbsp;
					<input type="radio" name="cate1commcnt" value="3" onClick="jsSpanShowHide('cate1commspan','tempspan');" <%=CHKIIF(vCate1Comment3<>"","checked","")%>> 3개<br />
					첫번째 입력란카피 : <input type="text" name="cate1precomm1" value="<%=vCate1precomm1%>" size="50" maxlength="48"> * #OO (예:#어디)<br />
					첫번째 코멘트 : &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<input type="text" name="cate1comment1" value="<%=vCate1Comment1%>" size="50" maxlength="48"><br />
					두번째 입력란카피 : <input type="text" name="cate1precomm2" value="<%=vCate1precomm2%>" size="50" maxlength="48"><br />
					두번째 코멘트 : &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<input type="text" name="cate1comment2" value="<%=vCate1Comment2%>" size="50" maxlength="48"><br />
					<span id="cate1commspan" style="display:none;">
						세번째 입력란카피 : <input type="text" name="cate1precomm3" value="<%=vCate1precomm3%>" size="50" maxlength="48"><br />
						세번째 코멘트 : &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<input type="text" name="cate1comment3" id="cate1comment3" value="<%=vCate1Comment3%>" size="50" maxlength="48"><br />
					</span>
					<script>
						<% If vCate1Comment3 <> "" Then %>
							jsSpanShowHide("cate1commspan","tempspan");
						<% Else %>
							jsSpanShowHide("tempspan","cate1commspan");
						<% End If %>
					</script>
				</td>
			</tr>
			<tr>
				<th width="15%">테그노출</th>
				<td height="30" style="padding:5px 5px 5px 5px;">
					<input type="checkbox" name="istagview" id="istagview" value="1" onClick="jsTagView();" <%=CHKIIF(vIsTagView,"checked","")%>> 노출&nbsp;&nbsp;&nbsp;
					<span id="tagspan" style="display:<%=CHKIIF(vIsTagView,"block","none")%>;">
					* 이벤트중테그 노출기간 : <input type="text" name="tagsdate" value="<%=vTagSDate%>" onClick="jsPopCal('tagsdate');" style="cursor:pointer;" size="10" maxlength="10" readonly>&nbsp;~&nbsp;
								<input type="text" name="tagedate" value="<%=vTagEDate%>" onClick="jsPopCal('tagedate');" style="cursor:pointer;" size="10" maxlength="10" readonly>&nbsp;&nbsp;
					* 당첨자발표테그 오픈일 : <input type="text" name="tagannouncedate" value="<%=vTagAnnounceDate%>" onClick="jsPopCal('tagannouncedate');" style="cursor:pointer;" size="10" maxlength="10" readonly>
					<br />
					* 리워드 카피 : <input type="text" name="cate1rewardcopy" id="cate1rewardcopy" value="<%=vCate1RewardCopy%>" size="120">
					</span>
				</td>
			</tr>
			<tr>
				<th width="15%">연결배너<br />(PC 이미지)<br /><font size="1" color="silver">(이미지구분 : 3)</font></th>
				<td height="30" style="padding:5px 5px 5px 5px;">
					<input type="button" value="이미지등록" onClick="jsUploadImg('cate1pclinkbanimg','cate1pclinkbanspan','3');" /><br /><br />
					<span id="cate1pclinkbanspan" style="padding:5px 5px 5px 0;"><%
						If vCate1PCLinkBanImg <> "" Then
							Response.Write "<img src='" & vCate1PCLinkBanImg & "' height='100' style='cursor:pointer;' onclick=jsImgView('" & vCate1PCLinkBanImg & "');>"
							Response.Write "<a href=javascript:jsDelImg('cate1pclinkbanimg','cate1pclinkbanspan');><img src='/images/icon_delete2.gif' border='0'></a>"
						End If
					%></span><br />
					<input type="hidden" name="cate1pclinkbanimg" value="<%=vCate1PCLinkBanImg%>">
					<br />
					링크주소 : <input type="text" name="cate1pclinkbanurl" id="cate1pclinkbanurl" value="<%=vCate1PCLinkBanURL%>" size="60" maxlength="195">
					<br />
					<font color="#707070">
					- <font color="red"><strong>app & mobile 공용</strong></font> - <br/>
					- <span style="cursor:pointer" onClick="putLinkText('pc','event','cate1pclinkbanurl')">이벤트 링크 : /event/eventmain.asp?eventid=<font color="darkred">이벤트코드</font></span><br/>
					- <span style="cursor:pointer" onClick="putLinkText('pc','itemid','cate1pclinkbanurl')">상품코드 링크 : /shopping/category_prd.asp?itemid=<font color="darkred">상품코드 (O)</font></span><br/>
					- <span style="cursor:pointer" onClick="putLinkText('pc','brand','cate1pclinkbanurl')">브랜드아이디 링크 : /street/street_brand_sub06.asp?makerid=<font color="darkred">브랜드아이디</font></span><br/>
					</font>
				</td>
			</tr>
			<tr>
				<th width="15%">연결배너<br />(Mo 이미지)<br /><font size="1" color="silver">(이미지구분 : 18)</font></th>
				<td height="30" style="padding:5px 5px 5px 5px;">
					<input type="button" value="이미지등록" onClick="jsUploadImg('cate1molinkbanimg','cate1molinkbanspan','18');" /><br /><br />
					<span id="cate1molinkbanspan" style="padding:5px 5px 5px 0;"><%
						If vCate1MoLinkBanImg <> "" Then
							Response.Write "<img src='" & vCate1MoLinkBanImg & "' height='100' style='cursor:pointer;' onclick=jsImgView('" & vCate1MoLinkBanImg & "');>"
							Response.Write "<a href=javascript:jsDelImg('cate1molinkbanimg','cate1molinkbanspan');><img src='/images/icon_delete2.gif' border='0'></a>"
						End If
					%></span><br />
					<input type="hidden" name="cate1molinkbanimg" value="<%=vCate1MoLinkBanImg%>">
					<br />
					링크주소 : <input type="text" name="cate1molinkbanurl" id="cate1molinkbanurl" value="<%=vCate1MoLinkBanURL%>" size="60" maxlength="195">
					<br />
					<font color="#707070">
					- <font color="red"><strong>app & mobile 공용</strong></font> - <br/>
					- <span style="cursor:pointer" onClick="putLinkText('m','event','cate1molinkbanurl')">이벤트 링크 : /event/eventmain.asp?eventid=<font color="darkred">이벤트코드</font></span><br/>
					- <span style="cursor:pointer" onClick="putLinkText('m','itemid','cate1molinkbanurl')">상품코드 링크 : /category/category_itemprd.asp?itemid=<font color="darkred">상품코드 (O)</font></span><br/>
					- <span style="cursor:pointer" onClick="putLinkText('m','brand','cate1molinkbanurl')">브랜드아이디 링크 : /street/street_brand.asp?makerid=<font color="darkred">브랜드아이디</font></span><br/>
					</font>
				</td>
			</tr>
		<% ElseIf vCate = "21" Then	'### INSPIRATION Design %>
			<tr>
				<th width="15%">컬러코드</th>
				<td height="30" style="padding-left:5px;">
					# <input type="text" name="mo_bgcolor" value="<%=vMoBGColor%>" size="10" maxlength="6"> * #을 제외한 6자 숫자로만 입력하세요.
				</td>
			</tr>
			<tr>
				<th width="15%">컨텐츠(이미지)<br /><font size="1" color="silver">(이미지구분 : 4)</font></th>
				<td height="30" style="padding:5px 5px 5px 5px;">
					<table width="100%">
					<tr>
						<%
						For l=1 To 5
						%>
						<td width="20%" style="border-width:1px; border-style:solid; border-collapse:collapse; border-color:#D3D3D3;">
							<input type="button" value="이미지등록" onClick="jsUploadImg('cate21Img<%=l%>','cate21Imgspan<%=l%>','4');" /><br /><br />
							<span id="cate21Imgspan<%=l%>" style="padding:5px 5px 5px 0;"><%
								If vCate21ImageURL(l) <> "" Then
									Response.Write "<img src='" & vCate21ImageURL(l) & "' height='100' style='cursor:pointer;' onclick=jsImgView('" & vCate21ImageURL(l) & "');>"
									Response.Write "<a href=javascript:jsDelImgDB('"&vCate21ImageIdx(l)&"');><img src='/images/icon_delete2.gif' border='0'></a>"
								End If
							%></span>
							<input type="hidden" name="cate21Img<%=l%>" value="<%=vCate21ImageURL(l)%>">
						</td>
						<% Next %>
					</tr>
					</table>
				</td>
			</tr>
			<tr>
				<th width="15%">상품코드</th>
				<td height="30" style="padding-left:5px;">
					<input type="text" name="cate21item" value="<%=vCate21Item%>" size="60">  * 콤마(,)로 구분
				</td>
			</tr>
		<% ElseIf vCate = "22" Then	'### INSPIRATION STYLE %>
			<tr>
				<th width="15%">컬러코드</th>
				<td height="30" style="padding-left:5px;">
					# <input type="text" name="mo_bgcolor" value="<%=vMoBGColor%>" size="10" maxlength="6"> * #을 제외한 6자 숫자로만 입력하세요.
				</td>
			</tr>
			<tr>
				<th width="15%">컨텐츠(이미지)<br /><font size="1" color="silver">(이미지구분 : 5)</font></th>
				<td height="30" style="padding:5px 5px 5px 5px;">
					<table width="100%">
					<tr>
						<%
						For l=1 To 5
						%>
						<td width="20%" style="border-width:1px; border-style:solid; border-collapse:collapse; border-color:#D3D3D3;">
							<input type="button" value="이미지등록" onClick="jsUploadImg('cate22Img<%=l%>','cate22Imgspan<%=l%>','5');" /><br /><br />
							<span id="cate22Imgspan<%=l%>" style="padding:5px 5px 5px 0;"><%
								If vCate22ImageURL(l) <> "" Then
									Response.Write "<img src='" & vCate22ImageURL(l) & "' height='100' style='cursor:pointer;' onclick=jsImgView('" & vCate22ImageURL(l) & "');>"
									Response.Write "<a href=javascript:jsDelImgDB('"&vCate22ImageIdx(l)&"');><img src='/images/icon_delete2.gif' border='0'></a>"
								End If
							%></span>
							<input type="hidden" name="cate22Img<%=l%>" value="<%=vCate22ImageURL(l)%>">
						</td>
						<% Next %>
					</tr>
					</table>
				</td>
			</tr>
			<tr>
				<th width="15%">상품코드</th>
				<td height="30" style="padding-left:5px;">
					<input type="text" name="cate22item" value="<%=vCate22Item%>" size="60">  * 콤마(,)로 구분
				</td>
			</tr>
		<% ElseIf vCate = "3" Then	'### AZIT %>
			<tr>
				<th width="15%">아이콘링크</th>
				<td height="30" style="padding-left:5px;">
					<input type="text" name="cate3icon" value="<%=vCate3Icon%>" size="60">
				</td>
			</tr>
			<tr>
				<th width="15%">컨텐츠<br />(PC 상단 이미지)<br /><font size="1" color="silver">(이미지구분 : 19)</font></th>
				<td height="30" style="padding:5px 5px 5px 5px;">
					<input type="button" value="이미지등록" onClick="jsUploadImg('cate3pcimg','cate3pcimgspan','19');" /><br /><br />
					<span id="cate3pcimgspan" style="padding:5px 5px 5px 0;"><%
						If vCate3PCImageURL <> "" Then
							Response.Write "<img src='" & vCate3PCImageURL & "' height='100' style='cursor:pointer;' onclick=jsImgView('" & vCate3PCImageURL & "');>"
							Response.Write "<a href=javascript:jsDelImg('cate3pcimg','cate3pcimgspan');><img src='/images/icon_delete2.gif' border='0'></a>"
						End If
					%></span>
					<input type="hidden" name="cate3pcimg" value="<%=vCate3PCImageURL%>"
				</td>
			</tr>
			<tr>
				<th width="15%">컨텐츠<br />(Mo 상단 이미지)<br /><font size="1" color="silver">(이미지구분 : 6)</font></th>
				<td height="30" style="padding:5px 5px 5px 5px;">
					<input type="button" value="이미지등록" onClick="jsUploadImg('cate3moimg','cate3moimgspan','6');" /><br /><br />
					<span id="cate3moimgspan" style="padding:5px 5px 5px 0;"><%
						If vCate3MoImageURL <> "" Then
							Response.Write "<img src='" & vCate3MoImageURL & "' height='100' style='cursor:pointer;' onclick=jsImgView('" & vCate3MoImageURL & "');>"
							Response.Write "<a href=javascript:jsDelImg('cate3moimg','cate3moimgspan');><img src='/images/icon_delete2.gif' border='0'></a>"
						End If
					%></span>
					<input type="hidden" name="cate3moimg" value="<%=vCate3MoImageURL%>"
				</td>
			</tr>
			<%
			tmp = 0
			For p=1 To 4
			%>
			<tr>
				<th width="15%">컨텐츠(장소<%=p%>)<br /><font size="1" color="silver">(이미지구분 : 7)</font>
					<br /><input type="button" value="장소<%=p%>삭제" onClick="jsDelAzit('<%=p%>');">
				</th>
				<td style="padding:5px 5px 5px 5px;">
					타이틀 : <input type="text" name="cate3P<%=p%>title" value="<%=vCate3Ptitle(p)%>" size="60"> * 타이틀 미입력시 장소<%=p%>은 저장되지 않습니다.<br />
					주소 : <input type="text" name="cate3P<%=p%>juso" value="<%=vCate3Pjuso(p)%>" size="62"><br />
					링크 : <input type="text" name="cate3P<%=p%>link" value="<%=vCate3Plink(p)%>" size="90"><br />
					<br />
					<table width="100%">
					<tr>
						<% For l=1 To 5
								tmp = tmp + 1
						%>
						<td width="20%" style="border-width:1px; border-style:solid; border-collapse:collapse; border-color:#D3D3D3;">
							<%=l%>.<br />
							<input type="button" value="이미지등록" onClick="jsUploadImg('cate3P<%=p%>Img<%=l%>','cate3P<%=p%>Imgspan<%=l%>','7');" /><br /><br />
							<span id="cate3P<%=p%>Imgspan<%=l%>" style="padding:5px 5px 5px 0;"><%
								If vCate3PImg(tmp) <> "" Then
									Response.Write "<img src='" & vCate3PImg(tmp) & "' height='100' style='cursor:pointer;' onclick=jsImgView('" & vCate3PImg(tmp) & "');>"
									Response.Write "<a href=javascript:jsDelImg('cate3P"&p&"Img"&l&"','cate3P"&p&"Imgspan"&l&"');><img src='/images/icon_delete2.gif' border='0'></a>"
								End If
							%></span>
							<input type="hidden" name="cate3P<%=p%>Img<%=l%>" value="<%=vCate3PImg(tmp)%>"><br />
							<textarea name="cate3P<%=p%>copy<%=l%>" rows="3" cols="20"><%=vCate3PCopy(tmp)%></textarea><br />
						</td>
						<% Next %>
					</tr>
					</table>
				</td>
			</tr>
			<%
			Next
			%>
			<tr>
				<th width="15%">참여(응모내용)</th>
				<td height="30" style="padding:5px 5px 5px 5px;"><textarea name="cate3entrycont" rows="3" cols="80"><%=vCate3EntryCont%></textarea><br /></td>
			</tr>
			<tr>
				<th width="15%">참여(응모기간)</th>
				<td height="30" style="padding:5px 5px 5px 5px;">
					응모기간 : <input type="text" name="cate3entrysdate" value="<%=vCate3EntrySDate%>" onClick="jsPopCal('cate3entrysdate');" style="cursor:pointer;" size="10" maxlength="10" readonly>&nbsp;~&nbsp;
								<input type="text" name="cate3entryedate" value="<%=vCate3EntryEDate%>" onClick="jsPopCal('cate3entryedate');" style="cursor:pointer;" size="10" maxlength="10" readonly><br />
					당첨자발표 : <input type="text" name="cate3announdate" value="<%=vCate3AnnounDate%>" onClick="jsPopCal('cate3announdate');" style="cursor:pointer;" size="10" maxlength="10" readonly>
				</td>
			</tr>
			<tr>
				<th width="15%">테그노출</th>
				<td height="30" style="padding:5px 5px 5px 5px;">
					<input type="checkbox" name="istagview" id="istagview" value="1" onClick="jsTagView();" <%=CHKIIF(vIsTagView,"checked","")%>> 노출&nbsp;&nbsp;&nbsp;
					<span id="tagspan" style="display:<%=CHKIIF(vIsTagView,"block","none")%>;">
					* 이벤트중테그 노출기간 : <input type="text" name="tagsdate" value="<%=vTagSDate%>" onClick="jsPopCal('tagsdate');" style="cursor:pointer;" size="10" maxlength="10" readonly>&nbsp;~&nbsp;
								<input type="text" name="tagedate" value="<%=vTagEDate%>" onClick="jsPopCal('tagedate');" style="cursor:pointer;" size="10" maxlength="10" readonly>&nbsp;&nbsp;
					* 당첨자발표테그 오픈일 : <input type="text" name="tagannouncedate" value="<%=vTagAnnounceDate%>" onClick="jsPopCal('tagannouncedate');" style="cursor:pointer;" size="10" maxlength="10" readonly>
					</span>
				</td>
			</tr>
			<tr>
				<th width="15%">응모표현방법</th>
				<td height="30" style="padding:5px 5px 5px 5px;">
					<input type="radio" name="entry_method" value="c" <%=CHKIIF(vCate3EntryMethod="c","checked","")%>> 코멘트형식&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
					<input type="radio" name="entry_method" value="b" <%=CHKIIF(vCate3EntryMethod="b","checked","")%>> 버튼형식(초기버전)
				</td>
			</tr>
			<tr>
				<th width="15%">유의사항</th>
				<td height="30" style="padding:5px 5px 5px 5px;"><textarea name="cate3notice" rows="3" cols="80"><%=vCate3Notice%></textarea></td>
			</tr>

		<%'// 2017.06.01 원승현 azit comma 스타일 추가 %>
		<% ElseIf vCate = "31" Then %>
			<tr>
				<th width="15%">컨텐츠<br />(PC 상단 이미지)<br /><font size="1" color="silver">(이미지구분 : 23)</font></th>
				<td height="30" style="padding:5px 5px 5px 5px;">
					<input type="button" value="이미지등록" onClick="jsUploadImg('cate31pctopimg','cate31pctopimgspan','23');" /><br /><br />
					<span id="cate31pctopimgspan" style="padding:5px 5px 5px 0;"><%
						If vCate31PCTopImageURL <> "" Then
							Response.Write "<img src='" & vCate31PCTopImageURL & "' height='100' style='cursor:pointer;' onclick=jsImgView('" & vCate31PCTopImageURL & "');>"
							Response.Write "<a href=javascript:jsDelImg('cate31pctopimg','cate31pctopimgspan');><img src='/images/icon_delete2.gif' border='0'></a>"
						End If
					%></span>
					<input type="hidden" name="cate31pctopimg" value="<%=vCate31PCTopImageURL%>"
				</td>
			</tr>
			<tr>
				<th width="15%">컨텐츠<br />(Mo 상단 이미지)<br /><font size="1" color="silver">(이미지구분 : 24)</font></th>
				<td height="30" style="padding:5px 5px 5px 5px;">
					<input type="button" value="이미지등록" onClick="jsUploadImg('cate31motopimg','cate31motopimgspan','24');" /><br /><br />
					<span id="cate31motopimgspan" style="padding:5px 5px 5px 0;"><%
						If vCate31MoTopImageURL <> "" Then
							Response.Write "<img src='" & vCate31MoTopImageURL & "' height='100' style='cursor:pointer;' onclick=jsImgView('" & vCate31MoTopImageURL & "');>"
							Response.Write "<a href=javascript:jsDelImg('cate31motopimg','cate31motopimgspan');><img src='/images/icon_delete2.gif' border='0'></a>"
						End If
					%></span>
					<input type="hidden" name="cate31motopimg" value="<%=vCate31MoTopImageURL%>"
				</td>
			</tr>
			<tr>
				<th width="15%">컨텐츠(제작자)</th>
				<td height="30" style="padding-left:5px;"><input type="text" name="cate31directer" value="<%=vCate31Directer%>" size="60" maxlength="146"></td>
			</tr>
			<tr>
				<th width="15%">컨텐츠(에디터)<br /><font size="1" color="silver">(이미지구분 : 25)</font></th>
				<td style="padding:5px 5px 5px 5px;">
				<%
				tmp = 0
				For l=1 To 5
						tmp = tmp + 1
				%>
					<%=l%>.&nbsp;&nbsp;
					<input type="button" value="이미지등록" onClick="jsUploadImg('cate31Img<%=l%>','cate31Imgspan<%=l%>','25');" /><br /><br />
					<span id="cate31Imgspan<%=l%>" style="padding:5px 5px 5px 0;"><%
						If vCate31Img(tmp) <> "" Then
							Response.Write "<img src='" & vCate31Img(tmp) & "' height='100' style='cursor:pointer;' onclick=jsImgView('" & vCate31Img(tmp) & "');>"
							Response.Write "<a href=javascript:jsDelImg('cate31Img"&l&"','cate31Imgspan"&l&"');><img src='/images/icon_delete2.gif' border='0'></a>"
						End If
					%></span>
					<input type="hidden" name="cate31Img<%=l%>" value="<%=vCate31Img(tmp)%>"><br />
					<textarea name="cate31copy<%=l%>" rows="3" cols="80"><%=vCate31Copy(tmp)%></textarea><br />
					<hr />
				<% Next %>
				</td>
			</tr>
			<tr>
				<th width="15%">연결배너<br />(PC 이미지)<br /><font size="1" color="silver">(이미지구분 : 26)</font></th>
				<td height="30" style="padding:5px 5px 5px 5px;">
					<input type="button" value="이미지등록" onClick="jsUploadImg('cate31pclinkbanimg','cate31pclinkbanspan','26');" /><br /><br />
					<span id="cate31pclinkbanspan" style="padding:5px 5px 5px 0;"><%
						If vCate31PCLinkBanImg <> "" Then
							Response.Write "<img src='" & vCate31PCLinkBanImg & "' height='100' style='cursor:pointer;' onclick=jsImgView('" & vCate31PCLinkBanImg & "');>"
							Response.Write "<a href=javascript:jsDelImg('cate31pclinkbanimg','cate31pclinkbanspan');><img src='/images/icon_delete2.gif' border='0'></a>"
						End If
					%></span><br />
					<input type="hidden" name="cate31pclinkbanimg" value="<%=vCate31PCLinkBanImg%>">
					<br />
					링크주소 : <input type="text" name="cate31pclinkbanurl" id="cate31pclinkbanurl" value="<%=vCate31PCLinkBanURL%>" size="60" maxlength="195">
					<br />
					<font color="#707070">
					- <font color="red"><strong>app & mobile 공용</strong></font> - <br/>
					- <span style="cursor:pointer" onClick="putLinkText('pc','event','cate31pclinkbanurl')">이벤트 링크 : /event/eventmain.asp?eventid=<font color="darkred">이벤트코드</font></span><br/>
					- <span style="cursor:pointer" onClick="putLinkText('pc','itemid','cate31pclinkbanurl')">상품코드 링크 : /shopping/category_prd.asp?itemid=<font color="darkred">상품코드 (O)</font></span><br/>
					- <span style="cursor:pointer" onClick="putLinkText('pc','brand','cate31pclinkbanurl')">브랜드아이디 링크 : /street/street_brand_sub06.asp?makerid=<font color="darkred">브랜드아이디</font></span><br/>
					</font>
				</td>
			</tr>
			<tr>
				<th width="15%">연결배너<br />(Mo 이미지)<br /><font size="1" color="silver">(이미지구분 : 27)</font></th>
				<td height="30" style="padding:5px 5px 5px 5px;">
					<input type="button" value="이미지등록" onClick="jsUploadImg('cate31molinkbanimg','cate31molinkbanspan','27');" /><br /><br />
					<span id="cate31molinkbanspan" style="padding:5px 5px 5px 0;"><%
						If vCate31MoLinkBanImg <> "" Then
							Response.Write "<img src='" & vCate31MoLinkBanImg & "' height='100' style='cursor:pointer;' onclick=jsImgView('" & vCate31MoLinkBanImg & "');>"
							Response.Write "<a href=javascript:jsDelImg('cate31molinkbanimg','cate31molinkbanspan');><img src='/images/icon_delete2.gif' border='0'></a>"
						End If
					%></span><br />
					<input type="hidden" name="cate31molinkbanimg" value="<%=vCate31MoLinkBanImg%>">
					<br />
					링크주소 : <input type="text" name="cate31molinkbanurl" id="cate31molinkbanurl" value="<%=vCate31MoLinkBanURL%>" size="60" maxlength="195">
					<br />
					<font color="#707070">
					- <font color="red"><strong>app & mobile 공용</strong></font> - <br/>
					- <span style="cursor:pointer" onClick="putLinkText('m','event','cate31molinkbanurl')">이벤트 링크 : /event/eventmain.asp?eventid=<font color="darkred">이벤트코드</font></span><br/>
					- <span style="cursor:pointer" onClick="putLinkText('m','itemid','cate31molinkbanurl')">상품코드 링크 : /category/category_itemprd.asp?itemid=<font color="darkred">상품코드 (O)</font></span><br/>
					- <span style="cursor:pointer" onClick="putLinkText('m','brand','cate31molinkbanurl')">브랜드아이디 링크 : /street/street_brand.asp?makerid=<font color="darkred">브랜드아이디</font></span><br/>
					</font>
				</td>
			</tr>

		<% ElseIf vCate = "41" Then	'### Thing. Thing. %>
			<tr>
				<th width="15%">PCWEB EXECUTE<br />FilePath<br />(개발자전용)</th>
				<td height="70" style="padding-left:5px;">
					<input type="radio" name="pc_isexec" value="1" <%=CHKIIF(vCate41PCIsExec=True,"checked","")%>> 사용함&nbsp;&nbsp;&nbsp;
					<input type="radio" name="pc_isexec" value="0" <%=CHKIIF(vCate41PCIsExec=False,"checked","")%>> 사용안함<br />
					<input type="text" name="pc_execfile" value="<%=vCate41PCExecFile%>" size="50" maxlength="98">
				</td>
			</tr>
			<tr>
				<th width="15%">PC HTML</th>
				<td height="30" style="padding:5px 5px 5px 5px;">
					<textarea name="pc_contents" rows="10" cols="100"><%=vCate41PCContent%></textarea>
				</td>
			</tr>
			<tr>
				<th width="15%">MOBILE EXECUTE<br />FilePath<br />(개발자전용)</th>
				<td height="70" style="padding-left:5px;">
					<input type="radio" name="mo_isexec" value="1" <%=CHKIIF(vCate41MoIsExec=True,"checked","")%>> 사용함&nbsp;&nbsp;&nbsp;
					<input type="radio" name="mo_isexec" value="0" <%=CHKIIF(vCate41MoIsExec=False,"checked","")%>> 사용안함<br />
					<input type="text" name="mo_execfile" value="<%=vCate41MoExecFile%>" size="50" maxlength="98">
				</td>
			</tr>
			<tr>
				<th width="15%">MOBILE HTML</th>
				<td height="30" style="padding:5px 5px 5px 5px;">
					<textarea name="mo_contents" rows="10" cols="100"><%=vCate41MoContent%></textarea>
				</td>
			</tr>
			<tr>
				<th width="15%">테그노출</th>
				<td height="30" style="padding:5px 5px 5px 5px;">
					<input type="checkbox" name="istagview" id="istagview" value="1" onClick="jsTagView();" <%=CHKIIF(vIsTagView,"checked","")%>> 노출&nbsp;&nbsp;&nbsp;
					<span id="tagspan" style="display:<%=CHKIIF(vIsTagView,"block","none")%>;">
					* 이벤트중테그 노출기간 : <input type="text" name="tagsdate" value="<%=vTagSDate%>" onClick="jsPopCal('tagsdate');" style="cursor:pointer;" size="10" maxlength="10" readonly>&nbsp;~&nbsp;
								<input type="text" name="tagedate" value="<%=vTagEDate%>" onClick="jsPopCal('tagedate');" style="cursor:pointer;" size="10" maxlength="10" readonly>&nbsp;&nbsp;
					* 당첨자발표테그 오픈일 : <input type="text" name="tagannouncedate" value="<%=vTagAnnounceDate%>" onClick="jsPopCal('tagannouncedate');" style="cursor:pointer;" size="10" maxlength="10" readonly>
					</span>
				</td>
			</tr>
		<% ElseIf vCate = "42" Then	'### Thing. ThingThing. %>
			<tr>
				<th width="15%">컬러코드</th>
				<td height="30" style="padding-left:5px;">
					# <input type="text" name="mo_bgcolor" value="<%=vMoBGColor%>" size="10" maxlength="6"> * #을 제외한 6자 숫자로만 입력하세요.
				</td>
			</tr>
			<tr>
				<th width="15%">컨텐츠(롤링이미지)<br /><font size="1" color="silver">(이미지구분 : 8)</font></th>
				<td style="padding:5px 5px 5px 5px;">
					<table width="100%">
					<tr>
						<%
						tmp = 0
						For l=1 To 3
								tmp = tmp + 1
						%>
						<td width="33%" style="border-width:1px; border-style:solid; border-collapse:collapse; border-color:#D3D3D3;">
							<input type="button" value="이미지등록" onClick="jsUploadImg('cate42Img<%=l%>','cate42Imgspan<%=l%>','7');" /><br /><br />
							<span id="cate42Imgspan<%=l%>" style="padding:5px 5px 5px 0;"><%
								If vCate42Img(tmp) <> "" Then
									Response.Write "<img src='" & vCate42Img(tmp) & "' height='100' style='cursor:pointer;' onclick=jsImgView('" & vCate42Img(tmp) & "');>"
									Response.Write "<a href=javascript:jsDelImg('cate42Img"&l&"','cate42Imgspan"&l&"');><img src='/images/icon_delete2.gif' border='0'></a>"
								End If
							%></span>
							<input type="hidden" name="cate42Img<%=l%>" value="<%=vCate42Img(tmp)%>">
						</td>
						<% Next %>
					</tr>
					</table>
				</td>
			</tr>
			<tr>
				<th width="15%">컨텐츠(응모기간)</th>
				<td height="30" style="padding:5px 5px 5px 5px;">
					응모기간 : <input type="text" name="cate42entrysdate" value="<%=vCate42EntrySDate%>" onClick="jsPopCal('cate42entrysdate');" style="cursor:pointer;" size="10" maxlength="10" readonly>&nbsp;~&nbsp;
								<input type="text" name="cate42entryedate" value="<%=vCate42EntryEDate%>" onClick="jsPopCal('cate42entryedate');" style="cursor:pointer;" size="10" maxlength="10" readonly><br />
					당첨자발표 : <input type="text" name="cate42announdate" value="<%=vCate42AnnounDate%>" onClick="jsPopCal('cate42announdate');" style="cursor:pointer;" size="10" maxlength="10" readonly>
				</td>
			</tr>
			<tr>
				<th width="15%">테그노출</th>
				<td height="30" style="padding:5px 5px 5px 5px;">
					<input type="checkbox" name="istagview" id="istagview" value="1" onClick="jsTagView();" <%=CHKIIF(vIsTagView,"checked","")%>> 노출&nbsp;&nbsp;&nbsp;
					<span id="tagspan" style="display:<%=CHKIIF(vIsTagView,"block","none")%>;">
					* 이벤트중테그 노출기간 : <input type="text" name="tagsdate" value="<%=vTagSDate%>" onClick="jsPopCal('tagsdate');" style="cursor:pointer;" size="10" maxlength="10" readonly>&nbsp;~&nbsp;
								<input type="text" name="tagedate" value="<%=vTagEDate%>" onClick="jsPopCal('tagedate');" style="cursor:pointer;" size="10" maxlength="10" readonly>&nbsp;&nbsp;
					* 당첨자발표테그 오픈일 : <input type="text" name="tagannouncedate" value="<%=vTagAnnounceDate%>" onClick="jsPopCal('tagannouncedate');" style="cursor:pointer;" size="10" maxlength="10" readonly>
					</span>
				</td>
			</tr>

			<%' 2017-07-17 유태욱 추가 %>
			<tr>
				<th width="15%">뱃지, 상단 블릿 이달의 뱃지</th>
				<td height="30" style="padding-left:5px;">
					<input type="text" name="badgetag" value="<%=vCate42Badgetag%>" size="16" maxlength="16">
				</td>
			</tr>

			<tr>
				<th width="15%">참여(코멘트)</th>
				<td height="30" style="padding:5px 5px 5px 5px;"><textarea name="cate42entrycopy" rows="3" cols="80"><%=vCate42EntryCopy%></textarea></td>
			</tr>
			<tr>
				<th width="15%">참여(유의사항)</th>
				<td height="30" style="padding:5px 5px 5px 5px;"><textarea name="cate42notice" rows="3" cols="80"><%=vCate42Notice%></textarea></td>
			</tr>
			<tr>
				<th width="15%">컨텐츠(당첨자)</th>
				<td height="30" style="padding-left:5px;"><input type="text" name="cate42winnertxt" value="<%=vCate42WinnerTxt%>" size="60" maxlength="146">
					* 예: 텐바이텐 X nayajuri**** 고객님
				</td>
			</tr>
			<tr>
				<th width="15%">컨텐츠(당첨이름)</th>
				<td height="30" style="padding-left:5px;"><input type="text" name="cate42winnervalue" value="<%=vCate42WinnerValue%>" size="60" maxlength="146">
					* 예: 김신발
				</td>
			</tr>
			<tr>
				<th width="15%">컨텐츠(그외이름)</th>
				<td height="30" style="padding-left:5px;">
				<%
				tmp = 0
				For l=1 To 10
						tmp = tmp + 1
				%>
					<input type="text" name="cate42value<%=l%>" value="<%=vCate42Value(tmp)%>" size="15" maxlength="96">
				<%
					If l=5 Then Response.Write "<br />" End If
				Next %>
				</td>
			</tr>
			<tr>
				<th width="15%">연결배너<br />(PC 이미지)<br /><font size="1" color="silver">(이미지구분 : 21)</font></th>
				<td height="30" style="padding:5px 5px 5px 5px;">
					<input type="button" value="이미지등록" onClick="jsUploadImg('cate42pclinkbanimg','cate42pclinkbanspan','21');" /><br /><br />
					<span id="cate42pclinkbanspan" style="padding:5px 5px 5px 0;"><%
						If vCate42PCLinkBanImg <> "" Then
							Response.Write "<img src='" & vCate42PCLinkBanImg & "' height='100' style='cursor:pointer;' onclick=jsImgView('" & vCate42PCLinkBanImg & "');>"
							Response.Write "<a href=javascript:jsDelImg('cate42pclinkbanimg','cate42pclinkbanspan');><img src='/images/icon_delete2.gif' border='0'></a>"
						End If
					%></span><br />
					<input type="hidden" name="cate42pclinkbanimg" value="<%=vCate42PCLinkBanImg%>">
					<br />
					링크주소 : <input type="text" name="cate42pclinkbanurl" id="cate42pclinkbanurl" value="<%=vCate42PCLinkBanURL%>" size="60" maxlength="195">
					<br />
					<font color="#707070">
					- <font color="red"><strong>app & mobile 공용</strong></font> - <br/>
					- <span style="cursor:pointer" onClick="putLinkText('pc','event','cate42pclinkbanurl')">이벤트 링크 : /event/eventmain.asp?eventid=<font color="darkred">이벤트코드</font></span><br/>
					- <span style="cursor:pointer" onClick="putLinkText('pc','itemid','cate42pclinkbanurl')">상품코드 링크 : /shopping/category_prd.asp?itemid=<font color="darkred">상품코드 (O)</font></span><br/>
					- <span style="cursor:pointer" onClick="putLinkText('pc','brand','cate42pclinkbanurl')">브랜드아이디 링크 : /street/street_brand_sub06.asp?makerid=<font color="darkred">브랜드아이디</font></span><br/>
					</font>
				</td>
			</tr>
			<tr>
				<th width="15%">연결배너<br />(Mo 이미지)<br /><font size="1" color="silver">(이미지구분 : 22)</font></th>
				<td height="30" style="padding:5px 5px 5px 5px;">
					<input type="button" value="이미지등록" onClick="jsUploadImg('cate42molinkbanimg','cate42molinkbanspan','22');" /><br /><br />
					<span id="cate42molinkbanspan" style="padding:5px 5px 5px 0;"><%
						If vCate42MoLinkBanImg <> "" Then
							Response.Write "<img src='" & vCate42MoLinkBanImg & "' height='100' style='cursor:pointer;' onclick=jsImgView('" & vCate42MoLinkBanImg & "');>"
							Response.Write "<a href=javascript:jsDelImg('cate42molinkbanimg','cate42molinkbanspan');><img src='/images/icon_delete2.gif' border='0'></a>"
						End If
					%></span><br />
					<input type="hidden" name="cate42molinkbanimg" value="<%=vCate42MoLinkBanImg%>">
					<br />
					링크주소 : <input type="text" name="cate42molinkbanurl" id="cate42molinkbanurl" value="<%=vCate42MoLinkBanURL%>" size="60" maxlength="195">
					<br />
					<font color="#707070">
					- <font color="red"><strong>app & mobile 공용</strong></font> - <br/>
					- <span style="cursor:pointer" onClick="putLinkText('m','event','cate42molinkbanurl')">이벤트 링크 : /event/eventmain.asp?eventid=<font color="darkred">이벤트코드</font></span><br/>
					- <span style="cursor:pointer" onClick="putLinkText('m','itemid','cate42molinkbanurl')">상품코드 링크 : /category/category_itemprd.asp?itemid=<font color="darkred">상품코드 (O)</font></span><br/>
					- <span style="cursor:pointer" onClick="putLinkText('m','brand','cate42molinkbanurl')">브랜드아이디 링크 : /street/street_brand.asp?makerid=<font color="darkred">브랜드아이디</font></span><br/>
					</font>
				</td>
			</tr>
			<tr>
				<th width="15%">상품코드</th>
				<td height="30" style="padding-left:5px;">
					<input type="text" name="cate42item" value="<%=vCate42Item%>" size="10">
				</td>
			</tr>
		<% ElseIf vCate = "43" Then	'### Thing. 배경화면 %>
			<tr>
				<th width="15%">컬러코드</th>
				<td height="30" style="padding-left:5px;">
					# <input type="text" name="mo_bgcolor" value="<%=vMoBGColor%>" size="10" maxlength="6"> * #을 제외한 6자 숫자로만 입력하세요.
				</td>
			</tr>
			<tr>
				<th width="15%">컨텐츠<br />(PC 상단 이미지)<br /><font size="1" color="silver">(이미지구분 : 20)</font></th>
				<td height="30" style="padding:5px 5px 5px 5px;">
					<input type="button" value="이미지등록" onClick="jsUploadImg('cate43pcimg','cate43pcimgspan','20');" /><br /><br />
					<span id="cate43pcimgspan" style="padding:5px 5px 5px 0;"><%
						If vCate43PCImageURL <> "" Then
							Response.Write "<img src='" & vCate43PCImageURL & "' height='100' style='cursor:pointer;' onclick=jsImgView('" & vCate43PCImageURL & "');>"
							Response.Write "<a href=javascript:jsDelImg('cate43pcimg','cate43pcimgspan');><img src='/images/icon_delete2.gif' border='0'></a>"
						End If
					%></span>
					<input type="hidden" name="cate43pcimg" value="<%=vCate43PCImageURL%>"
				</td>
			</tr>
			<tr>
				<th width="15%">컨텐츠<br />(Mo 상단 이미지)<br /><font size="1" color="silver">(이미지구분 : 9)</font></th>
				<td height="30" style="padding:5px 5px 5px 5px;">
					<input type="button" value="이미지등록" onClick="jsUploadImg('cate43moimg','cate43moimgspan','9');" /><br /><br />
					<span id="cate43moimgspan" style="padding:5px 5px 5px 0;"><%
						If vCate43MoImageURL <> "" Then
							Response.Write "<img src='" & vCate43MoImageURL & "' height='100' style='cursor:pointer;' onclick=jsImgView('" & vCate43MoImageURL & "');>"
							Response.Write "<a href=javascript:jsDelImg('cate43moimg','cate43moimgspan');><img src='/images/icon_delete2.gif' border='0'></a>"
						End If
					%></span>
					<input type="hidden" name="cate43moimg" value="<%=vCate43MoImageURL%>"
				</td>
			</tr>
			<tr>
				<th width="15%">컨텐츠<br />(PC 다운로드)</th>
				<td height="70" style="padding-left:5px;">
				<%
				tmp = 0
				For l=1 To 3
						tmp = tmp + 1
				%>
					대표기종 : <input type="text" name="cate43pcdown<%=l%>" value="<%=vCate43PCDown(tmp)%>" size="30" maxlength="96">&nbsp;&nbsp;&nbsp;
					링크 : <input type="text" name="cate43pclink<%=l%>" value="<%=vCate43PCLink(tmp)%>" size="40" maxlength="96"><br />
				<%	Next %>
				</td>
			</tr>
			<tr>
				<th width="15%">컨텐츠<br />(Mo 다운로드)</th>
				<td height="70" style="padding-left:5px;">
				<%
				tmp = 0
				For l=1 To 3
						tmp = tmp + 1
				%>
					대표기종 : <input type="text" name="cate43modown<%=l%>" value="<%=vCate43MoDown(tmp)%>" size="30" maxlength="96">&nbsp;&nbsp;&nbsp;
					링크 : <input type="text" name="cate43molink<%=l%>" value="<%=vCate43MoLink(tmp)%>" size="40" maxlength="96"> * 스크립트 복사 사용.<br />
				<%	Next %>
				</td>
			</tr>
			<tr>
				<th width="15%">컨텐츠(QR)<br /><font size="1" color="silver">(이미지구분 : 10)</font></th>
				<td height="30" style="padding:5px 5px 5px 5px;">
					<input type="text" name="cate43qrimg" value="<%=vCate43QRImageURL%>" size="55"> * 어드민 QR코드 관리에 이미지 주소(http 포함)를 입력하세요.
				</td>
			</tr>
		<% ElseIf vCate = "5" Then	'### COMMA %>
			<tr>
				<th width="15%">컨텐츠<br />(PC 상단 이미지)<br /><font size="1" color="silver">(이미지구분 : 12)</font></th>
				<td height="30" style="padding:5px 5px 5px 5px;">
					<input type="button" value="이미지등록" onClick="jsUploadImg('cate5pctopimg','cate5pctopimgspan','12');" /><br /><br />
					<span id="cate5pctopimgspan" style="padding:5px 5px 5px 0;"><%
						If vCate5PCTopImageURL <> "" Then
							Response.Write "<img src='" & vCate5PCTopImageURL & "' height='100' style='cursor:pointer;' onclick=jsImgView('" & vCate5PCTopImageURL & "');>"
							Response.Write "<a href=javascript:jsDelImg('cate5pctopimg','cate5pctopimgspan');><img src='/images/icon_delete2.gif' border='0'></a>"
						End If
					%></span>
					<input type="hidden" name="cate5pctopimg" value="<%=vCate5PCTopImageURL%>"
				</td>
			</tr>
			<tr>
				<th width="15%">컨텐츠<br />(Mo 상단 이미지)<br /><font size="1" color="silver">(이미지구분 : 13)</font></th>
				<td height="30" style="padding:5px 5px 5px 5px;">
					<input type="button" value="이미지등록" onClick="jsUploadImg('cate5motopimg','cate5motopimgspan','13');" /><br /><br />
					<span id="cate5motopimgspan" style="padding:5px 5px 5px 0;"><%
						If vCate5MoTopImageURL <> "" Then
							Response.Write "<img src='" & vCate5MoTopImageURL & "' height='100' style='cursor:pointer;' onclick=jsImgView('" & vCate5MoTopImageURL & "');>"
							Response.Write "<a href=javascript:jsDelImg('cate5motopimg','cate5motopimgspan');><img src='/images/icon_delete2.gif' border='0'></a>"
						End If
					%></span>
					<input type="hidden" name="cate5motopimg" value="<%=vCate5MoTopImageURL%>"
				</td>
			</tr>
			<tr>
				<th width="15%">컨텐츠(제작자)</th>
				<td height="30" style="padding-left:5px;"><input type="text" name="cate5directer" value="<%=vCate5Directer%>" size="60" maxlength="146"></td>
			</tr>
			<tr>
				<th width="15%">컨텐츠(에디터)<br /><font size="1" color="silver">(이미지구분 : 14)</font></th>
				<td style="padding:5px 5px 5px 5px;">
				<%
				tmp = 0
				For l=1 To 5
						tmp = tmp + 1
				%>
					<%=l%>.&nbsp;&nbsp;
					<input type="button" value="이미지등록" onClick="jsUploadImg('cate5Img<%=l%>','cate5Imgspan<%=l%>','14');" /><br /><br />
					<span id="cate5Imgspan<%=l%>" style="padding:5px 5px 5px 0;"><%
						If vCate5Img(tmp) <> "" Then
							Response.Write "<img src='" & vCate5Img(tmp) & "' height='100' style='cursor:pointer;' onclick=jsImgView('" & vCate5Img(tmp) & "');>"
							Response.Write "<a href=javascript:jsDelImg('cate5Img"&l&"','cate5Imgspan"&l&"');><img src='/images/icon_delete2.gif' border='0'></a>"
						End If
					%></span>
					<input type="hidden" name="cate5Img<%=l%>" value="<%=vCate5Img(tmp)%>"><br />
					<textarea name="cate5copy<%=l%>" rows="3" cols="80"><%=vCate5Copy(tmp)%></textarea><br />
					<hr />
				<% Next %>
				</td>
			</tr>
			<tr>
				<th width="15%">연결배너<br />(PC 이미지)<br /><font size="1" color="silver">(이미지구분 : 15)</font></th>
				<td height="30" style="padding:5px 5px 5px 5px;">
					<input type="button" value="이미지등록" onClick="jsUploadImg('cate5pclinkbanimg','cate5pclinkbanspan','15');" /><br /><br />
					<span id="cate5pclinkbanspan" style="padding:5px 5px 5px 0;"><%
						If vCate5PCLinkBanImg <> "" Then
							Response.Write "<img src='" & vCate5PCLinkBanImg & "' height='100' style='cursor:pointer;' onclick=jsImgView('" & vCate5PCLinkBanImg & "');>"
							Response.Write "<a href=javascript:jsDelImg('cate5pclinkbanimg','cate5pclinkbanspan');><img src='/images/icon_delete2.gif' border='0'></a>"
						End If
					%></span><br />
					<input type="hidden" name="cate5pclinkbanimg" value="<%=vCate5PCLinkBanImg%>">
					<br />
					링크주소 : <input type="text" name="cate5pclinkbanurl" id="cate5pclinkbanurl" value="<%=vCate5PCLinkBanURL%>" size="60" maxlength="195">
					<br />
					<font color="#707070">
					- <font color="red"><strong>app & mobile 공용</strong></font> - <br/>
					- <span style="cursor:pointer" onClick="putLinkText('pc','event','cate5pclinkbanurl')">이벤트 링크 : /event/eventmain.asp?eventid=<font color="darkred">이벤트코드</font></span><br/>
					- <span style="cursor:pointer" onClick="putLinkText('pc','itemid','cate5pclinkbanurl')">상품코드 링크 : /shopping/category_prd.asp?itemid=<font color="darkred">상품코드 (O)</font></span><br/>
					- <span style="cursor:pointer" onClick="putLinkText('pc','brand','cate5pclinkbanurl')">브랜드아이디 링크 : /street/street_brand_sub06.asp?makerid=<font color="darkred">브랜드아이디</font></span><br/>
					</font>
				</td>
			</tr>
			<tr>
				<th width="15%">연결배너<br />(Mo 이미지)<br /><font size="1" color="silver">(이미지구분 : 16)</font></th>
				<td height="30" style="padding:5px 5px 5px 5px;">
					<input type="button" value="이미지등록" onClick="jsUploadImg('cate5molinkbanimg','cate5molinkbanspan','16');" /><br /><br />
					<span id="cate5molinkbanspan" style="padding:5px 5px 5px 0;"><%
						If vCate5MoLinkBanImg <> "" Then
							Response.Write "<img src='" & vCate5MoLinkBanImg & "' height='100' style='cursor:pointer;' onclick=jsImgView('" & vCate5MoLinkBanImg & "');>"
							Response.Write "<a href=javascript:jsDelImg('cate5molinkbanimg','cate5molinkbanspan');><img src='/images/icon_delete2.gif' border='0'></a>"
						End If
					%></span><br />
					<input type="hidden" name="cate5molinkbanimg" value="<%=vCate5MoLinkBanImg%>">
					<br />
					링크주소 : <input type="text" name="cate5molinkbanurl" id="cate5molinkbanurl" value="<%=vCate5MoLinkBanURL%>" size="60" maxlength="195">
					<br />
					<font color="#707070">
					- <font color="red"><strong>app & mobile 공용</strong></font> - <br/>
					- <span style="cursor:pointer" onClick="putLinkText('m','event','cate5molinkbanurl')">이벤트 링크 : /event/eventmain.asp?eventid=<font color="darkred">이벤트코드</font></span><br/>
					- <span style="cursor:pointer" onClick="putLinkText('m','itemid','cate5molinkbanurl')">상품코드 링크 : /category/category_itemprd.asp?itemid=<font color="darkred">상품코드 (O)</font></span><br/>
					- <span style="cursor:pointer" onClick="putLinkText('m','brand','cate5molinkbanurl')">브랜드아이디 링크 : /street/street_brand.asp?makerid=<font color="darkred">브랜드아이디</font></span><br/>
					</font>
				</td>
			</tr>
		<% ElseIf vCate = "6" Then	'### HOWHOW %>
			<tr>
				<th width="15%">컬러코드</th>
				<td height="30" style="padding-left:5px;">
					# <input type="text" name="mo_bgcolor" value="<%=vMoBGColor%>" size="10" maxlength="6"> * #을 제외한 6자 숫자로만 입력하세요.
				</td>
			</tr>
			<tr>
				<th width="15%">컨텐츠(동영상)</th>
				<td height="30" style="padding-left:5px;"><input type="text" name="cate6videourl" value="<%=vCate6VideoURL%>" size="60" maxlength="146"></td>
			</tr>
			<tr>
				<th width="15%">컨텐츠(에디터)<br /><font size="1" color="silver">(이미지구분 : 17)</font></th>
				<td style="padding:5px 5px 5px 5px;">
				<%
				tmp = 0
				For l=1 To 4
						tmp = tmp + 1
				%>
					<%=l%>.&nbsp;&nbsp;
					<input type="button" value="이미지등록" onClick="jsUploadImg('cate6Img<%=l%>','cate6Imgspan<%=l%>','17');" /><br /><br />
					<span id="cate6Imgspan<%=l%>" style="padding:5px 5px 5px 0;"><%
						If vCate6Img(tmp) <> "" Then
							Response.Write "<img src='" & vCate6Img(tmp) & "' height='100' style='cursor:pointer;' onclick=jsImgView('" & vCate6Img(tmp) & "');>"
							Response.Write "<a href=javascript:jsDelImg('cate6Img"&l&"','cate6Imgspan"&l&"');><img src='/images/icon_delete2.gif' border='0'></a>"
						End If
					%></span>
					<input type="hidden" name="cate6Img<%=l%>" value="<%=vCate6Img(tmp)%>"><br />
					<textarea name="cate6copy<%=l%>" rows="3" cols="80"><%=vCate6Copy(tmp)%></textarea><br />
					<hr />
				<% Next %>
				</td>
			</tr>
			<tr>
				<th width="15%">연결배너<br />(텍스트)</th>
				<td height="140" style="padding-left:5px;">
					타이틀&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;: 
						<input type="text" name="cate6bannsub" value="<%=vCate6BannSub%>" size="60" maxlength="96"><br />
					메인카피&nbsp;&nbsp;&nbsp;&nbsp;: 
						<textarea name="cate6banntitle" rows="3" cols="80"><%=vCate6BannTitle%></textarea><br /><br />
					버튼타이틀 : <input type="text" name="cate6bannbtntitle" value="<%=vCate6BannBtnTitle%>" size="60" maxlength="96"><br />
					버튼링크&nbsp;&nbsp;&nbsp;&nbsp;: <input type="text" name="cate6bannbtnlink" value="<%=vCate6BannBtnLink%>" size="60" maxlength="196"><br />
				</td>
			</tr>
		<% End If %>
	</tbody>
</table>
<table width="100%">
<tr>
	<td style="padding-top:5px;float:right;"><input type="button" style="width:100px;height:30px;" value="저 장" onClick="goSavePlay();" /></td>
</tr>
</table>
<% End If %>
</form>
<span id="tempspan"></span>
</body>
</html>
<!-- #include virtual="/lib/db/dbclose.asp" -->