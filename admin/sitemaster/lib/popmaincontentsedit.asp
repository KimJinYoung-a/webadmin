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
'           2018-08-30 최종원 : pc, 모바일 상품상세 배너에 회원구분, os구분 추가
'			2019.09.27 정태훈 : 컬쳐스테이션 이벤트DB 이전 변경 적용
'			2019.11.20 정태훈 : 이미지 삭제 추가
'###############################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/classes/sitemasterclass/main_ContentsManageCls.asp" -->
<!-- #include virtual="/admin/eventmanage/common/event_function.asp"-->
<!-- #include virtual="/lib/classes/displaycate/displaycateCls.asp"-->
<%
dim isusing, fixtype, validdate, prevDate
dim idx, poscode, reload, gubun, edid
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

	if reload="on" then
			response.write "<script>opener.location.reload(); window.close();</script>"
			dbget.close()	:	response.End
	end if

	dim oMainContents
		set oMainContents = new CMainContents
		oMainContents.FRectIdx = idx
		oMainContents.GetOneMainContents

dim oposcode, defaultMapStr, defaultXMLMapStr
	set oposcode = new CMainContentsCode
	oposcode.FRectPosCode = poscode
	if poscode<>"" then
			oposcode.GetOneContentsCode

			defaultMapStr = "<map name='Map_" +oposcode.FOneItem.FPosvarName + "'>" + VbCrlf
			defaultMapStr = defaultMapStr + VbCrlf
			defaultMapStr = defaultMapStr + "</map>"

		defaultXMLMapStr = ""
			defaultXMLMapStr = defaultXMLMapStr + "<map name='Map_" +oposcode.FOneItem.FPosvarName + "'>"+ VbCrlf
			defaultXMLMapStr = defaultXMLMapStr + VbCrlf
		defaultXMLMapStr = defaultXMLMapStr + "</map>"
	end if

dim orderidx
	if oMainContents.FOneItem.forderidx = "" then
		orderidx = 99
	else
		orderidx = oMainContents.FOneItem.forderidx
		poscode = oMainContents.FOneItem.fposcode
	end if

	If gubun = "" Then
		gubun = "index"
	End If

	edid = oMainContents.FOneItem.Fworkeruserid
	If edid = "" Then
		If idx <> "" AND idx <> "0" Then
			edid = session("ssBctId")
		End If
	End If

	'// 컬쳐스테이션 불러오기
	Dim cultureContents, SqlStr
	Dim cultureEcode ,	cultureEtype ,cultureEname ,cultureEcomment , cultureEimagelist



	If culturecode<>"" Then
		sqlStr = "SELECT e.evt_code, d.eventtype_pc as evt_type, e.evt_name, e.evt_subcopyk, d.evt_mainimg" + vbcrlf
		sqlStr = sqlStr & " from db_event.dbo.tbl_event as e" + vbcrlf
		sqlStr = sqlStr & " LEFT JOIN [db_event].[dbo].[tbl_event_display] as d on d.evt_code=e.evt_code"
		sqlStr = sqlStr & " where e.evt_using='Y'" + vbcrlf
		sqlStr = sqlStr & " and e.evt_code="& culturecode		

		rsget.Open SqlStr, dbget, 1
		if Not rsget.Eof then
			cultureEcode		= rsget("evt_code")
			cultureEtype		= rsget("evt_type")
			cultureEname		= rsget("evt_name")
			cultureEcomment		= rsget("evt_subcopyk")
			'cultureEimagelist	= webImgUrl &"/culturestation/2009/list/" & rsget("image_list")
			cultureEimagelist	= rsget("evt_mainimg")
		end if
		rsget.close
	End If

'// 특정 코드에 링크텍스트 추가(IMG ALT 값 등)
dim IsLinkTextNeed
	IsLinkTextNeed = (InStr(",630,642,659,673,674,675,687,", ("," & poscode & ",")) > 0)

'//상품상세 배너 카테고리 설정
	dim cDisp, cateIndex, cateCodeArr, cateNameArr(), cateIdx, categoryOptions
	categoryOptions = oMainContents.FOneItem.FcategoryOptions
	cateCodeArr = split(categoryOptions, ",")
	redim preserve cateNameArr(ubound(cateCodeArr))

	SET cDisp = New cDispCate
	cDisp.FCurrPage = 1
	cDisp.FPageSize = 2000
	cDisp.FRectDepth = 1
	cDisp.GetDispCateList()

	For cateIndex = 0 To cDisp.FResultCount-1
		for cateIdx = 0 to ubound(cateCodeArr) - 1
			if Cint(cDisp.FItemList(cateIndex).FCateCode) = Cint(cateCodeArr(cateIdx)) then
				cateNameArr(cateIdx) = cDisp.FItemList(cateIndex).FCateName
			end if
		next
	next
	'response.write cateNameArr
%>
<script type="text/javascript" src="/js/jquery-1.7.2.min.js"></script>
<script type="text/javascript" src="/js/jsCal/js/jscal2.js"></script>
<script type="text/javascript" src="/js/jsCal/js/lang/ko.js"></script>
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/jscal2.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/border-radius.css" />
<script type="text/javascript">
<%
	'ecode 컬쳐스테이션이벤트id
	'maincopy 메인제목
	'subcopy 추가 코멘트내용
	'linktext3  내용 (설명)
	'xbtncolor 0/1	구분선택
	'file1 이미지 --  이미지 명만 넣고 조합해서 써야할듯
%>
	<% if culturecode <> "" then %>
	$(function(){
		var gubuncode = "<%=cultureEtype%>";
		var frm = document.frmcontents;
			frm.ecode.value = "<%=cultureEcode%>";
			frm.maincopy.value = "<%=cultureEname%>";
			frm.subcopy.value = "<%=cultureEcomment%>";
			if (gubuncode == "0"){
				frm.xbtncolor[0].value = "0";
				frm.xbtncolor[0].checked = true;
			}else{
				frm.xbtncolor[1].value = "1";
				frm.xbtncolor[1].checked = true;
			}
			frm.linkurl.value = "/culturestation/culturestation_event.asp?evt_code=<%=cultureEcode%>";
	});
	<% end if %>

	function SaveMainContents(frm){
			if (frm.poscode.value.length<1){
					alert('구분을 먼저 선택 하세요.');
					frm.poscode.focus();
					return;
			}

			if (frm.linkurl.value.length<1 && !$("#couponRadioBtn").is(':checked') && !$("#popupBnrBtn").is(':checked')){
					alert('링크 값을 입력 하세요.');
					frm.linkurl.focus();
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
		<% if poscode <> "562" and poscode <> "561" then  %>
		if (!frm.altname.value){
			alert('alt값을 입력 하세요.');
			frm.altname.focus();
			return;
		}
		<% end if %>

			var vstartdate = new Date(frm.startdate.value.substr(0,4), (1*frm.startdate.value.substr(5,2))-1, frm.startdate.value.substr(8,2));
			var venddate = new Date(frm.enddate.value.substr(0,4), (1*frm.enddate.value.substr(5,2))-1, frm.enddate.value.substr(8,2));

			if (vstartdate>venddate){
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

	function putLinkText(key) {
		var frm = document.frmcontents;
		switch(key) {
			case 'search':
				frm.linkurl.value='/search/search_item.asp?rect=검색어';
				break;
			case 'event':
				frm.linkurl.value='/event/eventmain.asp?eventid=이벤트번호';
				break;
			case 'itemid':
				frm.linkurl.value='/shopping/category_prd.asp?itemid=상품코드';
				break;
			case 'category':
				frm.linkurl.value='/shopping/category_list.asp?disp=카테고리';
				break;
			case 'brand':
				frm.linkurl.value='/street/street_brand.asp?makerid=브랜드아이디';
				break;
			case 'showbanner':
				frm.linkurl.value='/showbanner/show_view.asp?showidx=쇼배너아이디';
				break;
			case 'culture':
				frm.linkurl.value='/culturestation/culturestation_event.asp?evt_code=이벤트아이디';
				break;
			case 'ground':
				frm.linkurl.value='/play/playGround.asp?idx=그라운드번호&contentsidx=컨텐츠번호';
				break;
			case 'styleplus':
				frm.linkurl.value='/play/playStylePlus.asp?idx=스타일플러스번호&contentsidx=컨텐츠번호';
				break;
			case 'fingers':
				frm.linkurl.value='/play/playDesignFingers.asp?idx=핑거스번호&contentsidx=컨텐츠번호';
				break;
			case 'tepisode':
				frm.linkurl.value='/play/playTEpisode.asp?idx=티에피소드번호&contentsidx=컨텐츠번호';
				break;
			case 'gift':
				frm.linkurl.value='/gift/gifttalk/';
				break;
			case 'wish':
				frm.linkurl.value='/wish/index.asp';
				break;
			case 'hitchhiker':
				frm.linkurl.value='/hitchhiker/';
				break;
			case 'giftcard':
				frm.linkurl.value='/giftcard/';
				break;
			case 'coupon':
				frm.linkurl.value='/my10x10/couponbook.asp';
				break;
		}
	}

	function cultureloadpop(){
		winLast = window.open('pop_culturelist.asp?gubun=<%=gubun%>&poscode=<%=poscode%>&pidx=<%=idx%>','pLast','width=1200,height=600, scrollbars=yes')
		winLast.focus();
	}

	function fnSelectBannerType(bannertype){
		switch (bannertype) {
			case 1 :
				$("#bnimg3").hide();
				$("#bnalt3").hide();
				$("#bnimg2").hide();
				$("#bnalt2").hide();
				$("#bnbg1").hide();
				$("#bnbg2").hide();
				$("#bnlink2").hide();
				$("#bnlink3").hide();
				break;
			case 2 :
				$("#bnbg1").show();
				$("#bnbg2").show();
				$("#bnimg2").show();
				$("#bnalt2").show();
				$("#bnlink2").show();
				$("#bnlink3").hide();
				$("#bnimg3").hide();
				$("#bnalt3").hide();
				break;
			case 3 :
				$("#bnbg1").show();
				$("#bnbg2").show();
				$("#bnimg2").show();
				$("#bnalt2").show();
				$("#bnlink2").show();
				$("#bnimg3").show();
				$("#bnalt3").show();
				$("#bnlink3").show();
				break;
		}
	}

	$(function() {
		$('input:radio[name="etctag"]').click(function(){
			if($('input:radio[name="etctag"]:checked').val()==8 || $('input:radio[name="etctag"]:checked').val()==9)
			{
				alert('이벤트 코드를 입력 해주세요');
				$("#saleinfo2").focus();
			}
		});
	});

var selectedCategoryArr = [];
	<%
		if categoryOptions <> "" then
			for cateIdx = 0 to ubound(cateNameArr) - 1
	%>
			var tempObj = {
				categoryCode: '<%=cateCodeArr(cateIdx)%>',
				categoryName: '<%=cateNameArr(cateIdx)%>'
			}
			selectedCategoryArr.push(tempObj);
	<%
			next
		end if
	%>
$(function(){
	dispSelectedCateNames()
})
function addCategory(){
	var cateSelectBox = document.frmcontents.categoryCode;

	var selectedObj;
	var selectedCcode, selectedCname;

	selectedCcode = cateSelectBox.value;
	selectedCname = cateSelectBox.options[cateSelectBox.selectedIndex].text.replace(" ","");

	if(chkCategory(selectedCcode))return false;

	selectedObj = {
		categoryCode: selectedCcode,
		categoryName: selectedCname
	}
	selectedCategoryArr.push(selectedObj);

	dispSelectedCateNames();
	setCategoryValues();
}
function chkCategory(selectedCcode){
	var result = false;
	selectedCategoryArr.forEach(function(item, index){
		if(item.categoryCode == selectedCcode){
			alert("이미 추가돼있는 카테고리입니다.");
			result = true;
			return false;
		}
	});
	return result;
}
function dispSelectedCateNames(){

	var selectedCategoryNamesText="";

	selectedCategoryArr.forEach(function(item, index){
		selectedCategoryNamesText = selectedCategoryNamesText + "<span onclick='subCateObj("+item.categoryCode+")'>"+item.categoryName+", </span>";
	});
	$("#categoryDisplay").html(selectedCategoryNamesText);
}
function setCategoryValues(){

	var selectedCategoryCodes="";
	selectedCategoryArr.forEach(function(item, index){
		selectedCategoryCodes = selectedCategoryCodes + item.categoryCode+",";
	});
	document.frmcontents.categoryOptions.value = selectedCategoryCodes;
}
function subCateObj(selectedCode){
	selectedCategoryArr = selectedCategoryArr.filter(function(obj){
		return obj.categoryCode != selectedCode;
	});
	dispSelectedCateNames();
	setCategoryValues();
}
function chkWhiteSpace(obj){
	obj.value = obj.value.trim();
}

function fnDeleteImage(imgnum){
	$("#img"+imgnum).attr("src","");
	$("#file"+imgnum).val("");
	$("#imgurl"+imgnum).html("");
	$("#dfile"+imgnum).val("Y");
}
</script>

<table width="100%" cellpadding="2" cellspacing="1" class="a" bgcolor="#3d3d3d">
<form name="frmcontents" method="post" action="<%=uploadUrl%>/linkweb/doMainContentsRegNew.asp" onsubmit="return false;" enctype="multipart/form-data">
<input type="hidden" name="dfile1" id="dfile1">
<input type="hidden" name="dfile2" id="dfile2">
<input type="hidden" name="dfile3" id="dfile3">
	<tr bgcolor="#FFFFFF">
		<td width="150" bgcolor="#DDDDFF">Idx</td>
		<td>
			<% if oMainContents.FOneItem.Fidx<>"" then %>
				<%= oMainContents.FOneItem.Fidx %>
				<input type="hidden" name="idx" value="<%= oMainContents.FOneItem.Fidx %>">
			<% else %>
				<% '?? %>
			<% end if %>
		</td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td width="150" bgcolor="#DDDDFF">그룹구분</td>
		<td>
			<% if oMainContents.FOneItem.Fidx<>"" then %>
				<%= oMainContents.FOneItem.Fgubun %>
				<input type="hidden" name="gubun" value="<%= oMainContents.FOneItem.Fgubun %>">
			<% else %>
				<% call DrawGroupGubunCombo("gubun", gubun, "onChange='ChangeGroupGubun(this);'") %>
			<% end if %>
		</td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td width="150" bgcolor="#DDDDFF">구분명</td>
		<td>
			<% IF gubun <> "" Then %>
				<% if oMainContents.FOneItem.Fidx<>"" then %>
					<%= oMainContents.FOneItem.Fposname %> (<%= oMainContents.FOneItem.Fposcode %>)
					<input type="hidden" name="poscode" value="<%= oMainContents.FOneItem.Fposcode %>">
				<% else %>
					<% call DrawMainPosCodeCombo("poscode", poscode, "onChange='ChangeGubun(this);'", gubun) %>
				<% end if %>
			<% Else %>
				<font color="red">그룹구분을 먼저 선택하세요</font>
			<% End If %>
			<% If poscode = "714" Then %>
				<%'//[2018] 컬쳐스테이션%>
				&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<span><a href="" onclick="cultureloadpop();return false;">불러오기</a></span>
			<% End If %>
		</td>
	</tr>
<!-- ==================================쿠폰배너 추가 2019-08-21 ======================================-->
	<% If poscode = "716" or poscode = "715"  or poscode = "708" or poscode = "733" or poscode = "734" or poscode = "735" or poscode = "738" or poscode = "739" or poscode = "730" or poscode = "729" or poscode = "728" or poscode = "707" then %>
	<script>
	$(function(){
	<% if oMainContents.FOneItem.Fbannertype = 1 then %>
		setCouponRow(1)
	<% elseif oMainContents.FOneItem.Fbannertype = 2 then%>	
		setCouponRow(2)
		setAddButton()
	<% else %>		
		setCouponRow(3)
		setAddButton()
	<% end if %>
	
	})
	function setCouponRow(v){
		if(v == 1){			
			$(".coupon-row").css("display", "none")
			$(".lyr-row").css("display", "none")			
			$(".add-btn").css("display", "none")			
		}else if(v == 2){
			$(".coupon-row").show()
			$(".lyr-row").show()
			$(".add-btn").show()
		}else{
			$(".coupon-row").css("display", "none")
			$(".lyr-row").show()
			$(".add-btn").show()
		}
	}
	function setAddButton(){
		var isChk = $('input:checkbox[id="btnFlag"]').is(':checked')
		if(isChk){
			$(".btn-row").css("display", "")
		}else{
			$(".btn-row").css("display", "none")
		}
	}
	function jsLastEvent(){	
		var winLast = window.open('pop_coupon_list.asp','pLast','width=550,height=600, scrollbars=yes')
		winLast.focus();
	}	
	</script>
	<% if poscode = "716" then %>
	<tr bgcolor="#FFFFFF">
		<td width="150" bgcolor="#DDDDFF">채널</td>
		<td>
			<input type="radio" name="etctext" value="1" <%=ChkIIF(oMainContents.FOneItem.Fetctext = "1" or oMainContents.FOneItem.Fetctext ="", "checked", "")%>>앱, 모웹
			<input type="radio" name="etctext" value="2" <%=ChkIIF(oMainContents.FOneItem.Fetctext = "2", "checked", "")%>>앱
			<input type="radio" name="etctext" value="3" <%=ChkIIF(oMainContents.FOneItem.Fetctext = "3", "checked", "")%>>모웹
		</td>
	</tr>	
	<% end if %>
	<tr bgcolor="#FFFFFF">
		<td width="150" bgcolor="#DDDDFF">배너구분</td>
		<td>
			<input type="radio" name="bannerType" onclick="setCouponRow(1)" value="1" <%=ChkIIF(oMainContents.FOneItem.Fbannertype = 1 or oMainContents.FOneItem.Fbannertype ="", "checked", "")%>>링크 배너
			<input type="radio" name="bannerType" onclick="setCouponRow(2)" value="2" <%=ChkIIF(oMainContents.FOneItem.Fbannertype = 2, "checked", "")%> id="couponRadioBtn">쿠폰배너
			<input type="radio" name="bannerType" onclick="setCouponRow(3)" value="3" <%=ChkIIF(oMainContents.FOneItem.Fbannertype = 3, "checked", "")%> id="popupBnrBtn">팝업 배너
		</td>
	</tr>	
	<tr bgcolor="#FFFFFF" class="coupon-row" style="display:none">
		<td width="150" bgcolor="#DDDDFF">쿠폰번호</td>
		<td>
			<input type="number" name="couponidx" id="couponidx" value="<%=oMainContents.FOneItem.Fcouponidx%>">
			<button type="button" onclick="jsLastEvent()">쿠폰찾기</button>
		</td>
	</tr>
	<tr bgcolor="#FFFFFF" class="lyr-row">
		<td width="150" bgcolor="#DDDDFF">레이어 팝업 카피</td>
		<td>
			<input type="text" name="maincopy" value="<%=oMainContents.FOneItem.Fmaincopy%>" size="80" maxlength="200" class="text" /><br/>
		</td>
	</tr>
	<tr bgcolor="#FFFFFF" class="lyr-row">
		<td width="150" bgcolor="#DDDDFF">레이어 팝업 서브카피</td>
		<td>
			<input type="text" name="maincopy2" value="<%=oMainContents.FOneItem.Fmaincopy2%>" size="80" maxlength="200" class="text" /><br/>
		</td>
	</tr>
	<tr bgcolor="#FFFFFF"  class="add-btn">
		<td width="150" bgcolor="#DDDDFF">버튼 추가</td>
		<td>
			<input type="checkbox" name="etctag" id="btnFlag" onclick="setAddButton()" value="1" <%=ChkIIF(oMainContents.FOneItem.Fetctag = 1, "checked", "")%>>버튼 추가
		</td>
	</tr>			
	<tr bgcolor="#FFFFFF" class="btn-row" style="display:none">
		<td width="150" bgcolor="#DDDDFF">버튼 카피</td>
		<td>
			<input type="text" name="subcopy" value="<%=oMainContents.FOneItem.Fsubcopy%>" size="80" maxlength="200" class="text" /><br/>
		</td>
	</tr>
	<tr bgcolor="#FFFFFF" class="btn-row" style="display:none">
		<td width="150" bgcolor="#DDDDFF">버튼 렌딩url</td>
		<td>
			<input type="text" name="linkurl2" value="<%=oMainContents.FOneItem.Flinkurl2%>" size="80" maxlength="200" class="text" /><br/>
		</td>
	</tr>
	<% end if %>
	<% If poscode = "707" or poscode = "708" or poscode = "715" or poscode = "716" or poscode = "725" or poscode = "728" or poscode = "732" or poscode = "734" or poscode = "739" Then %>
	<%'//[2018] 상품상세 광고배너, [2018] 상품상세 광고 광고배너, [2018] pc상품상세 광고배너(쿠폰), [2018] mo상품상세 상단 (쿠폰)배너, 카테고리 좌측 리스트 배너, 이벤트 상세배너, [2018] app 상품상세 광고 광고배너 %>
		<tr bgcolor="#FFFFFF">
			<td width="150" bgcolor="#DDDDFF">대상</td>
			<td>
				<select name="targetType" class="formSlt">
					<option value="" <%=chkIIF(oMainContents.FOneItem.FtargetType="","selected","")%>>모든고객</option>
					<option value="0" <%=chkIIF(oMainContents.FOneItem.FtargetType="0","selected","")%>>white</option>
					<option value="1" <%=chkIIF(oMainContents.FOneItem.FtargetType="1","selected","")%>>red</option>
					<option value="2" <%=chkIIF(oMainContents.FOneItem.FtargetType="2","selected","")%>>vip</option>
					<option value="3" <%=chkIIF(oMainContents.FOneItem.FtargetType="3","selected","")%>>vip gold</option>
					<option value="4" <%=chkIIF(oMainContents.FOneItem.FtargetType="4","selected","")%>>vvip</option>
				</select>
				<span>
					<span style="padding-left: 120px">카테고리</span>
					<select name="categoryCode" class="formSlt">
						<% For cateIndex=0 To cDisp.FResultCount-1 %>
							<option value="<%=cDisp.FItemList(cateIndex).FCateCode%>"><%=" "&cDisp.FItemList(cateIndex).FCateName%></option>
						<% next %>
					</select>
					<button type="button" onclick="addCategory();">선택</button>
				</span>
				<br/>
				<span style="color:darkred">※ 카테고리를 선택하지 않으시면 전체 카테고리로 적용됩니다.</span>
			</td>
		</tr>
		<tr bgcolor="#FFFFFF">
			<td width="150" bgcolor="#DDDDFF">적용 카테고리</td>
			<td id="categoryDisplay"></td>
			<input type="hidden" name="categoryOptions" value="<%=categoryOptions%>">
		</tr>
	<% End If %>

	<% If poscode="729" or poscode="730" or poscode="733" or poscode="735" or poscode="738" Then %>
	<%'//[2018] 상품상세 광고배너(비회원), [2018] 상품상세 광고광고배너(비회원) %>
		<tr bgcolor="#FFFFFF">
			<td width="150" bgcolor="#DDDDFF">대상</td>
			<td>
				<select name="targetType" class="formSlt">
					<option value="99" selected>비회원</option>
				</select>
				<span>
					<span style="padding-left: 120px">카테고리</span>
					<select name="categoryCode" class="formSlt">
						<% For cateIndex=0 To cDisp.FResultCount-1 %>
							<option value="<%=cDisp.FItemList(cateIndex).FCateCode%>"><%=" "&cDisp.FItemList(cateIndex).FCateName%></option>
						<% next %>
					</select>
					<button type="button" onclick="addCategory();">선택</button>
				</span>
				<br/>
				<span style="color:darkred">※ 카테고리를 선택하지 않으시면 전체 카테고리로 적용됩니다.</span>
			</td>
		</tr>
		<tr bgcolor="#FFFFFF">
			<td width="150" bgcolor="#DDDDFF">적용 카테고리</td>
			<td id="categoryDisplay"></td>
			<input type="hidden" name="categoryOptions" value="<%=categoryOptions%>">
		</tr>
	<% End If %>

	<% If poscode = "708" or poscode = "716" or poscode = "739" Then %>
	<%'//[2018] 상품상세 광고 광고배너, [2018] 상품상세 광고 광고배너(테스트), [2018] app 상품상세 광고 광고배너%>
		<tr bgcolor="#FFFFFF">
			<td width="150" bgcolor="#DDDDFF">운영체제</td>
			<td>
				<select name="targetOS" class="formSlt">
					<option value="" <%=chkIIF(oMainContents.FOneItem.FtargetOS="","selected","")%>>전체</option>
					<option value="I" <%=chkIIF(oMainContents.FOneItem.FtargetOS="I","selected","")%>>iOS</option>
					<option value="A" <%=chkIIF(oMainContents.FOneItem.FtargetOS="A","selected","")%>>안드로이드</option>
				</select>
			</td>
		</tr>
	<% End If %>

	<tr bgcolor="#FFFFFF">
		<td width="150" bgcolor="#DDDDFF">링크구분</td>
		<td>
			<% IF gubun <> "" Then %>
				<% if oMainContents.FOneItem.Fidx<>"" then %>
					<%= oMainContents.FOneItem.getlinktypeName %>
					<input type="hidden" name="linktype" value="<%= oMainContents.FOneItem.Flinktype %>">
				<% else %>
					<% if poscode<>"" then %>
						<%= oposcode.FOneItem.getlinktypeName %>
						<input type="hidden" name="linktype" value="<%= oposcode.FOneItem.Flinktype %>">
					<% else %>
						<font color="red">구분을 먼저 선택하세요</font>
					<% end if %>
				<% end if %>
			<% Else %>
				<font color="red">그룹구분을 먼저 선택하세요</font>
			<% End If %>
		</td>
	</tr>

	<tr bgcolor="#FFFFFF">
		<td width="150" bgcolor="#DDDDFF">적용구분(반영주기)</td>
		<td>
			<% IF gubun <> "" Then %>
				<% if oMainContents.FOneItem.Fidx<>"" then %>
					<%= oMainContents.FOneItem.getfixtypeName %>
					<input type="hidden" name="fixtype" value="<%= oMainContents.FOneItem.Ffixtype %>">
				<% else %>
					<% if poscode<>"" then %>
						<%= oposcode.FOneItem.getfixtypeName %>
						<input type="hidden" name="fixtype" value="<%= oposcode.FOneItem.Ffixtype %>">
					<% else %>
						<font color="red">구분을 먼저 선택하세요</font>
					<% end if %>
				<% end if %>
			<% Else %>
				<font color="red">그룹구분을 먼저 선택하세요</font>
			<% End If %>
		</td>
	</tr>

	<tr bgcolor="#FFFFFF">
		<td width="150" bgcolor="#DDDDFF">우선순위</td>
		<td>
			<% IF gubun <> "" Then %>
				<% if oMainContents.FOneItem.Fidx<>"" then %>
					<input type="text" name="orderidx" size=5 value="<%= orderidx %>" class="text" />
				<% else %>
					<% if poscode<>"" then %>
						<input type="text" name="orderidx" size=5 value="<%= orderidx %>" class="text" />
					<% else %>
						<font color="red">구분을 먼저 선택하세요</font>
					<% end if %>
				<% end if %>
			<% Else %>
				<font color="red">그룹구분을 먼저 선택하세요</font>
			<% End If %>
		</td>
	</tr>


	<% If poscode = "727" Then %>
	<%'// 검색 상담 마케팅 배너(오른쪽) %>
		<tr bgcolor="#FFFFFF">
			<td width="150" bgcolor="#DDDDFF">검색 키워드(필수)</td>
			<td><textarea name="itemDesc" class="textarea" style="width:100%;height:80px;"><%= oMainContents.FOneItem.FitemDesc %></textarea></td>
		</tr>
	<% Else %>
		<tr bgcolor="#FFFFFF">
			<td width="150" bgcolor="#DDDDFF">작업 요청사항</td>
			<td><textarea name="itemDesc" class="textarea" style="width:100%;height:80px;"><%= oMainContents.FOneItem.FitemDesc %></textarea></td>
		</tr>
	<% End If %>



	<% If poscode = "706" or poscode="720" or poscode="722" or poscode="723" or poscode="724" or poscode="731" Then %>
	<%'// [2018] PC 헤더 최상단 띠배너, [2018] 메인빅이벤트배너1~3, 메인컨텐츠 상단배너, 메인컨텐츠 하단배너, 로그인배너, 모바일 로그인 배너 %>
		<tr bgcolor="#FFFFFF">
			<td width="150" bgcolor="#DDDDFF">배너 타입</td>
			<td>
				<input type="radio" name="bannertype" value="1"<% If oMainContents.FOneItem.Fbannertype="1" Or oMainContents.FOneItem.Fbannertype="" Then Response.write " checked" %> onclick="fnSelectBannerType(1);">1개&nbsp;&nbsp;
				<input type="radio" name="bannertype" value="2"<% If oMainContents.FOneItem.Fbannertype="2" Then Response.write " checked" %> onclick="fnSelectBannerType(2);">2개&nbsp;&nbsp;
				<%'// 메인컨텐츠 상,하단 배너, 로그인 배너는 3개 사용하지 않는다. %>
				<% If not(poscode="722" or poscode="723" or poscode="724" or poscode="731") Then %>
					<input type="radio" name="bannertype" value="3"<% If oMainContents.FOneItem.Fbannertype="3" Then Response.write " checked" %> onclick="fnSelectBannerType(3);">3개
				<% End If %>
			</td>
		</tr>
	<% End If %>

	<%
		'링크 텍스트 여부 확인
		dim chkText: chkText="N"
		IF gubun<>"" Then
			if oMainContents.FOneItem.Fidx<>"" then
				if oMainContents.FOneItem.FLinkType="T" then
					chkText="Y"
				End If
			elseif poscode<>"" then
				if oposcode.FOneItem.FLinkType="T" then
					chkText="Y"
				End If
			end if
		end if

		'2013/09/28 김진영 추가 poscode 얻기
		If oMainContents.FResultCount > 0 Then
			Dim oSQL
			oSQL = " SELECT poscode FROM [db_sitemaster].[dbo].tbl_main_contents where idx = '"&oMainContents.FOneItem.Fidx&"'  "
			rsget.open oSQL, dbget, 1
			poscode = rsget("poscode")
			rsget.close
		End If
	%>

	<% IF chkText="Y" or (IsLinkTextNeed = True) then %>
		<tr bgcolor="#FFFFFF">
			<td width="150" bgcolor="#DDDDFF"><%=chkIIF(poscode="630" or poscode="687","배경색","링크 텍스트")%></td>
			<td><input type="text" name="linkText" value="<%= oMainContents.FOneItem.FlinkText %>" size="32" maxlength="64" class="text" /> </td>
		</tr>

		<% if poscode="630" or poscode="687" then %>
			<tr bgcolor="#FFFFFF">
				<td width="150" bgcolor="#DDDDFF">텐바이텐 로고 형태</td>
				<td>
					<label><input type="radio" name="linkText2" value="wht" <%=chkIIF(oMainContents.FOneItem.FlinkText2="wht" or oMainContents.FOneItem.FlinkText2="","checked","")%> />화이트</label>
					<label><input type="radio" name="linkText2" value="red" <%=chkIIF(oMainContents.FOneItem.FlinkText2="red","checked","")%> />레드</label>
				</td>
			</tr>
			<tr bgcolor="#FFFFFF">
				<td width="150" bgcolor="#DDDDFF">배너 형식</td>
				<td>
					<label><input type="radio" name="linkText4" value="fix" <%=chkIIF(oMainContents.FOneItem.FlinkText4="fix" or oMainContents.FOneItem.FlinkText4="","checked","")%> />FIXED</label>
					<label><input type="radio" name="linkText4" value="wide" <%=chkIIF(oMainContents.FOneItem.FlinkText4="wide","checked","")%> />WIDE</label>
				</td>
			</tr>
		<% else %>
			<tr bgcolor="#FFFFFF">
				<td width="150" bgcolor="#DDDDFF">추가 텍스트 #1 (선택)</td>
				<td><input type="text" name="linkText2" value="<%= oMainContents.FOneItem.FlinkText2 %>" size="40" maxlength="128" class="text" /> </td>
			</tr>
			<tr bgcolor="#FFFFFF">
				<td width="150" bgcolor="#DDDDFF">추가 텍스트 #2 (선택)</td>
				<td><textarea name="linkText3" style="width:100%;height:60px;" class="text"><%= oMainContents.FOneItem.FlinkText3 %></textarea></td>
			</tr>
		<% end if %>
	<% end if %>

	<% if chkText<>"Y" then %>
		<% If poscode="688" Then %>
		<%'// [2015]라운드배너 %>
			<tr bgcolor="#FFFFFF">
				<td width="150" bgcolor="#DDDDFF">상단 타이틀(bold)</td>
				<td><input type="text" name="linkText2" value="<%= oMainContents.FOneItem.FlinkText2 %>" size="40" maxlength="128" class="text" /> </td>
			</tr>
			<tr bgcolor="#FFFFFF">
				<td width="150" bgcolor="#DDDDFF">하단 상품설명</td>
				<td><input type="text" name="linkText3" value="<%= oMainContents.FOneItem.FlinkText3 %>" size="40" maxlength="128" class="text" /></td>
			</tr>
			<tr bgcolor="#FFFFFF">
				<td width="150" bgcolor="#DDDDFF">할인율</td>
				<td><input type="text" name="linkText4" value="<%= oMainContents.FOneItem.FlinkText4 %>" size="40" maxlength="128" class="text" />
					<br>※ 할인율 작성시 하단 상품설명대신 할인율이 나옴
				</td>
			</tr>
		<% End If %>

		<% If poscode="689" Then %>
		<%'// [2015]JUST1DAY or 주말특가 %>
			<tr bgcolor="#FFFFFF">
				<td width="150" bgcolor="#DDDDFF">타이틀명</td>
				<td><input type="text" name="linkText2" value="<%= oMainContents.FOneItem.FlinkText2 %>" size="40" maxlength="128" class="text" />
					<br />※ 입력 안하면 기본값인 Just1Day나 주말특가 나옴<br/>※ 연휴특가 로 입력하면 배경색이 입혀진 연휴특가 글자가 출력됨.
				</td>
			</tr>
			<tr bgcolor="#FFFFFF">
				<td width="150" bgcolor="#DDDDFF">상세설명</td>
				<td><input type="text" name="linkText3" value="<%= oMainContents.FOneItem.FlinkText3 %>" size="40" maxlength="128" class="text" /></td>
			</tr>
		<% End If %>

		<% If poscode="690" Or poscode="691" Or poscode="692" Or poscode="693" Or poscode="699" Then %>
		<% '// [2015]상단배너2단#1, [2015]상단배너2단#2, [2015]상단배너2단#3, [2015]멀티왼쪽배너 %>
			<tr bgcolor="#FFFFFF">
				<td width="150" bgcolor="#DDDDFF">상단 타이틀(bold)</td>
				<td><input type="text" name="linkText2" value="<%= oMainContents.FOneItem.FlinkText2 %>" size="40" maxlength="128" class="text" /> </td>
			</tr>
			<tr bgcolor="#FFFFFF">
				<td width="150" bgcolor="#DDDDFF">하단 상품설명</td>
				<td><input type="text" name="linkText3" value="<%= oMainContents.FOneItem.FlinkText3 %>" size="40" maxlength="128" class="text" /></td>
			</tr>
		<% End If %>

		<% If poscode = "710" Then %>
		<%'// 2018 메인 롤링 %>
			<tr bgcolor="#FFFFFF">
				<td width="150" bgcolor="#DDDDFF">배경색</td>
				<td>
					좌 : # <input type="text" name="linkText" value="<%= oMainContents.FOneItem.FlinkText %>" size="20" maxlength="6" class="text" /><br/>
					우 : # <input type="text" name="bgcode" value="<%=oMainContents.FOneItem.Fbgcode%>" size="20" maxlength="6" class="text">
				</td>
			</tr>
			<tr bgcolor="#FFFFFF">
				<td width="150" bgcolor="#DDDDFF">배너 형식</td>
				<td>
					<label><input type="radio" name="linkText4" value="fix" <%=chkIIF(oMainContents.FOneItem.FlinkText4="fix" or oMainContents.FOneItem.FlinkText4="","checked","")%> />FIXED</label>
					<label><input type="radio" name="linkText4" value="wide" <%=chkIIF(oMainContents.FOneItem.FlinkText4="wide","checked","")%> />WIDE</label>
				</td>
			</tr>
			<tr bgcolor="#FFFFFF">
				<td width="150" bgcolor="#DDDDFF">폰트컬러선택</td>
				<td>
					<input type="radio" name="xbtncolor" value="0" <%=chkiif(oMainContents.FOneItem.Fxbtncolor = "" Or oMainContents.FOneItem.Fxbtncolor = "0" ,"checked","")%>/> : black
					<input type="radio" name="xbtncolor" value="1" <%=chkiif(oMainContents.FOneItem.Fxbtncolor = "1" ,"checked","")%>/> : white
				</td>
			</tr>
			<tr bgcolor="#FFFFFF">
				<td width="150" bgcolor="#DDDDFF">메인카피</td>
				<td>
					<input type="text" name="maincopy" value="<%=oMainContents.FOneItem.Fmaincopy%>" size="50" maxlength="25" class="text" /><br/>
					<input type="text" name="maincopy2" value="<%=oMainContents.FOneItem.Fmaincopy2%>" size="80" maxlength="60" class="text" />
				</td>
			</tr>
			<tr bgcolor="#FFFFFF">
				<td width="150" bgcolor="#DDDDFF">서브카피</td>
				<td><input type="text" name="subcopy" value="<%=oMainContents.FOneItem.Fsubcopy%>" size="80" maxlength="50" class="text" /></td>
			</tr>
			<tr bgcolor="#FFFFFF">
				<td width="150" bgcolor="#DDDDFF">태그</td>
				<td>
					<input type="radio" name="etctag" value="1" <%=chkiif(oMainContents.FOneItem.Fetctag="1" Or oMainContents.FOneItem.Fetctag="","checked","")%>> 없음
					<input type="radio" name="etctag" value="2" <%=chkiif(oMainContents.FOneItem.Fetctag="2","checked","")%>> 할인
					<input type="radio" name="etctag" value="3" <%=chkiif(oMainContents.FOneItem.Fetctag="3","checked","")%>> 쿠폰 <br/>
					<input type="radio" name="etctag" value="4" <%=chkiif(oMainContents.FOneItem.Fetctag="4","checked","")%>> GIFT
					<input type="radio" name="etctag" value="5" <%=chkiif(oMainContents.FOneItem.Fetctag="5","checked","")%>> 1+1
					<input type="radio" name="etctag" value="6" <%=chkiif(oMainContents.FOneItem.Fetctag="6","checked","")%>> 런칭
					<input type="radio" name="etctag" value="7" <%=chkiif(oMainContents.FOneItem.Fetctag="7","checked","")%>> 참여
					<input type="radio" name="etctag" value="8" <%=chkiif(oMainContents.FOneItem.Fetctag="8","checked","")%>> 할인율 자동 노출(A타입-중앙)
					<input type="radio" name="etctag" value="9" <%=chkiif(oMainContents.FOneItem.Fetctag="9","checked","")%>> 할인율 자동 노출(기존)
					<input type="radio" name="etctag" value="10" <%=chkiif(oMainContents.FOneItem.Fetctag="10","checked","")%>> 할인율 자동 노출(B타입-신규)
					 <br/>
					※ 한가지만 선택 하세요.<br/><br/>
					<input type="checkbox" name="tag_only" value="Y" <%=chkiif(oMainContents.FOneItem.Ftag_only="Y","checked","")%>> 단독<br/><br/>
					<input type="text" name="etctext" value="<%=oMainContents.FOneItem.Fetctext%>" size="20" maxlength="30" class="text" />※ 할인,쿠폰 일경우만 입력 하세요<br/>
				</td>
			</tr>
			<tr bgcolor="#FFFFFF">
				<td width="150" bgcolor="#DDDDFF">이벤트 코드</td>
				<td>
					<span><input type="text" id="saleinfo2" name="evt_code" value="<%= oMainContents.FOneItem.FEvt_Code %>" size="20" maxlength="10" class="text" /></span>
					<p class="tPad05"><span class="rMar10"><strong>※ 할인율 자동 노출 및 이벤트 상태 체크 (대기중 , 종료) 노출 X ※</strong></span></p>
				</td>
			</tr>
		<% End If %>

		<% If poscode="714" Then %>
		<%'// 2018 컬쳐스테이션%>
			<input type="hidden" name="ecode" value=""/><%' cultureidx %>
			<tr bgcolor="#FFFFFF">
				<td width="150" bgcolor="#DDDDFF">메인카피</td>
				<td><input type="text" name="maincopy" value="<%=oMainContents.FOneItem.Fmaincopy%>" size="50" maxlength="25" class="text" /></td>
			</tr>
			<tr bgcolor="#FFFFFF">
				<td width="150" bgcolor="#DDDDFF">서브카피</td>
				<td><input type="text" name="subcopy" value="<%=oMainContents.FOneItem.Fsubcopy%>" size="80" maxlength="60" class="text" /></td>
			</tr>
			<tr bgcolor="#FFFFFF">
				<td width="150" bgcolor="#DDDDFF">내용</td>
				<td><textarea name="linkText3" style="width:100%;height:60px;" class="text"><%= oMainContents.FOneItem.FlinkText3 %></textarea></td>
			</tr>
			<tr bgcolor="#FFFFFF">
				<td width="150" bgcolor="#DDDDFF">구분선택</td>
				<td>
					<input type="radio" name="xbtncolor" value="0" <%=chkiif(oMainContents.FOneItem.Fxbtncolor = "" Or oMainContents.FOneItem.Fxbtncolor = "0" ,"checked","")%>/> : 느껴봐
					<input type="radio" name="xbtncolor" value="1" <%=chkiif(oMainContents.FOneItem.Fxbtncolor = "1" ,"checked","")%>/> : 읽어봐
				</td>
			</tr>
		<% End If %>

		<tr bgcolor="#FFFFFF">
			<td width="150" bgcolor="#DDDDFF">이미지1</td>
			<td>
				<% If poscode <> "714" Then %>
				<%'// 2018 컬쳐스테이션이 아닐경우만 %>
					<input type="file" name="file1" value="" id="file1" size="32" maxlength="32" class="file">
				<% End If %>

				<% if oMainContents.FOneItem.GetImageUrl<>"" then %>
					<br>
					<img src="<%= oMainContents.FOneItem.GetImageUrl %>" id="img1" style="max-width:600px;" />
					<br><span id="imgurl1"> <%= oMainContents.FOneItem.GetImageUrl %>&nbsp;&nbsp;<input type="button" value=" 삭제 " onClick="fnDeleteImage('1');"></span>
				<% end if %>

				<% '컬쳐스테이션 %>
				<% If oMainContents.FOneItem.Fidx = "" And poscode = "714" Then %>
					<br>
					<img src="<%=cultureEimagelist %>" style="max-width:600px;" />
					<br> <%= cultureEimagelist %> <br/><br/> ※ 이미지 수정은 컬쳐스테이션 어드민에서 해주세요
				<% ElseIf oMainContents.FOneItem.Fidx <> "" And poscode = "714" Then %>
					<br>
					<img src="<%=oMainContents.FOneItem.Fcultureimage %>" style="max-width:600px;" />
					<br> <%=oMainContents.FOneItem.Fcultureimage %> <br/><br/> ※ 이미지 수정은 컬쳐스테이션 어드민에서 해주세요
				<% End If %>
			</td>
		</tr>

		<tr bgcolor="#FFFFFF">
			<td width="150" bgcolor="#DDDDFF">알트명1 (필수)</td>
			<td><input type="text" name="altname" value="<%=oMainContents.FOneItem.Faltname%>" size="20" maxlength="20"> </td>
		</tr>

		<% If poscode = "706" or poscode = "720" or poscode="722" or poscode="723" or poscode="724" or poscode="731" Then %>
		<%'// [2018] PC 헤더 최상단 띠배너, [2018] 메인빅이벤트배너1~3, 메인컨텐츠 상단배너, 메인컨텐츠 하단배너, 로그인배너 %>
			<tr bgcolor="#FFFFFF" id="bnimg2" style="display:<% If oMainContents.FOneItem.Fbannertype="1" or oMainContents.FOneItem.Fbannertype="" Then Response.write "none" %>">
				<td width="150" bgcolor="#DDDDFF">이미지2</td>
				<td>
					<input type="file" name="file2" id="file2" value="" size="32" maxlength="32" class="file">
					<% if oMainContents.FOneItem.GetImageUrl2<>"" then %>
						<br>
						<img src="<%= oMainContents.FOneItem.GetImageUrl2 %>" id="img2" style="max-width:600px;" />
						<br> <span id="imgurl2"> <%= oMainContents.FOneItem.GetImageUrl2 %>&nbsp;&nbsp;<input type="button" value=" 삭제 " onClick="fnDeleteImage('2');"></span>
					<% end if %>
				</td>
			</tr>
			<tr bgcolor="#FFFFFF" id="bnalt2" style="display:<% If oMainContents.FOneItem.Fbannertype="1" or oMainContents.FOneItem.Fbannertype="" Then Response.write "none" %>">
				<td width="150" bgcolor="#DDDDFF">알트명2 (필수)</td>
				<td><input type="text" name="altname2" value="<%=oMainContents.FOneItem.Faltname2%>" size="20" maxlength="20"> </td>
			</tr>
			<tr bgcolor="#FFFFFF" id="bnimg3" style="display:<% If oMainContents.FOneItem.Fbannertype<>"3" Or oMainContents.FOneItem.Fbannertype="" Then Response.write "none" %>">
				<td width="150" bgcolor="#DDDDFF">이미지3</td>
				<td>
					<input type="file" name="file3" id="file3" value="" size="32" maxlength="32" class="file">
					<% if oMainContents.FOneItem.GetImageUrl3<>"" then %>
						<br>
						<img src="<%= oMainContents.FOneItem.GetImageUrl3 %>" id="img3" style="max-width:600px;" />
						<br> <span id="imgurl3"> <%= oMainContents.FOneItem.GetImageUrl3 %>&nbsp;&nbsp;<input type="button" value=" 삭제 " onClick="fnDeleteImage('3');"></span>
					<% end if %>
				</td>
			</tr>
			<tr bgcolor="#FFFFFF" id="bnalt3" style="display:<% If oMainContents.FOneItem.Fbannertype<>"3" Or oMainContents.FOneItem.Fbannertype="" Then Response.write "none" %>">
				<td width="150" bgcolor="#DDDDFF">알트명3 (필수)</td>
				<td><input type="text" name="altname3" value="<%=oMainContents.FOneItem.Faltname3%>" size="20" maxlength="20"> </td>
			</tr>
		<% End If %>

		<% If gubun <> "PCbanner" and gubun <> "MAbanner" And poscode <> "706" And poscode <> "720" And poscode<>"722" And poscode<>"723" And poscode<>"724" And poscode<>"736" Then %>
		<%'// pc배너, 모바일앱배너, 헤더 최상단, 메인빅이벤트가 아닐경우 %>
			<tr bgcolor="#FFFFFF">
				<% If poscode = "721" then %>
					<% '// 메인 플로팅 배너일 경우만 타이틀명 변경 %>
					<td width="150" bgcolor="#DDDDFF">마우스 오버 시 이미지</td>
				<% Else %>
					<td width="150" bgcolor="#DDDDFF">이미지 (선택)</td>
				<% End If %>
				<td>
					<input type="file" name="file2" id="file2" value="" size="32" maxlength="32" class="file">
					<% if oMainContents.FOneItem.GetImageUrl2<>"" then %>
						<br>
						<img src="<%= oMainContents.FOneItem.GetImageUrl2 %>" id="img2" style="max-width:600px;" />
						<br> <span id="imgurl2"> <%= oMainContents.FOneItem.GetImageUrl2 %>&nbsp;&nbsp;<input type="button" value=" 삭제 " onClick="fnDeleteImage('2');"></span>
					<% end if %>
				</td>
			</tr>
		<% End If %>

		<tr bgcolor="#FFFFFF">
			<td width="150" bgcolor="#DDDDFF">이미지Width</td>
			<td>
				<% IF gubun <> "" Then %>
					<% if oMainContents.FOneItem.Fidx<>"" then %>
						<input type="text" name="imagewidth" value="<%= oMainContents.FOneItem.Fimagewidth %>" size="8" maxlength="16">
						<% If poscode="720" Then %>
						<%'// 메인빅이벤트 %>
							(이미지 1개 기준)
						<% End If %>
					<% else %>
						<% if poscode<>"" then %>
							<%= oposcode.FOneItem.Fimagewidth %>
							<% If poscode="720" Then %>
							<%'// 메인빅이벤트 %>
								(이미지 1개 기준)
							<% End If %>
						<% else %>
							<font color="red">구분을 먼저 선택하세요</font>
						<% end if %>
					<% end if %>
				<% Else %>
					<font color="red">그룹구분을 먼저 선택하세요</font>
				<% End If %>
			</td>
		</tr>
		<tr bgcolor="#FFFFFF">
			<td width="150" bgcolor="#DDDDFF">이미지Height</td>
			<td>
				<% IF gubun <> "" Then %>
					<% if oMainContents.FOneItem.Fidx<>"" then %>
						<input type="text" name="imageheight" value="<%= oMainContents.FOneItem.Fimageheight %>" size="8" maxlength="16">
					<% else %>
						<% if poscode<>"" then %>
							<%= oposcode.FOneItem.Fimageheight %>
						<% else %>
							<font color="red">구분을 먼저 선택하세요</font>
						<% end if %>
					<% end if %>
				<% Else %>
					<font color="red">그룹구분을 먼저 선택하세요</font>
				<% End If %>
			</td>
		</tr>
	<% End If %>

	<tr bgcolor="#FFFFFF">
		<td width="150" bgcolor="#DDDDFF">링크값1</td>
		<td>
			<% IF gubun <> "" Then %>
				<% if oMainContents.FOneItem.Fidx<>"" then %>
					<% if oMainContents.FOneItem.FLinkType="M" then %>
						<textarea name="linkurl" style="width:100%;height:120px;"><%= oMainContents.FOneItem.Flinkurl %></textarea>
					<% else %>
						<% if oMainContents.FOneItem.Fposcode = 539 Then%>
							<textarea name="linkurl" style="width:100%;height:120px;"><%= oMainContents.FOneItem.Flinkurl %></textarea>
						<% Else%>
							<input type="text" name="linkurl" value="<%= oMainContents.FOneItem.Flinkurl %>" maxlength="128" style="width:100%;" class="text">
						<% End If %>
					<% end if %>
				<% else %>
					<% if poscode<>"" then %>
						<% if oposcode.FOneItem.FLinkType="M" then %>
							<textarea name="linkurl" style="width:100%;height:120px;"><%= defaultMapStr %></textarea>
							<br>(이미지맵 변수값 변경 금지)
						<% elseif oposcode.FOneItem.FLinkType="B" then %>
							<input type="text" class="text_ro" name="linkurl" value="/" maxlength="128" size="40" readonly>
						<% elseif poscode="539" Then %>
							<textarea name="linkurl" style="width:100%;height:120px;"><%= defaultXMLMapStr %></textarea>
							<br>(이미지맵 변수값 변경 금지, href이하에 링크넣어주세요)
						<% Else %>
							<input type="text" name="linkurl" value="" maxlength="128" style="width:100%;" class="text">
							<br>ex)<br/>
							- <span style="cursor:pointer" onClick="putLinkText('event');">이벤트 링크 : /event/eventmain.asp?eventid=<span style="color:darkred">이벤트코드</span></span><br/>
							- <span style="cursor:pointer" onClick="putLinkText('itemid');">상품코드 링크 : /shopping/category_prd.asp?itemid=<span style="color:darkred">상품코드 (O)</span></span><br/>
							- <span style="cursor:pointer" onClick="putLinkText('category');">카테고리 링크 : /shopping/category_list.asp?disp=<span style="color:darkred">카테고리</span></span><br/>
							- <span style="cursor:pointer" onClick="putLinkText('brand');">브랜드아이디 링크 : /street/street_brand.asp?makerid=<span style="color:darkred">브랜드아이디</span></span><br/>
							- <span style="cursor:pointer" onClick="putLinkText('hitchhiker');">히치하이커 링크 : /hitchhiker/</span><br/>
							- <span style="cursor:pointer" onClick="putLinkText('giftcard');">기프트카드 링크 : /giftcard/</span><br/>
							- <span style="cursor:pointer" onClick="putLinkText('culture');">컬쳐스테이션 링크 : /culturestation/culturestation_event.asp?evt_code=<span style="color:darkred">이벤트아이디</span></span><br/>
							- <span style="cursor:pointer" onClick="putLinkText('coupon');">쿠폰함 링크 : /my10x10/couponbook.asp
						<% end if %>
					<% else %>
						<font color="red">구분을 먼저 선택하세요</font>
					<% end if %>
				<% end if %>
			<% Else %>
				<font color="red">그룹구분을 먼저 선택하세요</font>
			<% End If %>
		</td>
	</tr>

	<% If poscode = "706" or poscode="720" or poscode="722" or poscode="723" or poscode="724" or poscode="731" Then %>
	<%'// [2018] PC 헤더 최상단 띠배너, [2018] 메인빅이벤트배너1~3, 메인컨텐츠 상, 하단, 로그인 배너 %>
		<tr bgcolor="#FFFFFF" id="bnlink2" style="display:<% If oMainContents.FOneItem.Fbannertype="1" Or oMainContents.FOneItem.Fbannertype="" Then Response.write "none" %>">
			<td width="150" bgcolor="#DDDDFF">링크값2</td>
			<td>
				<input type="text" name="linkurl2" value="<%= oMainContents.FOneItem.Flinkurl2 %>" maxlength="128" style="width:100%;" class="text">
			</td>
		</tr>
		<tr bgcolor="#FFFFFF" id="bnlink3" style="display:<% If oMainContents.FOneItem.Fbannertype<>"3" Or oMainContents.FOneItem.Fbannertype="" Then Response.write "none" %>">
			<td width="150" bgcolor="#DDDDFF">링크값3</td>
			<td>
				<input type="text" name="linkurl3" value="<%= oMainContents.FOneItem.Flinkurl3 %>" maxlength="128" style="width:100%;" class="text">
			</td>
		</tr>

		<% If not(poscode = "720" or poscode="722" or poscode="723" or poscode="724" or poscode="731") Then %>
		<%'// 메인빅이벤트, 메인컨텐츠 상, 하단, 로그인 배너 사용안함 %>
			<tr bgcolor="#FFFFFF">
				<td width="150" bgcolor="#DDDDFF">좌우측 BG컬러코드</td>
				<td>
					<span  id="bnbg1" style="display:<% If oMainContents.FOneItem.Fbannertype="1" Or oMainContents.FOneItem.Fbannertype="" Then Response.write "none" %>">좌 : </span>#<input type="text" name="bgcode" value="<%=oMainContents.FOneItem.Fbgcode%>" size="20" maxlength="6">
					<div  id="bnbg2" style="display:<% If oMainContents.FOneItem.Fbannertype="1" Or oMainContents.FOneItem.Fbannertype="" Then Response.write "none" %>">우 : #<input type="text" name="bgcode2" value="<%=oMainContents.FOneItem.Fbgcode2%>" size="20" maxlength="6"></div>
				</td>
			</tr>
			<tr bgcolor="#FFFFFF">
				<td width="150" bgcolor="#DDDDFF">X버튼선택</td>
				<td>
					<input type="radio" name="xbtncolor" value="0" <%=chkiif(oMainContents.FOneItem.Fxbtncolor = "" Or oMainContents.FOneItem.Fxbtncolor = "0" ,"checked","")%>/> : 화이트
					<input type="radio" name="xbtncolor" value="1" <%=chkiif(oMainContents.FOneItem.Fxbtncolor = "1" ,"checked","")%>/> : black
				</td>
			</tr>
		<% End If %>
	<% End If %>

	<tr bgcolor="#FFFFFF">
		<td width="150" bgcolor="#DDDDFF">반영시작일</td>
		<td>
			<input id="startdate" name="startdate" value="<%=chkiif(idx=0,prevDate,Left(oMainContents.FOneItem.Fstartdate,10))%>" class="text" size="10" maxlength="10" />
			<img src="http://webadmin.10x10.co.kr/images/calicon.gif" id="startdate_trigger" border="0" style="cursor:pointer;" align="absbottom" />
			<% if oMainContents.FOneItem.Ffixtype="R" or poscode="687" Or gubun = "PCbanner" Or gubun = "MAbanner" then %> <!-- 실시간인경우 / 걍 일단위로 돌림 (나중에 시간단위로 돌릴때 False 제거)-->
				<input type="text" name="startdatetime" size="2" maxlength="2" value="<%= Format00(2,Hour(oMainContents.FOneItem.Fstartdate)) %>" />(시 00~23)
				<input type="text" name="dummy0" value="00:00" size="6" readonly class="text_ro" />
			<% else %>
				<input type="text" name="dummy0" value="00:00:00" size="8" readonly class="text_ro" />
			<% end if %>
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
		<td width="150" bgcolor="#DDDDFF">반영종료일</td>
		<td>
			<input id="enddate" name="enddate" value="<%=chkiif(idx=0,prevDate,Left(oMainContents.FOneItem.Fenddate,10)) %>" class="text" size="10" maxlength="10" />
			<img src="http://webadmin.10x10.co.kr/images/calicon.gif" id="enddate_trigger" border="0" style="cursor:pointer" align="absbottom" />
			<% if oMainContents.FOneItem.Ffixtype="R"  or poscode="687" Or gubun = "PCbanner" Or gubun = "MAbanner" then %> <!-- 실시간인경우 -->
				<input type="text" name="enddatetime" size="2" maxlength="2" value="<%= ChkIIF(oMainContents.FOneItem.Fenddate="","23",Format00(2,Hour(oMainContents.FOneItem.Fenddate))) %>">(시 00~23)
				<input type="text" name="dummy1" value="59:59" size="6" readonly class="text_ro" />
			<% else %>
				<input type="text" name="dummy1" value="23:59:59" size="8" readonly class="text_ro" />
			<% end if %>
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
		<td width="150" bgcolor="#DDDDFF">등록일</td>
		<td>
			<%= oMainContents.FOneItem.Fregdate %> (<%= oMainContents.FOneItem.Fregname %>)
		</td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td width="150" bgcolor="#DDDDFF">작업자</td>
		<td>
			<% If idx <> "" AND idx <> "0" Then %>
				최종 작업자 : <%=oMainContents.FOneItem.Fworkername%><input type="hidden" name="selDId" value="<%=session("ssBctId")%>">
				&nbsp;<strong><%=oMainContents.FOneItem.Flastupdate%></strong>
			<% Else %>
				<input type="hidden" name="selDId" value="">
			<% End If %>
		</td>
	</tr>

	<tr bgcolor="#FFFFFF">
		<td width="150" bgcolor="#DDDDFF">사용여부</td>
		<td>
			<% if oMainContents.FOneItem.Fisusing="N" then %>
				<input type="radio" name="isusing" value="Y">사용함
				<input type="radio" name="isusing" value="N" checked >사용안함
			<% else %>
				<input type="radio" name="isusing" value="Y" checked >사용함
				<input type="radio" name="isusing" value="N">사용안함
			<% end if %>
		</td>
	</tr>

	<tr bgcolor="#FFFFFF">
		<td colspan="2" align="center"><input type="button" value=" 저 장 " onClick="SaveMainContents(frmcontents);"></td>
	</tr>
</form>
</table>
<%
	set oposcode = Nothing
	set oMainContents = Nothing
%>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->