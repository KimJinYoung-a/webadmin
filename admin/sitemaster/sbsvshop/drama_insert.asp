<%@ language=vbscript %>
<% option explicit %>
<%
'###############################################
' PageName : drama_insert.asp
' Discription : 모바일 exhibition
' History : 2016.04.07 이종화
'###############################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/classes/sitemasterclass/sbsvshopCls.asp" -->
<%
Dim mode
Dim srcSDT , srcEDT
Dim mainStartDate, mainEndDate , lp , ii
Dim sDt, sTm, eDt, eTm 
dim listidx,dramaidx,title,contents,mainimage,videourl,subimage1,subimage2,subimage3,subimage4,subimage5,startdate,enddate,regdate,lastupdate,adminid,lastadminid,isusing,ordertext , kakaoshareimage
Dim evtcode, evtbnrimg, evtMainCopy, evtSubCopy, salePercentage, evtsDt, evteDt, bannerIsUsing


evtsDt = request("evtsDt")
evteDt = request("evteDt")

'테스트데이터

	listidx = requestCheckvar(request("listidx"),16)
	srcSDT = request("sDt")
	srcEDT = request("eDt")

If listidx = "" Then
	mode = "add"
Else
	mode = "modify"
End If

	isusing = 1

If listidx <> "" then
	dim dramaList
	set dramaList = new sbsvshop
	dramaList.FRectIdx = listidx
	dramaList.fnDramaContentsGet()

	listidx			=	dramaList.FOneItem.Flistidx
	dramaidx		=	dramaList.FOneItem.Fdramaidx	
	title			=	dramaList.FOneItem.Ftitle		
	contents		=	dramaList.FOneItem.Fcontents	
	mainimage		=	dramaList.FOneItem.Fmainimage	
	videourl		=	dramaList.FOneItem.Fvideourl	
	subimage1		=	dramaList.FOneItem.Fsubimage1	
	subimage2		=	dramaList.FOneItem.Fsubimage2	
	subimage3		=	dramaList.FOneItem.Fsubimage3	
	subimage4		=	dramaList.FOneItem.Fsubimage4	
	subimage5		=	dramaList.FOneItem.Fsubimage5	
	mainStartDate	=	dramaList.FOneItem.Fstartdate	
	mainEndDate		=	dramaList.FOneItem.Fenddate		
	regdate			=	dramaList.FOneItem.Fregdate		
	lastupdate		=	dramaList.FOneItem.Flastupdate	
	adminid			=	dramaList.FOneItem.Fadminid		
	lastadminid		=	dramaList.FOneItem.Flastadminid	
	isusing			=	dramaList.FOneItem.Fisusing		
	ordertext		=	dramaList.FOneItem.Fordertext
	kakaoshareimage	=	dramaList.FOneItem.Fkakaoshareimage
'20180731 최종원 추가	
	evtcode			=   dramaList.FOneItem.FeventCode
	evtbnrimg		=	dramaList.FOneItem.FBannerImg
	evtMainCopy		=	dramaList.FOneItem.FMainCopy
	evtSubCopy		=	dramaList.FOneItem.FSubCopy
	salePercentage	=	dramaList.FOneItem.FSalePer
	evtsDt			=	dramaList.FOneItem.FOpenDate
	evteDt			=	dramaList.FOneItem.FCloseFate
	bannerIsUsing	=   dramaList.FOneItem.FbannerIsUsing

	set dramaList = Nothing
End If

Dim oSubItemList
set oSubItemList = new sbsvshop
	oSubItemList.FPageSize = 100
	oSubItemList.FRectlistIdx = listidx
	If listidx <> "" then
		oSubItemList.fnDramaContentsItemList()
	End If


if Not(mainStartDate="" or isNull(mainStartDate)) then
	sDt = left(mainStartDate,10)
	sTm = Num2Str(hour(mainStartDate),2,"0","R")
else
	if srcSDT<>"" then
		sDt = left(srcSDT,10)
	else
		sDt = date
	end if
	sTm = "00"
end if

if Not(mainEndDate="" or isNull(mainEndDate)) then
	eDt = left(mainEndDate,10)
	eTm = Num2Str(hour(mainEndDate),2,"0","R")
else
	if srcEDT<>"" then
		eDt = left(srcEDT,10)
	else
		eDt = date
	end if
	eTm = "23"
end If
%>
<link href="/js/jqueryui/css/jquery-ui.css" rel="stylesheet">
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript" src="/js/jqueryui/jquery-ui-1.10.2.custom.min.js"></script>
<script type='text/javascript'>
	function deleteImg(){
		var frm = document.frm;
		frm.evtbnrimg.value = "";
		frm.bnrimg.src= "/images/admin_login_logo2.png";
	}
	function setImg(){
		var frm = document.frm;
		<% if evtbnrimg="" then %>
		frm.bnrimg.src= "/images/admin_login_logo2.png";
		<% else %>
		frm.bnrimg.src= "<%=evtbnrimg%>";
		<% end if %>
		frm.evtbnrimg.value = "<%=evtbnrimg%>";
		
	}	
	function showBannerRow(){		
		var frm = document.frm;
		var evtBannerRowObj = document.getElementById("evtBannerRow")		
		if(frm.isBannerUsing[1].checked === true){
			evtBannerRowObj.style.display="none";
		}else if(frm.isBannerUsing[0].checked === true){
			evtBannerRowObj.style.display="";
		}		
	}
	function jsLastEvent(){
	var valsdt , valedt , valgcode
		valsdt = document.frm.sDt.value;
		valedt = document.frm.eDt.value;

	var winLast,eKind;
	winLast = window.open('pop_event_list.asp?menupos=<%=menupos%>&sDt='+valsdt+'&eDt='+valedt,'pLast','width=550,height=600, scrollbars=yes')
	winLast.focus();
	}
	function jsSubmit(){
		var frm = document.frm;
//test
		if (frm.sTm.value.length != 2) {
			alert("시간을 정확히 입력하세요");
			frm.sTm.focus();
			return;
		}

		if (frm.eTm.value.length != 2) {
			alert("시간을 정확히 입력하세요");
			frm.eTm.focus();
			return;
		}

		if (!frm.dramaidx.value){
			alert("드라마를 선택 해주세요");
			frm.dramaidx.focus();
			return;
		}

		if (!frm.title.value){
			alert("컨텐츠 카피를 입력 해주세요");
			frm.title.focus();
			return;
		}

		if (GetByteLength(frm.title.value) > 30){
			alert("제한길이를 초과하였습니다. 15자 까지 작성 가능합니다.");
			frm.title.focus();
			return false;
		}

		if (!frm.contents.value){
			alert("컨텐츠를 입력 해주세요");
			frm.contents.focus();
			return;
		}
		if(frm.isBannerUsing[0].checked===true){
			if(frm.evtcode.value === ""){
				alert('이벤트 코드를 입력 해주세요');	
				return false;
			}
			if(frm.evtMainCopy.value === ""){
				alert('이벤트 메인 카피를 입력 해주세요');	
				return false;
			}
			if(frm.evtMainCopy.value === ""){
				alert('이벤트 서브 카피를 입력 해주세요');	
				return false;
			}			
			if(frm.salePercentage.value === ""){
				alert('이벤트 할인율을 입력 해주세요');	
				return false;
			}
			if(frm.evtsDt.value === ""){
				alert('이벤트 시작날짜를 입력 해주세요');	
				return false;
			}
			if(frm.evteDt.value === ""){
				alert('이벤트 종료 날짜를 입력 해주세요');	
				return false;
			}			
		}
		if (confirm('저장 하시겠습니까?')){
			//frm.target = "blank";
			frm.submit();
		}
	}
	
	function jsgolist(){
		self.location.href="/admin/sitemaster/sbsvshop/";
	}

	$(function(){
	showBannerRow();//이벤트 배너 창		
	setImg();//이미지 설정
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
      	<% if listidx<>"" then %>maxDate: "<%=eDt%>",<% end if %>
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
      	<% if listidx<>"" then %>minDate: "<%=sDt%>",<% end if %>
    	onClose: function( selectedDate ) {
    		$( "#sDt" ).datepicker( "option", "maxDate", selectedDate );
    	}
    });

	//라디오버튼
    $(".rdoUsing").buttonset().children().next().attr("style","font-size:11px;");


	$( "#subList").sortable({
		placeholder: "ui-state-highlight",
		start: function(event, ui) {
			ui.placeholder.html('<td height="54" colspan="10" style="border:1px solid #F9BD01;">&nbsp;</td>');
		},
		stop: function(){
			var i=99999;
			$(this).find("input[name^='sort']").each(function(){
				if(i>$(this).val()) i=$(this).val()
			});
			if(i<=0) i=1;
			$(this).find("input[name^='sort']").each(function(){
				$(this).val(i);
				i++;
			});
		}
	});
});

//소재
function popSubEdit(subidx) {
<% if listidx <>"" then %>
    var popwin = window.open('popSubItemEdit.asp?listIdx=<%=listidx%>&subIdx='+subidx,'popTemplateManage','width=800,height=300,scrollbars=yes,resizable=yes');
    popwin.focus();
<% else %>
	alert("템플릿 컨텐츠 정보를 먼저 등록해주세요.");
<% end if %>
}

// 상품검색 일괄 등록
function popRegSearchItem() {
<% if listidx <> "" then %>
    var popwin = window.open("/admin/itemmaster/pop_itemAddInfo.asp?sellyn=Y&usingyn=Y&defaultmargin=0&acURL=/admin/sitemaster/sbsvshop/doSubRegItemCdArray.asp?listidx=<%=listidx%>", "popup_item", "width=800,height=500,scrollbars=yes,resizable=yes");
    popwin.focus();
<% else %>
	alert("템플릿 컨텐츠 정보를 먼저 등록해주세요.");
<% end if %>
}

// 상품코드 일괄 등록
function popRegArrayItem() {
<% if listidx<>"" then %>
    var popwin = window.open('popSubRegItemCdArray.asp?listIdx=<%=listidx%>','popRegArray','width=600,height=300,scrollbars=yes,resizable=yes');
    popwin.focus();
<% else %>
	alert("템플릿 컨텐츠 정보를 먼저 등록해주세요.");
<% end if %>
}

function chkAllItem() {
	if($("input[name='chkIdx']:first").attr("checked")=="checked") {
		$("input[name='chkIdx']").attr("checked",false);
	} else {
		$("input[name='chkIdx']").attr("checked","checked");
	}
}

function saveList() {
	var chk=0;
	$("form[name='frmList']").find("input[name='chkIdx']").each(function(){
		if($(this).attr("checked")) chk++;
	});
	if(chk==0) {
		alert("수정하실 소재를 선택해주세요.");
		return;
	}
	if(confirm("지정하신 목록의 선택 정보를 저장하시겠습니까?")) {
		document.frmList.action="doListModify.asp";
		document.frmList.submit();
	}
}

//'아이템 삭제
function itemdel(v){
	if (confirm("상품이 삭제됩니다 삭제 하시겠습니까?")){
		document.frmdel.chkIdx.value = v;
		document.frmdel.mode.value = "itemdel";
		document.frmdel.action="doListModify.asp";
		document.frmdel.submit();
	}
}

function fileInfo(f,id){
	var file = f.files; // files 를 사용하면 파일의 정보를 알 수 있음

	var $el = $("#"+id);

	console.log($el);

	var reader = new FileReader(); // FileReader 객체 사용
	reader.onload = function(rst){ // 이미지를 선택후 로딩이 완료되면 실행될 부분

		$el.attr('src',rst.target.result);

		console.log($el.attr('src'));
//		$('#img_box').empty().html('<img src="' + rst.target.result + '">'); // append 메소드를 사용해서 이미지 추가
		// 이미지는 base64 문자열로 추가
		// 이 방법을 응용하면 선택한 이미지를 미리보기 할 수 있음
	}
	reader.readAsDataURL(file[0]); // 파일을 읽는다, 배열이기 때문에 0 으로 접근
}

function fnimgdelete(id){
	if(confirm("이미지를 삭제하시겠습니까?")){
		document.frmdelimg.imageno.value = id;
		document.frmdelimg.action ="doListModify.asp";
		document.frmdelimg.submit();
	}
}
</script>
<form name="frmdelimg" method="POST" action="">
<input type="hidden" name="mode" value="imagedel"/>
<input type="hidden" name="chkIdx" value="<%=listidx%>"/>
<input type="hidden" name="imageno" />
</form>
<form name="frmdel" method="POST" action="">
<input type="hidden" name="mode" />
<input type="hidden" name="chkIdx" />
</form>
<table width="1000" cellpadding="2" cellspacing="1" class="a" bgcolor="#3d3d3d">
<form name="frm" method="post" action="<%=uploadUrl%>/linkweb/mobile/sbsdramalist_proc.asp" onsubmit="return false;" enctype="multipart/form-data">
<input type="hidden" name="mode" value="<%=mode%>">
<input type="hidden" name="adminid" value="<%=session("ssBctId")%>">
<input type="hidden" name="listidx" value="<%=listidx%>">
<input type="hidden" name="evtbnrimg" value="">
<tr bgcolor="#FFFFFF">
	<% If listidx = ""  Then %>
	<td colspan="2" align="center" height="35">등록 진행 중 입니다.</td>
	<% Else %>
	<td bgcolor="#FFF999" colspan="2" align="center" height="35">수정 진행 중 입니다.</td>
	<% End If %>
</tr>
<% If listidx <> ""  Then %>
<tr bgcolor="#FFFFFF">
	<td bgcolor="#FFF999" align="center" width="15%">IDX</td>
	<td><%=listidx%></td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="#FFF999" align="center" width="15%">드라마명</td>
	<td><% Call getdramaname("dramaidx",dramaidx,"on") %></td>
</tr>
<% Else %>
<tr bgcolor="#FFFFFF">
	<td bgcolor="#FFF999" align="center" width="15%">드라마명</td>
	<td><% Call getdramaname("dramaidx",dramaidx,"on") %></td>
</tr>
<% End If %>
<tr bgcolor="#FFFFFF">
    <td bgcolor="#FFF999" align="center" width="15%">노출기간</td>
    <td>
		<input type="text" id="sDt" name="StartDate" size="10" value="<%=sDt%>" />
		<input type="text" name="sTm" size="2" value="<%=sTm%>" maxlength="2"/>:00:00 ~
		<input type="text" id="eDt" name="EndDate" size="10" value="<%=eDt%>" />
		<input type="text" name="eTm" size="2" value="<%=eTm%>" maxlength="2"/>:59:59
    </td>
</tr>
<tr bgcolor="#FFFFFF">
    <td bgcolor="#FFF999" align="center" width="15%">컨텐츠카피</td>
    <td>
		<input type="text" name="title" size="50" value="<%=title%>" maxlength="15"/>
    </td>
</tr>
<tr bgcolor="#FFFFFF">
    <td bgcolor="#FFF999" align="center" width="15%">컨텐츠설명</td>
    <td>
		<textarea name="contents" cols="80" rows="8" maxlength="60"/><%=contents%></textarea>
    </td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="#FFF999" align="center" width="10%">카카오 공유</td>
	<td align="left">
		<input type="file" name="kakaoshareimage" class="file" title="이벤트 #1" require="N" style="width:80%;" accept="image/*" onchange="fileInfo(this,'kakaoshareimage');"/>
		<% If kakaoshareimage <> "" Then %>
		<br/><img src="<%=kakaoshareimage%>" width="120" id="kakaoshareimage"/>
		<% Else %>
		<br/>
			<img src="/images/admin_login_logo2.png" width="120" id="kakaoshareimage"/>
		<% End If %>
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="#FFF999" align="center" width="10%">썸네일</td>
	<td align="left">
		<input type="file" name="mainimage" class="file" title="이벤트 #1" require="N" style="width:80%;" accept="image/*" onchange="fileInfo(this,'thumbimage');"/>
		<% If mainimage <> "" Then %>
		<br/><img src="<%=mainimage%>" width="120" id="thumbimage"/>
		<% Else %>
		<br/>
			<img src="/images/admin_login_logo2.png" width="120" id="thumbimage"/>
		<% End If %>
	</td>
</tr>
<tr bgcolor="#FFFFFF">
    <td bgcolor="#FFF999" align="center" width="15%">동영상URL</td>
    <td>
		<input type="text" name="videourl" size="100" value="<%=videourl%>" />
		<br/> <span style="color:red">※ 영상 URL을 넣어 주세요 ※</span>
    </td>
</tr>

<tr bgcolor="#FFFFFF">
	<td bgcolor="#FFF999" align="center" width="10%">swiper 이미지</td>
	<td align="left">
		<table>
			<tr>
				<td>
					<input type="file" name="subimage1" class="file" title="이벤트 #1" require="N" style="width:80%;" accept="image/*" onchange="fileInfo(this,'image1');"/>
					<% If subimage1 <> "" Then %>
					<br/>
					<div style="position:relative">
						<img src="<%=subimage1%>" width="120" id="image1"/>
						<div style="position:absolute;left:100px;top:100px;cursor:pointer" onclick="fnimgdelete('1');"><img src="/images/icon_delete2.gif" border="0" class="delImg"></div>
					</div>
					<% Else %>
					<br/><img src="/images/admin_login_logo2.png" width="120" id="image1"/>
					<% End If %>
				</td>
				<td>
					<input type="file" name="subimage2" class="file" title="이벤트 #1" require="N" style="width:80%;" accept="image/*" onchange="fileInfo(this,'image2');"/>
					<% If subimage2 <> "" Then %>
					<br/>
					<div style="position:relative">
						<img src="<%=subimage2%>" width="120" id="image2"/>
						<div style="position:absolute;left:100px;top:100px;cursor:pointer" onclick="fnimgdelete('2');"><img src="/images/icon_delete2.gif" border="0" class="delImg"></div>
					</div>
					<% Else %>
					<br/><img src="/images/admin_login_logo2.png" width="120" id="image2"/>
					<% End If %>
				</td>
				<td>
					<input type="file" name="subimage3" class="file" title="이벤트 #1" require="N" style="width:80%;" accept="image/*" onchange="fileInfo(this,'image3');"/>
					<% If subimage3 <> "" Then %>
					<br/>
					<div style="position:relative">
						<img src="<%=subimage3%>" width="120" id="image3"/>
						<div style="position:absolute;left:100px;top:100px;cursor:pointer" onclick="fnimgdelete('3');"><img src="/images/icon_delete2.gif" border="0" class="delImg"></div>
					</div>
					<% Else %>
					<br/><img src="/images/admin_login_logo2.png" width="120" id="image3"/>
					<% End If %>
				</td>
				<td>
					<input type="file" name="subimage4" class="file" title="이벤트 #1" require="N" style="width:80%;" accept="image/*" onchange="fileInfo(this,'image4');"/>
					<% If subimage4 <> "" Then %>
					<br/>
					<div style="position:relative">
						<img src="<%=subimage4%>" width="120" id="image4"/>
						<div style="position:absolute;left:100px;top:100px;cursor:pointer" onclick="fnimgdelete('4');"><img src="/images/icon_delete2.gif" border="0" class="delImg"></div>
					</div>
					<% Else %>
					<br/><img src="/images/admin_login_logo2.png" width="120" id="image4"/>
					<% End If %>
				</td>
				<td>
					<input type="file" name="subimage5" class="file" title="이벤트 #1" require="N" style="width:80%;" accept="image/*" onchange="fileInfo(this,'image5');"/>
					<% If subimage5 <> "" Then %>
					<br/>
					<div style="position:relative">
						<img src="<%=subimage5%>" width="120" id="image5"/>
						<div style="position:absolute;left:100px;top:100px;cursor:pointer" onclick="fnimgdelete('5');"><img src="/images/icon_delete2.gif" border="0" class="delImg"></div>
					</div>
					<% Else %>
					<br/><img src="/images/admin_login_logo2.png" width="120" id="image5"/>
					<% End If %>
				</td>
			</tr>
		</table>
	</td>
</tr>
<tr bgcolor="#FFFFFF"><%'2018-07-30 최종원 배너 추가%>
	<td bgcolor="#FFF999" align="center">이벤트 배너 사용여부</td>
	<td><div style="float:left;"><input type="radio" name="isBannerUsing" value="Y" <%if bannerIsUsing = "Y" then%>checked<% end if %>  onclick="showBannerRow();"/>사용함 &nbsp;&nbsp;&nbsp; <input type="radio" name="isBannerUsing" value="N" onclick="showBannerRow();" <%if bannerIsUsing <> "Y" then%>checked<% end if %> />사용안함</div> <div style="float:right;margin-top:5px;margin-right:10px;"></div></td>
</tr>
<tr bgcolor="#FFFFFF" id="evtBannerRow" style="display:none">
    <td bgcolor="#FFF999" align="center" width="15%">이벤트 배너 등록<br>(선택)</td>
    <td>
		<input type="text" name="evtcode" size="50" value="<%=evtcode%>" style="width:200px;" /><input type="button" value=" 이벤트 불러오기 " onClick="jsLastEvent();"/><br><br>		
		<img src="/images/admin_login_logo2.png" width="120" name="bnrimg"  style="border:1px solid gray;"/>
		<input type="button" value="삭제" onclick="deleteImg();"><br><br>
		메인카피 : <input type="text" name="evtMainCopy" style="width:400px" value="<%=evtMainCopy%>"><br><br>
		서브카피 : <input type="text" name="evtSubCopy" style="width:450px" value="<%=evtSubCopy%>"><br><br>
		할인율 : <input type="text" name="salePercentage" style="width:50px" value="<%=salePercentage%>"><br><br>
		시작일-종료일 : <input type="text" style="width:70px" name="evtsDt" value="<%=evtsDt%>"> ~  <input type="text" style="width:70px" name="evteDt" value="<%=evteDt%>">
    </td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="#FFF999" align="center">작업자 지시사항</td>
	<td><textarea name="ordertext" cols="80" rows="8"/><%=ordertext%></textarea></td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="#FFF999" align="center">사용여부</td>
	<td><div style="float:left;"><input type="radio" name="isusing" value="1" <%=chkiif(isusing,"checked","")%> checked />사용함 &nbsp;&nbsp;&nbsp; <input type="radio" name="isusing" value="0"  <%=chkiif(isusing,"","checked")%>/>사용안함</div> <div style="float:right;margin-top:5px;margin-right:10px;"></div></td>
</tr>
<tr bgcolor="#FFFFFF" align="center">
    <td colspan="2">
		<input type="button" value=" 취 소 " onClick="jsgolist();"/><input type="button" value=" 저 장 " onClick="jsSubmit();"/>
	</td>
</tr>
</form>
</table>

<%
	If listidx <> "" then
%>
<p><b>▶ 소재 정보</b></p>
<!-- // 등록된 소재 목록 --------->
<form name="frmList" method="POST" action="" style="margin:0;">
<input type="hidden" name="mode" value="sub">
<table width="1000" align="left" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>" style="margin-bottom:100px">
<tr bgcolor="#FFFFFF">
	<td colspan="10">
		<table width="100%" border="0" class="a" cellpadding="0" cellspacing="0">
		<tr>
		    <td align="left">
		    	총 <%=oSubItemList.FTotalCount%> 건 /
		    	<input type="button" value="전체선택" class="button" onClick="chkAllItem()">
		    	<input type="button" value="상태저장" class="button" onClick="saveList()" title="표시순서 및 사용여부를 일괄저장합니다.">
		    </td>
			<td align="right">
		    	<input type="button" value="상품 추가" class="button" onClick="popRegSearchItem()" />
		    </td>
		</tr>
		</table>
	</td>
</tr>

<tr align="center" bgcolor="#DDDDFF">
    <td>&nbsp;</td>
    <td>소재번호</td>
    <td>이미지</td>
    <td>상품코드</td>
    <td>상품명</td>
    <td>표시순서</td>
    <td>사용여부</td>
    <td>상품삭제</td>
</tr>

<tbody id="subList<%=ii%>">
<% For lp=0 to oSubItemList.FResultCount-1 %>
<tr align="center" bgcolor="<%=chkIIF(oSubItemList.FItemList(lp).FIsUsing="Y","#FFFFFF","#FFF6F9")%>#FFFFFF">
    <td><input type="checkbox" name="chkIdx" value="<%=oSubItemList.FItemList(lp).FsubIdx%>" /></td>
    <td onclick="popSubEdit(<%=oSubItemList.FItemList(lp).FsubIdx%>)" style="cursor:pointer;"><%=oSubItemList.FItemList(lp).FsubIdx%></td>
    <td onclick="popSubEdit(<%=oSubItemList.FItemList(lp).FsubIdx%>)" style="cursor:pointer;">
    <%
    	if Not(oSubItemList.FItemList(lp).FsmallImage="" or isNull(oSubItemList.FItemList(lp).FsmallImage)) then
    		Response.Write "<img src='" & oSubItemList.FItemList(lp).FsmallImage & "' height='50' />"
    	end if
    %>
    </td>
    <td>
    <%
    	if Not(oSubItemList.FItemList(lp).FItemid="0" or isNull(oSubItemList.FItemList(lp).FItemid) or oSubItemList.FItemList(lp).FItemid="") then
    		Response.Write "<input type='text' value='" & oSubItemList.FItemList(lp).FItemid & "' readonly size='5'/>"
    	end if
    %>
    </td>
	<td><input type="text" name="itemname<%=oSubItemList.FItemList(lp).FsubIdx%>" value="<%=oSubItemList.FItemList(lp).Fitemname%>" size="60"></td>
    <td><input type="text" name="sort<%=oSubItemList.FItemList(lp).FsubIdx%>" size="3" class="text" value="<%=oSubItemList.FItemList(lp).Fsortnum%>" style="text-align:center;" /></td>
    <td>
		<span class="rdoUsing">
		<input type="radio" name="use<%=oSubItemList.FItemList(lp).FsubIdx%>" id="rdoUsing<%=lp%>_1" value="Y" <%=chkIIF(oSubItemList.FItemList(lp).FIsUsing="Y","checked","")%> /><label for="rdoUsing<%=lp%>_1">사용</label><input type="radio" name="use<%=oSubItemList.FItemList(lp).FsubIdx%>" id="rdoUsing<%=lp%>_2" value="N" <%=chkIIF(oSubItemList.FItemList(lp).FIsUsing="N","checked","")%> /><label for="rdoUsing<%=lp%>_2">삭제</label>
		</span>
    </td>
	<td><input type="button" value="상품삭제" onclick="itemdel('<%=oSubItemList.FItemList(lp).FsubIdx%>');"/></td>
</tr>
<% Next %>
</tbody>
</table>
</form>
<%
	End If
	set oSubItemList = Nothing
%>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
