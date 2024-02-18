<%@ language=vbscript %>
<% option explicit %>
<%
'###############################################
' PageName : pop_mobile_addbanner.asp
' Discription : 모바일 slide insert
' History : 2016-02-16 이종화
'###############################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<%
Dim eCode , title , eFolder , topimg , btmimg , topaddimg 'floating img
Dim slideimg
Dim mode , idx , strSql , sqlStr , sDt , eDt

	eCode = requestCheckvar(request("eC"),16)
	title = "슬라이드 등록 팝업(M)"
	eFolder = eCode

	If eCode <> "" Then
		strSql = "SELECT evt_startdate , evt_enddate " + vbcrlf
		strSql = strSql & " from db_event.dbo.tbl_event where evt_code = '"& eCode &"' " 
		rsget.Open strSql,dbget,1
		IF Not rsget.Eof Then
			sDt		= rsget("evt_startdate")
			eDt		= rsget("evt_enddate")
		End If
		rsget.close()
	End If 
%>
<!-- #include virtual="/admin/lib/popheaderslide.asp"-->
<link rel="stylesheet" type="text/css" href="http://webadmin.10x10.co.kr/css/adminDefault.css" />
<link rel="stylesheet" type="text/css" href="http://webadmin.10x10.co.kr/css/adminCommon.css" />
<style type="text/css">
html {overflow:auto;}
</style>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type='text/javascript'>
//아이템 위로
$(function(){
	// console.log($("input[name='chkIdx']").val());
	setSeq();
	chkAllItem()
});

function setSeq(){	
	var idxArr = [];
	$("input[name='chkIdx']").each(function(){
		if($(this).val()!==""){
			idxArr.push($(this).val());
		}
		console.log($(this).val());		
	})
	idxArr.sort();
	$("input[name='seq']").each(function(idx, item){
		item.value=idxArr[idx];
	});	
	
	console.log(idxArr);
}
function moveUpItem(obj) {	
    var idStr = '#' + obj;
    var prevHtml = $(idStr).prev().html();
	// console.log(prevHtml);
    if( $(idStr).prev().attr("id")  ==  undefined) {
        alert("최상위 입니다.");
        return;
    }
    var prevobj = $(idStr).prev().attr("id");
    var currobj = $(idStr).attr("id");
    var currHtml = $(idStr).html();
       
    $(idStr).html(prevHtml);//값 변경 
    $(idStr).prev().html(currHtml);
    $(idStr).prev().attr("id","TEMP_TR");//id 값도 변경
    $(idStr).attr("id",prevobj);
    $("#TEMP_TR").attr("id",currobj);
	setSeq();
}
//아이템 밑으로 
function moveDownItem(obj) {     
    var idStr = '#' + obj;
    var nextHtml = $(idStr).next().html();
    if( $(idStr).next().attr("id")  ==  undefined ) {
        alert("최하위 입니다");
        return;
    }
    var nextobj = $(idStr).next().attr("id");
    var currobj = $(idStr).attr("id");
    var currHtml = $(idStr).html();
    $(idStr).next().html(currHtml);
 
    $(idStr).html(nextHtml);//값 변경 
    $(idStr).next().attr("id","TEMP_TR");//id 값도 변경
    $(idStr).attr("id",nextobj);
    $("#TEMP_TR").attr("id",currobj);
    setSeq()
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
	var frm = document.frmList;
	$("form[name='frmList']").find("input[name='chkIdx']").each(function(){
		if($(this).attr("checked")) chk++;
	});
	if(chk==0) {
		alert("수정하실 소재를 선택해주세요.");
		return;
	}

	if(confirm("지정하신 목록의 선택 정보를 저장하시겠습니까?")) {
		frm.action="pop_mobile_addbanner_proc.asp";
		frm.submit();
	}
}

//'아이템 삭제
function slideimgDel(v){
	if (confirm("배너가 삭제됩니다. 삭제 하시겠습니까?")){
		document.frmdel.chkIdx.value = v;
		document.frmdel.submit();
	}
}

function jsSetImg(sFolder, sImg, sName, sSpan){ 
	var winImg;
	winImg = window.open('/admin/eventmanage/common/pop_event_uploadimgV2.asp?yr=<%=Year(now())%>&sF='+sFolder+'&sImg='+sImg+'&sName='+sName+'&sSpan='+sSpan,'popImg','width=370,height=150');
	winImg.focus();
}

function jsDelImg(sName, sSpan){
	if(confirm("이미지를 삭제하시겠습니까?\n\n삭제 후 저장버튼을 눌러야 처리완료됩니다.")){
	   eval("document.all."+sName).value = "";
	   eval("document.all."+sSpan).style.display = "none";
	}
}
</script>
<script type="text/javascript">
//링크값선택
function showDrop(){
	$(".selectLink ul").show();
}

//선택입력
function populateTextBox(v){
	var val = v;
	$("#mblink").val(val);
	$("#blink").val(val);
	$(".selectLink ul").css("display","none");
}

function linkcopy(){
	var val = $("#mblink").val();
	$("#blink").attr("value",val);
	$(".selectLink ul").css("display","none");
}

//달력
function jsPopCal(sName){
	var winCal;
	winCal = window.open('/lib/common_cal.asp?DN='+sName,'popCal','width=250, height=200');
	winCal.focus();
}

function jsPopCal_2(sName,sChkname){
	// if (eval("document.all."+sChkname).checked){
	// 	alert("체크 박스 해제후 변경이 가능합니다");
	// 	return false;
	// }else{
		var winCal;
		winCal = window.open('/lib/common_cal.asp?DN='+sName,'popCal','width=250, height=200');
		winCal.focus();
	// }
}

function simgsubmit(){ 
	// 배너 등록 1 row등록
	var frm = document.slideimgfrm;
	
	if (!frm.gubun[0].checked&&!frm.gubun[1].checked&&!frm.gubun[2].checked){
		alert("위치를 선택해주세요");
		frm.gubun[0].focus();
		return false;
	}

	if (!frm.btitle.value){ alert("Alt값을 입력 해주세요");frm.btitle.focus();return false; }
	if (!frm.bst_date.value){ alert("시작일을 입력 해주세요");frm.bst_date.focus();return false; }
	if (!frm.bed_date.value){ alert("종료일을 입력 해주세요");frm.bed_date.focus();return false; }

	 if(frm.bst_date.value > frm.bed_date.value){ alert("종료일이 시작일보다 빠릅니다. 다시 입력해주세요"); frm.bed_date.focus(); return false; }

	if (confirm('저장 하시겠습니까?')){
		frm.submit();
	}
}

//시작일 종료일 복사
function jscalcopy(s,e,y){
	if (eval("document.all."+y).checked){
		if(confirm("이벤트 기간내 노출 설정을 하시겠습니까?")){
			eval("document.all."+s).value = "<%=sDt%>";
			eval("document.all."+e).value = "<%=eDt%>";
		}
	}
//	else{
//		if(confirm("날짜를 초기화 하시겠습니까?")){
//			eval("document.all."+s).value = "";
//			eval("document.all."+e).value = "";
//		}
//	}
}
</script>
</head>
<body>
<div class="slideRegister adminMob bnrRegister">
	<h1>배너 등록 (MOBILE)</h1>
	<div class="register">
		<dl>
			<dd>
				<form name="slideimgfrm" method="post" action="pop_mobile_addbanner_proc.asp" style="margin:0px;">
				<input type="hidden" name="eventid" value="<%=eCode%>"/>
				<input type="hidden" name="mode" value="SI"/>
				<input type="hidden" name="bimg" value=""/>
				<input type="hidden" name="blink" id="blink" value=""/>
				<input type="hidden" name="sDt" value="<%=sDt%>"/>
				<input type="hidden" name="eDt" value="<%=eDt%>"/>
				<div class="insertImg">
					<table class="tbType1 listTb">
						<colgroup>
							<col width="13%" /><col width="20%" /><col /><col width="28%" /><col width="9%" />
						</colgroup>
						<tbody>
						<tr>
							<td>
								<span><input type="radio" name="gubun" value="1" id="gt"/> <label for="gt">상</label></span>
								<span class="lMar05"><input type="radio" name="gubun" value="2" id="gm"/> <label for="gm">중</label></span>
								<span class="lMar05"><input type="radio" name="gubun" value="3" id="gb"/> <label for="gb">하</label></span>
							</td>
							<td>
								<input class="button" type="button" value="이미지 불러오기" name="mbimg" onClick="jsSetImg('<%=eFolder%>','','bimg','spanslideimg');"/>
								<div id="spanslideimg"></div>
								<div class="tMar10"><input type="text" name="btitle" placeholder="Alt값 입력" /></div>
							</td>
							<td>
								<div class="selectLink">
									<input type="text" value="링크값 입력(선택)" onclick="showDrop();" id="mblink" onkeyup="linkcopy();" />
									<ul style="display:none;">
										<li onclick="populateTextBox('');">선택안함</li>
										<li onclick="populateTextBox('#group그룹코드');">#group그룹코드</li>
										<li onclick="populateTextBox('/event/eventmain.asp?eventid=이벤트코드');">/event/eventmain.asp?eventid=이벤트코드</li>
										<li onclick="populateTextBox('/category/category_itemprd.asp?itemid=상품코드');">/category/category_itemprd.asp?itemid=상품코드 (O)</li>
										<li onclick="populateTextBox('/category/category_list.asp?disp=카테고리');">/category/category_list.asp?disp=카테고리</li>
										<li onclick="populateTextBox('/street/street_brand.asp?makerid=브랜드아이디');">/street/street_brand.asp?makerid=브랜드아이디</li>
										<li onclick="populateTextBox('/playing/view.asp?didx=플레잉번호');">/playing/view.asp?didx=플레잉번호</li>
									</ul>
								</div>
							</td>
							<td>
								<p>시작일 : <input type="text" onclick="jsPopCal('bst_date');" style="width:82px; cursor:pointer;" name="bst_date" readonly> ~ 종료일 : <input type="text" onclick="jsPopCal('bed_date');" style="width:82px; cursor:hand;" name="bed_date" readonly></p>
								<p class="tMar05"><input type="checkbox" name="bdate_flag" value="Y" onclick="jscalcopy('bst_date','bed_date','bdate_flag');"/> 이벤트 기간 내내 노출</p>
							</td>
							<td><input type="button" class="btn" value="등록" onclick="simgsubmit();"></td>
						</tr>
						</tbody>
					</table>
				</div>
				</form>

				<form name="frmList" method="POST" action="" style="margin:0;">
				<input type="hidden" name="mode" value="SU"/>
				<input type="hidden" name="device" value="M"/>
				<input type="hidden" name="eventid" value="<%=eCode%>"/>
				<input type="hidden" name="sDt" value="<%=sDt%>"/>
				<input type="hidden" name="eDt" value="<%=eDt%>"/>
				<div class="tMar20">
					<table class="tbType1 listTb">
						<colgroup>
							<col width="5%" /><col width="13%" /><col width="5%" /><col width="20%" /><col /><col width="28%" /><col width="9%" />
						</colgroup>
						<thead>
						<tr>
							<th>idx</th>
							<th>위치</th>
							<th>순서</th>
							<th>이미지</th>
							<th>링크(선택)</th>
							<th>시작일/종료일</th>
							<th>사용여부</th>							
						</tr>
						</thead>
						<tbody>
<!-- ===================================================상=================================================================-->						
						<tr style="border: 2px solid black">
							<td colspan="7">상</td>
						</tr>		
						<% 
							If eCode <> "" Then 

							sqlStr = "SELECT idx , gubun , bimg , btitle , blink , bdate_flag , bst_date , bed_date "
							sqlStr = sqlStr & " , case when bdate_flag = 'Y' then te.evt_startdate else bst_date end as bst_date "
							sqlStr = sqlStr & " , case when bdate_flag = 'Y' then te.evt_enddate else bed_date end as bed_date "
							sqlStr = sqlStr & " , isusing"
							sqlStr = sqlStr & " from db_event.dbo.tbl_event_mobile_addbanner "
							sqlStr = sqlStr & " CROSS APPLY ( "
							sqlStr = sqlStr & " 				SELECT convert(varchar(10),evt_startdate,120) as evt_startdate , convert(varchar(10),evt_enddate,120) as evt_enddate FROM db_event.dbo.tbl_event where evt_code = '"& eCode &"' "
							sqlStr = sqlStr & " 			) as te"
							sqlStr = sqlStr & " where evt_code = '"& eCode &"' and gubun=1 "
							sqlStr = sqlStr & " order by gubun asc , idx asc " 
							rsget.Open sqlStr,dbget,1
							if Not(rsget.EOF or rsget.BOF) Then
								Do Until rsget.eof
						%>
						<tr id="top<%=rsget("idx")%>" name="trObj" class="<%=chkIIF(rsget("isusing")="N" Or (CStr(Date()) > CStr(rsget("bed_date"))),"bgGry1","")%>" onmouseover="this.style.backgroundColor='#D8D8D8'" onmouseout="this.style.backgroundColor=''">
							<td>
								<%=rsget("idx")%>
								<input type="checkbox" name="chkIdx" value="<%=rsget("idx")%>" style="display:none;"/>
								<input type="hidden" name="seq" value="" />
							</td>							
							<td>
								<span><input type="radio" name="gubun<%=rsget("idx")%>" value="1" <%=chkiif(rsget("gubun")=1,"checked","")%> id="gt<%=rsget("idx")%>"/> <label for="gt<%=rsget("idx")%>">상</label></span>
								<span class="lMar05"><input type="radio" name="gubun<%=rsget("idx")%>" value="2" <%=chkiif(rsget("gubun")=2,"checked","")%> id="gm<%=rsget("idx")%>"/> <label for="gm<%=rsget("idx")%>">중</label></span>
								<span class="lMar05"><input type="radio" name="gubun<%=rsget("idx")%>" value="3" <%=chkiif(rsget("gubun")=3,"checked","")%> id="gb<%=rsget("idx")%>"/> <label for="gb<%=rsget("idx")%>">하</label></span>
							</td>
							<td><button type="button" onclick=moveUpItem('top<%=rsget("idx")%>')>△</button><button type="button" onclick=moveDownItem('top<%=rsget("idx")%>')>▽</button></td>
							<td>
								<input class="button" type="button" value="이미지 불러오기" name="mbimg<%=rsget("idx")%>" onClick="jsSetImg('<%=eFolder%>','','bimg<%=rsget("idx")%>','spanslideimg<%=rsget("idx")%>');"/>
								<input type="hidden" name="bimg<%=rsget("idx")%>" value="<%=rsget("bimg")%>"/><%' 이미지 %>
								<div id="spanslideimg<%=rsget("idx")%>">
									<img src="<%=rsget("bimg")%>" style="width:100px;" alt="<%=rsget("btitle")%>"/>
									<%IF rsget("bimg") <> "" THEN %>
									<a href="javascript:jsDelImg('bimg<%=rsget("idx")%>','spanslideimg<%=rsget("idx")%>');"><img src="/images/icon_delete2.gif" border="0" class="delImg"></a>
									<%END IF%>
								</div>
								<div class="tMar10"><input type="text" name="btitle<%=rsget("idx")%>" value="<%=rsget("btitle")%>" placeholder="Alt값 입력"/></div>
							</td>
							<td><input type="text" name="blink<%=rsget("idx")%>" value="<%=rsget("blink")%>" /></td>
							<td>
								<p>시작일 : <input type="text" onclick="jsPopCal_2('bst_date<%=rsget("idx")%>','bdate_flag<%=rsget("idx")%>');" value="<%=rsget("bst_date")%>" style="width:82px; cursor:pointer;" name="bst_date<%=rsget("idx")%>"> 
								~ 종료일 : <input type="text" onclick="jsPopCal_2('bed_date<%=rsget("idx")%>','bdate_flag<%=rsget("idx")%>');" value="<%=rsget("bed_date")%>" style="width:82px; cursor:pointer;" name="bed_date<%=rsget("idx")%>"></p>
								<p class="tMar05"><input type="checkbox" name="bdate_flag<%=rsget("idx")%>" value="Y" <%=chkiif(rsget("bdate_flag")="Y","checked","")%> onclick="jscalcopy('bst_date<%=rsget("idx")%>','bed_date<%=rsget("idx")%>','bdate_flag<%=rsget("idx")%>');"/> 이벤트 기간 내내 노출</p>
							</td>
							<td>
								<span><input type="radio" name="isusing<%=rsget("idx")%>" <%=chkiif(rsget("isusing")="Y","checked","")%> value="Y"/> Y</span>
								<span class="lMar10"><input type="radio" name="isusing<%=rsget("idx")%>" <%=chkiif(rsget("isusing")="N","checked","")%> value="N"/> N</span>
								<br />
								<input type="button" class="btn tMar05" value="삭제" onclick="slideimgDel(<%=rsget("idx")%>);"/>
							</td>
						</tr>
						<% 
								rsget.movenext
								Loop
							End If
							rsget.close

							End If
						%>						
<!-- ===================================================중=================================================================-->						
						<tr style="border: 2px solid black">
							<td colspan="7">중</td>
						</tr>
						<% 
							If eCode <> "" Then 

							sqlStr = "SELECT idx , gubun , bimg , btitle , blink , bdate_flag , bst_date , bed_date "
							sqlStr = sqlStr & " , case when bdate_flag = 'Y' then te.evt_startdate else bst_date end as bst_date "
							sqlStr = sqlStr & " , case when bdate_flag = 'Y' then te.evt_enddate else bed_date end as bed_date "
							sqlStr = sqlStr & " , isusing"
							sqlStr = sqlStr & " from db_event.dbo.tbl_event_mobile_addbanner "
							sqlStr = sqlStr & " CROSS APPLY ( "
							sqlStr = sqlStr & " 				SELECT convert(varchar(10),evt_startdate,120) as evt_startdate , convert(varchar(10),evt_enddate,120) as evt_enddate FROM db_event.dbo.tbl_event where evt_code = '"& eCode &"' "
							sqlStr = sqlStr & " 			) as te"
							sqlStr = sqlStr & " where evt_code = '"& eCode &"' and gubun=2 "
							sqlStr = sqlStr & " order by gubun asc , idx asc " 
							rsget.Open sqlStr,dbget,1
							if Not(rsget.EOF or rsget.BOF) Then
								Do Until rsget.eof
						%>
						<tr id="mid<%=rsget("idx")%>" name="trObj" class="<%=chkIIF(rsget("isusing")="N" Or (CStr(Date()) > CStr(rsget("bed_date"))),"bgGry1","")%>" onmouseover="this.style.backgroundColor='#D8D8D8'" onmouseout="this.style.backgroundColor=''">
							<td>
								<%=rsget("idx")%>
								<input type="checkbox" name="chkIdx" value="<%=rsget("idx")%>" style="display:none;"/>
								<input type="hidden" name="seq" value="" />
							</td>														
							<td>
								<span><input type="radio" name="gubun<%=rsget("idx")%>" value="1" <%=chkiif(rsget("gubun")=1,"checked","")%> id="gt<%=rsget("idx")%>"/> <label for="gt<%=rsget("idx")%>">상</label></span>
								<span class="lMar05"><input type="radio" name="gubun<%=rsget("idx")%>" value="2" <%=chkiif(rsget("gubun")=2,"checked","")%> id="gm<%=rsget("idx")%>"/> <label for="gm<%=rsget("idx")%>">중</label></span>
								<span class="lMar05"><input type="radio" name="gubun<%=rsget("idx")%>" value="3" <%=chkiif(rsget("gubun")=3,"checked","")%> id="gb<%=rsget("idx")%>"/> <label for="gb<%=rsget("idx")%>">하</label></span>
							</td>
							<td><button type="button" onclick=moveUpItem('mid<%=rsget("idx")%>')>△</button><button type="button" onclick=moveDownItem('mid<%=rsget("idx")%>')>▽</button></td>
							<td>
								<input class="button" type="button" value="이미지 불러오기" name="mbimg<%=rsget("idx")%>" onClick="jsSetImg('<%=eFolder%>','','bimg<%=rsget("idx")%>','spanslideimg<%=rsget("idx")%>');"/>
								<input type="hidden" name="bimg<%=rsget("idx")%>" value="<%=rsget("bimg")%>"/><%' 이미지 %>
								<div id="spanslideimg<%=rsget("idx")%>">
									<img src="<%=rsget("bimg")%>" style="width:100px;" alt="<%=rsget("btitle")%>"/>
									<%IF rsget("bimg") <> "" THEN %>
									<a href="javascript:jsDelImg('bimg<%=rsget("idx")%>','spanslideimg<%=rsget("idx")%>');"><img src="/images/icon_delete2.gif" border="0" class="delImg"></a>
									<%END IF%>
								</div>
								<div class="tMar10"><input type="text" name="btitle<%=rsget("idx")%>" value="<%=rsget("btitle")%>" placeholder="Alt값 입력"/></div>
							</td>
							<td><input type="text" name="blink<%=rsget("idx")%>" value="<%=rsget("blink")%>" /></td>
							<td>
								<p>시작일 : <input type="text" onclick="jsPopCal_2('bst_date<%=rsget("idx")%>','bdate_flag<%=rsget("idx")%>');" value="<%=rsget("bst_date")%>" style="width:82px; cursor:pointer;" name="bst_date<%=rsget("idx")%>"> 
								~ 종료일 : <input type="text" onclick="jsPopCal_2('bed_date<%=rsget("idx")%>','bdate_flag<%=rsget("idx")%>');" value="<%=rsget("bed_date")%>" style="width:82px; cursor:pointer;" name="bed_date<%=rsget("idx")%>"></p>
								<p class="tMar05"><input type="checkbox" name="bdate_flag<%=rsget("idx")%>" value="Y" <%=chkiif(rsget("bdate_flag")="Y","checked","")%> onclick="jscalcopy('bst_date<%=rsget("idx")%>','bed_date<%=rsget("idx")%>','bdate_flag<%=rsget("idx")%>');"/> 이벤트 기간 내내 노출</p>
							</td>
							<td>
								<span><input type="radio" name="isusing<%=rsget("idx")%>" <%=chkiif(rsget("isusing")="Y","checked","")%> value="Y"/> Y</span>
								<span class="lMar10"><input type="radio" name="isusing<%=rsget("idx")%>" <%=chkiif(rsget("isusing")="N","checked","")%> value="N"/> N</span>
								<br />
								<input type="button" class="btn tMar05" value="삭제" onclick="slideimgDel(<%=rsget("idx")%>);"/>
							</td>
						</tr>
						<% 
								rsget.movenext
								Loop
							End If
							rsget.close

							End If
						%>						
<!-- ===================================================하=================================================================-->												
						<tr style="border: 2px solid black">
							<td colspan="7">하</td>
						</tr>		
						<% 
							If eCode <> "" Then 

							sqlStr = "SELECT idx , gubun , bimg , btitle , blink , bdate_flag , bst_date , bed_date "
							sqlStr = sqlStr & " , case when bdate_flag = 'Y' then te.evt_startdate else bst_date end as bst_date "
							sqlStr = sqlStr & " , case when bdate_flag = 'Y' then te.evt_enddate else bed_date end as bed_date "
							sqlStr = sqlStr & " , isusing"
							sqlStr = sqlStr & " from db_event.dbo.tbl_event_mobile_addbanner "
							sqlStr = sqlStr & " CROSS APPLY ( "
							sqlStr = sqlStr & " 				SELECT convert(varchar(10),evt_startdate,120) as evt_startdate , convert(varchar(10),evt_enddate,120) as evt_enddate FROM db_event.dbo.tbl_event where evt_code = '"& eCode &"' "
							sqlStr = sqlStr & " 			) as te "
							sqlStr = sqlStr & " where evt_code = '"& eCode &"' and gubun=3 "
							sqlStr = sqlStr & " order by gubun asc , idx asc " 
							rsget.Open sqlStr,dbget,1
							if Not(rsget.EOF or rsget.BOF) Then
								Do Until rsget.eof
						%>
						<tr id="bot<%=rsget("idx")%>" name="trObj" class="<%=chkIIF(rsget("isusing")="N" Or (CStr(Date()) > CStr(rsget("bed_date"))),"bgGry1","")%>" onmouseover="this.style.backgroundColor='#D8D8D8'" onmouseout="this.style.backgroundColor=''">
							<td>
								<%=rsget("idx")%>
								<input type="checkbox" name="chkIdx" value="<%=rsget("idx")%>" style="display:none;"/>
								<input type="hidden" name="seq" value="" />
							</td>							
							<td>
								<span><input type="radio" name="gubun<%=rsget("idx")%>" value="1" <%=chkiif(rsget("gubun")=1,"checked","")%> id="gt<%=rsget("idx")%>"/> <label for="gt<%=rsget("idx")%>">상</label></span>
								<span class="lMar05"><input type="radio" name="gubun<%=rsget("idx")%>" value="2" <%=chkiif(rsget("gubun")=2,"checked","")%> id="gm<%=rsget("idx")%>"/> <label for="gm<%=rsget("idx")%>">중</label></span>
								<span class="lMar05"><input type="radio" name="gubun<%=rsget("idx")%>" value="3" <%=chkiif(rsget("gubun")=3,"checked","")%> id="gb<%=rsget("idx")%>"/> <label for="gb<%=rsget("idx")%>">하</label></span>
							</td>
							<td><button type="button" onclick=moveUpItem('bot<%=rsget("idx")%>')>△</button><button type="button" onclick=moveDownItem('bot<%=rsget("idx")%>')>▽</button></td>
							<td>
								<input class="button" type="button" value="이미지 불러오기" name="mbimg<%=rsget("idx")%>" onClick="jsSetImg('<%=eFolder%>','','bimg<%=rsget("idx")%>','spanslideimg<%=rsget("idx")%>');"/>
								<input type="hidden" name="bimg<%=rsget("idx")%>" value="<%=rsget("bimg")%>"/><%' 이미지 %>
								<div id="spanslideimg<%=rsget("idx")%>">
									<img src="<%=rsget("bimg")%>" style="width:100px;" alt="<%=rsget("btitle")%>"/>
									<%IF rsget("bimg") <> "" THEN %>
									<a href="javascript:jsDelImg('bimg<%=rsget("idx")%>','spanslideimg<%=rsget("idx")%>');"><img src="/images/icon_delete2.gif" border="0" class="delImg"></a>
									<%END IF%>
								</div>
								<div class="tMar10"><input type="text" name="btitle<%=rsget("idx")%>" value="<%=rsget("btitle")%>" placeholder="Alt값 입력"/></div>
							</td>
							<td><input type="text" name="blink<%=rsget("idx")%>" value="<%=rsget("blink")%>" /></td>
							<td>
								<p>
								시작일 : <input type="text" size=10 onclick="jsPopCal_2('bst_date<%=rsget("idx")%>','bdate_flag<%=rsget("idx")%>');" value="<%=rsget("bst_date")%>" style="width:82px; cursor:pointer;" name="bst_date<%=rsget("idx")%>">
								 ~ 종료일 : <input type="text" onclick="jsPopCal_2('bed_date<%=rsget("idx")%>','bdate_flag<%=rsget("idx")%>');" value="<%=rsget("bed_date")%>" style="width:82px; cursor:pointer;" name="bed_date<%=rsget("idx")%>">
								 </p>
								<p class="tMar05"><input type="checkbox" name="bdate_flag<%=rsget("idx")%>" value="Y" <%=chkiif(rsget("bdate_flag")="Y","checked","")%> onclick="jscalcopy('bst_date<%=rsget("idx")%>','bed_date<%=rsget("idx")%>','bdate_flag<%=rsget("idx")%>');"/> 이벤트 기간 내내 노출</p>
							</td>
							<td>
								<span><input type="radio" name="isusing<%=rsget("idx")%>" <%=chkiif(rsget("isusing")="Y","checked","")%> value="Y"/> Y</span>
								<span class="lMar10"><input type="radio" name="isusing<%=rsget("idx")%>" <%=chkiif(rsget("isusing")="N","checked","")%> value="N"/> N</span>
								<br />
								<input type="button" class="btn tMar05" value="삭제" onclick="slideimgDel(<%=rsget("idx")%>);"/>
							</td>
						</tr>
						<% 
								rsget.movenext
								Loop
							End If
							rsget.close

							End If
						%>																
						</tbody>
					</table>
					<p class="tMar20 ct">
						<!--<input type="button" class="btn" value="전체 선택" onclick="chkAllItem();">-->
						<input type="button" class="btn" value="상태 저장" onClick="saveList();" title="표시순서 및 사용여부를 일괄저장합니다.">
					</p>
				</div>
				</form>
			</dd>
		</dl>
	</div>
</div>
<form name="frmdel" method="POST" action="pop_mobile_addbanner_proc.asp" style="margin:0px;">
<input type="hidden" name="sDt" value="<%=sDt%>"/>
<input type="hidden" name="eDt" value="<%=eDt%>"/>
<input type="hidden" name="eventid" value="<%=eCode%>"/>
<input type="hidden" name="mode" value="SD"/>
<input type="hidden" name="chkIdx" />
</form>
</body>
</html>
<!-- #include virtual="/lib/db/dbclose.asp" -->