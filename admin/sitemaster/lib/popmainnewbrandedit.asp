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
'           2019.10.30 정태훈 : 브랜드 정보 불러오기로 수정
'###############################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/classes/sitemasterclass/main_enjoyContentsManageCls.asp" -->
<!-- #include virtual="/admin/eventmanage/common/event_function.asp"-->
<%
Dim isusing, fixtype, validdate, prevDate
Dim idx, poscode, reload, gubun, edid
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

	Response.write culturecode

	if reload="on" then
	    response.write "<script>opener.location.reload(); window.close();</script>"
	    dbget.close()	:	response.End
	end if

	dim oMainContents
		set oMainContents = new CMainEnjoyContents
		oMainContents.FRectIdx = idx
		oMainContents.GetOneNewBrandMainContents

	If gubun = "" Then
		gubun = "index"
	End If

%>
<script language='javascript' src="/js/jsCal/js/jscal2.js"></script>
<script language='javascript' src="/js/jsCal/js/lang/ko.js"></script>
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/jscal2.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/border-radius.css" />
<script language='javascript'>

	function SaveMainContents(frm){

		if (frm.BrandID.value==""){
	        alert('브랜드 ID를 입력 하세요.');
	        frm.BrandID.focus();
	        return;
	    }
		
		if (frm.MainCopy.value==""){
	        alert('메인카피를 입력 하세요.');
	        frm.MainCopy.focus();
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

	    if (frm.startdate.value>frm.enddate.value){
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

	function cultureloadpop(){
		winLast = window.open('pop_culturelist.asp','pLast','width=1200,height=600, scrollbars=yes')
		winLast.focus();
	}

	//색상코드 선택
	function selColorChip(bg,cd) {
		var i;
		document.frmcontents.BGColor.value= bg;
		for(i=1;i<=11;i++) {
			document.all("cline"+i).bgColor='#DDDDDD';
		}
		if(!cd) document.all("cline0").bgColor='#DD3300';
		else document.all("cline"+cd).bgColor='#DD3300';
	}

	//-- jsLastEvent : 지난 이벤트 불러오기 --//
	function jsLastEvent(num){
	  winLast = window.open('pop_event_lastlist.asp?num='+num,'pLast','width=800,height=600, scrollbars=yes')
	  winLast.focus();
	}

	//-- jsImgView : 이미지 확대화면 새창으로 보여주기 --//
	function jsImgView(sImgUrl){
	 var wImgView;
	 wImgView = window.open('/admin/eventmanage/common/pop_event_detailImg.asp?sUrl='+sImgUrl,'pImg','width=100,height=100');
	 wImgView.focus();
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

	//브랜드 ID 검색 팝업창
	function jsSearchBrandIDNew(frmName,compName){
		var compVal = "";
		try{
			compVal = eval("document.all." + frmName + "." + compName).value;
		}catch(e){
			compVal = "";
		}

		var popwin = window.open("popBrandSearch.asp?frmName=" + frmName + "&compName=" + compName + "&rect=" + compVal,"popBrandSearch","width=800 height=400 scrollbars=yes resizable=yes");

		popwin.focus();
	}
</script>

<table width="100%" cellpadding="2" cellspacing="1" class="a" bgcolor="#3d3d3d">
<form name="frmcontents" method="post" action="doMainNewBrandReg.asp" onsubmit="return false;">
<input type="hidden" name="Main_Image" value="<%=oMainContents.FOneItem.FMain_Image%>">
<tr bgcolor="#FFFFFF">
    <td width="100" bgcolor="#DDDDFF">Idx</td>
    <td>
        <% if oMainContents.FOneItem.Fidx<>"" then %>
        <%= oMainContents.FOneItem.Fidx %>
        <input type="hidden" name="idx" value="<%= oMainContents.FOneItem.Fidx %>">
        <% else %>

        <% end if %>
    </td>
</tr>
<tr bgcolor="#FFFFFF">
    <td width="100" bgcolor="#DDDDFF">브랜드</td>
    <td>
		<% NewDrawSelectBoxDesignerwithNameEvent "BrandID", oMainContents.FOneItem.FBrandID %>
    </td>
</tr>
<tr bgcolor="#FFFFFF">
    <td width="100" bgcolor="#DDDDFF">메인카피</td>
    <td height="60">
		<textarea name="MainCopy" rows="5" cols="30"><%=oMainContents.FOneItem.FMainCopy%></textarea>
    </td>
</tr>
<tr bgcolor="#FFFFFF" id="tmpstyle1">
	<td bgcolor="#DDDDFF" align="center" width="15%">브랜드 이미지</td>
	<td>
		<div id="newbrandimg" style="padding: 5 5 5 5">
			<%IF oMainContents.FOneItem.FMain_Image <> "" THEN %>
			<a href="javascript:jsImgView('<%=oMainContents.FOneItem.FMain_Image%>')"><img src="<%=oMainContents.FOneItem.FMain_Image%>" width="400" border="0" id="mainimg"></a>
			<% else %>
			<img src="" width="400" border="0" id="mainimg">
			<%END IF%>
		</div>
		<span id="imgurl"><%=oMainContents.FOneItem.FMain_Image%></span>
	</td>
</tr>
<tr bgcolor="#FFFFFF">
  <td width="100" bgcolor="#DDDDFF">시작일</td>
  <td>
  	<input type="text" name="StartDate" id="startdate" value="<%=chkiif(idx=0,prevDate,oMainContents.FOneItem.FStartDate)%>">
	<img src="http://scm.10x10.co.kr/images/calicon.gif" id="startdate_trigger" border="0" style="cursor:pointer;" align="absbottom" />
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
  <td width="100" bgcolor="#DDDDFF">종료일</td>
  <td>
  	<input type="text" name="EndDate" id="enddate" value="<%=chkiif(idx=0,prevDate,oMainContents.FOneItem.FEndDate)%>">
	<img src="http://scm.10x10.co.kr/images/calicon.gif" id="enddate_trigger" border="0" style="cursor:pointer;" align="absbottom" />
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
  <td width="100" bgcolor="#DDDDFF">우선순위</td>
  <td>
  	<input type="text" name="DispOrder" value="<%=oMainContents.FOneItem.FDispOrder%>">
  </td>
</tr>
<tr bgcolor="#FFFFFF">
  <td width="100" bgcolor="#DDDDFF">사용여부</td>
  <td>
  	<input type="radio" name="Isusing" value="Y"<% If oMainContents.FOneItem.FIsusing="Y" Or oMainContents.FOneItem.FIsusing="" Then Response.write " checked" %>> 사용함
	<input type="radio" name="Isusing" value="N"<% If oMainContents.FOneItem.FIsusing="N" Then Response.write " checked" %>> 사용안함
  </td>
</tr>
<% If oMainContents.FOneItem.FRegUser<>"" Then %>
<tr bgcolor="#FFFFFF">
  <td width="100" bgcolor="#DDDDFF">작업자</td>
  <td>
  	작업자 : <%=oMainContents.FOneItem.FRegUser %><br>
	최종작업자 : <%=oMainContents.FOneItem.FLastUser %>
  </td>
</tr>
<% End If %>
<tr bgcolor="#FFFFFF">
    <td colspan="2" align="center"><input type="button" value=" 저 장 " onClick="SaveMainContents(frmcontents);"></td>
</tr>
</form>
</table>
<%
set oMainContents = Nothing
%>

<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
