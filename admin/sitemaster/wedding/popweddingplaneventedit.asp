<%@ language=vbscript %>
<% option explicit %>
<%
'###############################################
' Discription : 웨딩 기획전 등록페이지
' History : 2018.04.10 정태훈
'###############################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/classes/sitemasterclass/wedding_ContentsManageCls.asp" -->
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

	dim oPlanEvent
		set oPlanEvent = new CWeddingContents
		oPlanEvent.FRectIdx = idx
		oPlanEvent.GetOnePlanEventContents

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
	    if (frm.Evt_Code.value==""){
	        alert('이벤트 번호를 입력 하세요.');
	        frm.Evt_Code.focus();
	        return;
	    }

		if (frm.Evt_Title.value==""){
	        alert('이벤트 메인카피를 입력 하세요.');
	        frm.Evt_Title.focus();
	        return;
	    }

		if (frm.Evt_Subcopy.value==""){
	        alert('이벤트 서브카피를 입력 하세요.');
	        frm.Evt_Subcopy.focus();
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

	function jsSetImg(sImg, sName, sSpan){ 
		var winImg;
		var sFolder=document.frmcontents.Evt_Code.value;
		if (sFolder=="")
		{
			alert("이벤트 검색 후 이미지를 등록해주세요.");
		}
		else
		{
		winImg = window.open('/admin/eventmanage/common/pop_event_uploadimgV2.asp?yr=<%=Year(now())%>&sF='+sFolder+'&sImg='+sImg+'&sName='+sName+'&sSpan='+sSpan,'popImg','width=370,height=150');
		winImg.focus();
		}
	}

</script>

<table width="100%" cellpadding="2" cellspacing="1" class="a" bgcolor="#3d3d3d">
<form name="frmcontents" method="post" action="doPlanEventReg.asp" onsubmit="return false;">
<input type="hidden" name="weddingban" value="<%=oPlanEvent.FOneItem.FEvt_Img%>">
<tr bgcolor="#FFFFFF">
    <td width="100" bgcolor="#DDDDFF">Idx</td>
    <td>
        <% if oPlanEvent.FOneItem.Fidx<>"" then %>
        <%= oPlanEvent.FOneItem.Fidx %>
        <input type="hidden" name="idx" value="<%= oPlanEvent.FOneItem.Fidx %>">
        <% else %>

        <% end if %>
    </td>
</tr>
<tr bgcolor="#FFFFFF">
    <td width="100" bgcolor="#DDDDFF">기획전</td>
    <td>
		<table>
		<tr bgcolor="#FFFFFF" height="30">
			<td>이벤트 코드 : </td>
			<td><input type="text" name="Evt_Code" value="<%=oPlanEvent.FOneItem.FEvt_Code%>"> <a href="javascript:jsLastEvent(1);">불러오기</a></td>
		</tr>
		<tr bgcolor="#FFFFFF" height="30">
			<td>메인카피 : </td>
			<td><input type="text" name="Evt_Title" value="<%=oPlanEvent.FOneItem.FEvt_Title%>" size="50"></td>
		</tr>
		<tr bgcolor="#FFFFFF" height="30">
			<td>할인율 : </td>
			<td><input type="text" name="Evt_Discount" value="<%=oPlanEvent.FOneItem.FEvt_Discount%>"></td>
		</tr>
		<tr bgcolor="#FFFFFF" height="30">
			<td>쿠폰 할인율 : </td>
			<td><input type="text" name="Evt_Coupon" value="<%=oPlanEvent.FOneItem.FEvt_Coupon%>"></td>
		</tr>
		<tr bgcolor="#FFFFFF" height="30">
			<td>서브카피 : </td>
			<td><input type="text" name="Evt_Subcopy" value="<%=oPlanEvent.FOneItem.FEvt_Subcopy%>" size="70"></td>
		</tr>
		<tr bgcolor="#FFFFFF" height="30">
			<td width="120">PC메인 이미지 등록 : </td>
			<td><input type="button" name="etcitem" value="대표배너등록" onClick="jsSetImg('<%=oPlanEvent.FOneItem.FEvt_Img%>','weddingban','etciitem')" class="button"></td>
		</tr>
		<tr bgcolor="#FFFFFF">
			<td>&nbsp;</td>
			<td>
					<div id="etciitem" style="padding: 5 5 5 5">
						<%IF oPlanEvent.FOneItem.FEvt_Img <> "" THEN %>
						<img  src="<%=oPlanEvent.FOneItem.FEvt_Img%>" width="50%" border="0">
						<a href="javascript:jsDelImg('weddingban','etciitem');"><img src="/images/icon_delete2.gif" border="0"></a>
						<%END IF%>
					</div>
			</td>
		</tr>
		</table>
    </td>
</tr>
<tr bgcolor="#FFFFFF">
  <td width="100" bgcolor="#DDDDFF">시작일</td>
  <td>
  	<input type="text" name="StartDate" id="startdate" value="<%=oPlanEvent.FOneItem.FStartDate%>">
	<img src="http://webadmin.10x10.co.kr/images/calicon.gif" id="startdate_trigger" border="0" style="cursor:pointer;" align="absbottom" />
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
  	<input type="text" name="EndDate" id="enddate" value="<%=oPlanEvent.FOneItem.FEndDate%>">
	<img src="http://webadmin.10x10.co.kr/images/calicon.gif" id="enddate_trigger" border="0" style="cursor:pointer;" align="absbottom" />
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
  	<input type="text" name="DispOrder" value="<%=oPlanEvent.FOneItem.FDispOrder%>">
  </td>
</tr>
<tr bgcolor="#FFFFFF">
  <td width="100" bgcolor="#DDDDFF">사용여부</td>
  <td>
  	<input type="radio" name="Isusing" value="Y"<% If oPlanEvent.FOneItem.FIsusing="Y" Or oPlanEvent.FOneItem.FIsusing="" Then Response.write " checked" %>> 사용함
	<input type="radio" name="Isusing" value="N"<% If oPlanEvent.FOneItem.FIsusing="N" Then Response.write " checked" %>> 사용안함
  </td>
</tr>
<% If oPlanEvent.FOneItem.FRegUser<>"" Then %>
<tr bgcolor="#FFFFFF">
  <td width="100" bgcolor="#DDDDFF">작업자</td>
  <td>
  	작업자 : <%=oPlanEvent.FOneItem.FRegUser %><br>
	최종작업자 : <%=oPlanEvent.FOneItem.FLastUser %>
  </td>
</tr>
<% End If %>
<tr bgcolor="#FFFFFF">
    <td colspan="2" align="center"><input type="button" value=" 저 장 " onClick="SaveMainContents(frmcontents);"></td>
</tr>
</form>
</table>
<%
set oPlanEvent = Nothing
%>

<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
