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
'###############################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/classes/sitemasterclass/main_ContentsManageCls.asp" -->
<!-- #include virtual="/admin/eventmanage/common/event_function.asp"-->
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
	Dim cultureContents
	Dim cultureEcode ,	cultureEtype ,cultureEname ,cultureEcomment , cultureEimagelist

	If idx <> "" And culturecode = "" Then culturecode = oMainContents.FOneItem.Fecode

	If culturecode <> "" Then
		Dim SqlStr
		sqlStr = "SELECT C.evt_code ,C.evt_type , C.evt_name , C.evt_comment , C.image_list" & vbcrlf
		sqlStr = sqlStr &" FROM db_culture_station.dbo.tbl_culturestation_event as C" & vbcrlf
		sqlStr = sqlStr & "WHERE C.evt_code = "& culturecode

        rsget.Open SqlStr, dbget, 1
		if Not rsget.Eof then
			cultureEcode		= rsget("evt_code")
			cultureEtype		= rsget("evt_type")
			cultureEname		= rsget("evt_name")
			cultureEcomment		= rsget("evt_comment")
			'cultureEimagelist	= webImgUrl &"/culturestation/2009/list/" & rsget("image_list")
			cultureEimagelist	= rsget("image_list")
		end if
        rsget.close
	End If

'// 특정 코드에 링크텍스트 추가(IMG ALT 값 등)
dim IsLinkTextNeed
	IsLinkTextNeed = (InStr(",630,642,659,673,674,675,687,", ("," & poscode & ",")) > 0)

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

	    if (frm.linkurl.value.length<1){
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
		}
	}

	function cultureloadpop(){
		winLast = window.open('pop_culturelist.asp?gubun=<%=gubun%>&poscode=<%=poscode%>&pidx=<%=idx%>','pLast','width=1200,height=600, scrollbars=yes')
		winLast.focus();
	}

	function fnSelectBannerType(bannertype){
		if(bannertype==1)
		{
			$("#bnimg2").hide();
			$("#bnalt2").hide();
			$("#bnbg1").hide();
			$("#bnbg2").hide();
			$("#bnlink2").hide();
		}
		else
		{
			$("#bnbg1").show();
			$("#bnbg2").show();
			$("#bnimg2").show();
			$("#bnalt2").show();
			$("#bnlink2").show();
		}
	}
</script>

<table width="100%" cellpadding="2" cellspacing="1" class="a" bgcolor="#3d3d3d">
<form name="frmcontents" method="post" action="<%=uploadUrl%>/linkweb/doMainContentsRegNew.asp" onsubmit="return false;" enctype="multipart/form-data">
<tr bgcolor="#FFFFFF">
    <td width="150" bgcolor="#DDDDFF">Idx</td>
    <td>
        <% if oMainContents.FOneItem.Fidx<>"" then %>
        <%= oMainContents.FOneItem.Fidx %>
        <input type="hidden" name="idx" value="<%= oMainContents.FOneItem.Fidx %>">
        <% else %>

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
		&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<span><a href="" onclick="cultureloadpop();return false;">불러오기</a></span>
		<% End If %>
    </td>
</tr>
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
<tr bgcolor="#FFFFFF">
  <td width="150" bgcolor="#DDDDFF">작업 요청사항</td>
  <td><textarea name="itemDesc" class="textarea" style="width:100%;height:80px;"><%= oMainContents.FOneItem.FitemDesc %></textarea></td>
</tr>
<% If poscode = "706" Then %>
<tr bgcolor="#FFFFFF">
  <td width="150" bgcolor="#DDDDFF">배너 타입</td>
  <td><input type="radio" name="bannertype" value="1"<% If oMainContents.FOneItem.Fbannertype="1" Or oMainContents.FOneItem.Fbannertype="" Then Response.write " checked" %> onclick="fnSelectBannerType(1);">1개&nbsp;&nbsp;<input type="radio" name="bannertype" value="2"<% If oMainContents.FOneItem.Fbannertype="2" Then Response.write " checked" %> onclick="fnSelectBannerType(2);">2개</td>
</tr>
<% End If %>
<%
	'링크 텍스트 여부 확인
	dim chkText: chkText="N"
	IF gubun<>"" Then
		if oMainContents.FOneItem.Fidx<>"" then
			if oMainContents.FOneItem.FLinkType="T" then chkText="Y"
		elseif poscode<>"" then
			if oposcode.FOneItem.FLinkType="T" then chkText="Y"
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
<%
	end if

	if chkText<>"Y" then
%>

	<% If poscode="688" Then %>
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
		<tr bgcolor="#FFFFFF">
		  <td width="150" bgcolor="#DDDDFF">상단 타이틀(bold)</td>
		  <td><input type="text" name="linkText2" value="<%= oMainContents.FOneItem.FlinkText2 %>" size="40" maxlength="128" class="text" /> </td>
		</tr>
		<tr bgcolor="#FFFFFF">
		  <td width="150" bgcolor="#DDDDFF">하단 상품설명</td>
		  <td><input type="text" name="linkText3" value="<%= oMainContents.FOneItem.FlinkText3 %>" size="40" maxlength="128" class="text" /></td>
		</tr>
	<% End If %>
<%'2018 메인 롤링 %>
<% If poscode = "710" Then %>
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
		<input type="radio" name="etctag" value="1" <%=chkiif(oMainContents.FOneItem.Fetctag="1" Or oMainContents.FOneItem.Fetctag="","checked","")%>> 없음 <input type="radio" name="etctag" value="2" <%=chkiif(oMainContents.FOneItem.Fetctag="2","checked","")%>> 할인 <input type="radio" name="etctag" value="3" <%=chkiif(oMainContents.FOneItem.Fetctag="3","checked","")%>> 쿠폰 <br/>
		<input type="radio" name="etctag" value="4" <%=chkiif(oMainContents.FOneItem.Fetctag="4","checked","")%>> GIFT <input type="radio" name="etctag" value="5" <%=chkiif(oMainContents.FOneItem.Fetctag="5","checked","")%>> 1+1 <input type="radio" name="etctag" value="6" <%=chkiif(oMainContents.FOneItem.Fetctag="6","checked","")%>> 런칭 <input type="radio" name="etctag" value="7" <%=chkiif(oMainContents.FOneItem.Fetctag="7","checked","")%>> 참여<br/>
		<input type="text" name="etctext" value="<%=oMainContents.FOneItem.Fetctext%>" size="20" maxlength="30" class="text" />※ 할인,쿠폰 일경우만 입력 하세요<br/>
		※ 한가지만 선택 하세요.
	</td>
</tr>
<% End If %>
<%'2018 컬쳐스테이션%>
<% If poscode="714" Then %>
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
	<input type="file" name="file1" value="" size="32" maxlength="32" class="file">
	<% End If %>
  <% if oMainContents.FOneItem.Fidx<>"" then %>
  <br>
  <img src="<%= oMainContents.FOneItem.GetImageUrl %>" style="max-width:600px;" />
  <br> <%= oMainContents.FOneItem.GetImageUrl %>
  <% end if %>
  <% '컬쳐스테이션 %>
  <% If oMainContents.FOneItem.Fidx = "" And poscode = "714" Then %>
  <br>
  <img src="<%= webImgUrl &"/culturestation/2009/list/" & cultureEimagelist %>" style="max-width:600px;" />
  <br> <%= webImgUrl &"/culturestation/2009/list/" & cultureEimagelist %> <br/><br/> ※ 이미지 수정은 컬쳐스테이션 어드민에서 해주세요
  <% ElseIf oMainContents.FOneItem.Fidx <> "" And poscode = "714" Then %>
  <br>
  <img src="<%= webImgUrl &"/culturestation/2009/list/" & cultureEimagelist %>" style="max-width:600px;" />
  <br> <%= webImgUrl &"/culturestation/2009/list/" & cultureEimagelist %> <br/><br/> ※ 이미지 수정은 컬쳐스테이션 어드민에서 해주세요
  <% End If %>
  </td>
</tr>
<tr bgcolor="#FFFFFF">
  <td width="150" bgcolor="#DDDDFF">알트명1 (필수)</td>
  <td><input type="text" name="altname" value="<%=oMainContents.FOneItem.Faltname%>" size="20" maxlength="20"> </td>
</tr>
<% If poscode = "706" Then %>
<tr bgcolor="#FFFFFF" id="bnimg2" style="display:none">
  <td width="150" bgcolor="#DDDDFF">이미지2</td>
  <td>
	<input type="file" name="file2" value="" size="32" maxlength="32" class="file">
  <% if oMainContents.FOneItem.GetImageUrl2<>"" then %>
  <br>
  <img src="<%= oMainContents.FOneItem.GetImageUrl2 %>" style="max-width:600px;" />
  <br> <%= oMainContents.FOneItem.GetImageUrl2 %>
  <% end if %>
  </td>
</tr>
<tr bgcolor="#FFFFFF" id="bnalt2" style="display:none">
  <td width="150" bgcolor="#DDDDFF">알트명2 (필수)</td>
  <td><input type="text" name="altname2" value="<%=oMainContents.FOneItem.Faltname2%>" size="20" maxlength="20"> </td>
</tr>
<% End If %>
<% If gubun <> "PCbanner" and gubun <> "MAbanner" And poscode <> "706" Then %>
<tr bgcolor="#FFFFFF">
  <td width="150" bgcolor="#DDDDFF">추가 이미지 (선택)</td>
  <td><input type="file" name="file2" value="" size="32" maxlength="32" class="file">
  <% if oMainContents.FOneItem.Fidx<>"" then %>
  <br>
  <img src="<%= oMainContents.FOneItem.GetImageUrl2 %>" style="max-width:600px;" />
  <br> <%= oMainContents.FOneItem.GetImageUrl2 %>
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
		  <% else %>
		        <% if poscode<>"" then %>
		        <%= oposcode.FOneItem.Fimagewidth %>
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
<% If poscode = "706" Then %>
<tr bgcolor="#FFFFFF" id="bnlink2" style="display:none">
    <td width="150" bgcolor="#DDDDFF">링크값2</td>
    <td>
		<input type="text" name="linkurl2" value="<%= oMainContents.FOneItem.Flinkurl2 %>" maxlength="128" style="width:100%;" class="text">
    </td>
</tr>
<tr bgcolor="#FFFFFF">
    <td width="150" bgcolor="#DDDDFF">좌우측 BG컬러코드</td>
	<td><span  id="bnbg1" style="display:none">좌 : </span>#<input type="text" name="bgcode" value="<%=oMainContents.FOneItem.Fbgcode%>" size="20" maxlength="6">
		<div  id="bnbg2" style="display:none">우 : #<input type="text" name="bgcode2" value="<%=oMainContents.FOneItem.Fbgcode2%>" size="20" maxlength="6"></div>
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
<tr bgcolor="#FFFFFF">
    <td width="150" bgcolor="#DDDDFF">반영시작일</td>
    <td>
        <input id="startdate" name="startdate" value="<%= Left(oMainContents.FOneItem.Fstartdate,10) %>" class="text" size="10" maxlength="10" />
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
        <input id="enddate" name="enddate" value="<%= Left(oMainContents.FOneItem.Fenddate,10) %>" class="text" size="10" maxlength="10" />
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
