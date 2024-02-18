<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  이벤트
' History : 2010.09.17 한용민 수정
'           2012.02.13 허진원- 미니달력 교체
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/academy/lib/academy_function.asp"-->
<!-- #include virtual="/academy/lib/classes/Event_cls.asp"-->
<%
dim evtId ,page, searchKey, searchString ,oEvent, i, lp , evtDivCd , evtTitle , lecturerID , evtType ,contImage, contImage2
dim evtCont , isComment , evtSdate , evtedate ,prizeDate ,listImage , esale,egift,ecoupon , elktype ,elkurl
	evtId = RequestCheckvar(request("evtId"),10)
	page = RequestCheckvar(request("page"),10)
	searchKey = RequestCheckvar(request("searchKey"),16)
	searchString = RequestCheckvar(request("searchString"),128)

	if page="" then page=1
	if searchKey="" then searchKey="evtTitleLong"
	IF elktype="" Then elktype="E" '//링크타입 기본값 설정
		
'// 내용 접수
set oEvent = new CEvent
	oEvent.FRectevtId = evtId
	
	'//수정일 경우에만 쿼리
	if evtId <> "" then 
		oEvent.GetNoitceRead()

		elkurl = oEvent.FEventList(0).FELinkURL
		elktype	= oEvent.FEventList(0).FELinkType			
		evtDivCd = oEvent.FEventList(0).FevtDivCd
		evtTitle = db2html(oEvent.FEventList(0).FevtTitle)
		evtDivCd = oEvent.FEventList(0).FevtDivCd
		lecturerID = oEvent.FEventList(0).FlecturerID
		evtType = oEvent.FEventList(0).FevtType
		contImage = oEvent.FEventList(0).FcontImage
		contImage2 = oEvent.FEventList(0).FcontImage2
		evtCont = db2html(oEvent.FEventList(0).FevtCont)
		isComment = oEvent.FEventList(0).FisComment
		evtSdate = oEvent.FEventList(0).FevtSdate
		evtedate = oEvent.FEventList(0).Fevtedate
		prizeDate = oEvent.FEventList(0).FprizeDate
		listImage = oEvent.FEventList(0).FlistImage
		esale = oEvent.FEventList(0).fissale
		egift = oEvent.FEventList(0).fisgift
		ecoupon = oEvent.FEventList(0).fiscoupon
	end if

if evtType = "" then evtType = "M"

public Sub SelectLecturerId(byval lecturer_id)
	dim sqlStr,i
	sqlStr = "select  c.userid,p.company_name,c.defaultmargine, c.regdate"
	sqlStr = sqlStr + " from [db_user].[dbo].tbl_user_c c"
	sqlStr = sqlStr + " left join [db_partner].[dbo].tbl_partner p on c.userid=p.id"
	sqlStr = sqlStr + " where c.userid<>''" + vbCrlf
	sqlStr = sqlStr + " and c.userdiv < 22" + vbcrlf
	sqlStr = sqlStr + " and c.userdiv='14'" + vbcrlf
	
	'response.write sqlStr & "<Br>"
	rsget.open sqlStr,dbget,1

	if not rsget.eof then
			response.write "<select name='lecturerID'>"
			response.write "<option value=''>선택</option>"
		for i=0 to rsget.recordcount-1
			if lecturer_id=db2html(rsget("userid")) then
			response.write "<option value='" & db2html(rsget("userid")) & "' selected>" & db2html(rsget("userid")) & "(" & db2html(rsget("company_name")) & ")</option>"
			else
			response.write "<option value='" & db2html(rsget("userid")) & "'>" & db2html(rsget("userid")) & "(" & db2html(rsget("company_name")) & ")</option>"
			end if
		rsget.movenext
		next
			response.write "</select>"
	end if
	rsget.close
end sub
%>
<script language='javascript' src="/js/jsCal/js/jscal2.js"></script>
<script language='javascript' src="/js/jsCal/js/lang/ko.js"></script>
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/jscal2.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/border-radius.css" />
<script language='javascript'>

	// 입력폼 검사
	function chk_form(frm)
	{
		if(!frm.evtDivCd.value)
		{
			alert("이벤트 구분을 선택해주십시오.");
			frm.evtDivCd.focus();
			return false;
		}

		if(!frm.evtTitle.value)
		{
			alert("제목을 입력해주십시오.");
			frm.evtTitle.focus();
			return false;
		}

		// 폼 전송
		return true;
	}

	// 이벤트 구분 변경
	function chgEvtDiv(dv) {
		if(dv=="J020") {
			//무료강좌 이벤트
			document.all.lyrLecUID.style.display='';
		} else {
			//일반 이벤트
			document.all.lyrLecUID.style.display='none';
		}
		if(dv == "J040"){
			document.all.lyrLecEvttype.style.display='none';
		}else{
			document.all.lyrLecEvttype.style.display='';
		}
	}

	// 이벤트 형태 변경
	function chgEvtType(tp) {
		if(tp=="M") {
			document.all.lyrImage.style.display='';
			document.all.lyrImage2.style.display='';
			document.all.lyrtext.style.display='';
			document.all.lyrTitle.innerHTML='Image Map';
			document.all.lyrgroup.style.display='none';
		} else if(tp=="H") {
			document.all.lyrImage.style.display='none';
			document.all.lyrImage2.style.display='none';
			document.all.lyrtext.style.display='';
			document.all.lyrTitle.innerHTML='HTML';
			document.all.lyrgroup.style.display='none';
		} else {
			iframG.location.href = "iframe_eventitem_group.asp?eC=<%=evtId%>";
			document.all.lyrImage.style.display='none';
			document.all.lyrImage2.style.display='none';
			document.all.lyrtext.style.display='none';
			document.all.lyrgroup.style.display='';
		}
	}

	// 배너 링크설정 Eable
	function jsEvtLink(bln){
		var d = document.getElementById('elurl');

		if (bln) {
			d.readOnly=true;
			d.className ="text_ro";
		}else{
			d.readOnly=false;
			d.className="text";
		}
	}
	
</script>

<!-- 쓰기 화면 시작 -->
<table width="100%" border="0" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
<form name="frm_write" method="POST" onSubmit="return chk_form(this)" action="<%=UploadImgFingers%>/linkweb/doEvent.asp" enctype="multipart/form-data">
<input type="hidden" name="evtId" value="<%=evtId%>">
<input type="hidden" name="mode" value="write">
<input type="hidden" name="menupos" value="<%=menupos%>">
<input type="hidden" name="page" value="<%=page%>">
<input type="hidden" name="searchKey" value="<%=searchKey%>">
<input type="hidden" name="searchString" value="<%=searchString%>">
<tr align="center" bgcolor="#F0F0FD">
	<td colspan="2" height="26" align="left"><b>이벤트 내용 수정</b></td>
</tr>
<tr>
	<td align="center" width="120" bgcolor="<%= adminColor("tabletop") %>"><font color="darkred">* </font>이벤트 코드</td>
	<td bgcolor="#FFFFFF">
		<%=evtId%>
	</td>
</tr>
<tr>
	<td align="center" width="120" bgcolor="<%= adminColor("tabletop") %>"><font color="darkred">* </font>이벤트 구분</td>
	<td bgcolor="#FFFFFF">
		<select name="evtDivCd" onchange="chgEvtDiv(this.value)">
			<option value="">::선택::</option>
			<% call sbOptCommCd(evtDivCd,"J000") %>
		</select>	
	</td>
</tr>
<tr>
	<td align="center" width="120" bgcolor="<%= adminColor("tabletop") %>"><font color="darkred">* </font>제목</td>
	<td bgcolor="#FFFFFF"><input type="text" name="evtTitle" size="40" maxlength="40" value="<%=evtTitle%>"></td>
</tr>
<tr>
	<td align="center"  bgcolor="<%= adminColor("tabletop") %>">이벤트 링크 타입</td>
	<td bgcolor="#FFFFFF">
		<label><input type="radio" name="elType" value="E" onclick="jsEvtLink(true);"  <% IF elktype="E" Then %>checked<% End IF %> >이벤트</label>
		<label><input type="radio" name="elType" value="I" onclick="jsEvtLink(false);" <% IF elktype="I" Then %>checked<% End IF %>>직접입력</label>
		&nbsp;<input type="text" name="elUrl" size="40" maxlength="128" value="<%= elkurl %>" <% IF elktype="E" THEN%>class="text_ro" readOnly<%ELSE%>class="text"<%END IF %>>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>" id="lyrLecUID" name="lyrLecUID" <% if evtDivCd <>"J020" then %>style="display:none;"<% end if %>>
	<td width="120">대상 강사</td>
	<td bgcolor="#FFFFFF" align="left">
		<% SelectLecturerId(lecturerID) %>
	</td>
</tr>
<tr>
	<td align="center" width="120" bgcolor="<%= adminColor("tabletop") %>">내용</td>
	<td bgcolor="#FFFFFF">
		<table width="100%" cellpadding="0" cellspacing="2" border="0" class="a">
		<tr>
			<td width="100" bgcolor="#F3F0F8" align="center">이벤트형태</td>
			<td>
				<input type="radio" name="evtType" value="M" <% if evtType = "M" then Response.Write "checked" %> onclick="chgEvtType(this.value)">일반형태
				<input type="radio" name="evtType" value="H" <% if evtType = "H" then Response.Write "checked" %> onclick="chgEvtType(this.value)">수작업 형태
				<input type="radio" name="evtType" value="G" <% if evtType = "G" then Response.Write "checked" %> onclick="chgEvtType(this.value)">그룹 형태
			</td>
		</tr>
		<tr id="lyrImage" name="lyrImage" <% if evtType = "H" or evtType = "G" then %>style="display:none"<% end if %>>
			<td bgcolor="#F3F0F8" align="center">메인 이미지</td>
			<td>
				<input type="file" name="contImage" size="60">
				<%
					if Not(contImage = "" or isNull(contImage)) then
						Response.Write "<font color=gray><br>◎ 현재 : " & contImage & "</font>"
					end if
					response.write "<br><font color=red>일반 이벤트 이미지는 가로 960px로 맞춰주세요</font><br><font color=red>DIY Book의 이미지는 가로 758px로 맞춰주세요</font>"
				%>
			</td>
		</tr>

		<tr id="lyrImage2" name="lyrImage2">
			<td bgcolor="#F3F0F8" align="center">코맨트 이미지</td>
			<td>
				<input type="file" name="contImage2" size="60">
				<%
					if Not(contImage2 = "" or isNull(contImage2)) then
						Response.Write "<font color=gray><br>◎ 현재 : " & contImage2 & "</font>"
					end if
					response.write "<br><font color=red>일반 이벤트 이미지는 가로 960px로 맞춰주세요</font><br><font color=red>DIY Book의 이미지는 가로 758px로 맞춰주세요</font>"
				%>
			</td>
		</tr>
		<tr id="lyrtext" name="lyrtext" <% if evtType = "G" then %>style="display:none"<% end if %>>
			<td id="lyrTitle" name="lyrTitle" bgcolor="#F3F0F8" align="center"><% if evtType = "M" then Response.Write "Image Map": else Response.Write "HTML": end if %></td>
			<td><textarea name="evtCont" rows="14" cols="80"><%=evtCont%></textarea><br>map name="evtMainImg" 입니다</td>
		</tr>
		<tr id="lyrgroup" name="lyrgroup" <% if evtType = "M" or evtType = "H" then %>style="display:none"<% end if %>>
			<td bgcolor="#F3F0F8" align="center">그룹형</td>
			<td>
				<iframe id="iframG" src="about:blank" frameborder="0" width="100%" onload="this.style.height=this.contentWindow.document.body.scrollHeight+50;"></iframe>
			</td>
		</tr>		
		<tr>
			<td width="100" bgcolor="#F3F0F8" align="center">선택사항</td>
			<td>
				<input type="radio" name="isComment" value="1" <% if isComment then Response.Write "checked" %>>코멘트 사용
				<input type="radio" name="isComment" value="0" <% if Not(isComment) then Response.Write "checked" %>>코멘트 사용안함
			</td>
		</tr>
   		<tr id="lyrLecEvttype" name="lyrLecEvttype" <% if evtDivCd ="J040" then %>style="display:none;"<% end if %>>
	   		<td width="100" bgcolor="#F3F0F8" align="center">이벤트 타입</td>
	   		<td bgcolor="#FFFFFF">
		   		<input type="checkbox" name="chSale" <%IF esale THEN%>checked<%END IF%> value="1">할인
		   		<input type="checkbox" name="chGift" <%IF egift  THEN%>checked<%END IF%> value="1">사은품
		   		<input type="checkbox" name="chCoupon" <%IF ecoupon THEN%>checked<%END IF%> value="1">쿠폰		   		
	   		</td>
		</tr>		
		</table>
	</td>
</tr>
<tr>
	<td align="center" width="120" bgcolor="<%= adminColor("tabletop") %>"><font color="darkred">* </font>기간</td>
	<td bgcolor="#FFFFFF">
        <input id="evtSdate" name="evtSdate" value="<%=evtsdate%>" class="text" size="10" maxlength="10" />
        <img src="http://webadmin.10x10.co.kr/images/calicon.gif" id="evtSdate_trigger" border="0" style="cursor:pointer" align="absmiddle" /> ~
        <input id="evtEdate" name="evtEdate" value="<%=evtEdate%>" class="text" size="10" maxlength="10" />
        <img src="http://webadmin.10x10.co.kr/images/calicon.gif" id="evtEdate_trigger" border="0" style="cursor:pointer" align="absmiddle" />
	</td>
</tr>
<tr>
	<td align="center" width="120" bgcolor="<%= adminColor("tabletop") %>"><font color="darkred">* </font>당첨자 발표일</td>
	<td bgcolor="#FFFFFF">
        <input id="prizeDate" name="prizeDate" value="<%=prizeDate%>" class="text" size="10" maxlength="10" />
        <img src="http://webadmin.10x10.co.kr/images/calicon.gif" id="prizeDate_trigger" border="0" style="cursor:pointer" align="absmiddle" />
		<script language="javascript">
		var CAL_Start = new Calendar({
			inputField : "evtSdate", trigger    : "evtSdate_trigger",
			onSelect: function() {
				var date = Calendar.intToDate(this.selection.get());
				CAL_End.args.min = date;
				CAL_End.redraw();
				this.hide();
			}, bottomBar: true, dateFormat: "%Y-%m-%d"
		});
		var CAL_End = new Calendar({
			inputField : "evtEdate", trigger    : "evtEdate_trigger",
			onSelect: function() {
				var date = Calendar.intToDate(this.selection.get());
				CAL_Start.args.max = date;
				CAL_Start.redraw();
				this.hide();
			}, bottomBar: true, dateFormat: "%Y-%m-%d"
		});
		var CAL_Prize = new Calendar({
			inputField : "prizeDate", trigger    : "prizeDate_trigger",
			onSelect: function() { this.hide(); }, bottomBar: true, dateFormat: "%Y-%m-%d"
		});
		</script>
	</td>
</tr>
<tr>
	<td align="center" width="120" bgcolor="<%= adminColor("tabletop") %>">목록 이미지</td>
	<td bgcolor="#FFFFFF">
		<input type="file" name="listImage" size="60" class="button">
		<%
			if Not(listImage = "" or isNull(listImage)) then
		%>		
			<br><img src='<%=imgFingers%>/contents/event/<%=listImage%>'>
		<%
			end if
		%>
	</td>
</tr>
<tr><td height="1" colspan="2" bgcolor="#D0D0D0"></td></tr>
<tr>
	<td colspan="2" height="32" bgcolor="#FAFAFA" align="center">
		<input type="image" src="/images/icon_save.gif" style="border:0px;cursor:pointer" align="absmiddle"> &nbsp;
		<img src="/images/icon_cancel.gif" onClick="history.back()" style="cursor:pointer" align="absmiddle">
	</td>
</tr>
</form>
</table>

<% 
set oEvent = nothing
%>

<script language="javascript">
	chgEvtType('<%=evtType%>');
</script>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->