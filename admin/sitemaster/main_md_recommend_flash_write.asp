<%@ language=vbscript %>
<% option explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
%>
<%
'####################################################
' Description : PC메인관리 MD픽
' History : 서동석 생성
'			2022.07.01 한용민 수정(isms취약점조치)
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/sitemasterclass/main_event_rotationcls.asp"-->
<%
dim idx,mode
idx = request("idx")
mode = request("mode")
%>
<script language='javascript' src="/js/jsCal/js/jscal2.js"></script>
<script language='javascript' src="/js/jsCal/js/lang/ko.js"></script>
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/jscal2.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/border-radius.css" />
<script type='text/javascript'>
function SubmitForm(){

	if (GetByteLength(document.SubmitFrm.textinfo.value) > 64){
		alert('TEXT 정보는 64자 이하로 입력 해주세요.\n(현재 ' + GetByteLength(document.SubmitFrm.textinfo.value) + '자 입력 됨)');
		document.SubmitFrm.textinfo.focus();
		return;
	}

//	if (document.SubmitFrm.linkinfo.value.length < 1){
//		alert('링크 정보를 입력 하세요');
//		document.SubmitFrm.linkinfo.focus();
//		return;
//	}

	if (document.SubmitFrm.disporder.value.length < 1){
		alert('전시 순서를 입력 하세요');
		document.SubmitFrm.disporder.focus();
		return;
	}

    if (document.SubmitFrm.startdate.value.length!=10){
        alert('시작일을 입력  하세요.');
        return;
    }

    if (document.SubmitFrm.enddate.value.length!=10){
        alert('종료일을 입력  하세요.');
        return;
    }

    var vstartdate = new Date(document.SubmitFrm.startdate.value.substr(0,4), (1*document.SubmitFrm.startdate.value.substr(5,2))-1, document.SubmitFrm.startdate.value.substr(8,2));
    var venddate = new Date(document.SubmitFrm.enddate.value.substr(0,4), (1*document.SubmitFrm.enddate.value.substr(5,2))-1, document.SubmitFrm.enddate.value.substr(8,2));

    if (vstartdate>venddate){
        alert('종료일이 시작일보다 빠르면 안됩니다.');
        return;
    }


	var ret = confirm('저장 하시겠습니까?');
	if (ret) {
		document.SubmitFrm.submit();
	}
}

</script>

<div style="padding:5px 3px; margin:5px 0; font-size:13px;background:#F0F0FF;"><strong>[메인] MD 추천상품 등록/수정</strong></div>
<form name="SubmitFrm" method="post" action="<%=uploadUrl%>/linkweb/doMainMdChoiceRotate.asp" onsubmit="return false;" enctype="multipart/form-data">
<input type="hidden" name="mode" value="<% = mode %>">
<input type="hidden" name="regID" value="<%=session("ssBctId")%>">
<%
	if mode = "modify" then
		dim mdchoicerotate
	
		set mdchoicerotate = new CMainMdChoiceRotate
		mdchoicerotate.FCurrPage = 1
		mdchoicerotate.FPageSize = 1
		mdchoicerotate.read idx
%>
<input type="hidden" name="idx" value="<% = idx %>">
<input type="hidden" name="updateID" value="<%=session("ssBctId")%>">
	<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr bgcolor="#FFFFFF">
		<td width="100" align="center" bgcolor="<%= adminColor("gray") %>">이미지</td>
		<td><input type="file" name="photoimg" value="" size="32" maxlength="32" class="text">
			<br>
			<img src="<%= mdchoicerotate.FItemList(0).Fphotoimg %>" style="max-width:550px; max-height:120px;"><br/>
			<font color="red">(119px × 135px GIF 혹은 JPG 이미지 / ※ 2013리뉴얼: 등록 이미지 없으면 상품이미지 사용)</font>
		</td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td width="100" align="center" bgcolor="<%= adminColor("gray") %>">전시순서</td>
		<td><input type="text" name="disporder" value="<% = mdchoicerotate.FItemList(0).Fdisporder  %>" size="2" class="text">
			<font color="red">(2자리 숫자)</font>
		</td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td width="100" align="center" bgcolor="<%= adminColor("gray") %>">상품코드</td>
		<td><input type="text" name="linkitemid" value="<%= ReplaceBracket(mdchoicerotate.FItemList(0).Flinkitemid)  %>" size="6" class="text"></td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td width="100" align="center" bgcolor="<%= adminColor("gray") %>">Text정보</td>
		<td>
			<textarea name="textinfo" class="textarea" style="width:90%; height:42px;"><%= ReplaceBracket(mdchoicerotate.FItemList(0).Ftextinfo) %></textarea>
		</td>
	</tr>

	<tr bgcolor="#FFFFFF">
		<td width="100" align="center" bgcolor="<%= adminColor("gray") %>">link정보</td>
		<td><input type="text" name="linkinfo" value="<% = mdchoicerotate.FItemList(0).Flinkinfo  %>" size="70" class="text">
			<br>
			<font color="red">(상대경로로 입력하세요 /shopping/category_prd.asp?itemid=72367)</font>
			<br><font color="red">(링크값을 넣지 않으면 상단 상품코드를 기반으로 자동으로 링크값이 대체됩니다.)</font>
		</td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td width="100" align="center" bgcolor="<%= adminColor("gray") %>">최저가 표기</td>
		<td>
			<label><input type="radio" name="lowestPrice" value="Y" <% If trim(mdchoicerotate.FItemList(0).FLowestPrice) = "Y" Then %>checked<% End If %> >사용</label>
			<label><input type="radio" name="lowestPrice" value="N" <% If trim(mdchoicerotate.FItemList(0).FLowestPrice) = "N" or isnull(mdchoicerotate.FItemList(0).FLowestPrice) Then %>checked<% End If %> >사용안함</label>
		</td>
	</tr>		

	<tr bgcolor="#FFFFFF">
	    <td width="150" bgcolor="#DDDDFF">반영시작일</td>
	    <td>
	        <input id="startdate" name="startdate" value="<%= Left(mdchoicerotate.FItemList(0).Fstartdate,10) %>" class="text" size="10" maxlength="10" />
	        <img src="http://webadmin.10x10.co.kr/images/calicon.gif" id="startdate_trigger" border="0" style="cursor:pointer;" align="absbottom" />
	        <input type="text" name="startdatetime" size="2" maxlength="2" value="<%= Format00(2,Hour(mdchoicerotate.FItemList(0).Fstartdate)) %>" />(시 00~23)
	        <input type="text" name="dummy0" value="00:00" size="6" readonly class="text_ro" />
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
	        <input id="enddate" name="enddate" value="<%= Left(mdchoicerotate.FItemList(0).Fenddate,10) %>" class="text" size="10" maxlength="10" />
	        <img src="http://webadmin.10x10.co.kr/images/calicon.gif" id="enddate_trigger" border="0" style="cursor:pointer" align="absbottom" />
	        <input type="text" name="enddatetime" size="2" maxlength="2" value="<%= ChkIIF(mdchoicerotate.FItemList(0).Fenddate="","23",Format00(2,Hour(mdchoicerotate.FItemList(0).Fenddate))) %>">(시 00~23)
	        <input type="text" name="dummy1" value="59:59" size="6" readonly class="text_ro" />
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
		<td width="100" align="center" bgcolor="<%= adminColor("gray") %>">사용여부</td>
		<td>
			<label><input type="radio" name="isusing" value="Y" <%=chkIIF(mdchoicerotate.FItemList(0).FIsUsing="Y" or mdchoicerotate.FItemList(0).FIsUsing="M" ,"checked","")%> >사용함</label>
			<!--<label><input type="radio" name="isusing" value="M" <%=chkIIF(mdchoicerotate.FItemList(0).FIsUsing="M","checked","")%> >PC웹+모바일 사용</label>-->
			<label><input type="radio" name="isusing" value="N" <%=chkIIF(mdchoicerotate.FItemList(0).FIsUsing="N","checked","")%> >사용안함</label>
		</td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td width="100" align="center" bgcolor="<%= adminColor("gray") %>">등록일</td>
		<td><% = mdchoicerotate.FItemList(0).FRegdate  %> (<% = mdchoicerotate.FItemList(0).Fregname  %>) </td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td width="100" align="center" bgcolor="<%= adminColor("gray") %>">최종작업자</td>
		<td><% = mdchoicerotate.FItemList(0).Fworkername  %></td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td colspan="2" align="center">
			<input type="button" value="저장" onClick="SubmitForm()">
		</td>
	</tr>
	</table>
<%
		set mdchoicerotate = Nothing
	else
%>
	<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr bgcolor="#FFFFFF">
		<td width="100" align="center" bgcolor="<%= adminColor("gray") %>">이미지</td>
		<td>
			<input type="file" name="photoimg" value="" size="32" maxlength="32" class="file"><br />
			<font color="red">(119px × 135px GIF 혹은 JPG 이미지 / ※ 2013리뉴얼: 등록 이미지 없으면 상품이미지 사용)</font>
		</td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td width="100" align="center" bgcolor="<%= adminColor("gray") %>">전시순서</td>
		<td><input type="text" name="disporder" value="99" size="2" class="text">
			<font color="red">(2자리 숫자)</font>
		</td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td width="100" align="center" bgcolor="<%= adminColor("gray") %>">상품코드</td>
		<td><input type="text" name="linkitemid" value="" size="6" class="text"></td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td width="100" align="center" bgcolor="<%= adminColor("gray") %>">Text정보</td>
		<td><textarea name="textinfo" class="textarea" style="width:90%; height:42px;"></textarea></td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td width="100" align="center" bgcolor="<%= adminColor("gray") %>">link정보</td>
		<td><input type="text" name="linkinfo" size="70"  class="text">
			<br>
			<font color="red">(상대경로로 입력하세요 /shopping/category_prd.asp?itemid=72367)</font>
			<br><font color="red">(링크값을 넣지 않으면 상단 상품코드를 기반으로 자동으로 링크값이 대체됩니다.)</font>
		</td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td width="100" align="center" bgcolor="<%= adminColor("gray") %>">최저가 표기</td>
		<td>
			<label><input type="radio" name="lowestPrice" value="Y">사용</label>
			<label><input type="radio" name="lowestPrice" value="N" checked >사용안함</label>		
		</td>
	</tr>	

	<tr bgcolor="#FFFFFF">
	    <td width="150" bgcolor="#DDDDFF">반영시작일</td>
	    <td>
	        <input id="startdate" name="startdate" value="" class="text" size="10" maxlength="10" />
	        <img src="http://webadmin.10x10.co.kr/images/calicon.gif" id="startdate_trigger" border="0" style="cursor:pointer;" align="absbottom" />
	        <input type="text" name="startdatetime" size="2" maxlength="2" value="00" />(시 00~23)
	        <input type="text" name="dummy0" value="00:00" size="6" readonly class="text_ro" />
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
	        <input id="enddate" name="enddate" value="" class="text" size="10" maxlength="10" />
	        <img src="http://webadmin.10x10.co.kr/images/calicon.gif" id="enddate_trigger" border="0" style="cursor:pointer" align="absbottom" />
	        <input type="text" name="enddatetime" size="2" maxlength="2" value="23">(시 00~23)
	        <input type="text" name="dummy1" value="59:59" size="6" readonly class="text_ro" />
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
		<td width="100" align="center" bgcolor="<%= adminColor("gray") %>">사용여부</td>
		<td>
			<label><input type="radio" name="isusing" value="Y" checked>사용함</label>
			<!--<label><input type="radio" name="isusing" value="M" checked >PC웹+모바일 사용</label>-->
			<label><input type="radio" name="isusing" value="N">사용안함</label>
		</td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td colspan="2" align="center">
			<input type="button" value="저장" onClick="SubmitForm()">
		</td>
	</tr>
	</table>
<%
	end if
%>
</form>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->