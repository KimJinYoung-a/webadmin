<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 설문관리
' Hieditor : 허진원 생성
'			 2022.07.08 한용민 수정(isms취약점보안조치, 표준코드로변경)
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<%
	Dim lp, srv_sn

	srv_sn = Request("ssn")
%>
<script type='text/javascript'>
<!--
	function chgQstType(tp) {
		if(tp=="1") {
			document.getElementById("trQstPoll").style.display="";
			document.getElementById("trQstDiv").style.display="none";
		} else if(tp=="9") {
			document.getElementById("trQstPoll").style.display="none";
			document.getElementById("trQstDiv").style.display="";
		} else {
			document.getElementById("trQstPoll").style.display="none";
			document.getElementById("trQstDiv").style.display="none";
		}
	}

	var total_link = 3;
	function fnAddPoll() {
		var oRow1 = tbl_poll.insertRow();
		var oRow2 = tbl_poll.insertRow();
		oRow1.style.backgroundColor="#FFFFFF";
		oRow1.style.textAlign="center";
		oRow2.style.backgroundColor="#FFFFFF";
		
		var oCell1 = oRow1.insertCell();
			oCell1.rowSpan = 2;
		var oCell2 = oRow1.insertCell();
			oCell2.colSpan = 2;
		var oCell3 = oRow2.insertCell();
		var oCell4 = oRow2.insertCell();
		
		oCell1.innerHTML = '지문 #'+total_link + '<input type="hidden" name="poll_sn" value="" />';
		oCell2.innerHTML = '<textarea name="poll_content" class="textarea" style="width:100%; height:32px;"></textarea>';
		oCell3.innerHTML = '추가의견 <select name="poll_isAddAnswer" class="select"><option value="N" selected >없음</option><option value="Y">있음</option></select>';
		oCell4.innerHTML = '관련문항 번호 : <input type="text" name="link_qst_sn" size="4" class="text">';

		total_link++;
	}

	//폼 실행
	function fnQstSubmit() {
		var frm = document.frm_Qst;
		if(!frm.qst_type.value) {
			alert("문항 형태를 선택해주세요.");
			frm.qst_type.focus();
			return;
		}

		if(frm.qst_content.value.length<2) {
			alert("문항 내용을 작성해주세요.");
			frm.qst_content.focus();
			return;
		}

		// 객관식 내용 확인
		if(frm.qst_type.value=="1") {
			var chkPollCnt=0;
			for(var i=0;i<frm.poll_content.length;i++) {
				if(frm.poll_content[i].value) chkPollCnt++;
			}
			if(chkPollCnt<2) {
				alert("지문의 내용을 입력해주세요.\n※지문은 최소 2개이상 등록해야됩니다.");
				return;
			}
		}

		if(confirm("입력한 내용으로 문항을 등록하시겠습니까?")) {
			frm.submit();
		} else {
			return;
		}
	}
//-->
</script>
<!-- 입력테이블 시작 -->
<form name="frm_Qst" method="POST" action="survey_qst_process.asp" onsubmit="return false;">
<input type="hidden" name="mode" value="qAdd" />
<input type="hidden" name="srv_sn" value="<%=srv_sn%>" />
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td colspan="2" bgcolor="#DDDDFF" align="left"><img src="/images/icon_star.gif" align="absmiddle"><b>설문 문항 등록</b></td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td width="20%" bgcolor="#EEEEEE">설문 번호</td>
	<td width="80%" align="left"><b><%=srv_sn%></b></td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td bgcolor="#EEEEEE">문항형태</td>
	<td align="left">
		<select name="qst_type" class="select" onchange="chgQstType(this.value)">
			<option value="">::형태선택::</option>
			<option value="1">객관식</option>
			<option value="2">주관식</option>
			<option value="3">단답형</option>
			<option value="9">구분자</option>
		</select>
		<script language="javascript">
		frm_Qst.qst_type.value="";
		</script>
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td bgcolor="#EEEEEE">문항 내용</td>
	<td align="left"><textarea name="qst_content" class="textarea" style="width:100%; height:50px;"></textarea></td>
</tr>
<tr id="trQstPoll" align="center" bgcolor="#FFFFFF" style="display:none;">
	<td bgcolor="#EEEEEE">
		지문<br>
		<span style="cursor:pointer" onclick="fnAddPoll()">[지문추가]</span>
	</td>
	<td align="left">
		<table width="100%" id="tbl_poll" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
		<tr align="center" bgcolor="#FFFFFF" >
			<td width="50" rowspan="2">지문 #1<input type="hidden" name="poll_sn" value="" /></td>
			<td colspan="2"><textarea name="poll_content" class="textarea" style="width:100%; height:32px;"></textarea></td>
		</tr>
		<tr bgcolor="#FFFFFF" >
			<td>
				추가의견
				<select name="poll_isAddAnswer" class="select">
					<option value="N" selected >없음</option>
					<option value="Y">있음</option>
				</select>
			</td>
			<td>관련문항 번호 : <input type="text" name="link_qst_sn" size="4" class="text"></td>
		</tr>
		<tr align="center" bgcolor="#FFFFFF" >
			<td width="50" rowspan="2">지문 #2<input type="hidden" name="poll_sn" value="" /></td>
			<td colspan="2"><textarea name="poll_content" class="textarea" style="width:100%; height:32px;"></textarea></td>
		</tr>
		<tr bgcolor="#FFFFFF" >
			<td>
				추가의견
				<select name="poll_isAddAnswer" class="select">
					<option value="N" selected >없음</option>
					<option value="Y">있음</option>
				</select>
			</td>
			<td>관련문항 번호 : <input type="text" name="link_qst_sn" size="4" class="text"></td>
		</tr>
		</table>
	</td>
</tr>
<tr id="trQstDiv" align="center" bgcolor="#FFFFFF" style="display:none;">
	<td bgcolor="#EEEEEE">&nbsp;</td>
	<td align="left">※ 구분자는 문항이 아닙니다. 문항들이 그룹일 경우 설명을 넣는 항목입니다.</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td bgcolor="#EEEEEE">필수여부</td>
	<td align="left">
		<label><input type="radio" name="qst_isNull" value="N" checked /> 답변필수</label>
		<label><input type="radio" name="qst_isNull" value="Y" /> 공란허용</label>
	</td>
</tr>
</table>
</form>
<!-- 입력테이블 끝 -->
<!-- 문항액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a">
<tr>
	<td align="right" style="padding:4 0 4 0"><input type="button" class="button" value="문항등록" onClick="fnQstSubmit()"></td>
</tr>
</table>
<!-- 문항액션 끝 -->
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->