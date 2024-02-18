<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  핑거스 강좌Pick 관리
' History : 2012.08.08 허진원 생성
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/academy/lib/academy_function.asp"-->
<!-- #include virtual="/academy/lib/classes/lecture_pickCls.asp"-->
<%
	dim iLp, sYYYY, sMM, sCDL, sLevel, page, i
	sYYYY = RequestCheckvar(Request("yyyy"),4)
	sMM = RequestCheckvar(Request("mm"),2)
	sCDL = RequestCheckvar(Request("cdl"),3)
	sLevel = RequestCheckvar(Request("level"),10)

	if sYYYY="" then sYYYY=year(date)
	if sMM="" then sMM=Num2Str(Month(date),2,"0","R")

	page = RequestCheckvar(request("page"),10)
	if page="" then page=1

	dim oLecPick
	set oLecPick = new CLecPick
		oLecPick.FCurrPage = page
		oLecPick.FPageSize=20

		oLecPick.FRectCDL = sCDL
		oLecPick.FRectLecLevel = sLevel
		if sYYYY<>"" and sMM<>"" then oLecPick.FRectYYYYMM = sYYYY & sMM

		oLecPick.GetLecPickList
%>
<script type="text/javascript">
function goPage(page){
	document.frm.method="GET";
	document.frm.action="";
	document.frm.page.value= page;
	document.frm.submit();
}

// 신규 강좌 등록
function fnNewLecReg() {
	var f = document.frm;

	if(!f.arrLecIdx.value) {
		alert("등록하실 강좌번호를 입력해주세요.\n\n※강좌번호는 콤마(,)로 구분해서 여러강좌를 동시에 등록하실 수 있습니다.");
		f.arrLecIdx.focus();
		return;
	}

	if(!f.cdl.value) {
		alert("등록하실 카테고리를 선택해주세요.");
		f.cdl.focus();
		return;
	}

	if(!f.level.value) {
		alert("등록하실 강좌의 난이도를 선택해주세요.");
		f.level.focus();
		return;
	}

	f.mode.value="add";
	f.method="POST";
	f.action="doLecPick.asp";
	f.submit();
}

// 선택된 강좌 삭제
function fnSelectLecDel() {
	var f = document.frm;
	var l = document.frmList;

	if(l.chkSel=="undefined") return;

	var arrIdx="", chk=false;
	if(!l.chkSel.length) {
		if(l.chkSel.checked) {
			chk=true;
			arrIdx = l.chkSel.value;
		}
	} else {
		for(var i=0;i<l.chkSel.length;i++) {
			if(l.chkSel[i].checked) {
				chk=true;
				if(arrIdx=="") {
					arrIdx = l.chkSel[i].value;
				} else {
					arrIdx += "," + l.chkSel[i].value;
				}
			}
		}
	}
	if(!chk) {
		alert("선택된 강좌가 없습니다.");
		return;
	}

	if(confirm("선택된 강좌를 삭제하시겠습니까?\n삭제가 완료되면 복구 할 수 없습니다.")) {
		f.mode.value="del";
		f.arrSn.value=arrIdx;
		f.method="POST";
		f.action="doLecPick.asp";
		f.submit();
	}
}
</script>
<!-- 검색 시작 -->
<form name="frm" method="get" action="" style="margin:0px;">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="page">
<input type="hidden" name="mode">
<input type="hidden" name="arrSn">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="#999999">
<tr align="center" bgcolor="#FFFFFF" >
	<td width="50" bgcolor="#EEEEEE">검색<br>조건</td>
	<td align="left">
<!--
		회차 :
		<select name="yyyy" class="select">
		<% for iLp=2012 to year(dateadd("yyyy",1,date)) %>
			<option value="<%=iLp%>"><%=iLp%></option>
		<% next %>
		</select>년
		<select name="mm" class="select">
		<% for iLp=1 to 12 %>
			<option value="<%=Num2Str(iLp,2,"0","R")%>"><%=iLp%></option>
		<% next %>
		</select>월
		&nbsp;
-->
		 카테고리 :
		<select name="cdl" class="select">
			<option value="">::선택::</option>
			<option value="10">만지기</option>
			<option value="20">꿰매기</option>
			<option value="30">꾸미기</option>
			<option value="40">맛보기</option>
			<option value="50">그리기</option>
			<option value="60">즐기기</option>
		</select>
		&nbsp;
		/ 난이도 :
		<select name="level" class="select">
			<option value="">::선택::</option>
			<option value="L">초급</option>
			<option value="M">중급</option>
			<option value="H">고급</option>
		</select>
		<script type="text/javascript">
			//document.frm.yyyy.value="<%=sYYYY%>";
			//document.frm.mm.value="<%=sMM%>";
			document.frm.cdl.value="<%=sCDL%>";
			document.frm.level.value="<%=sLevel%>";
		</script>
	</td>
	<td width="50" bgcolor="#EEEEEE">
		<input type="button" class="button_s" value="검색" onClick="javascript:document.frm.submit();">
	</td>
</tr>
</table>
<!-- 검색 끝 -->
<p>
<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
<tr>
	<td align="left">
		<input type="button" class="button" value="선택삭제" onClick="fnSelectLecDel()">
	</td>
	<td align="right">
		신규 등록
		<input type="text" name="arrLecIdx" size="60" style="text">
		<input type="button" class="button" value="저장" onClick="fnNewLecReg()">
	</td>
</tr>
</table>
</form>
<!-- 액션 끝 -->

<!-- 리스트 시작 -->
<form name="frmList" style="margin:0px;">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="#999999">
<tr height="25" bgcolor="#F6F6F6">
	<td colspan="6">
		검색결과 : <b><%= formatNumber(oLecPick.FTotalCount,0) %></b>
	</td>
</tr>
<tr align="center" bgcolor="<%=adminColor("sky")%>">
	<td width="40"></td>
	<td width="100">카테고리</td>
  	<td width="100">난이도</td>
  	<td width="100">강좌번호</td>
  	<td>강좌명</td>
  	<td width="100">등록일</td>
</tr>
<%
	if oLecPick.FresultCount>0 then
		for iLp=0 to oLecPick.FResultCount - 1
%>
<tr align="center" bgcolor="#FFFFFF">
	<td><input type="checkbox" name="chkSel" value="<%=oLecPick.FItemList(iLp).FpickSn%>"></td>
	<td><%=oLecPick.FItemList(iLp).FcdlNm%></td>
  	<td><%=oLecPick.FItemList(iLp).FlecLvName%></td>
  	<td><%=oLecPick.FItemList(iLp).FlecIdx%></td>
  	<td align="left"><%=oLecPick.FItemList(iLp).FlecTitle%></td>
  	<td><%=left(oLecPick.FItemList(iLp).Fregdate,10)%></td>
</tr>
<%		Next %>
<tr height="25" bgcolor="#F6F6F6">
	<td colspan="6" align="center">
	<% if oLecPick.HasPreScroll then %>
		<a href="javascript:goPage('<%= oLecPick.StartScrollPage-1 %>')">[pre]</a>
	<% else %>
		[pre]
	<% end if %>

	<% for i=0 + oLecPick.StartScrollPage to oLecPick.FScrollCount + oLecPick.StartScrollPage - 1 %>
		<% if i>oLecPick.FTotalpage then Exit for %>
		<% if CStr(page)=CStr(i) then %>
		<font color="red">[<%= i %>]</font>
		<% else %>
		<a href="javascript:goPage('<%= i %>')">[<%= i %>]</a>
		<% end if %>
	<% next %>

	<% if oLecPick.HasNextScroll then %>
		<a href="javascript:goPage('<%= i %>')">[next]</a>
	<% else %>
		[next]
	<% end if %>
	</td>
</tr>
<% else %>
<tr height="160" bgcolor="<%=adminColor("pink")%>">
	<td colspan="6" align="center">등록된 강좌가 없습니다. 상단의 신규등록으로 강좌를 등록해주세요.</td>
</tr>
<% end if %>
</table>
</form>
<%	set oLecPick = Nothing %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->