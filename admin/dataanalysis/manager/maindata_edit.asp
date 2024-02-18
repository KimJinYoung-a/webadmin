<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 데이터분석
' History : 2016.01.29 한용민 생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbAnalopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/dataanalysis/dataanalysis_cls.asp"-->

<%
dim i, menupos, searchgroupcd
	menupos = requestCheckVar(request("menupos"),10)
	searchgroupcd = requestCheckVar(request("searchgroupcd"),32)

dim cdata
SET cdata = New cdataanalysis
	cdata.FCurrPage = 1
	cdata.FPageSize = 1000
	cdata.frectgroupcd = searchgroupcd
	'cdata.frectisusing="Y"
	cdata.Getdataanalysis_maingroup_list()

%>
<link href="/js/jqueryui/css/jquery-ui.css" rel="stylesheet">
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript" src="/js/jqueryui/jquery-ui-1.10.2.custom.min.js"></script>
<script type="text/javascript">

function chkAllchartItem() {
	if($("input[name='mainidx']:first").attr("checked")=="checked") {
		$("input[name='mainidx']").attr("checked",false);
	} else {
		$("input[name='mainidx']").attr("checked","checked");
	}
}

function savemaindataList() {
	var chk=0;
	$("form[name='frmmaindatalist']").find("input[name='mainidx']").each(function(){
		if($(this).attr("checked")) chk++;
	});
	if(chk==0) {
		alert("수정하실 항목을 선택해주세요.");
		return;
	}

	if(confirm("지정하신 차트 정보를 저장 하시겠습니까?")) {
		//frmmaindatalist.target="ifproc";
		frmmaindatalist.mode.value="maindatalistedit";
		frmmaindatalist.action="/admin/dataanalysis/manager/manager_process.asp";
		frmmaindatalist.submit();
	}
}

function savemaindataone() {
	if (frmmaindataone.groupcd.value=='NEWREG'){
		if (frmmaindataone.groupcdnewreg.value==''){
			alert('그룹코드를 입력해 주세요.');
			frmmaindataone.groupcdnewreg.focus();
			return;
		}
		if (frmmaindataone.groupcdnamenewreg.value==''){
			alert('그룹코드명을 입력해 주세요.');
			frmmaindataone.groupcdnamenewreg.focus();
			return;
		}
	}
	if(frmmaindataone.groupsortno.value!=''){
		if (!IsDouble(frmmaindataone.groupsortno.value)){
			alert('그룹정렬은 숫자만 입력 가능 합니다.');
			frmmaindataone.groupsortno.focus();
			return;
		}
	}else{
		alert('정렬값을 입력해주세요.');
		return false;
	}
	var tmpisusing='';
	for (var i=0; i < frmmaindataone.isusing.length; i++){
		if (frmmaindataone.isusing[i].checked){
			tmpisusing = frmmaindataone.isusing[i].value;
		}
	}
	if (tmpisusing==''){
		alert('사용여부를 선택해 주세요.');
		return false;
	}
	if (frmmaindataone.measure.value==''){
		alert('측정값코드를 입력해 주세요.');
		frmmaindataone.measure.focus();
		return;
	}
	if (frmmaindataone.measurename.value==''){
		alert('측정값명을 입력해 주세요.');
		frmmaindataone.measurename.focus();
		return;
	}
	if (frmmaindataone.dimensiongubun.value!=''){
		if (!IsDouble(frmmaindataone.dimensiongubun.value)){
			alert('일정타입은 숫자만 입력 가능 합니다.');
			frmmaindataone.dimensiongubun.focus();
			return;
		}
	}
	if (frmmaindataone.pretypegubun.value!=''){
		if (!IsDouble(frmmaindataone.pretypegubun.value)){
			alert('비교타입은 숫자만 입력 가능 합니다.');
			frmmaindataone.pretypegubun.focus();
			return;
		}
	}
	if (frmmaindataone.shchannelgubun.value!=''){
		if (!IsDouble(frmmaindataone.shchannelgubun.value)){
			alert('채널타입은 숫자만 입력 가능 합니다.');
			frmmaindataone.shchannelgubun.focus();
			return;
		}
	}
	if (frmmaindataone.shmakeridgubun.value!=''){
		if (!IsDouble(frmmaindataone.shmakeridgubun.value)){
			alert('브랜드타입은 숫자만 입력 가능 합니다.');
			frmmaindataone.shmakeridgubun.focus();
			return;
		}
	}
	if (frmmaindataone.shdategubun.value!=''){
		if (!IsDouble(frmmaindataone.shdategubun.value)){
			alert('브랜드타입은 숫자만 입력 가능 합니다.');
			frmmaindataone.shdategubun.focus();
			return;
		}
	}
	if (frmmaindataone.shdatetermgubun.value!=''){
		if (!IsDouble(frmmaindataone.shdatetermgubun.value)){
			alert('브랜드타입은 숫자만 입력 가능 합니다.');
			frmmaindataone.shdatetermgubun.focus();
			return;
		}
	}
	if(confirm("차트 정보를 신규 저장 하시겠습니까?")) {
		//frmmaindataone.target="ifproc";
		frmmaindataone.mode.value="maindatareg";
		frmmaindataone.action="/admin/dataanalysis/manager/manager_process.asp";
		frmmaindataone.submit();
	}
}

function gochartreg(){
	$("#gochartreg").show();
}
function gochartregclose(){
	$("#gochartreg").hide();
}

function gosearch(groupcd){
	location.replace('/admin/dataanalysis/manager/maindata_edit.asp?searchgroupcd='+ groupcd +'&menupos=<%= menupos %>');
}

function chgroupcdnewreg(groupcd){
	if (groupcd=='NEWREG'){
		$("#groupcdnewreg").show();
	}else{
		$("#groupcdnewreg").hide();
	}
}

</script>

<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="FFFFFF">
	<td>
		<form name="frmmaindataone" method="POST" action="" style="margin:0;">
		<input type="hidden" name="mode" value="">
		<input type="hidden" name="menupos" value="<%= menupos %>">
		<table width="100%" align="left" cellpadding="2" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>" id="gochartreg" style="display:<% if cdata.FtotalCount<1 then %><% else %>none<% end if %>;">
		<tr align="center" bgcolor="FFFFFF">
			<td colspan="2">
				데이터 신규등록
			</td>
		</tr>
		<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
			<td>그룹</td>
			<td bgcolor="FFFFFF" align="left">
				<% drawSelectBoxdataanalysisgroup "groupcd", searchgroupcd, " onchange='chgroupcdnewreg(this.value);'", "NEW", "" %> 
				신규등록을 원하시면 신규등록을 선택하세요.<br>
				<div id="groupcdnewreg" style="display:none;">
					그룹코드 : <input type="text" name="groupcdnewreg" size=32 class="text" value="" /> ex) mkt
					<br>그룹코드명 : <input type="text" name="groupcdnamenewreg" size=32 class="text" value="" /> ex) 마케팅
				</div>
			</td>
		</tr>
		<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
			<td>그룹정렬</td>
			<td bgcolor="FFFFFF" align="left">
				<input type="text" name="groupsortno" size=3 class="text" value="100" />
			</td>
		</tr>
		<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
			<td>종류</td>
			<td bgcolor="FFFFFF" align="left">
				<input type="text" name="kind" size=32 class="text" value="" /> ex) mktall
			</td>
		</tr>
		<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
			<td>측정값코드</td>
			<td bgcolor="FFFFFF" align="left">
				<input type="text" name="measure" size=32 class="text" value="" /> ex) gavisit
			</td>
		</tr>
		<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
			<td>측정값명</td>
			<td bgcolor="FFFFFF" align="left">
				<input type="text" name="measurename" size=32 class="text" value="" /> ex) 방문수(GA)
			</td>
		</tr>
		<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
			<td>APiURL</td>
			<td bgcolor="FFFFFF" align="left">
				<input type="text" name="goapiurl" size=32 class="text" value="" /> ex) 미입력시 기본값 http://wapi.10x10.co.kr/anal/getque.asp
			</td>
		</tr>
		<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
			<td>일정타입</td>
			<td bgcolor="FFFFFF" align="left">
				1 : 일,시간,주,월,년,요일
				<br>2 : 일,주,월,년,요일
				<br>-입력안함 : 사용안함
				<br><input type="text" name="dimensiongubun" size=1 class="text" value="" />
			</td>
		</tr>
		<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
			<td>비교타입</td>
			<td bgcolor="FFFFFF" align="left">
				1 : 
				<br>&nbsp;&nbsp;일 선택시 : 전주,전월,전년,전년동요일
				<br>&nbsp;&nbsp;시간 선택시 : 전일,전주,전월,전년,전년동요일
				<br>&nbsp;&nbsp;주 선택시 : 전월,전년,전년동요일
				<br>&nbsp;&nbsp;월 선택시 : 전년,전년동요일
				<br>&nbsp;&nbsp;년 선택시 : 없음
				<br>&nbsp;&nbsp;요일 선택시 : 전주, 전월, 전년, 전년동요일"
				<br>-입력안함 : 사용안함
				<br><input type="text" name="pretypegubun" size=1 class="text" value="" />
			</td>
		</tr>
		<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
			<td>채널타입</td>
			<td bgcolor="FFFFFF" align="left">
				1 : WWW,모바일,모바일_제휴,APP,APP_제휴,제휴몰,3PL
				<br>-입력안함 : 사용안함
				<br><input type="text" name="shchannelgubun" size=1 class="text" value="" />
			</td>
		</tr>
		<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
			<td>브랜드타입</td>
			<td bgcolor="FFFFFF" align="left">
				1 : 브랜드검색사용함
				<br>-입력안함 : 사용안함
				<br><input type="text" name="shmakeridgubun" size=1 class="text" value="" />
			</td>
		</tr>
		<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
			<td>날짜종류타입</td>
			<td bgcolor="FFFFFF" align="left">
				1 : 주문일
				<br>2 : 주문일,결제일
				<br>3 : 주문일,결제일,출고일
				<br>-입력안함 : 기간
				<br><input type="text" name="shdategubun" size=1 class="text" value="" />
			</td>
		</tr>
		<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
			<td>날짜단위</td>
			<td bgcolor="FFFFFF" align="left">
				1 : YYYY
				<br>2 : YYYY-MM
				<br>-입력안함 : YYYY-MM-DD
				<br><input type="text" name="shdateunit" size=1 class="text" value="" />
			</td>
		</tr>
		<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
			<td>날짜기간타입</td>
			<td bgcolor="FFFFFF" align="left">
				1 : 당일
				<br>2 : 한달
				<br>-입력안함 : 일주일
				<br><input type="text" name="shdatetermgubun" size=1 class="text" value="" />
			</td>
		</tr>
		<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
			<td>정렬타입</td>
			<td bgcolor="FFFFFF" align="left">
				1 : 오름차순,내림차순
				<br>2 : 카테고리구분순,매출순,매출달성순
				<br>-입력안함 : 사용안함
				<br><input type="text" name="ordtypegubun" size=1 class="text" value="" />
			</td>
		</tr>
		<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
			<td>코맨트</td>
			<td bgcolor="FFFFFF" align="left">
				<textarea name="comment" cols=30 rows=2></textarea>
			</td>
		</tr>
		<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
			<td>사용여부</td>
			<td bgcolor="FFFFFF" align="left">
				<input type="radio" name="isusing" value="Y" checked />Y
				<input type="radio" name="isusing" value="N" />N
			</td>
		</tr>
		<tr align="center" bgcolor="FFFFFF">
			<td colspan="2">
				<input type="button" onClick="savemaindataone();" value="신규저장" class="button">
				&nbsp;
				<input type="button" onClick="gochartregclose();" value="닫기" class="button">
			</td>
		</tr>
		</table>
		</form>
	</td>
</tr>
<tr align="center" bgcolor="FFFFFF">
	<td>
		<br>
		<form name="frmmaindatalist" method="POST" action="" style="margin:0;">
		<input type="hidden" name="chkAll" value="N">
		<input type="hidden" name="mode" value="">
		<input type="hidden" name="menupos" value="<%= menupos %>">
		<table width="100%" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
		<tr height="25" bgcolor="FFFFFF">
			<td colspan="20" align="right">
				<input type="button" onClick="savemaindataList();" value="선택저장" class="button">
				&nbsp;
				<input type="button" onClick="gochartreg();" value="신규등록" class="button">
			</td>
		</tr>
		<tr height="25" bgcolor="FFFFFF">
			<td colspan="20">
				검색결과 : <b><%= cdata.FtotalCount %></b>
				<br>
				※검색※
				&nbsp;&nbsp;&nbsp;&nbsp;
				그룹 : <% drawSelectBoxdataanalysisgroup "searchgroupcd", searchgroupcd, " onchange='gosearch(this.value);'", "", "" %>
			</td>
		</tr>
		<% if cdata.FtotalCount>0 then %>
		<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		    <td width=30><input type="button" value="전체" class="button" onClick="chkAllchartItem();"></td>
		    <td width=50>메인idx</td>
		    <td width=40>그룹<br>정렬</td>
		    <td>종류</td>
		    <td>측정값코드</td>
		    <td>측정값명</td>
		    <td width=160>APiURL</td>
		    <td width=30>일정<br>타입</td>
		    <td width=30>비교<br>타입</td>
		    <td width=30>채널<br>타입</td>
		    <td width=40>브랜드<br>타입</td>
		    <td width=30>날짜<br>종류<br>타입</td>
		    <td width=30>날짜<br>단위</td>
		    <td width=30>날짜<br>기간<br>타입</td>
		    <td width=30>정렬<br>타입</td>
		    <td width=160>코맨트</td>
		    <td width=60>사용여부</td>
		</tr>
		<%	for i=0 to cdata.FResultCount - 1 %>
		<input type="hidden" name="groupcd_<%= cdata.FItemList(i).fmainidx %>" value="<%= cdata.FItemList(i).fgroupcd %>" />
		<tr bgcolor="<%=chkIIF(cdata.FItemList(i).fisusing="Y","#FFFFFF","#DDDDDD")%>" onmouseover=this.style.background="#f1f1f1"; onmouseout=this.style.background='<%=chkIIF(cdata.FItemList(i).fisusing="Y","#FFFFFF","#DDDDDD")%>';>
		    <td><input type="checkbox" name="mainidx" value="<%= cdata.FItemList(i).fmainidx %>" /></td>
		    <td><%= cdata.FItemList(i).fmainidx %></td>
		    <td>
		    	<input type="text" name="groupsortno_<%= cdata.FItemList(i).fmainidx %>" size=2 class="text" value="<%= cdata.FItemList(i).fgroupsortno %>" />
		    </td>
		    <td>
		    	<input type="text" name="kind_<%= cdata.FItemList(i).fmainidx %>" size=15 class="text" value="<%= cdata.FItemList(i).fkind %>" />
		    </td>
		    <td>
		    	<input type="text" name="measure_<%= cdata.FItemList(i).fmainidx %>" size=15 class="text" value="<%= cdata.FItemList(i).fmeasure %>" />
		    </td>
		    <td>
		    	<input type="text" name="measurename_<%= cdata.FItemList(i).fmainidx %>" size=15 class="text" value="<%= cdata.FItemList(i).fmeasurename %>" />
		    </td>
		    <td>
		    	<textarea name="goapiurl_<%= cdata.FItemList(i).fmainidx %>" cols=20 rows=3><%= cdata.FItemList(i).fapiurl %></textarea>
		    </td>
		    <td>
		    	<input type="text" name="dimensiongubun_<%= cdata.FItemList(i).fmainidx %>" size=1 class="text" value="<%= cdata.FItemList(i).fdimensiongubun %>" />
		    </td>
		    <td>
		    	<input type="text" name="pretypegubun_<%= cdata.FItemList(i).fmainidx %>" size=1 class="text" value="<%= cdata.FItemList(i).fpretypegubun %>" />
		    </td>
		    <td>
		    	<input type="text" name="shchannelgubun_<%= cdata.FItemList(i).fmainidx %>" size=1 class="text" value="<%= cdata.FItemList(i).fshchannelgubun %>" />
		    </td>
		    <td>
		    	<input type="text" name="shmakeridgubun_<%= cdata.FItemList(i).fmainidx %>" size=1 class="text" value="<%= cdata.FItemList(i).fshmakeridgubun %>" />
		    </td>
		    <td>
		    	<input type="text" name="shdategubun_<%= cdata.FItemList(i).fmainidx %>" size=1 class="text" value="<%= cdata.FItemList(i).fshdategubun %>" />
		    </td>
		    <td>
		    	<input type="text" name="shdateunit_<%= cdata.FItemList(i).fmainidx %>" size=1 class="text" value="<%= cdata.FItemList(i).fshdateunit %>" />
		    </td>
		    <td>
		    	<input type="text" name="shdatetermgubun_<%= cdata.FItemList(i).fmainidx %>" size=1 class="text" value="<%= cdata.FItemList(i).fshdatetermgubun %>" />
		    </td>
		    <td>
		    	<input type="text" name="ordtypegubun_<%= cdata.FItemList(i).fmainidx %>" size=1 class="text" value="<%= cdata.FItemList(i).fordtypegubun %>" />
		    </td>
		    <td>
		    	<textarea name="comment_<%= cdata.FItemList(i).fmainidx %>" cols=20 rows=3><%= cdata.FItemList(i).fcomment %></textarea>
		    </td>
		    <td align="center">
				<input type="radio" name="isusing_<%= cdata.FItemList(i).fmainidx %>" value="Y" <%=chkIIF(cdata.FItemList(i).fisusing="Y" or isnull(cdata.FItemList(i).fisusing) or cdata.FItemList(i).fisusing="","checked","")%> />Y
				<input type="radio" name="isusing_<%= cdata.FItemList(i).fmainidx %>" value="N" <%=chkIIF(cdata.FItemList(i).fisusing="N","checked","")%> />N
		    </td>
		</tr>
		<%	Next %>
		<% else %>
			<tr bgcolor="#FFFFFF">
				<td colspan="20" align="center">검색결과가 없습니다.</td>
			</tr>
		<% end if %>
		</table>
		</form>
	</td>
</tr>
</table>

<iframe id="ifproc" name="ifproc" width=0 height=0></iframe>

<%
set cdata=nothing
%>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbAnalclose.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->