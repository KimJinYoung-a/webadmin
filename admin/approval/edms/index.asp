<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 문서관리 문서리스트
' History : 2011.02.24 정윤정  생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/approval/edmsCls.asp"-->
<%
Dim clsedms, arrList, intLoop
Dim icateidx1, icateidx2
Dim sedmsname,blnUsing
Dim iTotCnt,iPageSize, iTotalPage,page

	iPageSize = 20
	page = requestCheckvar(Request("page"),10)
	if page="" then page=1

	icateidx1 = requestCheckvar(Request("selC1"),10)
	icateidx2 = requestCheckvar(Request("hidC2"),10)

	if icateidx1 = "" then icateidx1 = 0
	if icateidx2 = "" then icateidx2= 0
		
	sedmsname = 	requestCheckvar(Request("sen"),20)
	blnUsing= requestCheckvar(Request("selU"),1)
	
Set clsedms = new Cedms
	clsedms.Fcateidx1 	= icateidx1
	clsedms.Fcateidx2	= icateidx2
	clsedms.Fedmsname		= sedmsname
	clsedms.FisUsing    = blnUsing
	clsedms.FCurrPage 	= page
	clsedms.FPageSize 	= iPageSize 
	arrList = clsedms.fnGetEdmsList
	iTotCnt = clsedms.FTotCnt

	iTotalPage 	=  int((iTotCnt-1)/iPageSize) +1  '전체 페이지 수
%>
<script type="text/javascript" src="/js/jquery-1.6.2.min.js"> </script>
<script type="text/javascript" src="/js/ajax.js"></script>
<script language="javascript">
<!--
// 페이지 이동
function jsGoPage(pg)
	{
		document.frm.page.value=pg;
		document.frm.submit();
	}


//새로등록
function jsNewReg(){
	var winD = window.open("popedmsConts.asp","popD","width=880, height=800, resizable=yes, scrollbars=yes");
	winD.focus();
}
//수정
function jsModReg(edmsidx){
	var winD = window.open("popedmsConts.asp?ieidx="+edmsidx,"popD","width=880, height=600, resizable=yes, scrollbars=yes");
	winD.focus();
}

// 카테고리 ajax =========================================================================================================
    initializeReturnFunction("processAjax()");
    initializeErrorFunction("onErrorAjax()");

    var _divName = "CL";

    function processAjax(){
        var reTxt = xmlHttp.responseText;
        eval("document.all.div"+_divName).innerHTML = reTxt;
    }

    function onErrorAjax() {
            alert("ERROR : " + xmlHttp.status);
    }

    //선택한 카테고리에 대한 하위 카테고리 리스트 가져오기 Ajax
    function jsSetCategory(sMode){
      var ipcidx  = document.frm.selC1.value;
      var icidx   = $("#selC2").val();

        initializeURL('ajaxCategory.asp?sMode='+sMode+'&ipcidx='+ipcidx+'&icidx='+icidx);
    	startRequest();
    }

    //파일 다운로드
    function jsDownload(ieidx, sRFN, sFN){
    var winFD = window.open("<%=uploadImgUrl%>/linkweb/edms/procDownload.asp?ieidx="+ieidx+"&sRFN="+sRFN+"&sFN="+sFN,"popFD","");
    winFD.focus();
    }

    //파일첨부
	function jsAttachFile(ieidx){
	var winAF = window.open("popRegFile.asp?ieidx="+ieidx+"&iML=10&menupos=<%=menupos%>&page=<%=page%>&icateidx1=<%=icateidx1%>&icateidx2=<%=icateidx2%>","popAF","width=450, height=200, resizable=yes, scrollbars=yes");
	winAF.focus();
	}

 	 //파일삭제
	function jsDeleteFile(ieidx){
	if (confirm("양식파일을 삭제하시겠습니까?")){
		document.frmDel.ieidx.value = ieidx;
		document.frmDel.submit();
	}
	}

	//검색
	function jsSearch(){
		document.frm.hidC2.value = $("#selC2").val();
		document.frm.submit();
	}

	//문서폼 등록
	function jsAddForm(ieidx){
		var winAF = window.open("popEdmsForm.asp?ieidx="+ieidx,"popAFo","width=880, height=600, resizable=yes, scrollbars=yes");
	winAF.focus();
	}
//-->
</script>
<table width="100%" align="center" cellpadding="5" cellspacing="1" class="a">
<form name="frmDel" method="post" action="procEdms.asp">
<input type="hidden" name="hidM" value="A">
<input type="hidden" name="ieidx" value="">
<input type="hidden" name="menupos" value="<%=menupos%>">
<input type="hidden" name="page" value="<%=page%>">
<input type="hidden" name="icateidx1" value="<%=icateidx1%>">
<input type="hidden" name="icateidx2" value="<%=icateidx2%>">
</form>
<tr>
	<td>
		<table width="100%" align="center" cellpadding="5" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
			<form name="frm" method="get" action="">
			<input type="hidden" name="menupos" value="<%= menupos %>">
			<input type="hidden" name="page" value="">
			<input type="hidden" name="hidC2" value="<%=icateidx2%>">
			<tr align="center" bgcolor="#FFFFFF" >
				<td  width="100" height="50" bgcolor="<%= adminColor("gray") %>">검색 조건</td>
				<td align="left">
					대 카테고리 :
					<select name="selC1" id="selC1" onChange="jsSetCategory('CL')">
					<option value="0">--최상위--</option>
					<%clsedms.sbGetOptedmsCategory 1,0,icateidx1 %>
					</select>

					중 카테고리 :
					<span id="divCL">
					<select name="selC2" id="selC2">
					<option value="0">----</option>
				<% 	IF icateidx1 > 0 THEN	'대카테고리 선택 후 중카테고리 선택가능하게
						clsedms.sbGetOptedmsCategory 2,icateidx1,icateidx2
					END IF
				%>
					</select>
					</span>
					
					문서명:<input type="text" name="sen"  size="20" maxlength="64" value="<%=sedmsname%>">
					
					사용유무:
					<select name="selU">
						<option value="">--</option> 
						<option value="1" <%IF blnUsing="1" THEN%>selected<%END IF%>>Y</option>
						<option value="0" <%IF blnUsing="0" THEN%>selected<%END IF%>>N</option>
					</select>
					
				</td>	
				<td  width="50" bgcolor="<%= adminColor("gray") %>">
					<input type="button" class="button_s" value="검색" onClick="jsSearch();">
				</td>
			</tr>
			</form>
		</table>
	</td>
</tr>
<%Set clsedms = nothing %>
<!-- #include virtual="/lib/db/dbclose.asp" -->
<tr>
	<td><input type="button" class="button" value="신규등록" onClick="jsNewReg();"></td>
</tr>
<tr>
	<td>
		<!-- 상단 띠 시작 -->
		<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
			<tr height="25" bgcolor="FFFFFF">
				<td colspan="16">
					검색결과 : <b><%=iTotCnt%></b> &nbsp;
					페이지 : <b><%= page %> / <%=iTotalPage%></b>
				</td>
			</tr>
			<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
				<td>idx</td>
				<td>문서코드</td>
				<td>대카테고리</td>
				<td>중카테고리</td>
				<td>일련번호</td>
				<td>문서명</td>
				<td>표시순서</td>
				<td>결재유무</td>
				<td>전자결재여부</td>
				<td>합의자</td>
				<td>최종결재자</td>
				<!-- td>CFO합의</td -->
				<td>어드민연결여부</td>
				<td>결제요청서사용유무</td>
				<td>사용유무</td>
				<td>문서폼</td>
				<td>양식파일</td>
			</tr>
			<% Dim sFileName, sReFileName
			IF isArray(arrList) THEN
				For intLoop = 0 To UBound(arrList,2)
					IF arrList(9,intLoop) <> "" THEN
					sFileName = split(arrList(9,intLoop),"/")(Ubound(split(arrList(9,intLoop),"/")))
					sReFileName = arrList(7,intLoop)&"_"&arrList(6,intLoop)&"."&split(arrList(9,intLoop),".")(ubound(split(arrList(9,intLoop),".")))
					END IF
				%>
			<tr height=30 align="center" bgcolor="#FFFFFF">
				<td><%=arrList(0,intLoop)%></td>
				<td><a href="javascript:jsModReg(<%=arrList(0,intLoop)%>);"><%=arrList(7,intLoop)%></a></td>
				<td><%=arrList(2,intLoop)%></td>
				<td><%=arrList(4,intLoop)%></td>
				<td><%=arrList(5,intLoop)%></td>
				<td><%=arrList(6,intLoop)%></td>
				<td><%=arrList(8,intLoop)%></td>
				<td><%IF arrList(10,intLoop) THEN%><font color="blue">Y</font><%ELSE%><font color="red">N</font><%END IF%></td>
				<td><%IF arrList(11,intLoop) THEN%><font color="blue">Y</font><%ELSE%><font color="red">N</font><%END IF%></td>
				<td><%If arrList(21,intLoop) = "Y" THEN%><font color="blue">Y</font><%ELSE%><font color="red">N</font><%END IF%></td>
				<td><%=arrList(16,intLoop)%></td>
				<!-- td><%IF arrList(20,intLoop) THEN%><font color="blue">Y</font><%ELSE%><font color="red">N</font><%END IF%></td -->
				<td><%IF (arrList(13,intLoop) <> "") and (not isNull(arrList(13,intLoop)))  THEN %><font color="blue">Y</font><%ELSE%><font color="red">N</font><%END IF%></td>
				<td><%IF arrList(17,intLoop) THEN%><font color="blue">Y</font><%ELSE%><font color="red">N</font><%END IF%></td>
				<td><%IF arrList(18,intLoop) THEN%><font color="blue">Y</font><%ELSE%><font color="red">N</font><%END IF%></td>
				<td><input type="button" class="button" value="<%IF isNull(arrList(19,intLoop)) or arrList(19,intLoop)="" THEN %>등록<%ELSE%>수정<%END IF%>" onClick="jsAddForm('<%=arrList(0,intLoop)%>');"></td>
				<td><%IF arrList(9,intLoop) <> "" THEN%><a href="javascript:jsDownload('<%=arrList(0,intLoop)%>','<%=sReFileName%>','<%=sFileName%>');"><%=sReFileName%></a>&nbsp;&nbsp;<a href="javascript:jsDeleteFile(<%=arrList(0,intLoop)%>);"><img src="/images/icon_minus.gif" border="0" alt="파일삭제"></a> <%END IF%><A href="javascript:jsAttachFile(<%=arrList(0,intLoop)%>);"><img src="/images/icon_plus.gif" border="0" alt="파일첨부 - 기존파일 삭제 후 새 파일 추가"></a></td>
			</tr>
		<%	Next
			ELSE%>
			<tr height=30 align="center" bgcolor="#FFFFFF">
				<td colspan="17">등록된 내용이 없습니다.</td>
			</tr>
			<%END IF%>
		</table>
	</td>
</tr>
<!-- 페이지 시작 -->
		<%
		Dim iStartPage,iEndPage,iX,iPerCnt
		iPerCnt = 10

		iStartPage = (Int((page-1)/iPerCnt)*iPerCnt) + 1

		If (page mod iPerCnt) = 0 Then
			iEndPage = page
		Else
			iEndPage = iStartPage + (iPerCnt-1)
		End If
		%>
			<tr height="25" >
				<td colspan="15" align="center">
					<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
					    <tr valign="bottom" height="25">
					        <td valign="bottom" align="center">
					         <% if (iStartPage-1 )> 0 then %><a href="javascript:jsGoPage(<%= iStartPage-1 %>)" onfocus="this.blur();">[pre]</a>
							<% else %>[pre]<% end if %>
					        <%
								for ix = iStartPage  to iEndPage
									if (ix > iTotalPage) then Exit for
									if Cint(ix) = Cint(page) then
							%>
								<a href="javascript:jsGoPage(<%= ix %>)" class="menu_link3" onfocus="this.blur();"><font color="00abdf"><strong>[<%=ix%>]</strong></font></a>
							<%		else %>
								<a href="javascript:jsGoPage(<%= ix %>)" class="menu_link3" onfocus="this.blur();">[<%=ix%>]</a>
							<%
									end if
								next
							%>
					    	<% if Cint(iTotalPage) > Cint(iEndPage)  then %><a href="javascript:jsGoPage(<%= ix %>)" onfocus="this.blur();">[next]</a>
							<% else %>[next]<% end if %>
					        </td>
					    </tr>
					</table>
				</td>
			</tr>
</table>
<!-- 페이지 끝 -->
</body>
</html>




