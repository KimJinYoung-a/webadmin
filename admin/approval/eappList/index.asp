<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 결제요청서 리스트
' History : 2011.10.13 정윤정  생성
'' ToDo 비타민(급여) 전송불가(DB타입설정 or ) // 환불 연동..
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/approval/eappListCls.asp"-->
<!-- #include virtual="/lib/classes/approval/eappCls.asp"-->
<!-- #include virtual="/lib/classes/approval/edmsCls.asp"-->
<!-- #include virtual="/lib/classes/admin/tenbyten/TenByTenDepartmentCls.asp"-->
<%

dim research
Dim clsEapp, clsedms
Dim ireportstate ,sadminId
Dim ireportidx
Dim iCurrpage,ipagesize,iTotCnt,iTotalPage
Dim arrList,intLoop
Dim iarap_cd,sarap_nm
Dim searchsdate,searchedate, susername, sreportname
Dim sOrderType
Dim icateidx1, sdatetype,sedmscode
dim department_id, inc_subdepartment

	iPageSize = 30
	iCurrPage = requestCheckvar(Request("iCP"),10)
	if iCurrPage="" then iCurrPage=1

	sadminId =  session("ssBctId")
	icateidx1	= requestCheckvar(Request("icidx1"),10)
	sdatetype	= requestCheckvar(Request("selDT"),10)
	ireportidx	= requestCheckvar(Request("iridx"),10)
	searchsdate= requestCheckvar(Request("selSD"),10)
	searchedate= requestCheckvar(Request("selED"),10)
	iarap_cd		= requestCheckvar(Request("iaidx"),13)
	sarap_nm		= requestCheckvar(Request("selarap"),50)
  ireportstate= requestCheckvar(Request("selPRS"),4)
  sedmscode 	= requestCheckvar(Request("sec"),10)
	sUserName		= requestCheckvar(Request("sUnm"),30)
	sreportname = requestCheckvar(Request("sRnm"),120)
	sOrderType	= requestCheckvar(Request("selOT"),1)
	department_id = requestCheckvar(Request("department_id"),10)
	inc_subdepartment = requestCheckvar(Request("inc_subdepartment"),1)
	research = requestCheckvar(Request("research"),10)

	if (research = "") then
		''searchsdate = Left(DateAdd("m", -6, Now()), 10)
	end if

	'메뉴에 따른 기본 카테고리 지정
	if menupos="1402" and Not(icateidx1="5" or icateidx1="12") then
		icateidx1="5"	'지출품의
	elseif menupos="1617" and Not(icateidx1="3" or icateidx1="12") then
		icateidx1="3"	'인사
	end if

'결재 기본 폼 정보 가져오기
set clsEapp = new CEappList
	clsEapp.Fcateidx1				= icateidx1
	clsEapp.FdateType				= sdatetype
	clsEapp.FStartDate				= searchsdate
	clsEapp.FEndDate				= searchedate
	clsEapp.FUsername				= sUserName
	clsEapp.FreportName				= sreportname
	clsEapp.FreportState    		= ireportstate
	clsEapp.Fedmscode				= sedmscode
 	clsEapp.Farap_cd				= iarap_cd
 	clsEapp.Farap_nm				= sarap_nm
 	clsEapp.FOrderType				= sOrderType
	clsEapp.FCurrpage 				= iCurrpage
	clsEapp.FPagesize				= ipagesize
	clsEapp.Fdepartment_id 			= department_id
	clsEapp.Finc_subdepartment 		= inc_subdepartment

	arrList = clsEapp.fnGetEappList
	iTotCnt = clsEapp.FTotCnt
set clsEapp = nothing
	iTotalPage 	=  int((iTotCnt-1)/iPageSize) +1  '전체 페이지 수 

%>
<script language="javascript" src="/admin/approval/eapp/eapp.js"></script>
<script language="javascript">
<!--
	function jsView(iridx){
		var winR = window.open("/admin/approval/eapp/vieweapp.asp?iridx="+iridx,"popR","width=1000, height=600, resizable=yes, scrollbars=yes");
		winR.focus();
	}

	function jsSearch(){
	 document.frm.submit();
	}

	// 페이지 이동
function jsGoPage(iCP)
	{
		document.frm.iCP.value=iCP;
		document.frm.submit();
	}

 	//선택 수지항목 가져오기
 	function jsSetARAP(dAC, sANM,sACC,sACCNM){
 		document.frm.iaidx.value = dAC;
 		document.frm.selarap.value = sANM;
 	}

//-->
</script>
<style>
	FORM {display:inline;}
	</style>
<table width="100%" align="center" cellpadding="5" cellspacing="1" class="a">
<tr>
	<td>
		<form name="frm" method="get" action="index.asp">
			<input type="hidden" name="menupos" value="<%= menupos %>">
			<input type="hidden" name="iCP" value="">
			<input type="hidden" name="research" value="on">
		<table width="100%" align="center" cellpadding="5" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
			<tr align="center" bgcolor="#FFFFFF" >
				<td rowspan="2" width="100" height="50" bgcolor="<%= adminColor("gray") %>">검색 조건</td>
				<td align="left">
					부서NEW:
					<%= drawSelectBoxDepartment("department_id", department_id) %>
					<input type="checkbox" name="inc_subdepartment" value="N" <% if (inc_subdepartment = "N") then %>checked<% end if %> > 하위 부서직원 제외
					&nbsp;&nbsp;
					카테고리:
					<select name="icidx1" id="icidx1">
					<%
						IF menupos=1402 THEN
							'지출품의 메뉴일 경우
					%>
						<option value="5" <%=chkIIF(icateidx1="5","selected","")%>>RP-지출품의</option>
						<option value="12" <%=chkIIF(icateidx1="12","selected","")%>>DR-기안</option>
					<%
						ELSEIF menupos=1617 Then
							'인사 메뉴일 경우
					%>
						<option value="3" <%=chkIIF(icateidx1="3","selected","")%>>HR-인사</option>
						<option value="12" <%=chkIIF(icateidx1="12","selected","")%>>DR-기안</option>
					<%
						ELSE
							Response.Write "<option value=""0"">--최상위--</option>"
							Set clsedms = new Cedms
							clsedms.sbGetOptedmsCategory 1,0,icateidx1
							Set clsedms = nothing
						END IF
					%>
					</select>&nbsp;&nbsp;
						<select name="selDT">
							<option value="1" <%IF sDateType ="1" THEN%>selected<%END IF%>>작성일</option>
							<option value="2" <%IF sDateType ="2" THEN%>selected<%END IF%>>최종승인일</option>
						</select>:
						<input type="text" name="selSD" size="10" value="<%=searchSDate%>"><img src="/images/calicon.gif" align="absmiddle" border="0" onClick="jsPopCal('selSD');"  style="cursor:hand;">
						~
						<input type="text" name="selED" size="10" value="<%=searchEDate%>"><img src="/images/calicon.gif" align="absmiddle" border="0" onClick="jsPopCal('selED');"  style="cursor:hand;">
				 </td>
				<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
					<input type="button" class="button_s" value="검색" onClick="javascript:jsSearch();">
				</td>
			</tr>
			<tr bgcolor="#FFFFFF" >
				<td>
					  문서코드: <input type="text" name="sec" value="<%=sedmscode%>" size="10" maxlength="10">&nbsp;
					  품의서명: <input type="text" name="sRnm" size="20" value="<%=sreportname%>">&nbsp;
						수지항목: <input type="text" name="selarap" value="<%=sarap_nm%>" size="13"><input type="hidden" name="iaidx" value="<%=iarap_cd%>" >
						<input type="button" value="선택" class="button" onClick="jsGetARAP();" >&nbsp;
					 작성자:
					<input type="text" name="sUnm" size="8" value="<%=sUserName%>">&nbsp;
					결재상태:
					<select name="selPRS">
						<option value="">----</option>
						 <%sbOptReportState ireportstate%>
					</select>&nbsp;
						정렬:
					<select name="selOT">
						<option value="1" <%IF sOrderType ="1" THEN%>selected<%END IF%>>최종승인일</option>
						<option value="2" <%IF sOrderType ="2" THEN%>selected<%END IF%>>작성일</option>
					</select>
				</td>
			</tr>
		</table>
		</form>
	</td>
</tr>
<tr>
	<td> 검색결과: <b><%=formatnumber(iTotCnt,0)%></b>  &nbsp;&nbsp;페이지: <b><%=iCurrpage%>/<%=iTotalPage%></b>
		<!-- 상단 띠 시작 -->
		<Form name="frmAct" method="post" action="erpLink_Process.asp">
		<input type="hidden" name="LTp" value="A">
		<table width="100%" align="left" cellpadding="5" cellspacing="1" class="a"   border="0">
				<tr bgcolor="<%= adminColor("tabletop") %>" align="center">
					<td>Idx</td>
					<td>문서코드</td>
					<td>품의서명</td>
					<td>품의금액</td>
					<td>수지항목</td>
					<td>계정과목</td>
					<td>작성자</td>
					<td>합의자</td>
					<td>최종승인자</td>
					<td>작성일</td>
					<td>최종승인일</td>
					<td>결재상태</td>
					<td>결제요청여부</td>
				</tr>
				<%IF isArray(arrList) THEN
					For intLoop = 0 To UBound(arrList,2)
				%>
				<tr bgcolor="#FFFFFF" align="center">
					<td><a href="javascript:jsView(<%=arrList(0,intLoop)%>);"><%=arrList(0,intLoop)%></a></td>
					<td nowrap><%=arrList(11,intLoop)%></td>
					<td align="left"><%=arrList(1,intLoop)%></td>
					<td align="right"><%=formatnumber(arrList(2,intLoop),0)%></td>
					<td align="left"><%=arrList(7,intLoop)%></td>
					<td align="left"><%=arrList(17,intLoop)%></td>
					<td nowrap><%=arrList(8,intLoop)%></td>
					<td nowrap><%=arrList(15,intLoop)%></td>
					<td nowrap><%=arrList(14,intLoop)%></td>
					<td><%=arrList(5,intLoop)%></td>
					<td><%=arrList(9,intLoop)%></td>
					<td><%=fnGetReportState(arrList(6,intLoop))%></td>
					<td><%=arrList(16,intLoop)%></td>
				</tr>
				<%
					Next
					ELSE
				%>
				<tr bgcolor="#FFFFFF">
					<td colspan="12" align="center">등록된 내역이 없습니다.</td>
				</tr>
				<%END IF%>
				</table>
				 </form>
			</td>
		</tr>
<!-- 페이지 시작 -->
		<%
		Dim iStartPage,iEndPage,iX,iPerCnt
		iPerCnt = 10

		iStartPage = (Int((iCurrPage-1)/iPerCnt)*iPerCnt) + 1

		If (iCurrPage mod iPerCnt) = 0 Then
			iEndPage = iCurrPage
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
									if Cint(ix) = Cint(iCurrPage) then
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
	</td>
</tr>
</table>
</body>
</html>

<!-- #include virtual="/lib/db/dbclose.asp" -->
