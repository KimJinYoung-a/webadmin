<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  부서별업무분배
' History : 2012.7.02 정윤정 생성
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/linkedERP/bizSectionCls.asp" -->
<!-- #include virtual="/lib/classes/admin/tenbyten/TenByTenMemberCls.asp" -->
<!-- #include virtual="/lib/classes/admin/tenbyten/TenByTenDepartmentCls.asp"-->
<%
Dim intY,intM, dYear,dMonth,part_sn
Dim blnView, blnSale,sBS_NM,clsBS
Dim arrList, intLoop,arrData, intData
dim bizsection_Cd
dim isUsing, isRegularMember,SearchKey, SearchString
dim YYYYMM
dim department_id, inc_subdepartment

dYear = requestCheckvar(Request("selY"),10)
IF dYear = "" THEN dYear = year(date())
dMonth= requestCheckvar(Request("selM"),10)
IF dMonth = "" THEN dMonth = month(date())
blnView = "Y"
blnSale = "Y"
bizsection_Cd = requestCheckvar(Request("bizsection_Cd"),32)
isUsing = requestCheckvar(Request("isUsing"),10)
isRegularMember = requestCheckvar(Request("isRegularMember"),10)
SearchKey = requestCheckvar(Request("SearchKey"),1)
SearchString = requestCheckvar(Request("SearchString"),32)
part_sn = requestCheckvar(Request("part_sn"),10)
department_id = requestCheckvar(Request("department_id"),10)
inc_subdepartment = requestCheckvar(Request("inc_subdepartment"),1)

''if part_sn = "" then
''	part_sn = 0
''elseIF part_sn = "0" and SearchKey <> ""  then
	 part_sn =  1
''end if

 Set clsBS = new CBizSection
	clsBS.FBS_NM 	= sBS_NM
	clsBS.FUSE_YN = "Y"
	clsBS.FView		= blnView
	clsBS.FSale		= blnSale
	arrList = clsBS.fnGetBizSectionList

	clsBS.FYYYYMM =  dYear&"-"&Format00(2,dMonth)
	clsBS.Fpart_sn = part_sn
	clsBS.Fbizsection_Cd = bizsection_Cd
	clsBS.FUSE_YN = isUsing
	clsBS.FisRegularMember = isRegularMember
  	clsBS.FSearchType 	= searchKey
	clsBS.FSearchText 	= searchString

	clsBS.Fdepartment_id 		= department_id
	clsBS.Finc_subdepartment 	= inc_subdepartment

	arrData	= clsBS.fnGetBizSectionAllList
Set clsBS = nothing

YYYYMM = dYear & "-" & Format00(2,dMonth)

dim delAvail

%>
<script type="text/javascript">
//업무비율 등록
function jsSetUserBiz(sDate){
	var bcd = "0000000505";

	var winBiz = window.open("pop_userBiz_bizsection_reg.asp?sBcd=" + bcd + "&sD="+sDate,"popBiz","width=630 height=800 scrollbars=yes");
	winBiz.focus();
}

function jsSetOneUserBiz(sDate, empno, delAvail){
	var bcd = "0000000505";
	<%
	'// 로그인정보(등급)에 따라 권한 설정(파트선임 이상:3 및 시스템팀:7 제외)
	if Not(session("ssAdminLsn")<=3 or session("ssAdminPsn")=7) then
		''if (part_sn <> session("ssAdminPsn")) then
			%>
			alert("파트선임이상 권한이 있어야 합니다.");
			return;
			<%
		''end if
	end if
	%>
	if (empno != "") {
		bcd = "";
	}

	var winBiz = window.open("pop_member_bizsection_Reg.asp?sBcd=" + bcd + "&sD="+sDate+"&sEn=" + empno + "&delAvail=" + delAvail,"jsSetOneUserBiz","width=630 height=800 scrollbars=yes");
	winBiz.focus();
}

</script>
<table width="100%" border="0" class="a">
<tr>
	<td>
		<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
			<form name="frm" method="get" action="">
			<input type="hidden" name="menupos" value="<%= menupos %>">
			<input type="hidden" name="research" value="on">
			<input type="hidden" name="page" value="">
			<tr align="center" bgcolor="#FFFFFF" >
				<td  width="50" bgcolor="<%= adminColor("gray") %>" rowspan="2">검색<br>조건</td>
				<td align="left" height="30">
					<select name="selY" class="select">
						<%For intY = Year(date()) To 2011 STEP -1%>
						<option value="<%=intY%>" <%IF Cstr(intY) = Cstr(dYear) THEN%>selected<%END IF%>><%=intY%></option>
						<%Next%>
						</select>년
						 <select name="selM" class="select">
						<%For intM = 1 To 12%>
						<option value="<%=intM%>" <%IF Cstr(intM) = Cstr(dMonth) THEN%>selected<%END IF%>><%=intM%></option>
						<%Next%>
						</select>월
			 			&nbsp; &nbsp;
					부서NEW:
					<%= drawSelectBoxDepartment("department_id", department_id) %>
					<input type="checkbox" name="inc_subdepartment" value="N" <% if (inc_subdepartment = "N") then %>checked<% end if %> > 하위 부서직원 제외
					&nbsp; &nbsp;
					매출부서:
					<select class="select" name="bizsection_Cd">
						<option value=""></option>
					<%

				 if isArray(arrList) then
				 	For intLoop = 0 To UBound(arrList,2)
				 	%>
				 	<option value="<%=arrList(0,intLoop)%>" <% if (bizsection_Cd = arrList(0,intLoop)) then %>selected<% end if %> ><%=arrList(1,intLoop)%></option>
					<% next %>
				<% end if %>
					</select>
				</td>

				<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
					<input type="button" class="button_s" value="검색" onClick="javascript:document.frm.submit();">
				</td>
			</tr>
			<tr align="center" bgcolor="#FFFFFF" >
				<td align="left" height="30">
					재직여부:
					<select name="isUsing" class="select">
						<option value="">전체</option>
						<option value="Y">재직</option>
						<option value="N">퇴사</option>
					</select>
					&nbsp;
					계약구분:
					<select name="isRegularMember" class="select">
						<option value="">전체</option>
						<option value="Y">정규직</option>
						<option value="N">계약직</option>
					</select>
					&nbsp;
					사원검색:
				 <select name="SearchKey" class="select">
						<option value="">::구분::</option>
						<option value="1" >아이디</option>
						<option value="2">사용자명</option>
						<option value="3">사번</option>
				 </select>
			<input type="text" class="text" name="SearchString" size="17" value="<%=SearchString%>">
			<script language="javascript">
				document.frm.isUsing.value="<%= isUsing %>";
				document.frm.SearchKey.value="<%= SearchKey %>";
				document.frm.isRegularMember.value="<%= isRegularMember %>";
			</script>
				</td>
			</tr>
			</form>
		</table>
<!-- 검색 끝 -->
</td>
</tr>
<%IF C_ADMIN_AUTH or C_ManagerUpJob THEN%>
<tr>
	<td><input type="button" class="button" value="전사업무비율 등록" onClick="jsSetUserBiz('<%=dYear&"-"&format00(2,dMonth)%>')">
	</td>
</tr>
<%END IF%>
<tr>
	<td>
		<p>
		<table  width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
			<%
			Dim  arrPBiz(),arrBiz(),intB,intP	, intChk , oldPCD
			intB = 0
			intP = 0
			intChk = 0

			IF isArray(arrList) THEN
				'// 온라인(공통) 오프라인본사 대학로1층매장 대학로2층매장 인천CGV점 두타점 김포롯데점 명동영플라자점 아이띵소본사 아카데미 기타 29cm ----> 온라인사업부 오프라인사업부 아이띵소사업부 핑거스사업부 기타 29cm
				'// 두번째 부서배열 생성
				For intLoop = 0 To UBound(arrList,2)
					IF oldPCD <> arrList(2,intLoop) THEN
						intP = intP + 1
						redim preserve arrPBiz(1,intP)
						arrPBiz(1,intP) =  arrList(4,intLoop)
						IF intP> 1 THEN
							arrPBiz(0,intP-1) = intChk
						END IF
						intChk =0
					END IF

					intChk = intChk + 1

					redim preserve arrBiz(1,intLoop)
					arrBiz(1,intLoop) = arrList(1,intLoop)
					arrBiz(0,intLoop) = arrList(0,intLoop)

					IF intLoop = UBound(arrList,2)    THEN
						arrPBiz(0,intP) = intChk
					END IF
					oldPCD  = arrList(2,intLoop)
				Next
			END IF
			%>

			<tr   bgcolor="<%= adminColor("tabletop") %>"  align="center">
				<%IF searchkey <> "" THEN%>
					<td width="80" rowspan="2">년월</td>
				<%END IF%>
				<%IF part_sn <> 0 THEN%>
				<td width="100" rowspan="2">사번</td>
				<td width="50" rowspan="2">이름</td>
				<td width="80" rowspan="2">입사일</td>
				<td width="80" rowspan="2">퇴사일</td>
				<%END IF%>
				<td rowspan="2">부서</td>
				<%For intLoop = 1 To intP%>
					<td colspan="<%=arrPBiz(0,intLoop)%>"><%=arrPBiz(1,intLoop)%></td>
				<%Next%>
			</tr>

			<!--사업부문 부서 리스트 start-->
			<tr  bgcolor="<%= adminColor("tabletop") %>"  align="center">
			<%For intLoop=0 To UBound(arrList,2) %>
				<td><%=arrBiz(1,intLoop)%></td>
			<%Next%>
			</tr>
			<!--// 사업부문 부서 리스트 end -->

			<!--사업부문 부서 리스트 start-->
			<%
			Dim OldENo, oldYM, intC, intD,intDA

			IF isArray(arrData) THEN
				For intData = 0 To UBound(arrData,2)
				IF searchKey <> "" THEN
					IF oldYM <>  arrData(0,intData) THEN
						intC = 0
						IF intData > 0 THEN
							%>
							</td>
							</tr>
						<% END IF%>
						<tr bgcolor="#FFFFFF" align="center">
							<%
							delAvail = "N"
							if (YYYYMM < Left(arrData(9,intData), 7)) then
								delAvail = "Y"
							end if
							if Not IsNull(arrData(10,intData)) then
								if (YYYYMM > Left(arrData(10,intData), 7)) then
									delAvail = "Y"
								end if
							end if
							%>
							<td><%=arrData(0,intData)%></td>
							<td><a href="javascript:jsSetOneUserBiz('<%=dYear&"-"&format00(2,dMonth)%>', '<%=arrData(3,intData)%>', '<%= delAvail %>')"><%=arrData(3,intData)%></td>
							<td><a href="javascript:jsSetOneUserBiz('<%=dYear&"-"&format00(2,dMonth)%>', '<%=arrData(3,intData)%>', '<%= delAvail %>')"><%=arrData(4,intData)%></td>
							<td>
								<% if (YYYYMM < Left(arrData(9,intData), 7)) then %><font color="red"><% end if %>
								<%= Left(arrData(9,intData), 10) %>
							</td>
							<td>
								<% if Not IsNull(arrData(10,intData)) then %>
									<% if (YYYYMM > Left(arrData(10,intData), 7)) then %><font color="red"><% end if %>
									<%= Left(arrData(10,intData), 10) %>
								<% end if %>
							</td>
						<td><%=arrData(11,intData)%></td>
					<% END IF%>
					<%For intD = intC To UBound(arrList,2)
						intC = intC + 1
						IF arrData(1,intData) = arrBiz(0,intD) THEN
							%>
							<td><%=arrData(2,intData)%>%</td>
							<% IF intData< UBound(arrData,2) THEN
								if  arrData(0,intData+1) = arrData(0,intData) THEN	Exit For
							END IF
						ELSE
							%>
							<td>&nbsp;</td>
							<%
						END IF
					Next

					OldYM= arrData(0,intData)
				ELSE
					IF OldENo <>  arrData(3,intData) THEN
						intC = 0
						IF intData > 0 THEN
							%>
							</td>
							</tr>
						<% END IF%>
						<tr bgcolor="#FFFFFF" align="center">
						<%IF part_sn <> 0 THEN%>
							<%
							delAvail = "N"
							if (YYYYMM < Left(arrData(9,intData), 7)) then
								delAvail = "Y"
							end if
							if Not IsNull(arrData(10,intData)) then
								if (YYYYMM > Left(arrData(10,intData), 7)) then
									delAvail = "Y"
								end if
							end if
							%>
							<td><a href="javascript:jsSetOneUserBiz('<%=dYear&"-"&format00(2,dMonth)%>', '<%=arrData(3,intData)%>', '<%= delAvail %>')"><%=arrData(3,intData)%></td>
							<td><a href="javascript:jsSetOneUserBiz('<%=dYear&"-"&format00(2,dMonth)%>', '<%=arrData(3,intData)%>', '<%= delAvail %>')"><%=arrData(4,intData)%></td>
							<td>
								<% if (YYYYMM < Left(arrData(9,intData), 7)) then %><font color="red"><% end if %>
								<%= Left(arrData(9,intData), 10) %>
							</td>
							<td>
								<% if Not IsNull(arrData(10,intData)) then %>
									<% if (YYYYMM > Left(arrData(10,intData), 7)) then %><font color="red"><% end if %>
									<%= Left(arrData(10,intData), 10) %>
								<% end if %>
							</td>
						<%END IF%>
						<td><%=arrData(11,intData)%></td>
					<% END IF%>
					<%For intD = intC To UBound(arrList,2)
						intC = intC + 1
						IF arrData(1,intData) = arrBiz(0,intD) THEN
							%>
							<td><%IF part_sn = 0 THEN%><%=formatnumber(arrData(2,intData),2)%><%ELSE%><%=arrData(2,intData)%><%END IF%>%</td>
							<% IF intData< UBound(arrData,2) THEN
								if  arrData(3,intData+1) = arrData(3,intData) THEN	Exit For
							END IF
						ELSE
							%>
							<td>&nbsp;</td>
							<%
						END IF
					Next

					OldENo= arrData(3,intData)
				END IF
				Next
			END IF
			%>
			</tr>
				<!--사업부문 부서 리스트 End-->
		</table>
	</td>
</tr>
</table>
