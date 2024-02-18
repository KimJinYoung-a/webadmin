<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 운영비관리 팀  리스트
' History : 2011.05.31 정윤정 생성
'			2018.10.11 한용민 수정
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/expenses/OpExpPartCls.asp"-->
<%
Dim clsPart, arrList, intLoop, arrType
Dim sOpExppartName, ipartTypeIdx, iTotCnt,iPageSize, iTotalPage,iCurrPage, incNo
	iPageSize = 20
	iCurrPage = requestCheckvar(Request("iCP"),10)
	if iCurrPage="" then iCurrPage=1

 	sOpExppartName 	= requestCheckvar(Request("sOEPN"),60)
 	iPartTypeIdx	= requestCheckvar(Request("selPT"),10)
	incNo			= requestCheckvar(Request("incNo"),1)

Set clsPart = new COpExpPart
	clsPart.FPartTypeidx 	= iPartTypeIdx
	clsPart.FOpExpPartName 	= sOpExppartName
	clsPart.FRectIncNo 	= incNo
	clsPart.FCurrPage 	= iCurrPage
	clsPart.FPageSize 	= iPageSize
	arrList = clsPart.fnGetOpExpPartList
	iTotCnt = clsPart.FTotCnt

	arrType = clsPart.fnGetOpExpPartTypeList
Set clsPart = nothing

iTotalPage 	=  int((iTotCnt-1)/iPageSize) +1  '전체 페이지 수
%>

<script type="text/javascript">
<!--
// 페이지 이동
function jsGoPage(iCP)
	{
		document.frm.iCP.value=iCP;
		document.frm.submit();
	}

//새로등록
function jsNewReg(){
	var winP = window.open("popPart.asp","popP","width=800,height=960,scrollbars=yes,resizable=yes");
	winP.focus();
}

//수정
function jsMod(iOEP){
	var winP = window.open("popPart.asp?hidOEP="+iOEP,"popP","width=800,height=960,scrollbars=yes,resizable=yes");
	winP.focus();
}

//타입수정
function jsModType(){
var winP = window.open("popPartType.asp","popP","width=800,height=600,scrollbars=yes,resizable=yes");
	winP.focus();
}
//-->
</script>
<table width="100%" align="center" cellpadding="5" cellspacing="1" class="a">
	<tr>
	<td>
		<table width="100%" align="center" cellpadding="5" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
			<form name="frm" method="get" action="">
			<input type="hidden" name="menupos" value="<%= menupos %>">
			<input type="hidden" name="iCP" value="">
			<tr align="center" bgcolor="#FFFFFF" >
				<td width="100" height="50" bgcolor="<%= adminColor("gray") %>">검색 조건</td>
				<td align="left">
					 구분:
					 <select name="selPT">
					 <option value="">--선택--</option>
					 <% sbOptPartType arrType,ipartTypeIdx%>
					 </select>
					 &nbsp;&nbsp;
					 운영비사용처 :
					 <input type="text" name="sOEPN" size="20" maxlength="60" value="<%=sOpExppartName%>">
					 &nbsp;&nbsp;
					 <input type="checkbox" name="incNo" value="Y" <% if (incNo = "Y") then %>checked<% end if %> >
					 사용안함 포함

				</td>
				<td width="50" bgcolor="<%= adminColor("gray") %>">
					<input type="button" class="button_s" value="검색" onClick="javascript:document.frm.submit();">
				</td>
			</tr>
			</form>
		</table>
	</td>
</tr>
<!-- #include virtual="/lib/db/dbclose.asp" -->
<tr>
	<td>
		<table width="100%" border="0" cellpadding="0" cellspacing="0" class="a">
		<tr>
			<td align="left"><input type="button" class="button" value="신규등록" onClick="jsNewReg();"></td>
			<td align="right"><input type="button" value="구분수정" onClick="jsModType()" class="button"></td>
		</tr>
		</table>
	</td>
</tr>
<tr>
	<td>
		<!-- 상단 띠 시작 -->
		<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
			<tr height="25" bgcolor="FFFFFF">
				<td colspan="15">
					검색결과 : <b><%=iTotCnt%></b> &nbsp;
					페이지 : <b><%= iCurrPage %> / <%=iTotalPage%></b>
				</td>
			</tr>
			<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
				<td width="80">표시순서</td>
				<td>구분</td>
				<td>운영비사용처</td>
				<td>담당자</td>
				<td>부서명</td>
				<td>자금관리부서</td>
				<td>지급거래처</td> 
				<td>수지항목</td> 
				<td>처리</td>
			</tr>
			<%  
			IF isArray(arrList) THEN
				For intLoop = 0 To UBound(arrList,2) 
				%> 
			<tr height=30 align="center" bgcolor="#<% if (arrList(20,intLoop) = True) then %>FFFFFF<% else %>DDDDDD<% end if %>">
				<td><%=arrList(12,intLoop)%></td>
				<td><%=arrList(2,intLoop)%></td>
				<td><%=arrList(3,intLoop)%></td>
				<td><%=arrList(7,intLoop)%></td>
				<td align="left">
					&nbsp;
					<%
					if arrList(21,intLoop) = 1 then
						response.write arrList(22,intLoop)
					elseif arrList(21,intLoop) > 1 then
						response.write arrList(22,intLoop) + " 외 " + CStr(arrList(21,intLoop) - 1)
					end if
					%>
				</td>
				<td><%=arrList(14,intLoop)%></td> 
				<td><%=arrList(17,intLoop)%></td> 
				<td><%=arrList(15,intLoop)%></td> 
				<td><input type="button" value="운영비수정" class="button" onClick="jsMod(<%=arrList(0,intLoop)%>)"></td>
			</tr>
		<%      Next
			ELSE%>  
			<tr height=5 align="center" bgcolor="#FFFFFF">
				<td colspan="5">등록된 내용이 없습니다.</td>
			</tr>
			<%END IF%>
		</table>
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
<!-- 페이지 끝 -->
</body>
</html>
