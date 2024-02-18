<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 자금관리 부서
' History : 2011.04.21 정윤정  생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/linkedERP/bizSectionCls.asp"-->
<%
Dim clsBS
Dim arrList, intLoop
Dim sUSEYN,sBS_NM, sView,sSale
sBS_NM = requestCheckvar(Request("sBS_NM"),100)
sUSEYN = requestCheckvar(Request("sUSEYN"),3)
sView = requestCheckvar(Request("sView"),1)
sSale = requestCheckvar(Request("sSale"),1)

Set clsBS = new CBizSection
	clsBS.FBS_NM 	= sBS_NM
	clsBS.FUSE_YN = sUSEYN
	clsBS.FView		= sView
	clsBS.FSale		= sSale
	arrList = clsBS.fnGetBizSectionList
Set clsBS = nothing
%>

<script language="javascript">
<!--

	//수정
	function jsModReg(eapppartidx){
		var winC = window.open("popPart.asp?iepidx="+eapppartidx,"popC","width=600, height=600, resizable=yes, scrollbars=yes");
		winC.focus();
	}

   //검색
   function jsSearch(){
    document.frm.submit();
   }

   //매출이익부서/전시여부  설정변경
   function jsChType(iType, sBizSectionCD,blnV){
   	document.frmChange.iT.value = iType;
   	document.frmChange.blnV.value = blnV;
   	document.frmChange.sBCD.value = sBizSectionCD;
   	document.frmChange.submit();
  }

   //업데이트
   function jsUpdate(){
   	document.frmUpdate.submit();
  }
//-->
</script>
<form name="frmUpdate" method="post" action="procBiz.asp">
	<input type="hidden" name="sM" value="U">
</form>

<form name="frmChange" method="post" action="procBiz.asp">
	<input type="hidden" name="sM" value="C">
	<input type="hidden" name="iT" value="">
	<input type="hidden" name="blnV" value="">
	<input type="hidden" name="sBCD" value="">
</form>
<table width="100%" align="center" cellpadding="5" cellspacing="1" class="a">
<tr>
	<td>
		<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="#999999">
		<form name="frm" method="post" action="index.asp">
		<input type="hidden" name="menupos" value="<%=menupos%>">
		<tr align="center" bgcolor="#FFFFFF" >
			<td rowspan="2" width="100" bgcolor="#EEEEEE">검색 조건</td>
			<td align="left">&nbsp;
			 부서명: <input type="text" name="sBS_NM" size="20" value="<%=sBS_NM%>">
		 	&nbsp;
		 	<input type="checkbox" name="sUSEYN" value="A" <%IF cStr(sUSEYN) ="A" THEN%>checked<%END IF%>>비활동포함
		 	&nbsp;
		 	<input type="checkbox" name="sSale" value="Y" <%IF cStr(sSale) ="Y" THEN%>checked<%END IF%>>이익부서만
		 	&nbsp;
		 	<input type="checkbox" name="sView" value="Y" <%IF cStr(sView) ="Y" THEN%>checked<%END IF%>>전시만
			</td>
			<td rowspan="2" width="50" bgcolor="#EEEEEE">
				<input type="button" class="button_s" value="검색" onClick="jsSearch();">
			</td>
		</tr>
		</form>
		</table>
	</td>
</tr>
<tr>
	<td>
        <input type="button" class="button" value="update" onclick="jsUpdate();"><br />
        * <font color="red">세금계산서</font> 에 부서가 표시되지 않는 경우, 전시여부/이익부서여부 를 Y 로 설정하세요.
    </td>
</tr>
<tr>
	<td>
		<!-- 상단 띠 시작 -->
		<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
			<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
				<td>부서명</td>
				<td width="80">활동여부</td>
			  <td width="80">이익부서여부</td>
			  <td width="80">전시여부</td>
			</tr>
			<%  Dim oldPCD
			IF isArray(arrList) THEN
				For intLoop = 0 To UBound(arrList,2)
					IF oldPCD <> arrList(2,intLoop) THEN
				%>
				<tr height=30 align="center" bgcolor="<%IF arrList(3,intLoop) ="N" THEN%>#EFEFEF<%ELSE%>#FFFFFF<%END IF%>">
				<td align="left"><%=arrList(2,intLoop)%>&nbsp; <%=arrList(4,intLoop)%></td>
				<td></td>
				<td></td>
				<td></td>
			 </tr>
			<%	END IF%>
			<tr height=30 align="center" bgcolor="<%IF arrList(3,intLoop) ="N" THEN%>#EFEFEF<%ELSE%>#FFFFFF<%END IF%>">
				<td align="left">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
					└ <%=arrList(7,intLoop)%>&nbsp; <%=arrList(1,intLoop)%>
					<% if arrList(7,intLoop)<>arrList(0,intLoop) then %>
					&nbsp;<font color="#CCCCCC">(<%=arrList(0,intLoop)%>)</font>
					<% end if %>
				</td>
				<td><%=arrList(3,intLoop)%></td>
				<td><%IF not arrList(6,intLoop) THEN%>N<%ELSE%><font color="blue">Y</font><%END IF%> <a href="javascript:jsChType(1,'<%=arrList(0,intLoop)%>','<%IF  not arrList(6,intLoop) THEN%>1<%ELSE%>0<%END IF%>');"><img src="/images/icon_arrow_link.gif" border="0"></a> </td>
				<td><%IF not arrList(5,intLoop) THEN%>N<%ELSE%><font color="blue">Y</font><%END IF%> <a href="javascript:jsChType(2,'<%=arrList(0,intLoop)%>','<%IF  not arrList(5,intLoop)   THEN%>1<%ELSE%>0<%END IF%>');"><img src="/images/icon_arrow_link.gif" border="0"></a> </td>
			 </tr>
		<%		oldPCD  = arrList(2,intLoop)
				Next
			ELSE%>
			<tr height=5 align="center" bgcolor="#FFFFFF">
				<td colspan="4">등록된 내용이 없습니다.</td>
			</tr>
		<%END IF%>
		</table>
	</td>
</tr>
</table>
<!-- 페이지 끝 -->
</body>
</html>
 <!-- #include virtual="/lib/db/dbclose.asp" -->
