<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  관리/전시 카테고리 매칭
' History : 2014.5.16 정윤정 생성 
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/admin/lib/incPageFunction.asp"--> 
<!-- #include virtual="/lib/classes/displaycate/displaycateCls.asp"-->
<!-- #include virtual="/lib/classes/displaycate/categoryMatchingCls.asp"-->
<% 
dim cdl, cdm, cds, dispCate,blnNotMatching
dim clsCM, arrList, intLoop
dim iTotCnt, iCurrPage, iPageSize
dim parm

cdl = requestCheckvar(request("cdl"),10)
cdm = requestCheckvar(request("cdm"),10)
cds = requestCheckvar(request("cds"),10)
dispCate = requestCheckvar(request("disp"),16)
blnNotMatching= requestCheckvar(request("blnNotM"),1)
iCurrPage= requestCheckvar(request("iCP"),10)
 
if dispCate ="" then dispCate = 101
if iCurrPage = "" THEN iCurrPage = 1 
iPageSize = 30
 
set clsCM = new CCategoryMatching
	clsCM.FRectCateLarge = cdl
	clsCM.FRectCateMid = cdm
	clsCM.FRectCateSmall = cds
	clsCM.FRectDispCate = dispCate
	clsCM.FRectIsNotMatching = blnNotMatching
	clsCM.FPageSize = iPageSize
	clsCM.FCurrPage = iCurrPage 
	arrList = clsCM.fnGetCategoryList
	iTotCnt = clsCM.FTotCnt
set clsCM = nothing

 
%>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript">
	function jsSetMatching(dispcate){
		var winM = window.open("/admin/categoryMaster/matching/popMatching.asp?disp="+dispcate ,"popM","width=800,height=300,scrollbars=yes,resizable=yes");
		winM.focus();
	}

	function jsDelMatchingSelect(){
		if($(".selChkBox").is(":checked")) {
			if(confirm("선택한 매칭을 삭제하시겠습니까?")) {
				document.frmAct.submit();
			}
		} else {
			alert("삭제할 매칭을 선택해주세요.")
		}
	}

	function fnCheckAll(obj) {
		$(".selChkBox").not(":disabled").prop('checked',obj.checked);
	}
</script>
<table width="100%" align="center" cellpadding="5" cellspacing="1" class="a">
<tr>
	<td> 
		<!-- 검색 시작 -->
		<form name="frm" method="GET">
		<input type="hidden" name="menupos" value="<%= menupos %>">
		<input type="hidden" name="page" >
		<table width="100%" align="center" cellpadding="10" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>"> 
			<tr align="center" bgcolor="<%= adminColor("topbar") %>" >
				<td  width="80" bgcolor="<%= adminColor("gray") %>">검색조건</td>
				<td align="left">
					전시카테고리: <!-- #include virtual="/common/module/dispCateSelectBox.asp"-->  
					&nbsp;&nbsp; 관리<!-- #include virtual="/common/module/categoryselectbox.asp"--> 
					&nbsp;&nbsp;<input type="checkbox" name="blnNotM" value="Y" <%IF blnNotMatching = "Y" THEN%>checked<%END IF%>>미매칭내역만
				</td> 
				<td   width="50" bgcolor="<%= adminColor("gray") %>">
					<input type="button" class="button_s" value="검색" onClick="javascript:document.frm.submit();">
				</td>
			</tr>  
		</table>
		</form>
	</td>
</tr>
<tr>
	<td align="right"><input type="button" value="선택 삭제" class="button_auth" onClick="jsDelMatchingSelect()"></td>
</tr>
<tr>
	<td> 
		<form name="frmAct" method="POST" action="procMatching.asp" target="hidFrm">
		<input type="hidden" name="mode" value="del">
		<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>"> 
		<tr height="20" bgcolor="<%= adminColor("tabletop") %>" align="center">
			<td><input type="checkbox" onclick="fnCheckAll(this);" /></td>
			<td>전시 카테고리</td>
			<td>관리 카테고리</td> 
			<td>등록일</td>
			<td>매칭처리</td>
		</tr> 
		<%
			IF isArray(arrList) THEN
				For intLoop = 0 To UBound(arrList,2)
		%>
		<tr bgcolor="#ffffff">
			<td align="center"><input type="checkbox" name="chkdisp" class="selChkBox" value="<%=arrList(0,intLoop)%>" <%=chkIIF(arrList(3,intLoop)="" or isNull(arrList(3,intLoop)),"disabled","")%> /></td>
			<td width="500"><%=replace(arrList(2,intLoop),"^^"," > ")%></td>
			<td><%IF arrList(3,intLoop) <> "" THEN%><%=arrList(3,intLoop)%> > <%=arrList(4,intLoop)%> > <%=arrList(5,intLoop)%> <%END IF%></td> 
			<td align="center"><%if arrList(6,intLoop) <> "" then%><%=formatdate(arrList(6,intLoop),"0000-00-00")%><%end if%></td>
			<td align="center">
				<input type="button" value="매칭" class="button_s" onClick="jsSetMatching('<%=arrList(0,intLoop)%>')">
			</td>
		</tr>			
		<%
				Next
			END IF
		%>
		</table>
		</form>
	</td>
</tr>
</table>	
<iframe id="hidFrm" name="hidFrm" src="about:blank" frameborder="0" width="0" height="0"></iframe>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->