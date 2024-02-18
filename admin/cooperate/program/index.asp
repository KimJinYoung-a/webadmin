<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/cooperate/programchangeCls.asp"-->
<!-- #include virtual="/lib/classes/admin/tenbyten/TenByTenMemberCls.asp" -->

<%
	Dim cPrCh, i, iPageSize, iCurrentpage, arrList, iTotCnt, intLoop, iL
	Dim iStartPage, iEndPage, iTotalPage, ix,iPerCnt, arrL
	Dim vRegUserID, vTitle, v1Check, v2Check, vParam
	vRegUserID		= requestCheckVar(Request("reguserid"),50)
	vTitle			= requestCheckVar(Request("title"),100)
	v1Check			= requestCheckVar(Request("1check"),1)
	v2Check			= requestCheckVar(Request("2check"),1)
	iCurrentpage 	= NullFillWith(requestCheckVar(Request("iC"),10),1)
	iPageSize 		= 20
	iPerCnt 		= 10
	

	set cPrCh = new CProgramChange
	 	cPrCh.FCPage = iCurrentpage
	 	cPrCh.FPSize = iPageSize
	 	cPrCh.FRectRegUserID = vRegUserID
	 	cPrCh.FRectTitle = vTitle
	 	cPrCh.FRect1Check = v1Check
	 	cPrCh.FRect2Check = v2Check
		arrList = cPrCh.fnGetPrChList
		iTotCnt = cPrCh.FTotCnt
		
		arrL = cPrCh.fnGetMemList
	set cPrCh = nothing
	
	iTotalPage 	=  int((iTotCnt-1)/iPageSize) +1
	
	vParam = "&menupos="&request("menupos")&"&reguserid="&vRegUserID&"&title="&vTitle&"&1check="&v1Check&"&2check="&v2Check&""
%>
<link href="/js/jqueryui/css/jquery-ui.css" rel="stylesheet">
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript" src="/js/jqueryui/jquery-ui-1.10.2.custom.min.js"></script>
<script language="Javascript">
function jsGoPage(iP){
	frm.iC.value = iP;
	frm.submit();
}
function goWrite(pidx)
{
	location.href = "write.asp?pidx="+pidx+"&iC=<%=iCurrentpage%><%=vParam%>";
}
function Check_All()
{
	var chk = document.frmp.pidx; 
	var cnt = 0;
	var ischecked = ""
	if(document.getElementById("chkall").checked){
		ischecked = "checked"
	}else{
		ischecked = ""
	}
	if(cnt == 0 && chk.length != 0){
		for(i = 0; i < chk.length; i++){ chk.item(i).checked = ischecked; }
		cnt++;
	}
}
function jsCheckAllReg(){
	var i = "";
	$("input:checkbox[name='pidx']").each(
		function(){
			if (this.checked)
			{
				i = i + this.value + ",";
			}
		}
	)
	
	if(i == ""){
		alert("선택된 내역이 없습니다.");
		return;
	}else{
		$('input[name="allpidx"]').val(i);
		frmp.submit();
	}
}
</script>

<form name="frm" action="index.asp" method="get" style="margin:0px;">
<input type="hidden" name="menupos" value="<%=request("menupos")%>">
<input type="hidden" name="iC" value="">
<table  cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
	<td align="left">
		<table class="a">
		<tr>
			<td>
				<select name="reguserid" class="select" onChange="frm.submit();">
				<option value="" <%=CHKIIF(vRegUserID="","selected","")%>>-작성자-</option>
				<%
					IF isArray(arrL) THEN
						For iL =0 To UBound(arrL,2)
							Response.Write "<option value=""" & arrL(0,iL) & """ " & CHKIIF(vRegUserID=arrL(0,iL),"selected","") & ">" & arrL(1,iL) & "</option>" & vbCrLf
						Next
					End If
				%>
				</select>
				제목:
				<input type="text" name="title" value="<%=vTitle%>" size="30">
			</td>
			<td rowspan="2" style="padding:0 0 0 30px;" align="right" valign="top"><input type="submit" value=" 검  색 " class="button" style="width:70px;height:50px;" onfocus="this.blur();"></td>
		</tr>
		<tr>
			<td>
				<select name="1check" class="select" onChange="frm.submit();">
					<option value="" <%=CHKIIF(v1Check="","selected","")%>>-1차결제-</option>
					<option value="x" <%=CHKIIF(v1Check="x","selected","")%>>전</option>
					<option value="o" <%=CHKIIF(v1Check="o","selected","")%>>완료</option>
				</select>
				&nbsp;
				<select name="2check" class="select" onChange="frm.submit();">
					<option value="" <%=CHKIIF(v2Check="","selected","")%>>-2차결제-</option>
					<option value="x" <%=CHKIIF(v2Check="x","selected","")%>>전</option>
					<option value="o" <%=CHKIIF(v2Check="o","selected","")%>>완료</option>
				</select>
			</td>
		</tr>
		</table>
	</td>
</tr>
</table>
</form>

<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
<tr>
	<td align="left">
		<input type="button" class="button" value="신규등록" onClick="location.href='write.asp?iC=<%=iCurrentpage%><%=vParam%>'">
	</td>
</tr>
<% If session("ssBctId") = "kobula" Then %>
<tr height="40">
	<td align="left">
		<input type="button" class="button" value="체크한것 결제" onClick="jsCheckAllReg()">
	</td>
</tr>
<% End If %>
</table>
<p>
<!-- 리스트 시작 -->
<form name="frmp" action="proc.asp" method="post" style="margin:0px;">
<input type="hidden" name="menupos" value="<%=request("menupos")%>">
<input type="hidden" name="gubun" value="allsign">
<input type="hidden" name="allpidx" value="">
<table cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="20">
		검색결과 : <b><%= iTotCnt %></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>" height="30">
	<td width="30"><input type="checkbox" name="chkall" id="chkall" value="" onClick="Check_All()"></td>
	<td width="60" align="center">No.</td>
	<td width="60" align="center">작성자</td>
	<td>제목</td>
	<td width="220">내용</td>
	<td width="150">작성일</td>
	<td width="150"></td>
</tr>
<%
	'####### A.pidx, A.title, A.contents, C.username, A.regdate, A.sign1, A.sign2, A.sign1date, A.sign2date, A.filename
	IF isArray(arrList) THEN
		For intLoop =0 To UBound(arrList,2)
%>
	<tr align="center" bgcolor="#FFFFFF" height="30" onmouseout="this.style.backgroundColor='#FFFFFF'" onmouseover="this.style.backgroundColor='#F1F1F1'" style="cursor:pointer">
		<td width="30"><input type="checkbox" name="pidx" value="<%=arrList(0,intLoop)%>"></td>
		<td onClick="goWrite('<%=arrList(0,intLoop)%>')"><%=arrList(0,intLoop)%></td>
		<td onClick="goWrite('<%=arrList(0,intLoop)%>')"><%=arrList(3,intLoop)%></td>
		<td align="left" onClick="goWrite('<%=arrList(0,intLoop)%>')"><%=db2html(arrList(1,intLoop))%><br><span style="color:green;">파일명:<%=chrbyte(arrList(9,intLoop),100,"Y")%></span></td>
		<td align="left" onClick="goWrite('<%=arrList(0,intLoop)%>')"><span onmouseover="subCID<%=arrList(0,intLoop)%>.style.display='block';" onmouseout="subCID<%=arrList(0,intLoop)%>.style.display='none';"><%=chrbyte(db2html(arrList(2,intLoop)),36,"Y")%></span>
			<div id='subCID<%=arrList(0,intLoop)%>' style='display:none; position:absolute; border:solid 1px #000000; width:200px; padding:3px; background-color:#ffffff;'><%=db2html(arrList(2,intLoop))%></div>
		</td>
		<td align="left" onClick="goWrite('<%=arrList(0,intLoop)%>')"><%=arrList(4,intLoop)%></td>
		<td onClick="goWrite('<%=arrList(0,intLoop)%>')">
			1차결제<% If arrList(5,intLoop) = "" Then %> <strong>전</strong> ,<% Else %> <strong>완료</strong> ,<% End If %>
			2차결제<% If arrList(6,intLoop) = "" Then %> <strong>전</strong><% Else %> <strong>완료</strong><% End If %>
		</td>
	</tr>
<%
		Next
	Else
%>
	<tr bgcolor="#FFFFFF" height="30">
		<td colspan="20" align="center" class="page_link">[데이터가 없습니다.]</td>
	</tr>
<%
	End If


iStartPage = (Int((iCurrentpage-1)/iPerCnt)*iPerCnt) + 1

If (iCurrentpage mod iPerCnt) = 0 Then
	iEndPage = iCurrentpage
Else
	iEndPage = iStartPage + (iPerCnt-1)
End If
%>
	
<tr height="25" bgcolor="FFFFFF">
	<td colspan="20" align="center">
		<% if (iStartPage-1 )> 0 then %><a href="javascript:jsGoPage(<%= iStartPage-1 %>)" onfocus="this.blur();">[pre]</a>
		<% else %>[pre]<% end if %>
        <%
			for ix = iStartPage  to iEndPage
				if (ix > iTotalPage) then Exit for
				if Cint(ix) = Cint(iCurrentpage) then
		%>
			<a href="javascript:jsGoPage(<%= ix %>)" class="menu_link3" onfocus="this.blur();"><font color="red">[<%=ix%>]</font></a>
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
</form>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->