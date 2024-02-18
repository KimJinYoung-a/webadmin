<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/admin/tenbyten/TenByTenMemberCls.asp"-->
<%
Dim oMember, vArr, i, oldpart_sn, part_sn, vCID
Dim vEnd2, vEnd3, vEnd4
vCID = requestCheckvar(Request("cid"),32)
 
'사원리스트 가져오기
	Set oMember = new CTenByTenMember
	vArr = oMember.fnGetTeamPartList2017
	set oMember = nothing

	IF isArray(vArr) THEN
		For i = 0 To UBound(vArr,2)
			If vArr(1,i) = "2" Then
				vEnd2 = vArr(0,i)
			End If
			If vArr(1,i) = "3" Then
				vEnd3 = vArr(0,i)
			End If
			If vArr(1,i) = "4" Then
				vEnd4 = vArr(0,i)
			End If
		Next
	END IF

%>
<script type="text/javascript" src="/js/jquery-1.7.2.min.js"> </script>
<script type="text/javascript">
$(document).ready(function(){
<% If vCID = "" Then
	vCID = "1"
%>
	top.contents.location.href = "/common/pop_organization_chart_rightiframe.asp?cid=<%=vCID%>&default=o";
<% End If %>
});

function jsTeamPart(v){
	location.href = "/common/pop_organization_chart_leftiframe.asp?cid="+v+"";
	top.contents.location.href = "/common/pop_organization_chart_rightiframe.asp?cid="+v+"";
}
</script>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" border="0">
<tr>
	<td>
		<div id="divUL">
		<%
		'0~10 B.cid, depth, B.cid1, B.cid2, B.cid3, B.cid4, B.departmentname1, B.departmentname2, B.departmentname3, B.departmentname4, personcount
		Dim cid1, cid2, cid3, cid4, oldcid1, oldcid2, oldcid3, oldcid4, vOldDepth, vIsEndChk
		IF isArray(vArr) THEN
				For i = 0 To UBound(vArr,2)
					part_sn = vArr(0,i)
					cid1 = vArr(2,i)
					cid2 = vArr(3,i)
					cid3 = vArr(4,i)
					cid4 = vArr(5,i)
					
					
					if isnull(cid2) then cid2 = 0
					if isnull(cid3) then cid3 = 0
					if isnull(cid4) then cid4 = 0
					
					If i <> UBound(vArr,2) Then
						If vArr(1,i) <> vArr(1,i+1) Then
							vIsEndChk = True
						Else
							vIsEndChk = False
						End If
					End If
						

					if cid1 <> oldcid1 then
						if i <> 0 then
							if oldcid2 <> 0 then
								if oldcid3 <> 0 then
									if oldcid4 <> 0 then
										Response.Write "</div>"
									end if
									Response.Write "</div>"
								end if
								Response.Write "</div>"
							end if
							Response.Write "</div>"
						end if
		%>
						<div id="divP1<%=cid1%>">
							<span style="background:#<%=CHKIIF(CStr(vCID)=CStr(part_sn),"CDE8FF","FFFFFF")%>;cursor:pointer;" onClick="jsTeamPart('<%=part_sn%>');">
							<img src="/images/dtree/<%=CHKIIF(CStr(vCID)=CStr(part_sn),"openfolder","closedfolder")%>.png" align="absmiddle" id="Fimg<%=cid1%>"> <%=vArr(6,i)%> (<span id="tot<%=cid1%>"><%=vArr(10,i)%></span>)
							&nbsp;</span>
						</div>
						<div id="divB1<%=cid1%>">
		<%
					end if
				
	 	 		if cid2  <>  oldcid2 and cid2 <> 0 then
					if oldcid2<> 0 then
						if oldcid3 <> 0 then
							if oldcid4 <> 0 then
								Response.Write "</div>"
							end if
							Response.Write "</div>"
						end if
						Response.Write "</div>"
					end if
		%>
					<div id="divP2<%=cid2%>" style="padding:0 0 0 1;">
						<img src="/images/blank.png" align="absmiddle">
						<img src="/images/dtree/join<%=CHKIIF(vEnd2=part_sn,"bottom","")%>.gif" align="absmiddle" id="Timg<%=cid2%>">
						<span style="background:#<%=CHKIIF(CStr(vCID)=CStr(part_sn),"CDE8FF","FFFFFF")%>;cursor:pointer;" onClick="jsTeamPart('<%=part_sn%>');">
						<img src="/images/dtree/<%=CHKIIF(CStr(vCID)=CStr(part_sn),"openfolder","closedfolder")%>.png" align="absmiddle" id="Fimg<%=cid2%>"> <%=vArr(7,i)%> (<span id="tot<%=cid2%>"><%=vArr(10,i)%></span>)
						&nbsp;</span>
					</div>
					<div id="divB2<%=cid2%>">
					<script>
						var tot1 = parseInt($("#tot<%=cid1%>").text());
						tot1 = tot1 + <%=vArr(10,i)%>;
						$("#tot<%=cid1%>").text(tot1);
					</script>
		<%
				end if
				
				if cid3  <>  oldcid3 and cid3 <> 0 then
	 	 			if oldcid3<> 0 then
						if oldcid4 <> 0 then
							Response.Write "</div>"
						end if
						Response.Write "</div>"
					end if
		%>
					<div id="divP3<%=cid3%>" style="padding:0 0 0 1;">
						<img src="/images/blank.png" align="absmiddle">
						<img src="/images/<%=CHKIIF(vEnd2=cid2,"blank.png","dtree/line.gif")%>" align="absmiddle">
						<img src="/images/blank.png" align="absmiddle">
						<% If vEnd3 = cid3 Then %>
							<img src="/images/dtree/joinbottom.gif" align="absmiddle" id="Timg<%=cid3%>">
						<% Else %>
							<img src="/images/dtree/join<%=CHKIIF(cid2<>vArr(3,i+1),"bottom","")%>.gif" align="absmiddle" id="Timg<%=cid3%>">
						<% End If %>
						<span style="background:#<%=CHKIIF(CStr(vCID)=CStr(part_sn),"CDE8FF","FFFFFF")%>;cursor:pointer;" onClick="jsTeamPart('<%=part_sn%>');">
						<img src="/images/dtree/<%=CHKIIF(CStr(vCID)=CStr(part_sn),"openfolder","closedfolder")%>.png" align="absmiddle" id="Fimg<%=cid3%>"> <%=vArr(8,i)%> (<span id="tot<%=cid3%>"><%=vArr(10,i)%></span>)
						&nbsp;</span>
					</div>
					<div id="divB3<%=cid3%>">
					<script>
						var tot2 = parseInt($("#tot<%=cid2%>").text());
						tot2 = tot2 + <%=vArr(10,i)%>;
						$("#tot<%=cid2%>").text(tot2);
						
						var tot1 = parseInt($("#tot<%=cid1%>").text());
						tot1 = tot1 + <%=vArr(10,i)%>;
						$("#tot<%=cid1%>").text(tot1);
					</script>
		<%
				end if
				
				if cid4  <>  oldcid4 and cid4 <> 0 then
	 	 			if oldcid4<> 0 then
						Response.Write "</div>"
					end if
		%>
					<div id="divP4<%=cid4%>" style="padding:0 0 0 1;background:white;">
						<img src="/images/blank.png" align="absmiddle">
						<img src="/images/<%=CHKIIF(vEnd2=cid2,"blank.png","dtree/line.gif")%>" align="absmiddle">
						<img src="/images/blank.png" align="absmiddle">
						<img src="/images/<%=CHKIIF(vEnd3=cid3,"blank.png","dtree/line.gif")%>" align="absmiddle">
						<img src="/images/dtree/join<%=CHKIIF(vIsEndChk,"bottom","")%>.gif" align="absmiddle" id="Timg<%=cid4%>">
						<span style="background:#<%=CHKIIF(CStr(vCID)=CStr(part_sn),"CDE8FF","FFFFFF")%>;cursor:pointer;" onClick="jsTeamPart('<%=part_sn%>');">
						<img src="/images/dtree/<%=CHKIIF(CStr(vCID)=CStr(part_sn),"openfolder","closedfolder")%>.png" align="absmiddle" id="Fimg<%=cid4%>"> <%=vArr(9,i)%> (<%=vArr(10,i)%>)
						&nbsp;</span>
					</div>

					<script>
						var tot3 = parseInt($("#tot<%=cid3%>").text());
						tot3 = tot3 + <%=vArr(10,i)%>;
						$("#tot<%=cid3%>").text(tot3);

						var tot2 = parseInt($("#tot<%=cid2%>").text());
						tot2 = tot2 + <%=vArr(10,i)%>;
						$("#tot<%=cid2%>").text(tot2);
						
						var tot1 = parseInt($("#tot<%=cid1%>").text());
						tot1 = tot1 + <%=vArr(10,i)%>;
						$("#tot<%=cid1%>").text(tot1);
					</script>
					<div id="divB4<%=cid4%>">
		<%
				end if

				vOldDepth = vArr(1,i)
				oldcid1 = cid1
				oldcid2 = cid2
				oldcid3 = cid3
				oldcid4 = cid4
				
				If i = UBound(vArr,2) Then
					Response.Write "</div>"
				End If
			Next
		END IF
		%>
		</div>
	</td>
</tr>
</table>