<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Page : /admin/eventmanage/event/iframe_eventitem_group.asp
' Description :  이벤트 그룹보기
' History : 2007.02.22 정윤정 생성
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/eventmanage/common/event_function.asp"-->
<!-- #include virtual="/lib/classes/event/eventmanageCls.asp"-->
<%
Dim eCode : eCode = Request("eC")
Dim sType : sType = Request("T")
Dim ekind : ekind = Request("ekind")
Dim cEGroup, arrList,intg, vYear

 set cEGroup = new ClsEventGroup
 	cEGroup.FECode = eCode  	
 	cEGroup.FRectGroupDelInc ="N"
  	arrList = cEGroup.fnGetEventItemGroup		
  	vYear = cEGroup.FRegdate
 set cEGroup = nothing
 
dim isDelGrp 
%>
<script language="javascript" defer>
<!-- 
function jsAddGroup(eCode,gCode){
	var winG
	winG = window.open('pop_eventitem_group.asp?yr=<%=vYear%>&eC='+eCode+'&eGC='+gCode,'popG','width=600, height=500');
	winG.focus();
}

function jsDelGroup(eCode,gCode){
	if(confirm("삭제시 하위그룹들 모두 삭제됩니다. 삭제하시겠습니까? ")){
		document.frmD.eGC.value = gCode; 
		document.frmD.submit();
	}
}

function popRegItem(eCode, gCode){
	var wImgView;
	wImgView = window.open('/admin/eventmanage/event/eventitem_regist.asp?eC='+eCode+'&selG='+gCode,'pImg','width=900,height=800');
	wImgView.focus();
}

//-- jsImgView : 이미지 확대화면 새창으로 보여주기 --//
function jsImgView(sImgUrl){
	var wImgView;
	wImgView = window.open('/admin/eventmanage/common/pop_event_detailImg.asp?sUrl='+sImgUrl,'pImg','width=100,height=100');
	wImgView.focus();
}

//-->
</script>
<%IF sType="1" THEN%><body onunload="opener.location.href='eventitem_regist.asp?eC=<%=eCode%>';"><%END IF%>
<form name="frmD" method="post" action="event_process.asp">
<input type="hidden" name="imod"  value="gD">
<input type="hidden" name="eC" value="<%=eCode%>">
<input type="hidden" name="eGC" value="">
</form>
<table width="800" border="0" cellpadding="1" cellspacing="3" class="a">		   							   				

	
	<tr>
		<td>					   						
			<input type="button" value="그룹추가" onClick="jsAddGroup('<%=eCode%>','');" class="input_b">
		</td>
	</tr>
	<tr>
		<td>
			<%IF isArray(arrList) THEN %>
				<table width="100%" border="0" cellpadding="5" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
				<tr align="center">
					<td width="100" bgcolor="<%= adminColor("tabletop") %>">그룹코드</td>					
					<td width="100" bgcolor="<%= adminColor("tabletop") %>">상위그룹</td>
					<td bgcolor="<%= adminColor("tabletop") %>">그룹명</td>
					<td width="50" bgcolor="<%= adminColor("tabletop") %>">정렬순서</td>					
					<td width="100" bgcolor="<%= adminColor("tabletop") %>">이미지</td>
					<td width="100" bgcolor="<%= adminColor("tabletop") %>">관리</td>
				</tr>
				<%FOR intg = 0 To UBound(arrList,2)%>	
				<% isDelGrp = (arrList(8,intg)="N")  %>
				<tr >
					<td  align="center" <%= CHKIIF(isDelGrp,"bgcolor='#CCCCCC'","bgcolor='#FFFFFF'") %> ><%IF arrList(5,intg) <> 0 THEN%><img src="/images/L.png">&nbsp;<%END IF%><%=arrList(0,intg)%></td>						
					<td  align="center" <%= CHKIIF(isDelGrp,"bgcolor='#CCCCCC'","bgcolor='#FFFFFF'") %> ><%IF isnull(arrList(7,intg))THEN%>최상위<%ELSE%>[<%=arrList(5,intg)%>]<%=db2html(arrList(7,intg))%><%END IF%></td>	
					<td  align="center" <%= CHKIIF(isDelGrp,"bgcolor='#CCCCCC'","bgcolor='#FFFFFF'") %> ><%=db2html(arrList(1,intg))%></td>	
					<td  align="center" <%= CHKIIF(isDelGrp,"bgcolor='#CCCCCC'","bgcolor='#FFFFFF'") %> ><%=arrList(2,intg)%></td>									   									
					<td  align="center" <%= CHKIIF(isDelGrp,"bgcolor='#CCCCCC'","bgcolor='#FFFFFF'") %> ><%IF arrList(3,intg) <> "" THEN%><a href="javascript:jsImgView('<%=arrList(3,intg)%>');"><img src="<%=arrList(3,intg)%>" width="50" border="0"></a><%END IF%></td>					   									
					<td  align="center" <%= CHKIIF(isDelGrp,"bgcolor='#CCCCCC'","bgcolor='#FFFFFF'") %> >
					    <% if ( isDelGrp) then %>
					    삭제된 그룹
					    <% else %>
						<input type="button" name="btnU" value="수정" onclick="jsAddGroup('<%=eCode%>','<%=arrList(0,intg)%>')" class="button">
						<input type="button" name="btnD" value="삭제" onclick="jsDelGroup('<%=eCode%>','<%=arrList(0,intg)%>')"  class="button">
						<input type="button" name="btnD" value="상품등록" onclick="popRegItem('<%=eCode%>','<%=arrList(0,intg)%>')"  class="button">
						<% IF arrList(5,intg) = 0 THEN %>
						<%
							'이벤트 종류에 따른 프론트링크 페이지 선택
							Select Case ekind
								Case "26"		'모바일
									Response.Write "<a href='" & vmobileUrl & "/event/eventmain.asp?eventid=" & eCode & "&eGC="& arrList(0,intg) &"' target='_blank'>미리보기</a>"
								Case Else		'쇼핑찬스 및 기타
									Response.Write "<a href='" & vwwwUrl & "/event/eventmain.asp?eventid=" & eCode & "&eGC="& arrList(0,intg)&" ' target='_blank'>미리보기</a>"
							End Select
						%>
						<% END IF %>
					    <% end if %>
					</td>					   									
				</tr>
				<%NEXT%>
				</table>
			<%END IF%>	
		</td>
	</tr>				
</table>
	
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/admin/lib/poptail.asp"-->