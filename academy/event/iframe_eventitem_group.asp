<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  이벤트 그룹보기
' History : 2010.09.28 한용민 생성
'####################################################
%>
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/academy/lib/academy_function.asp"-->
<!-- #include virtual="/academy/lib/classes/Event_cls.asp"-->
<%
Dim eCode : eCode = RequestCheckvar(Request("eC"),10)
Dim sType : sType = RequestCheckvar(Request("T"),10)
Dim cEGroup, arrList,intg

set cEGroup = new ClsEventGroup
	cEGroup.FECode = eCode  	
	arrList = cEGroup.fnGetEventItemGroup		
set cEGroup = nothing 
%>

<script language="javascript" defer>

  function jsAddGroup(eCode,gCode){
  	var winG
  	
  	<% if eCode = "" then %>
  	alert('이벤트 처음 등록시에는 그룹형을 선택 하실수 없습니다. \n이벤트를 우선 저장하신후 수정으로 그룹형을 지정해 주세요');
  	return;
  	<% end if %>
  	
  	winG = window.open('pop_eventitem_group.asp?eC='+eCode+'&eGC='+gCode,'popG','width=600, height=500');
  	winG.focus();
  }
  
  function jsDelGroup(eCode,gCode){
   if(confirm("삭제시 하위그룹들 모두 삭제됩니다. 삭제하시겠습니까? ")){
  	document.frmD.eGC.value = gCode; 
  	document.frmD.submit();
  	}
  }
  
  //-- jsImgView : 이미지 확대화면 새창으로 보여주기 --//
	function jsImgView(sImgUrl){
	 var wImgView;
	 wImgView = window.open('/admin/eventmanage/common/pop_event_detailImg.asp?sUrl='+sImgUrl,'pImg','width=100,height=100');
	 wImgView.focus();
	}
	

</script>

<%IF sType="1" THEN%><body onunload="opener.location.href='eventitem_regist.asp?eC=<%=eCode%>';"><%END IF%>
<form name="frmD" method="post" action="event_process.asp">
<input type="hidden" name="imod"  value="gD">
<input type="hidden" name="eC" value="<%=eCode%>">
<input type="hidden" name="eGC" value="">
</form>
<table width="650" border="0" cellpadding="1" cellspacing="3" class="a">		   							   					
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
			<tr>
				<td  align="center" bgcolor="#FFFFFF"><%IF arrList(5,intg) <> 0 THEN%><img src="/images/L.png">&nbsp;<%END IF%><%=arrList(0,intg)%></td>						
				<td  align="center" bgcolor="#FFFFFF"><%IF isnull(arrList(7,intg))THEN%>최상위<%ELSE%>[<%=arrList(5,intg)%>]<%=db2html(arrList(7,intg))%><%END IF%></td>	
				<td  align="center" bgcolor="#FFFFFF"><%=db2html(arrList(1,intg))%></td>	
				<td  align="center" bgcolor="#FFFFFF"><%=arrList(2,intg)%></td>									   									
				<td  align="center" bgcolor="#FFFFFF"><%IF arrList(3,intg) <> "" THEN%><a href="javascript:jsImgView('<%=arrList(3,intg)%>');"><img src="<%=arrList(3,intg)%>" width="50" border="0"></a><%END IF%></td>					   									
				<td  align="center" bgcolor="#FFFFFF">
					<input type="button" name="btnU" value="수정" onclick="jsAddGroup('<%=eCode%>','<%=arrList(0,intg)%>')" class="button">
					<input type="button" name="btnD" value="삭제" onclick="jsDelGroup('<%=eCode%>','<%=arrList(0,intg)%>')"  class="button">
				</td>					   									
			</tr>
			<%NEXT%>
			</table>
		<%END IF%>	
	</td>
</tr>				
</table>
	
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->