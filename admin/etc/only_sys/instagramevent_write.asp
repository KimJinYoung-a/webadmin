<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 인스타그램 이벤트용 수동 등록페이지
' Hieditor : 2016.06.23 유태욱 생성
'###########################################################
%>
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/classes/event/eventmanageCls.asp"-->
<!-- #include virtual="/admin/eventmanage/common/event_function.asp"-->
<!-- #include virtual="/admin/etc/only_sys/instagrameventCls.asp"-->

<%
Dim i, mode, contentsidx , ecode
dim instaidx,  evt_code, imgurl, instauserid, linkurl, isusing
	contentsidx = request("contentsidx")

dim oinsta
	set oinsta = new CInstagramevent
		oinsta.Frectcontentsidx=contentsidx
		if contentsidx <> "" then
			oinsta.fnGetinstagramevent_oneitem
			if oinsta.FResultCount > 0 then
				instaidx = oinsta.Foneitem.Fcontentsidx
				evt_code = oinsta.Foneitem.Fevt_code
				imgurl = oinsta.Foneitem.Fimgurl
				instauserid = oinsta.Foneitem.Fuserid
				linkurl = oinsta.Foneitem.Flinkurl
				isusing = oinsta.Foneitem.FIsusing
			end if
		end if

		
'만약 idx값이 없을경우(신규등록) NEW, 아닐경우(수정) EDIT	
if instaidx = "" then 
	mode="NEW"
	ecode = contentsidx
else
	mode="EDIT"
	ecode =	evt_code
end if
%>

<link href="/js/jqueryui/css/jquery-ui.css" rel="stylesheet">
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript" src="/js/jqueryui/jquery-ui-1.10.2.custom.min.js"></script>
<script type='text/javascript'>
	
function frmedit(){
	if(frm.evt_code.value==""){
		alert('이벤트 코드를 입력해 주세요');
		frm.evt_code.focus();
		return;
	}

	if(frm.userid.value==""){
		alert('게시자 ID를 입력해 주세요');
		frm.userid.focus();
		return;
	}

	if(frm.imgurl.value==""){
		alert('이미지 주소를 입력해 주세요');
		frm.imgurl.focus();
		return;
	}
	
	if(frm.linkurl.value==""){
		alert('게시물 링크를 입력해 주세요');
		frm.linkurl.focus();
		return;
	}
	
	frm.submit();
}

</script>

<form name="frm" method="post" action="instagramevent_proc.asp">
<input type="hidden" name="mode" value="<%=mode %>">
<input type="hidden" name="menupos" value="<%=menupos %>">
<input type="hidden" name="contentsidx" value="<%=contentsidx %>">

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr>
		<td align="left">
			<b>※인스타그램 이벤트 컨텐츠 수동 등록</b>
		</td>
	</tr>
</table>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<% IF contentsidx <> "" THEN%>
	<tr bgcolor="#FFFFFF">
		<td bgcolor="<%= adminColor("tabletop") %>" align="center">번호</td>
		<td colspan="2"><%=instaidx%></td>
	</tr>
	<% End if %>

	<tr bgcolor="#FFFFFF">
		<td bgcolor="<%= adminColor("tabletop") %>" align="center">이벤트코드</td>
		<td colspan="2">
			<input type="text" name="evt_code" size="10" value="<%= ecode %>"/>
		</td>
	</tr>
	
	<tr bgcolor="#FFFFFF">
		<td bgcolor="<%= adminColor("tabletop") %>" align="center">게시자ID</td>
		<td colspan="2">
			<input type="text" name="userid" size="25" value="<% if mode="NEW" then response.write "10x10" else response.write instauserid end if %>"/>
		</td>
	</tr>
	
	<tr bgcolor="#FFFFFF">
		<td bgcolor="<%= adminColor("tabletop") %>" align="center">이미지URL</td>
		<td colspan="2">
				<input type="text" name="imgurl" size="100" value="<%= imgurl %>" />
		</td>
	</tr>

	<tr bgcolor="#FFFFFF">
		<td bgcolor="<%= adminColor("tabletop") %>" align="center">게시물링크</td>
		<td colspan="2">
				<input type="text" name="linkurl" size="100" value="<%= linkurl %>" />
		</td>
	</tr>
	
	<tr bgcolor="#FFFFFF">
		<td bgcolor="<%= adminColor("tabletop") %>" align="center"> 사용여부 </td>
		<td colspan="2">
			<input type="radio" name="isusing" value="Y" <%=chkiif(isusing = "Y","checked","")%> checked />사용함 &nbsp;&nbsp;&nbsp; 
			<input type="radio" name="isusing" value="N"  <%=chkiif(isusing = "N","checked","")%>/>사용안함
		</td>
	</tr>

	<tr align="center" bgcolor="#FFFFFF">
		<td colspan="3">
			<% if mode = "EDIT" or mode = "NEW" then %>
				<input type="button" name="editsave" value="저장" onclick="frmedit()" />	
			<% end if %>
			
			<input type="button" name="editclose" value="취소" onclick="self.close()" />
		</td>
	</tr>
</table>
</form>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->