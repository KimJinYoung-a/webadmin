<%@ language=vbscript %>
<% option explicit %>
<%
	Response.AddHeader "Cache-Control","no-cache"
	Response.AddHeader "Expires","0"
	Response.AddHeader "Pragma","no-cache"
%>
<%
'###########################################################
' Description : 스타일픽 관리
' Hieditor : 2011.04.06 한용민 생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/stylepick/stylepick_cls.asp"-->

<%
dim menupos ,oevent ,catename
dim evtidx,title,subcopy,state,banner_img,startdate,enddate,isusing,regdate,comment
dim lastadminid,cd1,opendate,closedate,partMDid,partWDid
	evtidx = request("evtidx")
	menupos = request("menupos")

'//이벤트정보
set oevent = new cstylepick
	oevent.frectevtidx = evtidx
	
	if evtidx <> "" then
		oevent.fnGetEvent_item()
		
		if oevent.ftotalcount > 0 then			
			title = oevent.foneitem.ftitle
			subcopy = oevent.foneitem.fsubcopy
			state = oevent.foneitem.fstate
			banner_img = oevent.foneitem.fbanner_img
			startdate = left(oevent.foneitem.fstartdate,10)
			enddate = left(oevent.foneitem.fenddate,10)
			isusing = oevent.foneitem.fisusing
			regdate = oevent.foneitem.fregdate
			comment = oevent.foneitem.fcomment
			lastadminid = oevent.foneitem.flastadminid
			cd1 = oevent.foneitem.fcd1
			opendate = oevent.foneitem.fopendate
			closedate = oevent.foneitem.fclosedate
			partMDid = oevent.foneitem.fpartMDid
			partWDid = oevent.foneitem.fpartWDid
			catename = oevent.foneitem.fcatename
		end if	
	end if
set oevent = nothing
	
if isusing = "" then isusing = "Y"
%>

<script language="javascript">

	//-- jsPopCal : 달력 팝업 --//
	function jsPopCal(sName){
		var winCal;

		winCal = window.open('/lib/common_cal.asp?DN='+sName,'pCal','width=250, height=200');
		winCal.focus();
	}

	//이미지 확대화면 새창으로 보여주기
	function jsImgView(sImgUrl){
	 var wImgView;
	 wImgView = window.open('pop_event_detailImg.asp?sUrl='+sImgUrl,'pImg','width=100,height=100');
	 wImgView.focus();
	}

	function jsDelImg(sName, sSpan){
		if(confirm("이미지를 삭제하시겠습니까?\n\n삭제 후 저장버튼을 눌러야 처리완료됩니다.")){
		   eval("document.all."+sName).value = "";
		   eval("document.all."+sSpan).style.display = "none";
		}
	}

	function jsSetImg(sImg, sName, sSpan){	
		document.domain ="10x10.co.kr";
		
		var winImg;
		winImg = window.open('pop_event_uploadimg.asp?sImg='+sImg+'&sName='+sName+'&sSpan='+sSpan,'popImg','width=370,height=150');
		winImg.focus();
	}

	function jsNowDate(){
	var mydate=new Date() 
		var year=mydate.getYear() 
		    if (year < 1000) 
		        year+=1900 
		
		var day=mydate.getDay() 
		var month=mydate.getMonth()+1 
		    if (month<10) 
		        month="0"+month 
		
		var daym=mydate.getDate() 
		    if (daym<10) 
		        daym="0"+daym
		        
		return year+"-"+month+"-"+ daym      
	}
		
	//저장
	function jsEvtSubmit(){

		if(!frm.cd1.value){
			alert("카테고리를 선택해주세요");
			frm.cd1.focus();
			return;
		}

		if(!frm.title.value){
			alert("제목을 입력해주세요");
			frm.title.focus();
			return;
		}
		
		if(!frm.state.value){
			alert("상태를 선택해주세요");
			frm.state.focus();
			return;
		}
	
		if(!frm.startdate.value || !frm.enddate.value ){
			alert("이벤트 기간을 입력해주세요");
			return;
		}
	
		if(frm.startdate.value > frm.enddate.value){
			alert("종료일이 시작일보다 빠릅니다. 다시 입력해주세요");
			frm.evt_enddate.focus();
			return;
		}
	
		if(!frm.partmdid.value){
			alert("담당 MD를 선택하세요.");
			frm.partmdid.focus();
			return;
		}

		if(!frm.partwdid.value){
			alert("담당 WD를 선택하세요.");
			frm.partwdid.focus();
			return;
		}

		if(!frm.isusing.value){
			alert("사용여부를 선택하세요.");
			frm.isusing.focus();
			return;
		}
					
		var nowDate = jsNowDate();
	
		<%
		'//수정일경우
		if evtidx <> "" then
		%>
	
			if(<%=state%>==7 || <%=state%> ==9){
				if(frm.opendate.value != ""){
					nowDate = '<%IF opendate <> "" THEN%><%=FormatDate(opendate,"0000-00-00")%><%END IF%>';
				}
			}
	
			//if(<%=state%>==7 || <%=state%> ==9){
			//	if(frm.startdate.value > nowDate){
			//		alert("시작일이 오픈일보다  늦으면 안됩니다. 시작일을 다시 선택해주세요");
			//	  	frm.startdate.focus();
			//	  	return;
			//	}
			//}
	
			//if(frm.enddate.value < jsNowDate()){
			//	alert("종료일이 현재날짜보다 빠르면 안됩니다. 종료된 이벤트는 수정되지 않습니다");
			//	frm.evt_enddate.focus();
			//	return;
			//}
	
		<%
		'//신규등록
		else
		%>
	
			//if(frm.startdate.value < nowDate){
			//	alert("시작일이 현재일보다  빠르면 안됩니다. 시작일을 다시 선택해주세요");
			//	frm.enddate.focus();
			//	return false;
			//}
	
		<% end if %>
	
		frm.submit();
	}
	
</script>

<table width="100%" border="0" cellpadding="2" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" action="/admin/stylepick/stylepick_event_process.asp" method="post">
<input type="hidden" name="mode" value="eventedit">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="banner_img" value="<%=banner_img%>">
<input type="hidden" name="opendate" value="<%=opendate%>">
<input type="hidden" name="closedate" value="<%=closedate%>">
<tr>
	<td bgcolor="<%= adminColor("tabletop") %>" align="center">기획전번호</td>
	<td bgcolor="#FFFFFF"><%= evtidx %><input type="hidden" name="evtidx" value="<%=evtidx%>"></td>
</tr>	
<tr>
	<td bgcolor="<%= adminColor("tabletop") %>" align="center">스타일</td>
	<td bgcolor="#FFFFFF"><% Drawcategory "cd1",cd1,"","CD1" %></td>
</tr>	
<tr>
	<td bgcolor="<%= adminColor("tabletop") %>" align="center">제목</td>
	<td bgcolor="#FFFFFF"><input type="text" size=64 maxlength=64 name="title" value="<%=title%>"></td>
</tr>	
<tr>
	<td bgcolor="<%= adminColor("tabletop") %>" align="center">서브카피</td>
	<td bgcolor="#FFFFFF"><input type="text" size=64 maxlength=64 name="subcopy" value="<%=subcopy%>"></td>
</tr>
<tr>
	<td bgcolor="<%= adminColor("tabletop") %>" align="center">상태</td>
	<td bgcolor="#FFFFFF">
		<% Draweventstate "state" , state ,"" %>		
	</td>
</tr>		
<tr>
	<td bgcolor="<%= adminColor("tabletop") %>" align="center">기간</td>
	<td bgcolor="#FFFFFF">
   		<%IF state = "9" THEN%>
   			시작일 : <%=startdate%><input type="hidden" name="startdate" size=10 maxlength=10 value="<%=startdate%>">
   			~ 종료일 : <%=enddate%> <input type="hidden" name="enddate" size=10 maxlength=10 value="<%=enddate%>">
   		<%ELSE%>
   			시작일 : <input type="text" name="startdate" size=10 maxlength=10 value="<%=startdate%>" onClick="jsPopCal('startdate');"  style="cursor:hand;">
   			~ 종료일 : <input type="text" name="enddate" value="<%=enddate%>" size=10 maxlength=10 onClick="jsPopCal('enddate');" style="cursor:hand;">
   		<%END IF%>
   		<%
		if opendate <> "1900-01-01" and opendate <> "" then response.write " 오픈처리일 : " & opendate
		if closedate <> "1900-01-01" and closedate <> "" then response.write " 종료처리일 : " & closedate
		%>
	</td>
</tr>
<tr>
	<td bgcolor="<%= adminColor("tabletop") %>" align="center">담당MD</td>
	<td bgcolor="#FFFFFF">
		<% sbGetpartid "partmdid",partmdid,"","11" %>
	</td>
</tr>
<tr>
	<td bgcolor="<%= adminColor("tabletop") %>" align="center">담당WD</td>
	<td bgcolor="#FFFFFF">
		<% sbGetpartid "partwdid",partwdid,"","12" %>
	</td>
</tr>
<tr>
	<td bgcolor="<%= adminColor("tabletop") %>" align="center">사용여부</td>
	<td bgcolor="#FFFFFF">
		<% drawSelectBoxUsingYN "isusing", isusing %>
	</td>
</tr>
<tr>
	<td bgcolor="<%= adminColor("tabletop") %>" align="center">작업전달사항</td>
	<td bgcolor="#FFFFFF">
		<textarea rows=10 cols=100 name="comment"><%=nl2br(comment)%></textarea>
	</td>
</tr>
<tr>
	<td bgcolor="<%= adminColor("tabletop") %>" align="center">배너이미지</td>
	<td bgcolor="#FFFFFF">		
		<input type="button" name="btnBan2011" value="배너이미지등록" onClick="jsSetImg('<%=banner_img%>','banner_img','banner_imgdiv')" class="button">
		<div id="banner_imgdiv" style="padding: 5 5 5 5">
			<%IF banner_img <> "" THEN %>			
				<img src="<%=banner_img%>" border="0" width=100 height=100 onclick="jsImgView('<%=banner_img%>');" alt="누르시면 확대 됩니다">
				<a href="javascript:jsDelImg('banner_img','banner_imgdiv');"><img src="/images/icon_delete2.gif" border="0"></a>
			<%END IF%>
		</div>
	</td>
</tr>
<tr>
	<td bgcolor="#FFFFFF" colspan="2" align="center"><input type="button" onclick="jsEvtSubmit();" class="button" value="저장"></td>
</tr>	
</form>
</table>

<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
