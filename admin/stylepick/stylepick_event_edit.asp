<%@ language=vbscript %>
<% option explicit %>
<%
	Response.AddHeader "Cache-Control","no-cache"
	Response.AddHeader "Expires","0"
	Response.AddHeader "Pragma","no-cache"
%>
<%
'###########################################################
' Description : ��Ÿ���� ����
' Hieditor : 2011.04.06 �ѿ�� ����
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

'//�̺�Ʈ����
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

	//-- jsPopCal : �޷� �˾� --//
	function jsPopCal(sName){
		var winCal;

		winCal = window.open('/lib/common_cal.asp?DN='+sName,'pCal','width=250, height=200');
		winCal.focus();
	}

	//�̹��� Ȯ��ȭ�� ��â���� �����ֱ�
	function jsImgView(sImgUrl){
	 var wImgView;
	 wImgView = window.open('pop_event_detailImg.asp?sUrl='+sImgUrl,'pImg','width=100,height=100');
	 wImgView.focus();
	}

	function jsDelImg(sName, sSpan){
		if(confirm("�̹����� �����Ͻðڽ��ϱ�?\n\n���� �� �����ư�� ������ ó���Ϸ�˴ϴ�.")){
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
		
	//����
	function jsEvtSubmit(){

		if(!frm.cd1.value){
			alert("ī�װ��� �������ּ���");
			frm.cd1.focus();
			return;
		}

		if(!frm.title.value){
			alert("������ �Է����ּ���");
			frm.title.focus();
			return;
		}
		
		if(!frm.state.value){
			alert("���¸� �������ּ���");
			frm.state.focus();
			return;
		}
	
		if(!frm.startdate.value || !frm.enddate.value ){
			alert("�̺�Ʈ �Ⱓ�� �Է����ּ���");
			return;
		}
	
		if(frm.startdate.value > frm.enddate.value){
			alert("�������� �����Ϻ��� �����ϴ�. �ٽ� �Է����ּ���");
			frm.evt_enddate.focus();
			return;
		}
	
		if(!frm.partmdid.value){
			alert("��� MD�� �����ϼ���.");
			frm.partmdid.focus();
			return;
		}

		if(!frm.partwdid.value){
			alert("��� WD�� �����ϼ���.");
			frm.partwdid.focus();
			return;
		}

		if(!frm.isusing.value){
			alert("��뿩�θ� �����ϼ���.");
			frm.isusing.focus();
			return;
		}
					
		var nowDate = jsNowDate();
	
		<%
		'//�����ϰ��
		if evtidx <> "" then
		%>
	
			if(<%=state%>==7 || <%=state%> ==9){
				if(frm.opendate.value != ""){
					nowDate = '<%IF opendate <> "" THEN%><%=FormatDate(opendate,"0000-00-00")%><%END IF%>';
				}
			}
	
			//if(<%=state%>==7 || <%=state%> ==9){
			//	if(frm.startdate.value > nowDate){
			//		alert("�������� �����Ϻ���  ������ �ȵ˴ϴ�. �������� �ٽ� �������ּ���");
			//	  	frm.startdate.focus();
			//	  	return;
			//	}
			//}
	
			//if(frm.enddate.value < jsNowDate()){
			//	alert("�������� ���糯¥���� ������ �ȵ˴ϴ�. ����� �̺�Ʈ�� �������� �ʽ��ϴ�");
			//	frm.evt_enddate.focus();
			//	return;
			//}
	
		<%
		'//�űԵ��
		else
		%>
	
			//if(frm.startdate.value < nowDate){
			//	alert("�������� �����Ϻ���  ������ �ȵ˴ϴ�. �������� �ٽ� �������ּ���");
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
	<td bgcolor="<%= adminColor("tabletop") %>" align="center">��ȹ����ȣ</td>
	<td bgcolor="#FFFFFF"><%= evtidx %><input type="hidden" name="evtidx" value="<%=evtidx%>"></td>
</tr>	
<tr>
	<td bgcolor="<%= adminColor("tabletop") %>" align="center">��Ÿ��</td>
	<td bgcolor="#FFFFFF"><% Drawcategory "cd1",cd1,"","CD1" %></td>
</tr>	
<tr>
	<td bgcolor="<%= adminColor("tabletop") %>" align="center">����</td>
	<td bgcolor="#FFFFFF"><input type="text" size=64 maxlength=64 name="title" value="<%=title%>"></td>
</tr>	
<tr>
	<td bgcolor="<%= adminColor("tabletop") %>" align="center">����ī��</td>
	<td bgcolor="#FFFFFF"><input type="text" size=64 maxlength=64 name="subcopy" value="<%=subcopy%>"></td>
</tr>
<tr>
	<td bgcolor="<%= adminColor("tabletop") %>" align="center">����</td>
	<td bgcolor="#FFFFFF">
		<% Draweventstate "state" , state ,"" %>		
	</td>
</tr>		
<tr>
	<td bgcolor="<%= adminColor("tabletop") %>" align="center">�Ⱓ</td>
	<td bgcolor="#FFFFFF">
   		<%IF state = "9" THEN%>
   			������ : <%=startdate%><input type="hidden" name="startdate" size=10 maxlength=10 value="<%=startdate%>">
   			~ ������ : <%=enddate%> <input type="hidden" name="enddate" size=10 maxlength=10 value="<%=enddate%>">
   		<%ELSE%>
   			������ : <input type="text" name="startdate" size=10 maxlength=10 value="<%=startdate%>" onClick="jsPopCal('startdate');"  style="cursor:hand;">
   			~ ������ : <input type="text" name="enddate" value="<%=enddate%>" size=10 maxlength=10 onClick="jsPopCal('enddate');" style="cursor:hand;">
   		<%END IF%>
   		<%
		if opendate <> "1900-01-01" and opendate <> "" then response.write " ����ó���� : " & opendate
		if closedate <> "1900-01-01" and closedate <> "" then response.write " ����ó���� : " & closedate
		%>
	</td>
</tr>
<tr>
	<td bgcolor="<%= adminColor("tabletop") %>" align="center">���MD</td>
	<td bgcolor="#FFFFFF">
		<% sbGetpartid "partmdid",partmdid,"","11" %>
	</td>
</tr>
<tr>
	<td bgcolor="<%= adminColor("tabletop") %>" align="center">���WD</td>
	<td bgcolor="#FFFFFF">
		<% sbGetpartid "partwdid",partwdid,"","12" %>
	</td>
</tr>
<tr>
	<td bgcolor="<%= adminColor("tabletop") %>" align="center">��뿩��</td>
	<td bgcolor="#FFFFFF">
		<% drawSelectBoxUsingYN "isusing", isusing %>
	</td>
</tr>
<tr>
	<td bgcolor="<%= adminColor("tabletop") %>" align="center">�۾����޻���</td>
	<td bgcolor="#FFFFFF">
		<textarea rows=10 cols=100 name="comment"><%=nl2br(comment)%></textarea>
	</td>
</tr>
<tr>
	<td bgcolor="<%= adminColor("tabletop") %>" align="center">����̹���</td>
	<td bgcolor="#FFFFFF">		
		<input type="button" name="btnBan2011" value="����̹������" onClick="jsSetImg('<%=banner_img%>','banner_img','banner_imgdiv')" class="button">
		<div id="banner_imgdiv" style="padding: 5 5 5 5">
			<%IF banner_img <> "" THEN %>			
				<img src="<%=banner_img%>" border="0" width=100 height=100 onclick="jsImgView('<%=banner_img%>');" alt="�����ø� Ȯ�� �˴ϴ�">
				<a href="javascript:jsDelImg('banner_img','banner_imgdiv');"><img src="/images/icon_delete2.gif" border="0"></a>
			<%END IF%>
		</div>
	</td>
</tr>
<tr>
	<td bgcolor="#FFFFFF" colspan="2" align="center"><input type="button" onclick="jsEvtSubmit();" class="button" value="����"></td>
</tr>	
</form>
</table>

<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
