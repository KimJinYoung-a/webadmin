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
' Hieditor : 2011.04.07 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/stylepick/stylepick_cls.asp"-->

<%
dim menupos ,omain ,catename ,mainimagelink ,contentsyn , omaincontents ,i , comment
dim mainidx,cd1,mainimage,state,startdate,enddate,isusing,regdate,lastadminid,opendate,closedate,partMDid,partWDid
	mainidx = request("mainidx")
	menupos = request("menupos")
	
'//������
set omain = new cstylepick
	omain.frectmainidx = mainidx
	
	if mainidx <> "" then
		omain.fnGetmain_item()
		
		if omain.ftotalcount > 0 then
			mainimagelink = omain.foneitem.fmainimagelink			
			mainidx = omain.foneitem.fmainidx
			cd1 = omain.foneitem.fcd1
			mainimage = omain.foneitem.fmainimage
			state = omain.foneitem.fstate
			startdate = left(omain.foneitem.fstartdate,10)
			enddate = left(omain.foneitem.fenddate,10)
			isusing = omain.foneitem.fisusing
			regdate = omain.foneitem.fregdate
			lastadminid = omain.foneitem.flastadminid
			opendate = omain.foneitem.fopendate
			closedate = omain.foneitem.fclosedate
			partMDid = omain.foneitem.fpartMDid
			partWDid = omain.foneitem.fpartWDid
			contentsyn = omain.foneitem.fcontentsyn
			comment = omain.foneitem.fcomment
		end if	
	end if
set omain = nothing
	
if isusing = "" then isusing = "Y"
if mainimagelink = "" then mainimagelink = "<map name='Mapmainimage'></map>"	
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
		winImg = window.open('pop_event_mainimg.asp?sImg='+sImg+'&sName='+sName+'&sSpan='+sSpan,'popImg','width=370,height=150');
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

	//�۾����� ȭ�� ����
	function jsChkDisp(){
		var tmp = '<%=mainidx%>';
		
		if (tmp==''){
			alert('�۾����� �߰������ �űԵ�Ͽ��� ���� �ϽǼ� �����ϴ�.\n������ �������� �߰�����ϼ���')
			document.frm.contentsyn.checked = false
			return;
		}
		
		if(document.frm.contentsyn.checked){
			eDetail.style.display = "";
		}else{
			eDetail.style.display = "none";
		}
	}
		
	//����
	function jsEvtSubmit(){

		if(!frm.cd1.value){
			alert("��Ÿ�� ī�װ��� �������ּ���");
			frm.cd1.focus();
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
			frm.enddate.focus();
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
		if mainidx <> "" then
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
			//	frm.enddate.focus();
			//	return;
			//}
			
			if (frm.contentsyn.checked){			
				var gubun;	var gubunvaluetmp;
				gubun = document.getElementsByName("gubun")
				gubunvaluetmp = document.getElementsByName("gubunvalue")
	
				for (var i=0; i < gubun.length; i++){			
					if ( gubun[i].value == ''){
						alert("������ �����ϼ���");
						gubun[i].focus();
						return;
					}
				}
											
				for (var i=0; i < gubunvaluetmp.length; i++){			
					if ( gubunvaluetmp[i].value == ''){
						alert("��ȹ���ڵ峪 ��ǰ�ڵ带 �Է����ּ���");
						gubunvaluetmp[i].focus();
						return;
					}
				}
			}
				
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
	
	//tr�߰�
	function AutoInsert() {
		var f = document.all;
	
		var rowLen = f.div1.rows.length;
		var r  = f.div1.insertRow(rowLen++);
		var c0 = r.insertCell(0);
		
		var Html;
	
		c0.innerHTML = "&nbsp;";
		var inHtml = "&nbsp;&nbsp;&nbsp;<img src='http://fiximage.10x10.co.kr/web2009/common/cmt_del.gif' border='0' style='cursor:pointer' onClick='clearRow(this)'><table width='100%' align='center' cellpadding=0 cellspacing=0 border=0 class='a'><tr align='center'><td rowspan=3 valign='top'>�����μҽ�"+rowLen+"</td><td>����</td><td align='left'><select name='gubun' onchange='searchcode(this.value,"+rowLen+");'><option value=''>�����ϼ���</option><option value='1'>��Ÿ���� ��ȹ��</option><option value='2'>��ǰ</option></select><div id='divsub"+rowLen+"'>��ȹ���ڵ� & ��ǰ�ڵ� : <input type='text' name='gubunvalue' size=10 maxlength=10></div></td></tr><tr align='center'><td>ī��</td><td align='left'><input type='text' name='copy' size=90 maxlength=50></td></tr><tr align='center'><td>��ũ��</td><td align='left'><input type='text' name='link' size=90 maxlength=50></td></tr></table>";
		c0.innerHTML = inHtml;
	}
	
	//tr����
	function clearRow(tdObj) {
		if(confirm("������ �����ž� ������ �Ϸ� �˴ϴ�\n �����Ͻ� ������ �����Ͻðڽ��ϱ�?") == true) {
			var tblObj = tdObj.parentNode.parentNode.parentNode;
			var trIdx = tdObj.parentNode.parentNode.rowIndex;
		
			tblObj.deleteRow(trIdx);
		} else {
			return false;
		}
	}

	//��Ÿ�� & ��ǰ �˻�
	function searchcode(gubun,num){
		
		if (gubun.value==''){
			alert('������ �����ϼ���');			
			return;
		}
				
		if (frm.cd1.value==''){
			alert('ī�װ��� �����ϼ���');
			frm.cd1.focus();
			return;
		}
			
		//��Ÿ���� ��ȹ�� 
		if (gubun=='1'){
			var searchcode = window.open('/admin/stylepick/stylepick_main_search_event.asp?num='+num+'&cd1=<%=cd1%>','searchcode','width=1024,height=768,scrollbars=yes,resizable=yes');
		
		//��ǰ
		}else if (gubun=='2'){
			var searchcode = window.open('/admin/stylepick/stylepick_main_search_item.asp?num='+num+'&cd1=<%=cd1%>','searchcode','width=1024,height=768,scrollbars=yes,resizable=yes');
		}				
	}
	
</script>

<table width="100%" border="0" cellpadding="2" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="post" action="/admin/stylepick/stylepick_main_process.asp">
<input type="hidden" name="mode" value="mainedit">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="mainimage" value="<%=mainimage%>">
<input type="hidden" name="opendate" value="<%=opendate%>">
<input type="hidden" name="closedate" value="<%=closedate%>">
<tr>
	<td bgcolor="<%= adminColor("tabletop") %>" align="center">��ȣ</td>
	<td bgcolor="#FFFFFF"><%= mainidx %><input type="hidden" name="mainidx" value="<%=mainidx%>"></td>
</tr>	
<tr>
	<td bgcolor="<%= adminColor("tabletop") %>" align="center">��Ÿ��</td>
	<td bgcolor="#FFFFFF"><% Drawcategory "cd1",cd1,"","CD1" %></td>
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
	<td bgcolor="<%= adminColor("tabletop") %>" align="center">�����̹���</td>
	<td bgcolor="#FFFFFF">		
		<input type="button" name="btnBan2011" value="����̹������" onClick="jsSetImg('<%=mainimage%>','mainimage','mainimagediv')" class="button">
		<div id="mainimagediv" style="padding: 5 5 5 5">
			<%IF mainimage <> "" THEN %>			
				<img src="<%=mainimage%>" border="0" width=100 height=100 onclick="jsImgView('<%=mainimage%>');" alt="�����ø� Ȯ�� �˴ϴ�">
				<a href="javascript:jsDelImg('mainimage','mainimagediv');"><img src="/images/icon_delete2.gif" border="0"></a>
			<%END IF%>
		</div>
	</td>
</tr>
<tr>
	<td bgcolor="<%= adminColor("tabletop") %>" align="center">�����̹�����</td>
	<td bgcolor="#FFFFFF">
		�� �� �̸� ���� ���� ������<br>
		<textarea name="mainimagelink" cols="80" rows="6"><%=mainimagelink%></textarea>
	</td>
</tr>
<tr>
	<td bgcolor="<%= adminColor("tabletop") %>" align="center">�۾����޻���</td>
	<td bgcolor="#FFFFFF">
		<textarea cols="80" rows="6" name="comment"><%=nl2br(comment)%></textarea>
	</td>
</tr>
<tr>
	<td bgcolor="<%= adminColor("tabletop") %>" colspan=2>		 
		�۾����� �߰���� <input type="checkbox" name="contentsyn" onClick="jsChkDisp();" <%IF contentsyn= "Y" THEN%>checked<%END IF%>>
		<br><br>�� ����Ʈ ��ũ ����
		<br>- ����
		<br>&nbsp; &nbsp; &nbsp; (ī�װ��ڵ帵ũ) : &nbsp; /stylepick/index.asp?cd1=��Ÿ��ī�װ��ڵ�
		<br>&nbsp; &nbsp; &nbsp; (���ι�ȣ��ũ) : &nbsp; /stylepick/index.asp?mainidx=���ι�ȣ
		<br>- ��ȹ��
		<br>&nbsp; &nbsp; &nbsp; (ī�װ��ڵ帵ũ) : &nbsp; /stylepick/stylepick_collect.asp?cd1=��Ÿ��ī�װ��ڵ�
		<br>&nbsp; &nbsp; &nbsp; (��ȹ����ȣ��ũ) : &nbsp; /stylepick/stylepick_collect.asp?evtidx=��ȹ����ȣ
		<br>- ��ǰ����Ʈ
		<br>&nbsp; &nbsp; &nbsp;  : &nbsp; /stylepick/stylepick_list.asp?cd1=��Ÿ��ī�װ��ڵ�&cd2=�з�ī�װ��ڵ�
		<br>- ��ǰ������
		<br>&nbsp; &nbsp; &nbsp;  : &nbsp; /shopping/category_prd_stylepick.asp?itemid=��ǰ��ȣ
	</td>
</tr>
<%
'/������忡���� ��������
if mainidx <> "" then
%>
<tr align="center"  bgcolor="#FFFFFF" id="eDetail" style="display:<%IF contentsyn="N" THEN%>none;<%END IF%>">
	<td colspan=2>
		<table width="100%" cellpadding="0" cellspacing="0" border=0 class="a" id="div1">
		<%
		set omaincontents = new cstylepick
			omaincontents.frectisusing = "Y"
			omaincontents.frectmainidx = mainidx	
			omaincontents.fnGetmainctList()
		
		'/�������� ����
		if omaincontents.fresultcount > 0 then
		
		for i = 0 to omaincontents.fresultcount - 1
		%>
		<tr>
			<td><br>&nbsp;&nbsp;&nbsp;<img src='http://fiximage.10x10.co.kr/web2009/common/cmt_del.gif' border='0' style='cursor:pointer' onClick='clearRow(this)'>
				<table width='100%' align='center' cellpadding=0 cellspacing=0 border=0 class='a'>
				<tr align='center'>
					<td rowspan=3 valign='top'>�����μҽ�<%=i+1%></td>
					<td>����</td>
					<td align='left'>
						<select name='gubun' onchange='searchcode(this.value,<%=i+1%>);'>
							<option value='' <% if omaincontents.FItemList(i).fgubun="" then response.write " selected" %>>�����ϼ���</option>
							<option value='1' <% if omaincontents.FItemList(i).fgubun="1" then response.write " selected" %>>��Ÿ���� ��ȹ��</option>
							<option value='2' <% if omaincontents.FItemList(i).fgubun="2" then response.write " selected" %>>��ǰ</option>
						</select>						
						<div id='divsub<%=i+1%>'>��ȹ���ڵ� & ��ǰ�ڵ� : <input type='text' name='gubunvalue' value='<%=omaincontents.FItemList(i).fgubunvalue%>' size=10 maxlength=10></div>
					</td>
				</tr>
				<tr align='center'>
					<td>ī��</td>
					<td align='left'><input type='text' name='copy' value='<%=trim(omaincontents.FItemList(i).fcopy)%>' size=90 maxlength=50></td>
				</tr>
				<tr align='center'>
					<td>��ũ��</td>
					<td align='left'><input type='text' name='link' value='<%=trim(omaincontents.FItemList(i).flink)%>' size=90 maxlength=50></td>
				</tr>
				</table>
			</td>
		</tr>
		<%
		next
		
		'/�ű�
		else
		%>
		<tr>
			<td><br>
				<table width='100%' align='center' cellpadding=0 cellspacing=0 border=0 class='a'>
				<tr align='center'>
					<td rowspan=3 valign='top'>�����μҽ�1</td>
					<td>����</td>
					<td align='left'>
						<select name='gubun' onchange='searchcode(this.value,1);'><option value=''>�����ϼ���</option><option value='1'>��Ÿ���� ��ȹ��</option><option value='2'>��ǰ</option></select>						
						<div id='divsub1'>��ȹ���ڵ� & ��ǰ�ڵ� : <input type='text' name='gubunvalue' size=10 maxlength=10></div>
					</td>
				</tr>
				<tr align='center'>
					<td>ī��</td>
					<td align='left'><input type='text' name='copy' size=90 maxlength=50></td>
				</tr>
				<tr align='center'>
					<td>��ũ��</td>
					<td align='left'><input type='text' name='link' size=90 maxlength=50></td>
				</tr>
				</table>
			</td>
		</tr>
		<% end if %>			
		</table><br>
		<table width="100%" cellpadding="0" cellspacing="0" border=0 class="a">
		<tr>	
			<td bgcolor="#FFFFFF" colspan=2>	
				<input type="button" value="�����μҽ� 1�� �߰�" onClick="AutoInsert()" class="button">
			</td>
		</tr>
		</table>		
	</td>
</tr>
<% end if %>
<tr>
	<td bgcolor="#FFFFFF" colspan="2" align="center"><input type="button" onclick="jsEvtSubmit();" class="button" value="����"></td>
</tr>	
</form>
</table>

<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->