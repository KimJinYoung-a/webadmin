<%@ language=vbscript %>
<% option explicit %>
<%
'#############################################################
'	Description : HITCHHIKER ADMIN
'	History		: 2014.07.09 유태욱 생성
'#############################################################
%>
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/classes/event/eventmanageCls.asp"-->
<!-- #include virtual="/admin/eventmanage/common/event_function.asp"-->
<!-- #include virtual="/lib/classes/hitchhiker/about_hitchhiker_contentsCls.asp"-->
<%
Dim mode, i
dim deviceidx, device_name, contents_size, sortnum, isusing
	isusing = request("isusing")
	deviceidx = request("deviceidx")
	mode = requestCheckvar(Request("mode"),10)
	sortnum = requestCheckvar(Request("sortnum"),5)
	device_name = requestCheckvar(Request("device_name"),32)
	contents_size = requestCheckvar(Request("contents_size"),32)

Dim ohitpc
set ohitpc = new CAbouthitchhiker
	ohitpc.Frectgubun="1"
	ohitpc.fnGetDeviceList
	
dim ohitm
set ohitm = new CAbouthitchhiker
	ohitm.Frectgubun="2"
	ohitm.fnGetDeviceList
%>

<script type='text/javascript'>
//신규입력 펑션	
function frmreg(gubun){
	if (gubun==""){
		alert("장비구분이 지정되지 않았습니다.관리자 문의 해주세요.");
		return;
	}
	
	//장비구분에 따른 처리임
	//PC
	if (gubun=="1"){
		frm.contents_size.value=frm.newpcsizetextbox.value;
		if (frm.newpcsizetextbox.value==""){
			alert("사이즈를 입력해주세요");
			frm.newpcsizetextbox.focus();
			return;
		}
		
		var tempchkvalue = "";
		for (var i=0;i<frm.newpcsizeisusing.length;i++) {
			if (frm.newpcsizeisusing[i].checked==true) {
				tempchkvalue=frm.newpcsizeisusing[i].value;
			}
		}
		if(tempchkvalue==""){
			alert("사용여부를 선택 해주세요");
			return;
		}
		frm.isusing.value=tempchkvalue;
		frm.sortnum.value=frm.newpcsizesortnum.value;
		frm.contents_size.value=frm.newpcsizetextbox.value;
		frm.device_name.value=frm.newmobiledevicetextbox.value;
	//모바일
	}else{
		if (frm.newmobiledevicetextbox.value==""){
			alert("대표기종을 입력해주세요");
			frm.newmobiledevicetextbox.focus();
			return;
		}
		
		if (frm.newmobilesizetextbox.value==""){
			alert("사이즈를 입력해주세요");
			frm.newmobilesizetextbox.focus();
			return;
		}
		
		var temmobilechkvalue = "";
		for (var i=0;i<frm.newmobilesizeisusing.length;i++) {
			if (frm.newmobilesizeisusing[i].checked==true) {
				temmobilechkvalue=frm.newmobilesizeisusing[i].value;
			}
		}
		if(temmobilechkvalue==""){
			alert("사용여부를 선택 해주세요");
			return;
		}
		
		frm.isusing.value=temmobilechkvalue;
		frm.sortnum.value=frm.newmobilesizesortnum.value;
		frm.contents_size.value=frm.newmobilesizetextbox.value;
		frm.device_name.value=frm.newmobiledevicetextbox.value;
	}
	
	frm.deviceidx.value="";
	frm.gubun.value=gubun;
	frm.mode.value="sizeedit"
	frm.submit();
}

//사이즈 수정 펑션	
function frmedit(gubun,ix){
	if (gubun==""){
		alert("장비구분이 지정되지 않았습니다.관리자 문의 해주세요.");
		return;
	}
	if (ix==""){
		alert("수정구분이 지정되지 않았습니다.관리자 문의 해주세요.");
		return;
	}

	//장비구분에 따른 처리임
	//PC
	if (gubun=="1"){
		var tmpdeviceidx = eval("frm.pcsizedeviceidx_"+ix);  //gubun=1(PC)idx
		frm.deviceidx.value=tmpdeviceidx.value;

		var tmppcsizetextbox = eval("frm.pcsizetextbox_"+ix); //gubun=1(PC) 월페이퍼 사이즈
		frm.contents_size.value=tmppcsizetextbox.value;
		if (tmppcsizetextbox.value==""){
			alert("사이즈를 입력해주세요");
			eval("frm.pcsizetextbox_"+ix).focus();
			return;
		}

		var tmpsortnum = eval("frm.pcsizesortnum_"+ix); //gubun=1(PC) 우선순위
		frm.sortnum.value=tmpsortnum.value;
		if (tmpsortnum.value==""){
			alert("우선순위를 입력해 주세요");
			eval("frm.pcsizesortnum_"+ix).focus();
			return;
		}
		
		var tmppcsizeisusing = eval("frm.pcsizeisusing_"+ix); //gubun=1(PC)월페이퍼 사용여부
		var tempchkvalue = "";
		for (var i=0;i<tmppcsizeisusing.length;i++) {
			if (tmppcsizeisusing[i].checked==true) {
				tempchkvalue=tmppcsizeisusing[i].value;
			}
		}
		if(tempchkvalue==""){
			alert("사용여부를 선택 해주세요");
			return;
		}

		frm.isusing.value=tempchkvalue;
	//모바일
	}else{
		var tmpdeviceidx = eval("frm.mobiledeviceidx_"+ix); //gubun=2(모바일)idx
		frm.deviceidx.value=tmpdeviceidx.value;
		
		var tmpdevicename = eval("frm.mobiledevicetextbox_"+ix); //모바일기기명
		frm.device_name.value=tmpdevicename.value;
		if (tmpdevicename.value==""){
			alert("모바일 기기명을 입력해 주세요");
			eval("frm.mobiledevicetextbox_"+ix).focus();
			return;
		}

		var tmpmsizetextbox = eval("frm.mobilesizetextbox_"+ix); //gubun=2(모바일)월페이퍼 사이즈
		frm.contents_size.value=tmpmsizetextbox.value;
		if (tmpdevicename.value==""){
			alert("모바일 기기명을 입력해 주세요");
			eval("frm.mobiledevicetextbox_"+ix).focus();
			return;
		}

		var tmpsortnum = eval("frm.mobilesizesortnum_"+ix); //gubun=2(모바일)우선순위
		frm.sortnum.value=tmpsortnum.value;
		if (tmpsortnum.value==""){
			alert("우선순위를 입력해 주세요");
			eval("frm.mobilesizesortnum_"+ix).focus();
			return;
		}


		var tmpmsizeisusing = eval("frm.mobilesizeisusing_"+ix); //gubun=2(모바일)월페이퍼 사용여부
		var temmobilechkvalue = "";
		for (var i=0;i<tmpmsizeisusing.length;i++) {
			if (tmpmsizeisusing[i].checked==true) {
				temmobilechkvalue=tmpmsizeisusing[i].value;
			}
		}
		if(tempchkvalue==""){
			alert("사용여부를 선택 해주세요");
			return;
		}

		frm.isusing.value=temmobilechkvalue;
	}
	
	frm.gubun.value=gubun;
	frm.mode.value="sizeedit"
	frm.submit();
}
	
	function onlyNumDecimalInput(){  //한글 입력 안되게
	var code = window.event.keyCode; 
	
	if ((code >= 48 && code <= 57) || (code >= 96 && code <= 105) || code == 110 || code == 190 || code == 8 || code == 9 || code == 13 || code == 46){ 
		window.event.returnValue = true; 
		return; 
	} 
	window.event.returnValue = false; 
	}

</script>

<form name="frm" method="post" action="about_size_proc.asp" >
<input type="hidden" name="mode" >
<input type="hidden" name="gubun" >
<input type="hidden" name="isusing" >
<input type="hidden" name="sortnum" >
<input type="hidden" name="deviceidx" >
<input type="hidden" name="device_name" >
<input type="hidden" name="contents_size" >
<input type="hidden" name="menupos" value="<%=menupos %>">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr>
		<td align="left">
			<b>※월페이퍼 기본 사이즈</b>
		</td>
	</tr>
</table>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr bgcolor="<%=adminColor("tabletop")%>">
		<td>→ PC 월페이퍼 사이즈</td>
	</tr>
	
<!--PC월페이퍼 신규 사이즈 등록-->
	<tr bgcolor="FFFFFF">
		<td>
			사이즈	 <input type="text" name="newpcsizetextbox" value="" />
			우선순위 <input type="text" name="newpcsizesortnum" value="99" size="3" maxlength="3" onkeydown='onlyNumDecimalInput();' style="ime-mode:disabled" />
			사용여부 <input type="radio" name="newpcsizeisusing" value="Y" /> Y <input type="radio" name="newpcsizeisusing" value="N" /> N
					 <input type="button" value="신규등록" class="button" onclick="frmreg('1')" />
		</td>
	</tr>
<!--PC월페이퍼 신규 사이즈 등록 끝-->
	
	<tr bgcolor="FFFFFF">
		<td height=10></td>
	</tr>
	
<!--기존PC월페이퍼 사이즈 리스트-->
	<% if ohitpc.FResultCount > 0 then %>
		<% for i = 0 to ohitpc.FResultCount - 1 %>
			<tr bgcolor="FFFFFF">
				<td<% if ohitpc.FItemList(i).FIsusing = "N" then %> bgcolor="CCCCCC" <% else %> bgcolor="FFFFFF" <% end if %>>
							<input type="hidden" name="pcsizedeviceidx_<%= i %>" value="<%= ohitpc.FItemList(i).FDeviceidx %>" />
					사이즈	<input type="text" name="pcsizetextbox_<%= i %>" value="<%= trim(ohitpc.FItemList(i).FContentsSize) %>" />
					우선순위	<input type="text" name="pcsizesortnum_<%= i %>" value="<%= trim(ohitpc.FItemList(i).FSortnum) %>" size="3" maxlength="3" onkeydown='onlyNumDecimalInput();' style="ime-mode:disabled" />
					사용여부	<input type="radio" name="pcsizeisusing_<%= i %>" value="Y" <% if ohitpc.FItemList(i).FIsusing = "Y" then response.write " checked" %> /> Y
							<input type="radio" name="pcsizeisusing_<%= i %>" value="N" <% if ohitpc.FItemList(i).FIsusing = "N" then response.write " checked" %> /> N
							<input type="button" value="수정" class="button" onclick="frmedit('1','<%=i%>')"/>
				</td>
			</tr>
		<% next %>
	<% end if %>
<!--기존PC월페이퍼 사이즈 리스트 끝-->
</table>

<br>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr bgcolor="<%=adminColor("tabletop")%>">
		<td>→ MOBILE 월페이퍼 기종 및 사이즈</td>
	</tr>	
	
<!-- 모바일 신규 기기명 및 월페이퍼 사이즈 등록-->
	<tr bgcolor="FFFFFF">
		<td>
			대표기종 <input type="text" name="newmobiledevicetextbox" value="" />
			사이즈	 <input type="text" name="newmobilesizetextbox" value="" />
			우선순위 <input type="text" name="newmobilesizesortnum" value="99" size="3" maxlength="3" onkeydown='onlyNumDecimalInput();' style="ime-mode:disabled" />
			사용여부 <input type="radio" name="newmobilesizeisusing" value="Y" /> Y <input type="radio" name="newmobilesizeisusing" value="N" /> N
					 <input type="button" value="신규등록"  class="button" onclick="frmreg('2')"/>
		</td>
	</tr>
<!-- 모바일 신규 기기명 및 월페이퍼 사이즈 등록 끝-->
	
	<tr bgcolor="FFFFFF">
		<td height=10></td>
	</tr>
	
<!--기존모바일 기기명 및 월페이퍼 사이즈 리스트-->
	<% if ohitm.FResultCount > 0 then %>
		<% for i = 0 to ohitm.FResultCount - 1 %>
			<tr bgcolor="FFFFFF">
				<td <% if ohitm.FItemList(i).FIsusing = "N" then %> bgcolor="CCCCCC" <% else %> bgcolor="FFFFFF" <% end if %>>		
							<input type="hidden" name="mobiledeviceidx_<%= i %>" value="<%= ohitm.FItemList(i).FDeviceidx %>" />
					대표기종	<input type="text" name="mobiledevicetextbox_<%= i %>" value="<%= trim(ohitm.FItemList(i).FDevicename) %>" />
					사이즈	<input type="text" name="mobilesizetextbox_<%= i %>" value="<%= trim(ohitm.FItemList(i).FContentsSize) %>" />
					우선순위	<input type="text" name="mobilesizesortnum_<%= i %>" value="<%= trim(ohitm.FItemList(i).FSortnum) %>" size="3" maxlength="3" onkeydown='onlyNumDecimalInput();' style="ime-mode:disabled"/>
					사용여부	<input type="radio" name="mobilesizeisusing_<%= i %>" value="Y" <% if ohitm.FItemList(i).FIsusing = "Y" then response.write " checked" %> /> Y
							<input type="radio" name="mobilesizeisusing_<%= i %>" value="N" <% if ohitm.FItemList(i).FIsusing = "N" then response.write " checked" %> /> N
							<input type="button" value="수정"  class="button" onclick="frmedit('2','<%=i%>')"/>
				</td>
			</tr>
		<% next %>
	<% end if %>
<!--기존모바일 기기명 및 월페이퍼 사이즈 리스트 끝-->
</table>
</form>

<%
set ohitpc = nothing
set ohitm = nothing
%>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->