<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbCTopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/admin/etc/between/mainCls.asp"-->
<%
Dim idx, imgurl, mode, sortno, isusing, imglink
Dim mainStartDate, mainEndDate, gender
Dim sDt, eDt

idx = requestCheckvar(request("idx"),16)

If idx = "" Then
	mode = "I"
Else
	mode = "U"
End If

If idx <> "" then
	Dim o3ban
	SET o3ban = new cMain
		o3ban.FRectIdx = idx
		o3ban.GetOne3Banner()

		imgurl			= o3ban.FItemList(0).FImgurl
		sortno			= o3ban.FItemList(0).FSortno
		mainStartDate	= Left(o3ban.FItemList(0).FStartdate, 10)
		mainEndDate		= Left(o3ban.FItemList(0).FEnddate, 10)
		isusing			= o3ban.FItemList(0).FIsusing
		imgurl			= o3ban.FItemList(0).FImgurl
		imglink			= o3ban.FItemList(0).FImglink
		gender			= o3ban.FItemList(0).FGender
	SET o3ban = Nothing
End If
%>
<link href="/js/jqueryui/css/jquery-ui.css" rel="stylesheet">
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript" src="/js/jqueryui/jquery-ui-1.10.2.custom.min.js"></script>
<script type='text/javascript'>
function jsSubmit(){
	var frm = document.frm;
	if (confirm('���� �Ͻðڽ��ϱ�?')){
		frm.submit();
	}
}

function jsgolist(){
	self.location.href="/admin/etc/between/main/3banner/index.asp";
}

$(function(){
	//�޷´�ȭâ ����
	var arrDayMin = ["��","��","ȭ","��","��","��","��"];
	var arrMonth = ["1��","2��","3��","4��","5��","6��","7��","8��","9��","10��","11��","12��"];
    $("#sDt").datepicker({
		dateFormat: "yy-mm-dd",
		prevText: '������', nextText: '������', yearSuffix: '��',
		dayNamesMin: arrDayMin,
		monthNames: arrMonth,
		showMonthAfterYear: true,
    	numberOfMonths: 2,
    	showCurrentAtPos: 1,
      	showOn: "button",
      	<% if Idx<>"" then %>maxDate: "<%=eDt%>",<% end if %>
    	onClose: function( selectedDate ) {
    		$( "#eDt" ).datepicker( "option", "minDate", selectedDate );
    	}
    });
    $("#eDt").datepicker({
		dateFormat: "yy-mm-dd",
		prevText: '������', nextText: '������', yearSuffix: '��',
		dayNamesMin: arrDayMin,
		monthNames: arrMonth,
		showMonthAfterYear: true,
    	numberOfMonths: 2,
      	showOn: "button",
      	<% if Idx<>"" then %>minDate: "<%=sDt%>",<% end if %>
    	onClose: function( selectedDate ) {
    		$( "#sDt" ).datepicker( "option", "maxDate", selectedDate );
    	}
    });
});

function putLinkText(key,gubun) {
	var frm = document.frm;
	var urllink
	if (gubun == "3" ){
		urllink = frm.imglink;
	}
	switch(key) {
		case 'search':
			urllink.value='/apps/appCom/between/project/?pjt_code=�ڵ�';
			break;
	}
}
function jsSetImg(sFolder, sImg, sName, sSpan){
	document.domain ="10x10.co.kr";
	var winImg;
	winImg = window.open('/admin/etc/between/main/3banner/pop_3Banner_uploadimg.asp?sF='+sFolder+'&sImg='+sImg+'&sName='+sName+'&sSpan='+sSpan,'popImg','width=370,height=150');
	winImg.focus();
}
function jsDelImg(sName, sSpan){
	if(confirm("�̹����� �����Ͻðڽ��ϱ�?\n\n���� �� �����ư�� ������ ó���Ϸ�˴ϴ�.")){
	   eval("document.all."+sName).value = "";
	   eval("document.all."+sSpan).style.display = "none";
	}
}
</script>
<table width="900" cellpadding="2" cellspacing="1" class="a" bgcolor="#3d3d3d">
<form name="frm" method="post" action="3ban_process.asp" style="margin:0px;">
<input type="hidden" name="mode" value="<%=mode%>">
<input type="hidden" name="idx" value="<%=idx%>">
<input type="hidden" name="ban" value="<%=imgurl%>">
<input type="hidden" name="menupos" value="<%=menupos%>">
<tr bgcolor="#FFFFFF">
    <td bgcolor="#FFF999" align="center" width="15%">����</td>
    <td colspan="3">
    	<select name="gender" class="select">
    		<option value="M" <%= Chkiif(gender="M", "selected", "") %> >����</option>
    		<option value="F" <%= Chkiif(gender="F", "selected", "") %> >����</option>
    	</select>
    </td>
</tr>
<tr bgcolor="#FFFFFF">
    <td bgcolor="#FFF999" align="center" width="15%">����Ⱓ</td>
    <td colspan="3">
		<input type="text" id="sDt" name="startDate" size="10" value="<%=mainStartDate%>" readonly />
		<input type="text" name="sTm" size="8" value="00:00:00" disabled /> ~
		<input type="text" id="eDt" name="endDate" size="10" value="<%=mainEndDate%>" readonly />
		<input type="text" name="eTm" size="8" value="23:59:59" disabled />
    </td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="#FFF999" align="center" width="15%">�̹���</td>
	<td width="45%">
	<input type="button" name="btnBan" value="�̹��� ���" onClick="jsSetImg('<%=idx%>','<%= imgurl %>','ban','spanban')" class="button">
		<div id="spanban" style="padding: 5 5 5 5">
		<% If imgurl <> "" Then %>
			<img src="<%=imgurl%>" border="0">
			<a href="javascript:jsDelImg('ban','spanban');"><img src="/images/icon_delete2.gif" border="0"></a>
		<% End If %>
		</div>
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="#FFF999" align="center">�̹��� Link</td>
	<td colspan="3"><input type="text" name="imglink" size="80" value="<%=imglink%>"/>
	<br/><br/>ex)
		<font color="#707070">
		- <span style="cursor:pointer" onClick="putLinkText('search','3')">�˻���� ��ũ : /apps/appCom/between/project/?pjt_code=<font color="darkred">�ڵ�</font></span><br>
		</font>
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="#FFF999" align="center">��뿩��</td>
	<td colspan="3"><div style="float:left;"><input type="radio" name="isusing" value="Y" <%=chkiif(isusing = "Y","checked","")%> checked />����� &nbsp;&nbsp;&nbsp; <input type="radio" name="isusing" value="N"  <%=chkiif(isusing = "N","checked","")%>/>������</div> <div style="float:right;margin-top:5px;margin-right:10px;"></div></td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="#FFF999" align="center">���Ĺ�ȣ</td>
	<td colspan="3"><input type="text" name="sortno" value="<%=chkiif(sortno="","0",sortno)%>" size="2"/></td>
</tr>
<tr bgcolor="#FFFFFF" align="center">
    <td colspan="4"><input type="button" value=" �� �� " onClick="jsgolist();"/><input type="button" value=" �� �� " onClick="jsSubmit();"/></td>
</tr>
</form>
</table>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbCTclose.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->