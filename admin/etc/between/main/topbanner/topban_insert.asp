<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbCTopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/admin/etc/between/mainCls.asp"-->
<%
Dim idx, imgurl, mode, isusing, linkURL, BanBGColor, partnerNmColor, BanTxtColor, bantext1, bantext2
Dim gender, pjt_kind
Dim sDt, eDt

idx = requestCheckvar(request("idx"),16)

If idx = "" Then
	mode = "I"
Else
	mode = "U"
End If

If idx <> "" then
	Dim otopban
	SET otopban = new cMain
		otopban.FRectIdx = idx
		otopban.getOneTopBanner()
		gender			= otopban.FItemList(0).FGender
		pjt_kind		= otopban.FItemList(0).FPjt_kind
		imgurl			= otopban.FItemList(0).FImgurl
		linkURL			= otopban.FItemList(0).FLinkURL
		BanBGColor		= otopban.FItemList(0).FBanBGColor
		partnerNmColor	= otopban.FItemList(0).FPartnerNmColor
		BanTxtColor		= otopban.FItemList(0).FBanTxtColor
		bantext1		= otopban.FItemList(0).FBantext1
		bantext2		= otopban.FItemList(0).FBantext2
		isusing			= otopban.FItemList(0).FIsusing
	SET otopban = Nothing
End If
%>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type='text/javascript'>
function jsSubmit(){
	var frm = document.frm;
	if (confirm('���� �Ͻðڽ��ϱ�?')){
		frm.submit();
	}
}

function jsgolist(){
	self.location.href="/admin/etc/between/main/topbanner/index.asp";
}


function jsSetImg(sFolder, sImg, sName, sSpan){
	document.domain ="10x10.co.kr";
	var winImg;
	winImg = window.open('/admin/etc/between/main/topbanner/pop_topbanner_uploadimg.asp?sF='+sFolder+'&sImg='+sImg+'&sName='+sName+'&sSpan='+sSpan,'popImg','width=370,height=150');
	winImg.focus();
}
function jsDelImg(sName, sSpan){
	if(confirm("�̹����� �����Ͻðڽ��ϱ�?\n\n���� �� �����ư�� ������ ó���Ϸ�˴ϴ�.")){
	   eval("document.all."+sName).value = "";
	   eval("document.all."+sSpan).style.display = "none";
	}
}
function putLinkText(key,gubun) {
	var frm = document.frm;
	var urllink
	if (gubun == "3" ){
		urllink = frm.linkURL;
	}
	switch(key) {
		case 'search':
			urllink.value='/apps/appCom/between/project/?project_idx=�ڵ�';
			break;
	}
}
</script>
<table width="900" cellpadding="2" cellspacing="1" class="a" bgcolor="#3d3d3d">
<form name="frm" method="post" action="topban_process.asp" style="margin:0px;">
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
    <td bgcolor="#FFF999" align="center" width="15%">��ȹ�� ����</td>
    <td colspan="3">
    	<% sbGetOptProjectCodeValue "pjt_kind",pjt_kind,"" %>
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
	<td bgcolor="#FFF999" align="center">Link URL</td>
	<td colspan="3">
		<input type="text" name="linkURL" size="80" value="<%=linkURL%>"/><br>
		<font color="#707070">
		- <span style="cursor:pointer" onClick="putLinkText('search','3')">�˻���� ��ũ : /apps/appCom/between/project/?project_idx=<font color="darkred">�ڵ�</font></span><br>
		</font>
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="#FFF999" align="center">��� ����</td>
	<td colspan="3"><input type="text" name="BanBGColor" maxlength="7" size="7" value="<%=BanBGColor%>"/></td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="#FFF999" align="center">Text 1���� ����</td>
	<td colspan="3"><input type="text" name="partnerNmColor" maxlength="7" size="7" value="<%=partnerNmColor%>"/></td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="#FFF999" align="center">Text 2���� ����</td>
	<td colspan="3"><input type="text" name="BanTxtColor" maxlength="7" size="7" value="<%=BanTxtColor%>"/></td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="#FFF999" align="center">�ؽ�Ʈ 1</td>
	<td colspan="3"><input type="text" name="bantext1" size="80" value="<%=bantext1%>"/></td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="#FFF999" align="center">�ؽ�Ʈ 2</td>
	<td colspan="3"><input type="text" name="bantext2" size="80" value="<%=bantext2%>"/></td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="#FFF999" align="center">��뿩��</td>
	<td colspan="3"><div style="float:left;"><input type="radio" name="isusing" value="Y" <%=chkiif(isusing = "Y","checked","")%> checked />����� &nbsp;&nbsp;&nbsp; <input type="radio" name="isusing" value="N"  <%=chkiif(isusing = "N","checked","")%>/>������</div> <div style="float:right;margin-top:5px;margin-right:10px;"></div></td>
</tr>
<tr bgcolor="#FFFFFF" align="center">
    <td colspan="4"><input type="button" value=" �� �� " onClick="jsgolist();"/><input type="button" value=" �� �� " onClick="jsSubmit();"/></td>
</tr>
</form>
</table>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbCTclose.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->