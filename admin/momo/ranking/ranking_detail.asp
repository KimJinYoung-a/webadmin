<%@ language=vbscript %>
<% option explicit %>

<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/momo/momo_rankingCls.asp"-->

<%
	Dim cRankingMng, vIdx, vDIdx, vTitle, vSDate, vEDate, vIsusing, vRegdate, arrRankingDetailList, i
	Dim vOrderNum, vItemid, vItemName, vItemDetail, vItemImg1, vItemImg2, vDIsusing, vTitleImg
	vIdx = Request("idx")
	vDIdx = Request("didx")
	vOrderNum = 1
	
	If vIdx <> "" Then
		set cRankingMng = new ClsMomoRanking
		cRankingMng.FIdx = vIdx
		cRankingMng.FRankingMasterView
		
		vTitle 		= cRankingMng.FOneItem.ftitle
		vTitle 		= Replace(vTitle,chr(34),"&#34;")
		
		vTitleImg	= cRankingMng.FOneItem.ftitle_img
		vSDate 		= cRankingMng.FOneItem.fstartdate
		vEDate 		= cRankingMng.FOneItem.fenddate
		vIsusing	= cRankingMng.FOneItem.fisusing
		vRegdate 	= cRankingMng.FOneItem.fregdate
		
		
		arrRankingDetailList = cRankingMng.FRankingDetailViewList
		
		vOrderNum = UBound(arrRankingDetailList,2)+2
		
		If vDIdx <> "" Then
			cRankingMng.FDIdx = vDIdx
			cRankingMng.FRankingDetailView
			
			vOrderNum	= cRankingMng.FOneItem.fordernum
			vItemid		= cRankingMng.FOneItem.fitemid
			
			vItemName	= cRankingMng.FOneItem.fitemname
			vItemName 	= Replace(vItemName,chr(34),"&#34;")
			
			vItemDetail	= cRankingMng.FOneItem.fitemdetail
			vItemDetail = Replace(vItemDetail,chr(34),"&#34;")
			
			vItemImg1	= cRankingMng.FOneItem.fitemimg1
			vItemImg2	= cRankingMng.FOneItem.fitemimg2
			vDIsusing	= cRankingMng.FOneItem.fisusing
		End If
		set cRankingMng = nothing
	End If
%>
<script language='javascript' src="/js/jsCal/js/jscal2.js"></script>
<script language='javascript' src="/js/jsCal/js/lang/ko.js"></script>
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/jscal2.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/border-radius.css" />
<script language="javascript">
document.domain = "10x10.co.kr";

function itemimgpop(divnm,iptNm,vPath,Fsize,Fwidth,thumb)
{
	if(frm.itemid.value == "0" || divnm == "titleimg")
	{
		window.open('','imginput','width=350,height=300,menubar=no,toolbar=no,scrollbars=no,status=yes,resizable=yes,location=no');
		document.imginputfrm.divName.value=divnm;
		document.imginputfrm.inputname.value=iptNm;
		document.imginputfrm.ImagePath.value = vPath;
		document.imginputfrm.maxFileSize.value = Fsize;
		document.imginputfrm.maxFileWidth.value = Fwidth;
		document.imginputfrm.makeThumbYn.value = thumb;
		document.imginputfrm.orgImgName.value = eval("document.getElementById('"+iptNm+"')").value;
		document.imginputfrm.target='imginput';
		document.imginputfrm.action='PopImgInput.asp';
		document.imginputfrm.submit();
	}
	else
	{
		alert("��ǰ��ȣ�� 0�ΰ�츸 ���ε� �� �� �ֽ��ϴ�.");
	}
}

function checkform(frm)
{
	if(frm.title.value == "" || frm.title.value == " ")
	{
		alert("������ �Է��ϼ���.");
		frm.title.value = "";
		frm.title.focus();
		return false;
	}
	
	if(frm.title_img.value == "")
	{
		alert("�����̹����� �Է��ϼ���.");
		return false;
	}
	
	if(frm.sdate.value == "")
	{
		alert("�������� �Է��ϼ���.");
		return false;
	}
	
	if(frm.edate.value == "")
	{
		alert("�������� �Է��ϼ���.");
		return false;
	}
	
	if(!frm.isusing[0].checked && !frm.isusing[1].checked)
	{
		alert("������ ��뿩�θ� �����ϼ���.");
		return false;
	}
	
	if(frm.ordernum.value == "")
	{
		alert("���� ������ �Է��ϼ���.");
		frm.ordernum.focus();
		return false;
	}
	
	if(isNaN(frm.ordernum.value))
	{
		alert("���� ������ ���ڷθ� �Է��ϼ���.");
		frm.ordernum.value = "";
		frm.ordernum.focus();
		return false;
	}
	
	if(frm.itemid.value == "")
	{
		alert("��ǰ ��ȣ�� �Է��ϼ���.");
		frm.itemid.focus();
		return false;
	}
	
	if(isNaN(frm.itemid.value))
	{
		alert("��ǰ ��ȣ�� ���ڷθ� �Է��ϼ���.");
		frm.itemid.value = "";
		frm.itemid.focus();
		return false;
	}
	
	if(frm.itemname.value == "" || frm.itemname.value == " ")
	{
		alert("��ǰ���� �Է��ϼ���.");
		frm.itemname.value = "";
		frm.itemname.focus();
		return false;
	}
	
	if(frm.itemid.value == "0" && frm.itemimg1.value == "")
	{
		alert("��ǰ��ȣ�� 0�� ����\n300 x 300 �̹����� �Է��ؾ� �մϴ�.");
		return false;
	}
	
	if(frm.itemid.value == "0" && frm.itemimg2.value == "")
	{
		alert("��ǰ��ȣ�� 0�� ����\n50 x 50 �̹����� �Է��ؾ� �մϴ�.");
		return false;
	}
	
	if(!frm.disusing[0].checked && !frm.disusing[1].checked)
	{
		alert("��ǰ ��뿩�θ� �����ϼ���.");
		return false;
	}
	
	if(frm.itemdetail.value == "" || frm.itemdetail.value == " ")
	{
		alert("��ǰ�󼼸� �Է��ϼ���.");
		frm.itemdetail.value = "";
		frm.itemdetail.focus();
		return false;
	}
	
	return true;
}
</script>

<form name="frm" method="post" action="ranking_proc.asp" onSubmit="return checkform(this);">
<input type="hidden" name="idx" value="<%=vIdx%>">
<input type="hidden" name="didx" value="<%=vDIdx%>">
* <b>�����Է�</b>
<table width="100%" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<% If vIdx <> "" Then %>
<tr align="center" bgcolor="#FFFFFF" height="25">
	<td width="70" bgcolor="<%= adminColor("gray") %>">idx</td>
	<td align="left" width="300"><%=vIdx%></td>
</tr>
<% End If %>
<tr align="center" bgcolor="#FFFFFF" height="25">
	<td width="70" bgcolor="<%= adminColor("gray") %>">����</td>
	<td align="left" width="300"><input type="text" name="title" value="<%=vTitle%>" size="50" maxlength="70"></td>
</tr>
<tr align="center" bgcolor="#FFFFFF" height="25">
	<td width="70" bgcolor="<%= adminColor("gray") %>">�����̹���<br><input type="button" value="img up" onClick="itemimgpop('titleimg','title_img','title','2000','1000','false');"></td>
	<td align="left" width="300"><input type="hidden" name="title_img" value="<%=vTitleImg%>"><div align="center" id="titleimg"><% If vTitleImg <> "" Then %><img src="<%=vTitleImg%>" height="30"><% End If %></div></td>
</tr>
<tr align="center" bgcolor="#FFFFFF" height="25">
	<td width="70" bgcolor="<%= adminColor("gray") %>">�Ⱓ</td>
	<td align="left" width="400">
		<input id="sdate" name="sdate" value="<%=vSDate%>" class="text" size="10" maxlength="10" /><img src="http://webadmin.10x10.co.kr/images/calicon.gif" id="sdate_trigger" border="0" style="cursor:pointer" align="absmiddle" /> ~
		<input id="edate" name="edate" value="<%=vEDate%>" class="text" size="10" maxlength="10" /><img src="http://webadmin.10x10.co.kr/images/calicon.gif" id="edate_trigger" border="0" style="cursor:pointer" align="absmiddle" />
		<script language="javascript">
			var CAL_Start = new Calendar({
				inputField : "sdate", trigger    : "sdate_trigger",
				onSelect: function() {
					var date = Calendar.intToDate(this.selection.get());
					CAL_End.args.min = date;
					CAL_End.redraw();
					this.hide();
				}, bottomBar: true, dateFormat: "%Y-%m-%d"
			});
			var CAL_End = new Calendar({
				inputField : "edate", trigger    : "edate_trigger",
				onSelect: function() {
					var date = Calendar.intToDate(this.selection.get());
					CAL_Start.args.max = date;
					CAL_Start.redraw();
					this.hide();
				}, bottomBar: true, dateFormat: "%Y-%m-%d"
			});
		</script>
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" height="25">
	<td width="70" bgcolor="<%= adminColor("gray") %>">��뿩��</td>
	<td align="left" width="300">
		<input type="radio" name="isusing" value="Y" <% If vIsusing = "Y" Then Response.Write "checked" End If %>>Y&nbsp;&nbsp;&nbsp;
		<input type="radio" name="isusing" value="N" <% If vIsusing = "N" Then Response.Write "checked" End If %>>N
	</td>
</tr>
<% If vIdx <> "" Then %>
<tr align="center" bgcolor="#FFFFFF" height="25">
	<td width="70" bgcolor="<%= adminColor("gray") %>">�����</td>
	<td align="left">
		<table width="100%" cellpadding="0" cellspacing="0" border="0" class="a">
		<tr>
			<td><%=vRegdate%></td>
			<td align="right"><input type="button" value="������ ����" onClick="frm.didx.value='';frm.submit();"></td>
		</tr>
		</table>
	</td>
</tr>
<% End If %>
</table>

<br>
* <b>��ǰ�Է�</b><br>
�� ���ļ�:1�� ���� ����, ������ ���ʴ��.&nbsp;&nbsp;&nbsp;��ǰ��ȣ:�ٹ����� ��ǰ�� �ƴҰ��� 0
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<input type="button" value="�ٹ����ٻ�ǰ���" onClick="window.open('pop_additemlist.asp','findProd','width=900,height=600,scrollbars=yes');">
<table width="100%" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="#FFFFFF">
	<td width="50" align="center" bgcolor="<%= adminColor("gray") %>">���ļ�</td>
	<td width="70" align="center" bgcolor="<%= adminColor("gray") %>">��ǰ��ȣ</td>
	<td width="230" align="center" bgcolor="<%= adminColor("gray") %>">��ǰ��</td>
	<td width="90" align="center" bgcolor="<%= adminColor("gray") %>"><input type="button" value="300 img" onClick="itemimgpop('img1','itemimg1','300','2000','300','false');"></td>
	<td width="90" align="center" bgcolor="<%= adminColor("gray") %>"><input type="button" value="50 img" onClick="itemimgpop('img2','itemimg2','50','2000','50','false');"></td>
	<td width="100" align="center" bgcolor="<%= adminColor("gray") %>">��뿩��</td>
</tr>
<tr height="50" bgcolor="#FFFFFF">
	<td align="center"><input type="text" size="5" name="ordernum" value="<%=vOrderNum%>"></td>
	<td align="center"><input type="text" size="8" name="itemid" value="<%=vItemid%>"></td>
	<td align="center"><input type="text" size="30" maxlength="15" name="itemname" value="<%=vItemName%>"></td>
	<td align="center"><input type="hidden" name="itemimg1" value="<%=vItemImg1%>"><div align="center" id="img1"><% If vItemImg1 <> "" Then %><img src="<%=vItemImg1%>" width="50"><% End If %></div></td>
	<td align="center"><input type="hidden" name="itemimg2" value="<%=vItemImg2%>"><div align="center" id="img2"><% If vItemImg1 <> "" Then %><img src="<%=vItemImg2%>" width="50"><% End If %></div></td>
	<td align="center">
		<input type="radio" name="disusing" value="Y" <% If vDIsusing = "Y" Then Response.Write "checked" End If %>>Y&nbsp;&nbsp;&nbsp;
		<input type="radio" name="disusing" value="N" <% If vDIsusing = "N" Then Response.Write "checked" End If %>>N
	</td>
</tr>
<tr height="50" bgcolor="#FFFFFF">
	<td align="center" colspan="6"><textarea name="itemdetail" cols="90" rows="3"><%=vItemDetail%></textarea></td>
</tr>
</table>

<table width="100%" cellpadding="0" cellspacing="0" class="a">
<tr height="25">
	<td align="left"><input type="button" value="���۾���" onClick="location.href='/admin/momo/ranking/ranking_detail.asp?idx=<%=vIdx%>';"></td>
	<td align="right"><input type="submit" value="����,��ǰ �������"></td>
</tr>
</table>
</form>

<%
	IF isArray(arrRankingDetailList) THEN
%>
	<table width="100%" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr height="25" bgcolor="#FFFFFF">
		<td width="50" align="center" bgcolor="<%= adminColor("gray") %>">���ļ�</td>
		<td width="70" align="center" bgcolor="<%= adminColor("gray") %>">��ǰ��ȣ</td>
		<td align="center" bgcolor="<%= adminColor("gray") %>">��ǰ��</td>
		<td width="90" align="center" bgcolor="<%= adminColor("gray") %>">300 img</td>
		<td width="90" align="center" bgcolor="<%= adminColor("gray") %>">50 img</td>
	</tr>
<%
		For i = 0 To UBound(arrRankingDetailList,2)
%>
			<tr bgcolor="#FFFFFF">
				<td align="center"><%=arrRankingDetailList(0,i)%></td>
				<td align="center"><%=arrRankingDetailList(1,i)%></td>
				<td><%=Replace(arrRankingDetailList(2,i),"chr(34)",chr(34))%></td>
				<td align="center"><img src="<%=arrRankingDetailList(3,i)%>" width="50" height="50"></td>
				<td align="center"><img src="<%=arrRankingDetailList(4,i)%>" width="50" height="50"></td>
			</tr>
			<tr bgcolor="#FFFFFF">
				<td align="center" colspan="2">
					<% If arrRankingDetailList(7,i) = "Y" Then Response.Write "���" Else Response.Write "����" End If %>
					<br>(����ǥ:<%=arrRankingDetailList(5,i)%>, UP:<%=arrRankingDetailList(6,i)%>)
				</td>
				<td colspan="3">
					<table width="100%" cellpadding="0" cellspacing="0" border="0" class="a">
					<tr>
						<td><%=Replace(arrRankingDetailList(9,i),vbCrLf,"<br>")%></td>
						<td align="right"><input type="button" value="����" onClick="location.href='/admin/momo/ranking/ranking_detail.asp?idx=<%=vIdx%>&didx=<%=arrRankingDetailList(8,i)%>';"></td>
					</tr>
					</table>
				</td>
			</tr>
<%
		Next
		Response.Write "</table>"
	End If
%>

<form name="imginputfrm" method="post" action="">
<input type="hidden" name="YearUse" value="<%=year(now)%>">
<input type="hidden" name="divName" value="">
<input type="hidden" name="orgImgName" value="">
<input type="hidden" name="inputname" value="">
<input type="hidden" name="ImagePath" value="">
<input type="hidden" name="maxFileSize" value="">
<input type="hidden" name="maxFileWidth" value="">
<input type="hidden" name="makeThumbYn" value="">
</form>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
