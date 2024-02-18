<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/classes/sitemasterclass/look2018.asp" -->
<%
Dim idx, subidx, srcSDT, srcEDT, prevDate, paramisusing, mode, orderby, itemid, displaytype, regdate, itemimage, isusing
Dim orderbylist, isusinglist, copyimageurl, bgcolor, mainStartDate, mainEndDate, displaysale

	idx = requestCheckvar(request("idx"),16)
	subidx = requestCheckvar(request("subidx"),16)
	srcSDT = request("sDt")
	srcEDT = request("eDt")
	prevDate = request("prevDate")
	paramisusing = request("paramisusing")
	menupos = request("menupos")


If subidx = "" Then 
	mode = "subadd" 
Else 
	mode = "submodify" 
End If

If subidx <> "" then
	dim lookItem
	set lookItem = new Clook
	lookItem.FRectSubIdx = subidx
	lookItem.GetOneSubItem()
	itemid			= lookItem.FOneItem.FItemid
	orderby			= lookItem.FOneItem.Forderby
	isusing			= lookItem.FOneItem.Fisusing
	itemimage			= lookItem.FOneItem.Fitemimage
	displaytype			= lookItem.FOneItem.Fdisplaytype
	regdate			= lookItem.FOneItem.Fregdate
	displaysale			= lookItem.FOneItem.Fdisplaysale
	set lookItem = Nothing
Else
	Dim lookItemMaxOrderNum
	Set lookItemMaxOrderNum = New Clook
	lookItemMaxOrderNum.FRectIdx = idx
	lookItemMaxOrderNum.GetMaxSubItemOrderByNum()
	orderby = lookItemMaxOrderNum.FOneItem.Forderby
End If 

If idx <> "" Then
	dim lookList
	set lookList = new Clook
	lookList.FRectIdx = idx
	lookList.GetOneContents()

	copyimageurl	= lookList.FOneItem.Fcopyimageurl
	bgcolor		= lookList.FOneItem.Fbgcolor
	orderbylist		= lookList.FOneItem.Forderby
	mainStartDate	=	lookList.FOneItem.Fstartdate '// ������
	mainEndDate		=	lookList.FOneItem.Fenddate '// ������
	isusinglist		=	lookList.FOneItem.Fisusing '// ��뿩��

	set lookList = Nothing
End If

If Trim(isusing)="" Then
	isusing = "Y"
End If

If Trim(displaytype)="" Then
	displaytype = "D"
End If

If Trim(displaysale)="" Then
	displaysale = "Y"
End If

If isnull(displaysale) Then
	displaysale = "N"
End If

%>
<link href="/js/jqueryui/css/jquery-ui.css" rel="stylesheet">
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript" src="/js/jqueryui/jquery-ui-1.10.2.custom.min.js"></script>
<script type='text/javascript'>
function jsSubmit(){
	var frm = document.frm;

	if (!frm.displaytype[0].checked && !frm.displaytype[1].checked)
	{
		alert("Ÿ���� �����ϼ���!")
		return false;
	}

	if (!frm.itemimage.value){
		alert("��ǰ�̹����� ������ּ���.");
		return;
	}
	if (!frm.itemid.value){
		alert("��ǰ�ڵ带 �Է����ּ���.");
		frm.itemid.focus();
		return;
	}
	if (!frm.orderby.value){
		alert("���Ĺ�ȣ�� �Է����ּ���.");
		frm.orderby.focus();
		return;
	}
	if (!frm.isusing[0].checked && !frm.isusing[1].checked)
	{
		alert("��뿩�θ� �����ϼ���!")
		return false;
	}

	if (confirm('���� �Ͻðڽ��ϱ�?')){
		//frm.target = "blank";
		frm.submit();
	}
}


//-- jsImgView : �̹��� Ȯ��ȭ�� ��â���� �����ֱ� --//
function jsImgView(sImgUrl){
 var wImgView;
 wImgView = window.open('/admin/eventmanage/common/pop_event_detailImg.asp?sUrl='+sImgUrl,'pImg','width=100,height=100');
 wImgView.focus();
}


function jsSetImg(sFolder, sImg, sName, sSpan){ 
	var winImg;
	winImg = window.open('/admin/eventmanage/common/pop_event_uploadimgV2.asp?yr=<%=Year(now())%>&sF='+sFolder+'&sImg='+sImg+'&sName='+sName+'&sSpan='+sSpan,'popImg','width=370,height=150');
	winImg.focus();
}

function jsDelImg(sName, sSpan){
	if(confirm("�̹����� �����Ͻðڽ��ϱ�?\n\n���� �� �����ư�� ������ ó���Ϸ�˴ϴ�.")){
	   eval("document.all."+sName).value = "";
	   eval("document.all."+sSpan).style.display = "none";
	}
}
</script>
<form name="frm" method="post" action="dolook.asp" style="margin:0px;">
<input type="hidden" name="mode" value="<%=mode%>">
<input type="hidden" name="adminid" value="<%=session("ssBctId")%>">
<input type="hidden" name="idx" value="<%=idx%>">
<input type="hidden" name="prevDate" value="<%=prevDate%>">
<input type="hidden" name="paramisusing" value="<%=paramisusing%>">
<input type="hidden" name="itemimage" value="<%=itemimage%>">
<input type="hidden" name="menupos" value="<%=menupos%>">
<input type="hidden" name="subidx" value="<%=subidx%>">
<table width="750" cellpadding="2" cellspacing="1" class="a" bgcolor="#3d3d3d">
<tr bgcolor="#FFFFFF">
	<% If subidx = ""  Then %>
	<td colspan="2" align="center" height="35">��� ���� �� �Դϴ�.</td>
	<% Else %>
	<td bgcolor="#FFF999" colspan="2" align="center" height="35">���� ���� �� �Դϴ�.</td>
	<% End If %>
</tr>
<% If idx <> ""  Then %>
<tr bgcolor="#FFFFFF">
	<td bgcolor="#FFF999" align="center" width="15%">idx</td>
	<td><%=idx%></td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="#FFF999" align="center" width="15%">����ī���̹���</td>
	<td><img src="<%=copyimageurl%>" ></td>
</tr>
<% End If %>
<tr bgcolor="#FFFFFF">
	<td bgcolor="#FFF999" align="center">Ÿ��</td>
	<td><div style="float:left;"><input type="radio" name="displaytype" value="U" <%=chkiif(Trim(displaytype) = "U","checked","")%> />�� &nbsp;&nbsp;&nbsp; <input type="radio" name="displaytype" value="D"  <%=chkiif(Trim(displaytype) = "D","checked","")%>/>�Ʒ�</div> <div style="float:right;margin-top:5px;margin-right:10px;"></div></td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="#FFF999" align="center">����ǥ�ÿ���</td>
	<td><div style="float:left;"><input type="radio" name="displaysale" value="Y" <%=chkiif(Trim(displaysale) = "Y","checked","")%> />ǥ�� &nbsp;&nbsp;&nbsp; <input type="radio" name="displaysale" value="N"  <%=chkiif(Trim(displaysale) = "N" Or Trim(displaysale) = "","checked","")%>/>ǥ�þ���</div> <div style="float:right;margin-top:5px;margin-right:10px;"></div></td>
</tr>
<tr bgcolor="#FFFFFF" id="tmpstyle1">
	<td bgcolor="#DDDDFF" align="center" width="15%">��ǰ�̹���</td>
	<td><input type="button" name="limg" value="��ǰ �̹��� ���" onClick="jsSetImg('pcmainlookitem','<%=itemimage%>','itemimage','lookitemimg')" class="button">
		<div id="lookitemimg" style="padding: 5 5 5 5">
			<%IF itemimage <> "" THEN %>
			<a href="javascript:jsImgView('<%=itemimage%>')"><img  src="<%=itemimage%>" width="400" border="0"></a>
			<a href="javascript:jsDelImg('itemimage','lookitemimg');"><img src="/images/icon_delete2.gif" border="0"></a>
			<%END IF%>
		<%=itemimage%>
		</div>
	</td>
</tr>
<tr bgcolor="#FFFFFF" id="tmpstyle2">
	<td bgcolor="#DDDDFF"  align="center" width="15%">��ǰ�ڵ�</td>
	<td>
		<input type="text" name="itemid" value="<%=itemid%>" style="width:20%;" />
	</td>
</tr>
<tr bgcolor="#FFFFFF" id="tmpstyle2">
	<td bgcolor="#DDDDFF"  align="center" width="15%">���Ĺ�ȣ</td>
	<td>
		<input type="text" name="orderby" value="<%=orderby%>" style="width:10%;" /> <font color="red">���Ĺ�ȣ�� �������� �Դϴ�.</font>
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="#FFF999" align="center">��뿩��</td>
	<td><div style="float:left;"><input type="radio" name="isusing" value="Y" <%=chkiif(Trim(isusing) = "Y","checked","")%> />����� &nbsp;&nbsp;&nbsp; <input type="radio" name="isusing" value="N"  <%=chkiif(Trim(isusing) = "N","checked","")%>/>������</div> <div style="float:right;margin-top:5px;margin-right:10px;"></div></td>
</tr>

<tr bgcolor="#FFFFFF" align="center">
    <td colspan="2"><input type="button" value=" �� �� " onClick="window.close();"/><input type="button" value=" �� �� " onClick="jsSubmit();"/></td>
</tr>
</table>
</form>

<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->