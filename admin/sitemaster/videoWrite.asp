<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/sitemasterclass/videoInfoCls.asp"-->
<%
'###############################################
' PageName : videoWrite.asp
' Discription : ������ ���� ���/����
' History : 2009.09.29 ������ ����
'###############################################

dim videoSn, mode, i
mode=request("mode")
videoSn=request("videoSn")

dim fmainitem
set fmainitem = New Cvideo
fmainitem.FCurrPage = 1
fmainitem.FPageSize=1
fmainitem.FRectVSN=videoSn
fmainitem.GetVideoList
%>
<link href="/js/jqueryui/css/jquery-ui.css" rel="stylesheet">
<script type="text/javascript" src="/js/jquery-1.7.2.min.js"></script>
<script type="text/javascript" src="/js/jqueryui/jquery-ui-1.10.2.custom.min.js"></script>
<script language="javascript">
<!--
$(document).ready(function(){
    $('#videoDiv').change(function(){
        if($('#videoDiv').val() == "mov"){
			$("#mlink").show();
			$("#mdate").show();
			$("#msize").hide();
		}
		else{
			$("#mlink").hide();
			$("#mdate").hide();
			$("#msize").show();
		}
    });

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
      	<% if videoSn<>"" then %>maxDate: "<%=fmainitem.FItemList(0).FendDate%>",<% end if %>
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
      	<% if videoSn<>"" then %>maxDate: "<%=fmainitem.FItemList(0).FstartDate%>",<% end if %>
    	onClose: function( selectedDate ) {
    		$( "#sDt" ).datepicker( "option", "maxDate", selectedDate );
    	}
    });
});

function editcont(){
    //���µ��� ���� ������ ���;;
    var frm=document.inputfrm;
    
    if (confirm('���� �Ͻðڽ��ϱ�?')){
        frm.sale_code.value="";
        frm.submit();
    }
    
}

function subcheck(){
	var frm=document.inputfrm;

	if(!frm.videoDiv.value) {
		alert("������ ������ �������ּ���!");
		frm.videoDiv.focus();
		return;
	}

	if(!frm.videoTitle.value) {
		alert("������ ������ �Է����ּ���!");
		frm.videoTitle.focus();
		return;
	}

	if(!frm.videoFile.value&&!frm.videoSn.value) {
		alert("FLV������ ������ ������ �������ּ���!");
		frm.videoFile.focus();
		return;
	}
	if($('#videoDiv').val()!="mov"){
		if(!frm.videoWidth.value||frm.videoWidth.value=='0') {
			alert("������ �ʺ� �Է����ּ���!");
			frm.videoWidth.focus();
			return;
		}

		if(!frm.videoHeight.value||frm.videoHeight.value=='0') {
			alert("������ ���̸� �Է����ּ���!");
			frm.videoHeight.focus();
			return;
		}
	}
	else{
		if(frm.linkgubun.value=="") {
			alert("��ũ ������ �������ּ���!");
			frm.linkgubun.focus();
			return;
		}
		if(frm.linkinfo.value=="") {
			alert("��ũ ������ȣ�� �Է����ּ���!");
			frm.linkinfo.focus();
			return;
		}
		if(frm.startDate.value=="") {
			alert("�������� �������ּ���!");
			frm.startDate.focus();
			return;
		}
		if(frm.endDate.value=="") {
			alert("�������� �������ּ���!");
			frm.endDate.focus();
			return;
		}
	}

	frm.submit();
}

function delitems(){
	var frm = document.inputfrm;
	if (confirm('�� �������� �����Ͻðڽ��ϱ�?')) {
		frm.mode.value="delete";
		frm.submit();
	}
}
//-->
</script>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="inputfrm" method="post" action="<%= uploadImgUrl %>/linkweb/sitemaster/doVideoProcess.asp" enctype="MULTIPART/FORM-DATA">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="mode" value="<% =mode %>">
<tr height="30">
	<td colspan="2" bgcolor="#FFFFFF">
		<img src="/images/icon_star.gif" align="absmiddle">
		<font color="red"><b>������ ���/����</b></font>
	</td>
</tr>
<% if mode="add" then %>
<input type="hidden" name="videoSn" value="">
<tr>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">������ ����</td>
	<td bgcolor="#FFFFFF"><%=drawVDivSelect("videoDiv","")%></td>
</tr>
<tr>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">����</td>
	<td bgcolor="#FFFFFF">
		<input type="text" class="text" name="videoTitle" value="" size="60">
	</td>
</tr>
<tr>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">������</td>
	<td bgcolor="#FFFFFF">
		<input type="file" class="text" name="videoFile" value="" size="40"> (�� MP4/FLV/MP3����, �ִ� 20MB ����)
	</td>
</tr>
<tr>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">�̸����� �̹���</td>
	<td bgcolor="#FFFFFF">
		<input type="file" class="text" name="videoThumb" value="" size="40"> (�� JPG,GIF �̹���, �ִ� 300KB ����)
	</td>
</tr>
<tr id="msize">
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">������ ũ��</td>
	<td bgcolor="#FFFFFF">
		���� <input type="text" class="text" name="videoWidth" value="0" size="3" style="text-align:right">px �� ���� <input type="text" class="text" name="videoHeight" value="0" size="3" style="text-align:right">px 
		<br>�� 0 �Է½� �ʺ� ����
		<br>�� ���̾ : 450 �� 280 / ��Ÿ ���� : �ʿ信���� ����
	</td>
</tr>
<tr id="mdate" style="display:none">
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">����Ⱓ</td>
	<td bgcolor="#FFFFFF">
		<input type="text" id="sDt" name="startDate" size="10" />
		~
		<input type="text" id="eDt" name="endDate" size="10" />
	</td>
</tr>
<tr id="mlink" style="display:none">
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">��ũ ����</td>
	<td bgcolor="#FFFFFF">
		����
		<select name="linkgubun" class="select">
			<option value="">��ũ ����</option>
			<option value="1">��ǰ</option>
			<option value="2">�̺�Ʈ</option>
			<option value="3">��ġ����Ŀ</option>
			<option value="4">�귣��</option>
			<option value="5">�ٲ�TV</option>
		</select>
		&nbsp;&nbsp;��ũ ������ȣ <input type="text" class="text" name="linkinfo" size="15">
		<br><br>������ȣ ����)
		<br> ��ǰ : http://www.10x10.co.kr/shopping/category_prd.asp?itemid=<font color="red">2687010</font>
		<br>�̺�Ʈ : http://www.10x10.co.kr/event/eventmain.asp?eventid=<font color="red">102198</font>
		<br>��ġ����Ŀ : ������ȣ ���ʿ�
		<br>�귣�� : http://www.10x10.co.kr/street/street_brand.asp?makerid=<font color="red">tenten10000</font>
		<br>�ٲ�TV : http://www.10x10.co.kr/diarystory2020/daccutv_detail.asp?cidx=<font color="red">39</font>
	</td>
</tr>
<% elseif mode="edit" then %>
<tr>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">��ȣ</td>
	<td bgcolor="#FFFFFF">
		<b><%=fmainitem.FItemList(0).FvideoSn%></b>
		<input type="hidden" name="videoSn" value="<%=fmainitem.FItemList(0).FvideoSn%>">
	</td>
</tr>
<tr>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">������ ����</td>
	<td bgcolor="#FFFFFF"><%=drawVDivSelect("videoDiv",fmainitem.FItemList(0).FvideoDiv)%></td>
</tr>
<tr>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">����</td>
	<td bgcolor="#FFFFFF">
		<input type="text" class="text" name="videoTitle" value="<%=fmainitem.FItemList(0).FvideoTitle%>" size="60">
	</td>
</tr>
<tr>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">������</td>
	<td bgcolor="#FFFFFF">
		<input type="file" class="text" name="videoFile" value="" size="40"> (�� FLV����)
		<%
			if Not(fmainitem.FItemList(0).FvideoFile="" or isNull(fmainitem.FItemList(0).FvideoFile)) then
				Response.Write "<br>(����:" & fmainitem.FItemList(0).FvideoFile & ")"
			end if
		%>
	</td>
</tr>
<tr>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">�̸����� �̹���</td>
	<td bgcolor="#FFFFFF">
		<input type="file" class="text" name="videoThumb" value="" size="40"> (�� JPG,GIF �̹���)
		<%
			if Not(fmainitem.FItemList(0).FvideoThumb="" or isNull(fmainitem.FItemList(0).FvideoThumb)) then
				Response.Write "<br>(����:" & fmainitem.FItemList(0).FvideoThumb & ")"
			end if
		%>
	</td>
</tr>
<tr id="mwidth"<% if fmainitem.FItemList(0).FvideoDiv="mov" then %> style="display:none"<% end if %>>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">������ ũ��</td>
	<td bgcolor="#FFFFFF">
		���� <input type="text" class="text" name="videoWidth" value="<%=fmainitem.FItemList(0).FvideoWidth%>" size="3" style="text-align:right">px ��
		���� <input type="text" class="text" name="videoHeight" value="<%=fmainitem.FItemList(0).FvideoHeight%>" size="3" style="text-align:right">px 
		<br>�� 0 �Է½� �ʺ� ����
	</td>
</tr>
<tr id="mdate"<% if fmainitem.FItemList(0).FvideoDiv<>"mov" then %> style="display:none"<% end if %>>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">����Ⱓ</td>
	<td bgcolor="#FFFFFF">
		<input type="text" id="sDt" name="startDate" size="10" value="<%=fmainitem.FItemList(0).FstartDate%>" />
		~
		<input type="text" id="eDt" name="endDate" size="10" value="<%=fmainitem.FItemList(0).FendDate%>" />
	</td>
</tr>
<tr id="mlink"<% if fmainitem.FItemList(0).FvideoDiv<>"mov" then %> style="display:none"<% end if %>>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">��ũ ����</td>
	<td bgcolor="#FFFFFF">
		����
		<select name="linkgubun" class="select">
			<option value="">��ũ ����</option>
			<option value="1"<% if fmainitem.FItemList(0).Flinkgubun="1" then response.write " selected" %>>��ǰ</option>
			<option value="2"<% if fmainitem.FItemList(0).Flinkgubun="2" then response.write " selected" %>>�̺�Ʈ</option>
			<option value="3"<% if fmainitem.FItemList(0).Flinkgubun="3" then response.write " selected" %>>��ġ����Ŀ</option>
			<option value="4"<% if fmainitem.FItemList(0).Flinkgubun="4" then response.write " selected" %>>�귣��</option>
			<option value="5"<% if fmainitem.FItemList(0).Flinkgubun="5" then response.write " selected" %>>�ٲ�TV</option>
		</select>
		&nbsp;&nbsp;��ũ ������ȣ <input type="text" class="text" name="linkinfo" value="<%=fmainitem.FItemList(0).Flinkinfo %>" size="15">
		<br><br>������ȣ ����)
		<br> ��ǰ : http://www.10x10.co.kr/shopping/category_prd.asp?itemid=<font color="red">2687010</font>
		<br>�̺�Ʈ : http://www.10x10.co.kr/event/eventmain.asp?eventid=<font color="red">102198</font>
		<br>��ġ����Ŀ : ������ȣ ���ʿ�
		<br>�귣�� : http://www.10x10.co.kr/street/street_brand.asp?makerid=<font color="red">tenten10000</font>
		<br>�ٲ�TV : http://www.10x10.co.kr/diarystory2020/daccutv_detail.asp?cidx=<font color="red">39</font>
	</td>
</tr>
<% end if %>
<tr bgcolor="#FFFFFF" >
	<td colspan="2" align="center">
		<input type="button" value=" ���� " class="button" onclick="subcheck();"> &nbsp;&nbsp;
		<% if mode="edit" then %><input type="button" value=" ���� " class="button" onclick="delitems();"> &nbsp;&nbsp;<% end if %>
		<input type="button" value=" ��� " class="button" onclick="history.back();">
	</td>
</tr>
</form>
</table>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
