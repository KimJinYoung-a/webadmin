<%@ language=vbscript %>
<% option explicit %>
<%
'###############################################
' Discription : ���� ��ȹ�� ���������(�����)
' History : 2018.04.16 ������
'###############################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/classes/sitemasterclass/wedding_ContentsManageCls.asp" -->
<!-- #include virtual="/admin/eventmanage/common/event_function.asp"-->
<%
Dim isusing, fixtype, validdate, prevDate
Dim idx, poscode, reload, gubun, edid
Dim culturecode
	idx = request("idx")
	poscode = request("poscode")
	reload = request("reload")
	gubun = request("gubun")

	isusing = request("isusing")
	fixtype = request("fixtype")
	validdate= request("validdate")
	prevDate = request("prevDate")

	culturecode = request("eC")

	if idx="" then idx=0

	Response.write culturecode

	if reload="on" then
	    response.write "<script>opener.location.reload(); window.close();</script>"
	    dbget.close()	:	response.End
	end if

	dim oPlanEvent
		set oPlanEvent = new CWeddingContents
		oPlanEvent.FRectIdx = idx
		oPlanEvent.GetOneKitContents

	If gubun = "" Then
		gubun = "index"
	End If

%>
<script language='javascript' src="/js/jsCal/js/jscal2.js"></script>
<script language='javascript' src="/js/jsCal/js/lang/ko.js"></script>
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/jscal2.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/border-radius.css" />
<script language='javascript'>

	function SaveMainContents(frm){
	    
		if (frm.copy1.value==""){
	        alert('��� ī�Ǹ� �Է� �ϼ���.');
	        frm.copy1.focus();
	        return;
	    }

		if (frm.copy2.value==""){
	        alert('���� ī�Ǹ� �Է� �ϼ���.');
	        frm.copy2.focus();
	        return;
	    }

		if (frm.copy3.value==""){
	        alert('�ϴ� ī�Ǹ� �Է� �ϼ���.');
	        frm.copy3.focus();
	        return;
	    }

		if (frm.copy4.value==""){
	        alert('PC �ѿ��� ī�Ǹ� �Է� �ϼ���.');
	        frm.copy4.focus();
	        return;
	    }
		
		if (frm.itemid.value==""){
	        alert('������ ��ȣ�� �Է� �ϼ���.');
	        frm.itemid.focus();
	        return;
	    }

		if (frm.upload_img1.value==""){
	        alert('�̹��� 1���� ���ε� ���ּ���.');
	        frm.upload_img1.focus();
	        return;
	    }

		if (frm.DispOrder.value==""){
	        alert('�� ������ �Է� �ϼ���.');
	        frm.DispOrder.focus();
	        return;
	    }

	    if (confirm('���� �Ͻðڽ��ϱ�?')){
	        frm.submit();
	    }
	}

	function ChangeLinktype(comp){
	    if (comp.value=="M"){
	       document.all.link_M.style.display = "";
	       document.all.link_L.style.display = "none";
	    }else{
	       document.all.link_M.style.display = "none";
	       document.all.link_L.style.display = "";
	    }
	}

	//function getOnLoad(){
	//    ChangeLinktype(frmcontents.linktype.value);
	//}

	//window.onload = getOnLoad;

	function ChangeGubun(comp){
	    location.href = "?gubun=<%=gubun%>&poscode=" + comp.value;
	    // nothing;
	}


	function ChangeGroupGubun(comp){
	    location.href = "?gubun=" + comp.value;
	    // nothing;
	}

	function jsSetImg(sImg, sName, sSpan){ 
		var winImg;
		var sFolder=document.frmcontents.Evt_Code.value;
		if (sFolder=="")
		{
			alert("�̺�Ʈ �˻� �� �̹����� ������ּ���.");
		}
		else
		{
		winImg = window.open('/admin/eventmanage/common/pop_event_uploadimgV2.asp?yr=<%=Year(now())%>&sF='+sFolder+'&sImg='+sImg+'&sName='+sName+'&sSpan='+sSpan,'popImg','width=370,height=150');
		winImg.focus();
		}
	}

	// ����ǰ �߰� �˾�
	function addnewItem(itemnum){
			var popwin; 
			popwin = window.open("pop_additemlist.asp?itemnum="+itemnum, "popup_item", "width=1024,height=768,scrollbars=yes,resizable=yes");
			popwin.focus();
	}

	function jsDelImg(sName, sSpan){
		if(confirm("�̹����� �����Ͻðڽ��ϱ�?\n\n���� �� �����ư�� ������ ó���Ϸ�˴ϴ�.")){
		   eval("document.all."+sName).value = "";
		   eval("document.all."+sSpan).style.display = "none";
		}
	}
</script>

<table width="100%" cellpadding="2" cellspacing="1" class="a" bgcolor="#3d3d3d">
<form name="frmcontents" method="post" action="doWeddingKitReg.asp" onsubmit="return false;">
<input type="hidden" name="Evt_Code" value="30000">
<input type="hidden" name="idx" value="<%= oPlanEvent.FOneItem.FIdx %>">
<tr bgcolor="#FFFFFF">
    <td width="150" bgcolor="#DDDDFF">���ī��</td>
    <td>
        <input type="text" name="copy1" size="50" value="<%=oPlanEvent.FOneItem.FCopy1%>">
    </td>
</tr>
<tr bgcolor="#FFFFFF">
    <td width="150" bgcolor="#DDDDFF">����ī��</td>
    <td>
        <input type="text" name="copy2" size="50" value="<%=oPlanEvent.FOneItem.FCopy2%>">
    </td>
</tr>
<tr bgcolor="#FFFFFF">
    <td width="150" bgcolor="#DDDDFF">�ϴ�ī��</td>
    <td>
        <input type="text" name="copy3" size="50" value="<%=oPlanEvent.FOneItem.FCopy3%>">
    </td>
</tr>
<tr bgcolor="#FFFFFF">
    <td width="100" bgcolor="#DDDDFF">PC�ѿ��� ī��</td>
    <td>
		<textarea name="copy4" rows="5" cols="50"><%=oPlanEvent.FOneItem.FCopy4%></textarea>
    </td>
</tr>
<tr bgcolor="#FFFFFF">
    <td width="100" bgcolor="#DDDDFF">������</td>
    <td>
		<input type="text" name="itemid" value="<%=oPlanEvent.FOneItem.FItemid%>"> <a href="javascript:addnewItem('');">�ҷ�����</a>&nbsp;<span id="iteminfo"></span>
    </td>
</tr>
<tr bgcolor="#FFFFFF">
  <td width="100" bgcolor="#DDDDFF">������ �̹���1 ���ε�</td>
  <td>
  	<input type="button" name="etcitem1" value="�̹��� ���" onClick="jsSetImg('<%=oPlanEvent.FOneItem.FUpload_img1%>','upload_img1','item1')" class="button"> (PC,����� ��� ���簢 �̹��� ������ ����� ����)
	<input type="hidden" name="upload_img1" value="<%=oPlanEvent.FOneItem.FUpload_img1%>">
		<div id="item1" style="padding: 5 5 5 5">
			<%IF oPlanEvent.FOneItem.FUpload_img1 <> "" THEN %>
			<img  src="<%=oPlanEvent.FOneItem.FUpload_img1%>" width="50%" border="0">
			<a href="javascript:jsDelImg('upload_img1','item1');"><img src="/images/icon_delete2.gif" border="0"></a>
			<%END IF%>
		</div>
  </td>
</tr>
<tr bgcolor="#FFFFFF">
  <td width="100" bgcolor="#DDDDFF">������ �̹���2 ���ε�</td>
  <td>
  	<input type="button" name="etcitem2" value="�̹��� ���" onClick="jsSetImg('<%=oPlanEvent.FOneItem.FUpload_img2%>','upload_img2','item2')" class="button"> (PC ���簢 �̹���)
	<input type="hidden" name="upload_img2" value="<%=oPlanEvent.FOneItem.FUpload_img2%>">
		<div id="item2" style="padding: 5 5 5 5">
			<%IF oPlanEvent.FOneItem.FUpload_img2 <> "" THEN %>
			<img  src="<%=oPlanEvent.FOneItem.FUpload_img2%>" width="50%" border="0">
			<a href="javascript:jsDelImg('upload_img2','item2');"><img src="/images/icon_delete2.gif" border="0"></a>
			<%END IF%>
		</div>
  </td>
</tr>
<tr bgcolor="#FFFFFF">
    <td width="100" bgcolor="#DDDDFF">�� ����</td>
    <td>
		<input type="text" name="DispOrder" value="<%=oPlanEvent.FOneItem.FDispOrder %>">
    </td>
</tr>
<tr bgcolor="#FFFFFF">
    <td colspan="2" align="center"><input type="button" value=" �� �� " onClick="SaveMainContents(frmcontents);"></td>
</tr>
</form>
</table>
<%
set oPlanEvent = Nothing
%>

<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
