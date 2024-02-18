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
		oPlanEvent.GetOneShoppingListMoContents

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
	    if (frm.itemid.value==""){
	        alert('������ ��ȣ�� �Է� �ϼ���.');
	        frm.itemid.focus();
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
<form name="frmcontents" method="post" action="doWeddingShoppingListMoReg.asp" onsubmit="return false;">
<input type="hidden" name="Evt_Code" value="10000">
<input type="hidden" name="WeddingStepID" value="<%= oPlanEvent.FOneItem.FWeddingStepID %>">
<tr bgcolor="#FFFFFF">
    <td width="150" bgcolor="#DDDDFF">Idx</td>
    <td>
        [<%= oPlanEvent.FOneItem.FWeddingStepID %>]<%= oPlanEvent.FOneItem.GetDDayTitleMo %>
    </td>
</tr>
<tr bgcolor="#FFFFFF">
    <td width="100" bgcolor="#DDDDFF">������</td>
    <td>
		<input type="text" name="itemid" value="<%=oPlanEvent.FOneItem.FItemid%>"> <a href="javascript:addnewItem('');">�ҷ�����</a> <span id="iteminfo"></span>
    </td>
</tr>
<tr bgcolor="#FFFFFF">
  <td width="100" bgcolor="#DDDDFF">������ �̹��� ���ε�</td>
  <td>
  	<input type="button" name="etcitem" value="�̹��� ���" onClick="jsSetImg('<%=oPlanEvent.FOneItem.FUpload_img%>','upload_img','item1')" class="button">
	<input type="hidden" name="upload_img" value="<%=oPlanEvent.FOneItem.FUpload_img%>">
		<div id="item1" style="padding: 5 5 5 5">
			<%IF oPlanEvent.FOneItem.FUpload_img <> "" THEN %>
			<img  src="<%=oPlanEvent.FOneItem.FUpload_img%>" width="50%" border="0">
			<a href="javascript:jsDelImg('upload_img','item1');"><img src="/images/icon_delete2.gif" border="0"></a>
			<%END IF%>
		</div>
  </td>
</tr>
<tr bgcolor="#FFFFFF">
    <td width="100" bgcolor="#DDDDFF">����</td>
    <td>
		<textarea name="Contents" rows="5" cols="50"><%=oPlanEvent.FOneItem.FContents%></textarea>
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
