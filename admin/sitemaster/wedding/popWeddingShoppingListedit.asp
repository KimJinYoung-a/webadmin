<%@ language=vbscript %>
<% option explicit %>
<%
'###############################################
' Discription : ���� ��ȹ�� ���������
' History : 2018.04.10 ������
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
		oPlanEvent.GetOneShoppingListContents

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
	    if (frm.itemid1.value==""){
	        alert('������1 �Է� �ϼ���.');
	        frm.itemid1.focus();
	        return;
	    }

		<% If idx="4" Or idx="5" Or idx="6" Or idx="13" Then %>
		if (frm.upload_img1.value==""){
	        alert('������1 �̹����� ����� �ּ���.');
	        frm.upload_img1.focus();
	        return;
	    }
		<% End If %>

		<% If idx="5" Or idx="6" Then %>
		if (frm.upload_img2.value==""){
	        alert('������2 �̹����� ����� �ּ���.');
	        frm.upload_img2.focus();
	        return;
	    }
		<% End If %>

		<% If idx="5" Or idx="6" Then %>
		if (frm.upload_img3.value==""){
	        alert('������3 �̹����� ����� �ּ���.');
	        frm.upload_img3.focus();
	        return;
	    }
		<% End If %>

		<% If idx="5" Then %>
		if (frm.upload_img4.value==""){
	        alert('������4 �̹����� ����� �ּ���.');
	        frm.upload_img4.focus();
	        return;
	    }
		<% End If %>

		<% If idx="5" Then %>
		if (frm.upload_img5.value==""){
	        alert('������5 �̹����� ����� �ּ���.');
	        frm.upload_img5.focus();
	        return;
	    }
		<% End If %>

		<% If idx="5" Then %>
		if (frm.upload_img6.value==""){
	        alert('������6 �̹����� ������ �ּ���.');
	        frm.upload_img6.focus();
	        return;
	    }
		<% End If %>

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
<form name="frmcontents" method="post" action="doWeddingShoppingListReg.asp" onsubmit="return false;">
<input type="hidden" name="Evt_Code" value="10000">
<input type="hidden" name="WeddingStepID" value="<%= oPlanEvent.FOneItem.FWeddingStepID %>">
<tr bgcolor="#FFFFFF">
    <td width="150" bgcolor="#DDDDFF">Idx</td>
    <td>
        [<%= oPlanEvent.FOneItem.FWeddingStepID %>]<%= oPlanEvent.FOneItem.GetDDayTitle %>
    </td>
</tr>
<tr bgcolor="#FFFFFF">
    <td width="100" bgcolor="#DDDDFF">������1</td>
    <td>
		<input type="text" name="itemid1" value="<%=oPlanEvent.FOneItem.FItemid1%>"> <a href="javascript:addnewItem(1);">�ҷ�����</a> <span id="iteminfo1"></span>
    </td>
</tr>
<tr bgcolor="#FFFFFF">
  <td width="180" bgcolor="#DDDDFF"><% If idx="4" Or idx="5" Or idx="6" Or idx="13" Then %><font color="red">(���ʼ�)</font><% End If %>������1 �̹��� ���ε�</td>
  <td>
  	<input type="button" name="etcitem" value="�̹��� ���" onClick="jsSetImg('<%=oPlanEvent.FOneItem.FUpload_img1%>','upload_img1','item1')" class="button">
	<input type="hidden" name="upload_img1" value="<%=oPlanEvent.FOneItem.FUpload_img1%>">
		<div id="item1" style="padding: 5 5 5 5">
			<%IF oPlanEvent.FOneItem.FUpload_img1 <> "" THEN %>
			<img  src="<%=oPlanEvent.FOneItem.FUpload_img1%>" width="50%" border="0">
			<a href="javascript:jsDelImg('upload_img1','item1');"><img src="/images/icon_delete2.gif" border="0"></a>
			<%END IF%>
		</div>
  </td>
</tr>
<% If idx="2" Or idx="5" Or idx="6" Or idx="9" Or idx="12" Then %>
<tr bgcolor="#FFFFFF">
    <td width="100" bgcolor="#DDDDFF">������2</td>
    <td>
		<input type="text" name="itemid2" value="<%=oPlanEvent.FOneItem.FItemid2%>"> <a href="javascript:addnewItem(2);">�ҷ�����</a> <span id="iteminfo2"></span>
    </td>
</tr>
<tr bgcolor="#FFFFFF">
  <td width="180" bgcolor="#DDDDFF"><% If idx="5" Or idx="6" Then %><font color="red">(���ʼ�)</font><% End If %>������2 �̹��� ���ε�</td>
  <td>
  	<input type="button" name="etcitem" value="�̹��� ���" onClick="jsSetImg('<%=oPlanEvent.FOneItem.FUpload_img2%>','upload_img2','item2')" class="button">
	<input type="hidden" name="upload_img2" value="<%=oPlanEvent.FOneItem.FUpload_img2%>">
		<div id="item2" style="padding: 5 5 5 5">
			<%IF oPlanEvent.FOneItem.FUpload_img2 <> "" THEN %>
			<img  src="<%=oPlanEvent.FOneItem.FUpload_img2%>" width="50%" border="0">
			<a href="javascript:jsDelImg('upload_img2','item2');"><img src="/images/icon_delete2.gif" border="0"></a>
			<%END IF%>
		</div>
  </td>
</tr>
<% End If %>
<% If idx="2" Or idx="5" Or idx="6" Or idx="9" Then %>
<tr bgcolor="#FFFFFF">
    <td width="100" bgcolor="#DDDDFF">������3</td>
    <td>
		<input type="text" name="itemid3" value="<%=oPlanEvent.FOneItem.FItemid3%>"> <a href="javascript:addnewItem(3);">�ҷ�����</a> <span id="iteminfo3"></span>
    </td>
</tr>
<tr bgcolor="#FFFFFF">
  <td width="180" bgcolor="#DDDDFF"><% If idx="5" Or idx="6" Then %><font color="red">(���ʼ�)</font><% End If %>������3 �̹��� ���ε�</td>
  <td>
  	<input type="button" name="etcitem" value="�̹��� ���" onClick="jsSetImg('<%=oPlanEvent.FOneItem.FUpload_img3%>','upload_img3','item3')" class="button">
	<input type="hidden" name="upload_img3" value="<%=oPlanEvent.FOneItem.FUpload_img3%>">
		<div id="item3" style="padding: 5 5 5 5">
			<%IF oPlanEvent.FOneItem.FUpload_img3 <> "" THEN %>
			<img  src="<%=oPlanEvent.FOneItem.FUpload_img3%>" width="50%" border="0">
			<a href="javascript:jsDelImg('upload_img3','item3');"><img src="/images/icon_delete2.gif" border="0"></a>
			<%END IF%>
		</div>
  </td>
</tr>
<% End If %>
<% If idx="5" Or idx="9" Then %>
<tr bgcolor="#FFFFFF">
    <td width="100" bgcolor="#DDDDFF">������4</td>
    <td>
		<input type="text" name="itemid4" value="<%=oPlanEvent.FOneItem.FItemid4%>"> <a href="javascript:addnewItem(4);">�ҷ�����</a> <span id="iteminfo4"></span>
    </td>
</tr>
<tr bgcolor="#FFFFFF">
  <td width="180" bgcolor="#DDDDFF"><% If idx="5" Then %><font color="red">(���ʼ�)</font><% End If %>������4 �̹��� ���ε�</td>
  <td>
  	<input type="button" name="etcitem" value="�̹��� ���" onClick="jsSetImg('<%=oPlanEvent.FOneItem.FUpload_img4%>','upload_img4','item4')" class="button">
	<input type="hidden" name="upload_img4" value="<%=oPlanEvent.FOneItem.FUpload_img4%>">
		<div id="item4" style="padding: 5 5 5 5">
			<%IF oPlanEvent.FOneItem.FUpload_img4 <> "" THEN %>
			<img  src="<%=oPlanEvent.FOneItem.FUpload_img4%>" width="50%" border="0">
			<a href="javascript:jsDelImg('upload_img4','item4');"><img src="/images/icon_delete2.gif" border="0"></a>
			<%END IF%>
		</div>
  </td>
</tr>
<% End If %>
<% If idx="5" Then %>
<tr bgcolor="#FFFFFF">
    <td width="100" bgcolor="#DDDDFF">������5</td>
    <td>
		<input type="text" name="itemid5" value="<%=oPlanEvent.FOneItem.FItemid5%>"> <a href="javascript:addnewItem(5);">�ҷ�����</a> <span id="iteminfo5"></span>
    </td>
</tr>
<tr bgcolor="#FFFFFF">
  <td width="180" bgcolor="#DDDDFF"><% If idx="5" Then %><font color="red">(���ʼ�)</font><% End If %>������5 �̹��� ���ε�</td>
  <td>
  	<input type="button" name="etcitem" value="�̹��� ���" onClick="jsSetImg('<%=oPlanEvent.FOneItem.FUpload_img5%>','upload_img5','item5')" class="button">
	<input type="hidden" name="upload_img5" value="<%=oPlanEvent.FOneItem.FUpload_img5%>">
		<div id="item5" style="padding: 5 5 5 5">
			<%IF oPlanEvent.FOneItem.FUpload_img5 <> "" THEN %>
			<img  src="<%=oPlanEvent.FOneItem.FUpload_img5%>" width="50%" border="0">
			<a href="javascript:jsDelImg('upload_img5','item5');"><img src="/images/icon_delete2.gif" border="0"></a>
			<%END IF%>
		</div>
  </td>
</tr>
<% End If %>
<% If idx="5" Then %>
<tr bgcolor="#FFFFFF">
    <td width="100" bgcolor="#DDDDFF">������6</td>
    <td>
		<input type="text" name="itemid6" value="<%=oPlanEvent.FOneItem.FItemid6%>"> <a href="javascript:addnewItem(6);">�ҷ�����</a> <span id="iteminfo6"></span>
    </td>
</tr>
<tr bgcolor="#FFFFFF">
  <td width="180" bgcolor="#DDDDFF"><% If idx="5" Then %><font color="red">(���ʼ�)</font><% End If %>������6 �̹��� ���ε�</td>
  <td>
  	<input type="button" name="etcitem" value="�̹��� ���" onClick="jsSetImg('<%=oPlanEvent.FOneItem.FUpload_img6%>','upload_img6','item6')" class="button">
	<input type="hidden" name="upload_img6" value="<%=oPlanEvent.FOneItem.FUpload_img6%>">
		<div id="item5" style="padding: 5 5 5 5">
			<%IF oPlanEvent.FOneItem.FUpload_img6 <> "" THEN %>
			<img  src="<%=oPlanEvent.FOneItem.FUpload_img6%>" width="50%" border="0">
			<a href="javascript:jsDelImg('upload_img6','item6');"><img src="/images/icon_delete2.gif" border="0"></a>
			<%END IF%>
		</div>
  </td>
</tr>
<% End If %>
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
