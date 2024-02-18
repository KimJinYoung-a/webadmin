<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/classes/mobile/new_mdpickCls.asp" -->
<%
'###############################################
' PageName : popSubItemEdit.asp
' Discription : ���������� ���/����
' History : 2013.12.17 ����ȭ : �ű� ����
'		  : 2018.09.06 ������ :frontimage �ο� �߰�
'###############################################

'// ���� ����
Dim listidx , subIdx , subItemid , itemName , sortnum , isusing
Dim mode , smallImage , subImage1 , extraurl , gubun , topview, Frontimage, lowestPrice

'// �Ķ���� ����
listidx = request("listidx")
subidx = request("subidx")
gubun = request("gubun")
topview = request("topview")

If subidx = "" Then
	sortnum = "0"
	mode = "subadd"
Else
	mode = "submodify"
End If 

If gubun = topview Then
	topview = 1
Else
	topview = 0
End If 

If subidx <> "" then
	dim mdpickItem
	set mdpickItem = new Cmdpick
	mdpickItem.FRectSubIdx = subidx
	mdpickItem.GetOneSubItem()

	subIdx			=	mdpickItem.FOneItem.FsubIdx
	listidx			=	mdpickItem.FOneItem.Flistidx
	subItemid		=	mdpickItem.FOneItem.FItemid
	sortnum			=	mdpickItem.FOneItem.Fsortnum
	isusing			=	mdpickItem.FOneItem.Fisusing
	itemName		=	mdpickItem.FOneItem.FitemName
	smallImage		=	mdpickItem.FOneItem.FsmallImage
	Frontimage 	 	=   mdpickItem.FOneItem.FFrontImage
	lowestPrice  	=   mdpickItem.FOneItem.FLowestPrice	

	set mdpickItem = Nothing
End If 
%>
<link href="/js/jqueryui/css/jquery-ui.css" rel="stylesheet">
<link href="/js/jqueryui/css/evol.colorpicker.css" rel="stylesheet">
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript" src="/js/jqueryui/jquery-ui-1.10.2.custom.min.js"></script>
<script type="text/javascript" src="/js/jqueryui/evol.colorpicker.min.js"></script>
<script type="text/javascript">
$(function(){
	//������ư
	$("#rdoUsing").buttonset().children().next().attr("style","font-size:11px;");
	//������ư
	$("#rdoLowestPrice").buttonset().children().next().attr("style","font-size:11px;");	
	//�÷���Ŀ
	$("input[name='subBGColor']").colorpicker();
});

// ���˻�
function SaveForm(frm) {
	//
	if (confirm('���� �Ͻðڽ��ϱ�?')){
		//frm.target = "parent";
		frm.submit();
	}
}

// ��ǰ���� ����
function fnGetItemInfo(iid) {
	$.ajax({
		type: "GET",
		url: "/admin/sitemaster/wcms/act_iteminfo.asp?itemid="+iid,
		dataType: "xml",
		cache: false,
		async: false,
		timeout: 5000,
		beforeSend: function(x) {
			if(x && x.overrideMimeType) {
				x.overrideMimeType("text/xml;charset=euc-kr");
			}
		},
		success: function(xml) {
			if($(xml).find("itemInfo").find("item").length>0) {
				var rst = "<img src='" + $(xml).find("itemInfo").find("item").find("smallImage").text() + "' height='50' />"
					rst += $(xml).find("itemInfo").find("item").find("itemname").text();
				$("#lyItemInfo").fadeIn();
				$("#lyItemInfo").html(rst);
			} else {
				$("#lyItemInfo").fadeOut();
			}
		},
		error: function(xhr, status, error) {
			$("#lyItemInfo").fadeOut();
			/*alert(xhr + '\n' + status + '\n' + error);*/
		}
	});
}
function jsSetImg(sFolder, sImg, sName, sSpan){ 
	var winImg;
	winImg = window.open('/admin/eventmanage/common/pop_event_uploadimgV2.asp?yr=<%=Year(now())%>&sF='+sFolder+'&sImg='+sImg+'&sName='+sName+'&sSpan='+sSpan,'popImg','width=450,height=300');
	winImg.focus();
}
//-- jsImgView : �̹��� Ȯ��ȭ�� ��â���� �����ֱ� --//
function jsImgView(sImgUrl){
 var wImgView;
 wImgView = window.open('/admin/eventmanage/common/pop_event_detailImg.asp?sUrl='+sImgUrl,'pImg','width=100,height=100');
 wImgView.focus();
}
function jsDelImg(sName, sSpan){
	if(confirm("�̹����� �����Ͻðڽ��ϱ�?\n\n���� �� �����ư�� ������ ó���Ϸ�˴ϴ�.")){
	   eval("document.all."+sName).value = "";
	   eval("document.all."+sSpan).style.display = "none";
	}
}
</script>
<center>
<form name="frm" method="post" action="domdpick.asp">
<input type="hidden" name="adminid" value="<%=session("ssBctId")%>">
<input type="hidden" name="mode" value="<%=mode%>" />
<input type="hidden" name="listidx" value="<%=listidx%>" />
<input type="hidden" name="gubun" value="<%=gubun%>" />
<input type="hidden" name="topview" value="<%=topview%>" />
<input type="hidden" name="Frontimage" value="<%=Frontimage%>">
<table width="100%" cellpadding="2" cellspacing="1" class="a" bgcolor="#3d3d3d" style="table-layout: fixed;">
<tr bgcolor="#FFFFFF">
    <td height="25" colspan="4" bgcolor="#F8F8F8"><b>���� ���� ���/����</b></td>
</tr>
<colgroup>
	<col width="100" />
	<col width="*" />
	<col width="100" />
	<col width="*" />
</colgroup>
<tr bgcolor="#FFFFFF">
    <td bgcolor="#DDDDFF">���� ��ȣ</td>
    <td colspan="3">
        <%=subIdx %>
        <input type="hidden" name="subIdx" value="<%=subIdx%>" />
    </td>
</tr>
<tr bgcolor="#FFFFFF">
    <td bgcolor="#DDDDFF">��ǰ�ڵ�</td>
    <td colspan="3">
        <input type="text" name="subItemid" value="<%= subItemid %>" size="8" maxlength="8" class="text" require="N" onblur="fnGetItemInfo(this.value)" title="��ǰ�ڵ�" />
        <div id="lyItemInfo" style="display:<%=chkIIF(subItemid="","none","")%>;">
        <%
        	if Not(itemName="" or isNull(itemName)) then
        		Response.Write "<img src='" & smallImage & "' height='50' />"
        		Response.Write itemName
        	end if
        %>
        </div>
    </td>
</tr>
<tr bgcolor="#FFFFFF">
    <td bgcolor="#DDDDFF">FrontImage</td>
	<td colspan="3"><input type="button" name="wimg" value="Front �̹��� ���" onClick="jsSetImg('mdspick','<%=Frontimage%>','Frontimage','frontimg')" class="button">
		<div id="frontimg" style="padding: 5 5 5 5">
			<%IF Frontimage <> "" THEN %>
			<a href="javascript:jsImgView('<%=Frontimage%>')"><img  src="<%=Frontimage%>" width="400" border="0"></a>
			<a href="javascript:jsDelImg('Frontimage','frontimg');"><img src="/images/icon_delete2.gif" border="0"></a>
			<%END IF%>
		</div>
		<%=Frontimage%>
	</td>
</tr>
<tr bgcolor="#FFFFFF">
    <td bgcolor="#DDDDFF">���ļ���</td>
    <td>
        <input type="text" name="sortnum" class="text" size="4" value="<%=sortnum%>" />
    </td>
    <td bgcolor="#DDDDFF">��뿩��</td>
    <td>
		<span id="rdoUsing">
		<input type="radio" name="isusing" id="rdoUsing1" value="Y" <%=chkIIF(isusing="Y" or isusing="","checked","")%> /><label for="rdoUsing1">���</label>
		<input type="radio" name="isusing" id="rdoUsing2" value="N" <%=chkIIF(isusing="N","checked","")%> /><label for="rdoUsing2">����</label>
		</span>
    </td>
</tr>
<tr bgcolor="#FFFFFF">
    <td bgcolor="#DDDDFF">������ ǥ��</td>
    <td colspan="3">
		<span id="rdoLowestPrice">
		<input type="radio" name="islowestprice" id="rdoLowestPrice1" value="Y" <%=chkIIF(lowestPrice="Y","checked","")%> /><label for="rdoLowestPrice1">���</label>
		<input type="radio" name="islowestprice" id="rdoLowestPrice2" value="N" <%=chkIIF(lowestPrice="N" or lowestPrice="","checked","")%> /><label for="rdoLowestPrice2">������</label>
		</span>
    </td>
</tr>
<tr bgcolor="#FFFFFF">
    <td colspan="4" align="center"><input type="button" value=" �� �� " onClick="SaveForm(this.form);"></td>
</tr>
</table>
</form>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->