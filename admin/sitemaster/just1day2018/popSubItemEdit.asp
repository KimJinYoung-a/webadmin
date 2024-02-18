<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/classes/sitemasterclass/just1DayCls2018New.asp" -->
<%
'// ���� ����
Dim listidx , subIdx , Itemid, title, pclinkUrl, mobilelinkUrl, pcimage, mobileimage, price, saleper, sortnum, isusing
Dim mode, regdate
Dim itemdiv, Frontimage, sqlStr

'// �Ķ���� ����
listidx = request("listidx")
subidx = request("subidx")
itemid = request("itemid")
itemdiv = request("itemdiv")

If subidx = "" Then
	sortnum = "0"
	mode = "subadd"
Else
	mode = "submodify"
End If

If itemid <> "" Then
	sqlStr = " select itemname from db_item.dbo.tbl_item Where itemid='"&itemid&"' "
	rsget.Open sqlStr, dbget, 1
		title = rsget("itemname")
	rsget.close
End If

If subidx <> "" then
	dim just1dayItem
	set just1dayItem = new Cjust1Day
	just1dayItem.FRectSubIdx = subidx
	just1dayItem.GetOneSubItem()

	subIdx			=	just1dayItem.FOneItem.FsubIdx
	listidx			=	just1dayItem.FOneItem.Flistidx
	Itemid			=	just1dayItem.FOneItem.FItemid
	sortnum			=	just1dayItem.FOneItem.Fsortnum
	isusing			=	just1dayItem.FOneItem.Fisusing
	title			=	just1dayItem.FOneItem.FTitle
	Frontimage		=	just1dayItem.FOneItem.FitemFrontimage
	price			=	just1dayItem.FOneItem.FitemPrice
	saleper			=	just1dayItem.FOneItem.FitemSaleper
	regdate			=	just1dayItem.FOneItem.Fregdate
	itemdiv			=	just1dayItem.FOneItem.Fitemdiv

	set just1dayItem = Nothing
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

//-- jsImgView : �̹��� Ȯ��ȭ�� ��â���� �����ֱ� --//
function jsImgView(sImgUrl){
 var wImgView;
 wImgView = window.open('/admin/eventmanage/common/pop_event_detailImg.asp?sUrl='+sImgUrl,'pImg','width=100,height=100');
 wImgView.focus();
}


function jsSetImg(sFolder, sImg, sName, sSpan){ 
	var winImg;
	winImg = window.open('/admin/eventmanage/common/pop_event_uploadimgV2.asp?yr=<%=Year(now())%>&sF='+sFolder+'&sImg='+sImg+'&sName='+sName+'&sSpan='+sSpan,'popImg','width=450,height=300');
	winImg.focus();
}

function jsDelImg(sName, sSpan){
	if(confirm("�̹����� �����Ͻðڽ��ϱ�?\n\n���� �� �����ư�� ������ ó���Ϸ�˴ϴ�.")){
	   eval("document.all."+sName).value = "";
	   eval("document.all."+sSpan).style.display = "none";
	}
}

function popRegSearchItem() {
<% if listidx <> "" then %>
    var popwinsub = window.open("/admin/itemmaster/pop_itemAddInfo.asp?sellyn=Y&usingyn=Y&defaultmargin=0&acURL=/admin/sitemaster/just1day2018/doSubRegItemCdArray.asp?listidx=<%=listidx%>&ptype=just1day", "popup_itemsub", "width=800,height=500,scrollbars=yes,resizable=yes");
    popwinsub.focus();
<% else %>
	alert("���ø� ������ ������ ���� ������ּ���.");
<% end if %>
}

</script>
<center>
<form name="frm" method="post" action="dojust1day.asp">
<input type="hidden" name="adminid" value="<%=session("ssBctId")%>">
<input type="hidden" name="mode" value="<%=mode%>" />
<input type="hidden" name="listidx" value="<%=listidx%>" />
<input type="hidden" name="subIdx" value="<%=subIdx%>" />
<input type="hidden" name="Frontimage" value="<%=Frontimage%>">
<table width="100%" cellpadding="2" cellspacing="1" class="a" bgcolor="#3d3d3d" style="table-layout: fixed;">
<tr bgcolor="#FFFFFF">
    <td height="25" colspan="4" bgcolor="#F8F8F8"><b>Just1Day ��ǰ ���� ���/����</b></td>
</tr>
<colgroup>
	<col width="150" />
	<col width="*" />
</colgroup>
<tr bgcolor="#FFFFFF">
    <td bgcolor="#DDDDFF">��ǰ�ڵ�</td>
    <td colspan="3">
        <input type="text" name="itemid" value="<%= itemid %>" size="8" maxlength="8" class="text" title="��ǰ�ڵ�" readonly />&nbsp;<input type="button" value="��ǰ ����" class="button" onClick="popRegSearchItem()" />
    </td>
</tr>
<tr bgcolor="#FFFFFF">
    <td bgcolor="#DDDDFF">�����</td>
    <td colspan="3">
        <input type="text" name="title" value="<%= title %>" size="60" maxlength="60" class="text" title="�����" />
    </td>
</tr>
<tr bgcolor="#FFFFFF">
    <td bgcolor="#DDDDFF">FrontImage</td>
	<td colspan="3"><input type="button" name="wimg" value="Front �̹��� ���" onClick="jsSetImg('today','<%=Frontimage%>','Frontimage','frontimg')" class="button">
		<div id="frontimg" style="padding: 5 5 5 5">
			<%IF Frontimage <> "" THEN %>
			<a href="javascript:jsImgView('<%=Frontimage%>')"><img  src="<%=Frontimage%>" width="400" border="0"></a>
			<a href="javascript:jsDelImg('Frontimage','frontimg');"><img src="/images/icon_delete2.gif" border="0"></a>
			<%END IF%>
		</div>
		<%=Frontimage%>
	</td>
</tr>

<% If Trim(itemdiv)="21" Then %>
	<tr bgcolor="#FFFFFF">
		<td bgcolor="#DDDDFF">����</td>
		<td colspan="3">
			<input type="text" name="price" value="<%= price %>" size="20" maxlength="60" class="text" title="����" /> ex ) 14,000�� or 9,900��~
		</td>
	</tr>
<% End If %>		
	<tr bgcolor="#FFFFFF">
		<td bgcolor="#DDDDFF">������</td>
		<td colspan="3">
			<input type="text" name="saleper" value="<%= saleper %>" size="20" maxlength="60" class="text" title="����" /> ex ) 19% or ~51% (�Ϲݻ�ǰ:�ֿ켱 �Է°�,������ ����N->�Ϸ�Ư��,����Y->������)
		</td>
	</tr>


<tr bgcolor="#FFFFFF">
    <td bgcolor="#DDDDFF">���ļ���</td>
    <td colspan="3">
        <input type="text" name="sortnum" class="text" size="4" value="<%=sortnum%>" />
    </td>
</tr>
<tr bgcolor="#FFFFFF">
    <td bgcolor="#DDDDFF">��뿩��</td>
    <td colspan="3">
		<span id="rdoUsing">
		<input type="radio" name="isusing" id="rdoUsing1" value="Y" <%=chkIIF(isusing="Y" or isusing="","checked","")%> /><label for="rdoUsing1">���</label>
		<input type="radio" name="isusing" id="rdoUsing2" value="N" <%=chkIIF(isusing="N","checked","")%> /><label for="rdoUsing2">������</label>
		</span>
    </td>
</tr>
<tr bgcolor="#FFFFFF">
	<% If subidx<>"" Then %>
	    <td colspan="4" align="center"><input type="button" value=" �� �� " onClick="SaveForm(this.form);"></td>
	<% Else %>
	    <td colspan="4" align="center"><input type="button" value=" �� �� " onClick="SaveForm(this.form);"></td>
	<% End If %>
</tr>
</table>
</form>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->