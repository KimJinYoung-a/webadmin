<%@ language=vbscript %>
<% option explicit %>
<%
'###############################################
' PageName : main_manager.asp
' Discription : ����Ʈ ���� ����
' History : 2008.04.11 ������ : �Ǽ������� ����
'			2009.04.19 �ѿ�� 2009�� �°� ����
'           2009.12.21 ������ : ���ں� �÷��� ���� ��� �߰�
'			2012.02.08 ������ : �̴ϴ޷� ��ü
'           2013.09.28 ������ : 2013������ - �߰����� �ʵ� �߰�
'           2015.04.07 ������ : 2015������ - �߰����� �ʵ� �߰�
'           2018-01-15 ����ȭ : ���� PC��� ���� �߰�
'           2019.10.30 ������ : �귣�� ���� �ҷ������ ����
'###############################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/classes/sitemasterclass/main_enjoyContentsManageCls.asp" -->
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

	dim oMainContents
		set oMainContents = new CMainEnjoyContents
		oMainContents.FRectIdx = idx
		oMainContents.GetOneNewBrandMainContents

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

		if (frm.BrandID.value==""){
	        alert('�귣�� ID�� �Է� �ϼ���.');
	        frm.BrandID.focus();
	        return;
	    }
		
		if (frm.MainCopy.value==""){
	        alert('����ī�Ǹ� �Է� �ϼ���.');
	        frm.MainCopy.focus();
	        return;
	    }

	    if (frm.startdate.value.length!=10){
	        alert('�������� �Է�  �ϼ���.');
	        return;
	    }

	    if (frm.enddate.value.length!=10){
	        alert('�������� �Է�  �ϼ���.');
	        return;
	    }

	    if (frm.startdate.value>frm.enddate.value){
	        alert('�������� �����Ϻ��� ������ �ȵ˴ϴ�.');
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

	function cultureloadpop(){
		winLast = window.open('pop_culturelist.asp','pLast','width=1200,height=600, scrollbars=yes')
		winLast.focus();
	}

	//�����ڵ� ����
	function selColorChip(bg,cd) {
		var i;
		document.frmcontents.BGColor.value= bg;
		for(i=1;i<=11;i++) {
			document.all("cline"+i).bgColor='#DDDDDD';
		}
		if(!cd) document.all("cline0").bgColor='#DD3300';
		else document.all("cline"+cd).bgColor='#DD3300';
	}

	//-- jsLastEvent : ���� �̺�Ʈ �ҷ����� --//
	function jsLastEvent(num){
	  winLast = window.open('pop_event_lastlist.asp?num='+num,'pLast','width=800,height=600, scrollbars=yes')
	  winLast.focus();
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

	//�귣�� ID �˻� �˾�â
	function jsSearchBrandIDNew(frmName,compName){
		var compVal = "";
		try{
			compVal = eval("document.all." + frmName + "." + compName).value;
		}catch(e){
			compVal = "";
		}

		var popwin = window.open("popBrandSearch.asp?frmName=" + frmName + "&compName=" + compName + "&rect=" + compVal,"popBrandSearch","width=800 height=400 scrollbars=yes resizable=yes");

		popwin.focus();
	}
</script>

<table width="100%" cellpadding="2" cellspacing="1" class="a" bgcolor="#3d3d3d">
<form name="frmcontents" method="post" action="doMainNewBrandReg.asp" onsubmit="return false;">
<input type="hidden" name="Main_Image" value="<%=oMainContents.FOneItem.FMain_Image%>">
<tr bgcolor="#FFFFFF">
    <td width="100" bgcolor="#DDDDFF">Idx</td>
    <td>
        <% if oMainContents.FOneItem.Fidx<>"" then %>
        <%= oMainContents.FOneItem.Fidx %>
        <input type="hidden" name="idx" value="<%= oMainContents.FOneItem.Fidx %>">
        <% else %>

        <% end if %>
    </td>
</tr>
<tr bgcolor="#FFFFFF">
    <td width="100" bgcolor="#DDDDFF">�귣��</td>
    <td>
		<% NewDrawSelectBoxDesignerwithNameEvent "BrandID", oMainContents.FOneItem.FBrandID %>
    </td>
</tr>
<tr bgcolor="#FFFFFF">
    <td width="100" bgcolor="#DDDDFF">����ī��</td>
    <td height="60">
		<textarea name="MainCopy" rows="5" cols="30"><%=oMainContents.FOneItem.FMainCopy%></textarea>
    </td>
</tr>
<tr bgcolor="#FFFFFF" id="tmpstyle1">
	<td bgcolor="#DDDDFF" align="center" width="15%">�귣�� �̹���</td>
	<td>
		<div id="newbrandimg" style="padding: 5 5 5 5">
			<%IF oMainContents.FOneItem.FMain_Image <> "" THEN %>
			<a href="javascript:jsImgView('<%=oMainContents.FOneItem.FMain_Image%>')"><img src="<%=oMainContents.FOneItem.FMain_Image%>" width="400" border="0" id="mainimg"></a>
			<% else %>
			<img src="" width="400" border="0" id="mainimg">
			<%END IF%>
		</div>
		<span id="imgurl"><%=oMainContents.FOneItem.FMain_Image%></span>
	</td>
</tr>
<tr bgcolor="#FFFFFF">
  <td width="100" bgcolor="#DDDDFF">������</td>
  <td>
  	<input type="text" name="StartDate" id="startdate" value="<%=chkiif(idx=0,prevDate,oMainContents.FOneItem.FStartDate)%>">
	<img src="http://scm.10x10.co.kr/images/calicon.gif" id="startdate_trigger" border="0" style="cursor:pointer;" align="absbottom" />
	<script type="text/javascript">
	var CAL_Start = new Calendar({
		inputField : "startdate",
		trigger    : "startdate_trigger",
		onSelect: function() {
			var date = Calendar.intToDate(this.selection.get());
			CAL_End.args.min = date;
			CAL_End.redraw();
			this.hide();
		},
		bottomBar: true,
		dateFormat: "%Y-%m-%d"
	});
	</script>
  </td>
</tr>
<tr bgcolor="#FFFFFF">
  <td width="100" bgcolor="#DDDDFF">������</td>
  <td>
  	<input type="text" name="EndDate" id="enddate" value="<%=chkiif(idx=0,prevDate,oMainContents.FOneItem.FEndDate)%>">
	<img src="http://scm.10x10.co.kr/images/calicon.gif" id="enddate_trigger" border="0" style="cursor:pointer;" align="absbottom" />
	<script type="text/javascript">
	var CAL_End = new Calendar({
		inputField : "enddate",
		trigger    : "enddate_trigger",
		onSelect: function() {
			var date = Calendar.intToDate(this.selection.get());
			CAL_Start.args.max = date;
			CAL_Start.redraw();
			this.hide();
		},
		bottomBar: true,
		dateFormat: "%Y-%m-%d"
	});
	</script>
  </td>
</tr>
<tr bgcolor="#FFFFFF">
  <td width="100" bgcolor="#DDDDFF">�켱����</td>
  <td>
  	<input type="text" name="DispOrder" value="<%=oMainContents.FOneItem.FDispOrder%>">
  </td>
</tr>
<tr bgcolor="#FFFFFF">
  <td width="100" bgcolor="#DDDDFF">��뿩��</td>
  <td>
  	<input type="radio" name="Isusing" value="Y"<% If oMainContents.FOneItem.FIsusing="Y" Or oMainContents.FOneItem.FIsusing="" Then Response.write " checked" %>> �����
	<input type="radio" name="Isusing" value="N"<% If oMainContents.FOneItem.FIsusing="N" Then Response.write " checked" %>> ������
  </td>
</tr>
<% If oMainContents.FOneItem.FRegUser<>"" Then %>
<tr bgcolor="#FFFFFF">
  <td width="100" bgcolor="#DDDDFF">�۾���</td>
  <td>
  	�۾��� : <%=oMainContents.FOneItem.FRegUser %><br>
	�����۾��� : <%=oMainContents.FOneItem.FLastUser %>
  </td>
</tr>
<% End If %>
<tr bgcolor="#FFFFFF">
    <td colspan="2" align="center"><input type="button" value=" �� �� " onClick="SaveMainContents(frmcontents);"></td>
</tr>
</form>
</table>
<%
set oMainContents = Nothing
%>

<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
