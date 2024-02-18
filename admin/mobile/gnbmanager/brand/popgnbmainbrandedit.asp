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
'###############################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/classes/sitemasterclass/gnb_main_ContentsManageCls.asp" -->
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
		set oMainContents = new CGNBContents
		oMainContents.FRectIdx = idx
		oMainContents.GetOneBrandMainContents
	
	If gubun = "" Then
		gubun = "index"
	End If

%>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
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

		if (frm.BrandName.value==""){
	        alert('�귣����� �Է� �ϼ���.');
	        frm.BrandName.focus();
	        return;
	    }

		if (frm.MainCopy.value==""){
	        alert('�귣�� ����ī�Ǹ� �Է� �ϼ���.');
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

	//�귣�� ID �˻� �˾�â
	function jsSearchBrandIDThis(frmName,compName){
		var compVal = "";
		try{
			compVal = eval("document.all." + frmName + "." + compName).value;
		}catch(e){
			compVal = "";
		}

		var popwin = window.open("popBrandSearch.asp?frmName=" + frmName + "&compName=" + compName + "&rect=" + compVal,"popBrandSearch","width=800 height=400 scrollbars=yes resizable=yes");

		popwin.focus();
	}


	function jsSetImg(sFolder, sImg, sName, sSpan){ 
		var winImg;
		winImg = window.open('/admin/eventmanage/common/pop_event_uploadimgV2.asp?yr=<%=Year(now())%>&sF='+sFolder+'&sImg='+sImg+'&sName='+sName+'&sSpan='+sSpan,'popImg','width=370,height=150');
		winImg.focus();
	}

	/**
	 * ����Ʈ ���� �Է°��� ���ڼ� üũ
	 * 
	 * @param id : tag id 
	 * @param title : tag title
	 * @param maxLength : �ִ� �Է°��� �� (byte)
	 * @returns {Boolean}
	 */
	function maxLengthCheck(id, title, maxLength){
		 var obj = $("#"+id);
		 if(maxLength == null) {
			 maxLength = obj.attr("maxLength") != null ? obj.attr("maxLength") : 1000;
		 }
		 
		 if(Number(byteCheck(obj)) > Number(maxLength)) {
			 alert(title + "��(��) �Է°��ɹ��ڼ��� �ʰ��Ͽ����ϴ�.\n(����, ����, �Ϲ� Ư������ : " + maxLength + " / �ѱ�, ����, ��Ÿ Ư������ : " + parseInt(maxLength/2, 10) + ").");
			 obj.focus();
			 return false;
		 } else {
			 return true;
		}
	}

	/**
	 * ����Ʈ�� ��ȯ  
	 * 
	 * @param el : tag jquery object
	 * @returns {Number}
	 */
	function byteCheck(el){
		var codeByte = 0;
		for (var idx = 0; idx < el.val().length; idx++) {
			var oneChar = escape(el.val().charAt(idx));
			if ( oneChar.length == 1 ) {
				codeByte ++;
			} else if (oneChar.indexOf("%u") != -1) {
				codeByte += 2;
			} else if (oneChar.indexOf("%") != -1) {
				codeByte ++;
			}
		}
		return codeByte;
	}

</script>

<table width="100%" cellpadding="2" cellspacing="1" class="a" bgcolor="#3d3d3d">
<form name="frmcontents" method="post" action="doMainBrandReg.asp" onsubmit="return false;">
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
<tr bgcolor="#FFFFFF" height="30">
	<td bgcolor="#DDDDFF">�귣��ID</td>
	<td><input type="text" name="BrandID" value="<%=oMainContents.FOneItem.FBrandID%>"> <a href="javascript:jsSearchBrandIDThis('frmcontents','BrandID');">�ҷ�����</a></td>
</tr>
<tr bgcolor="#FFFFFF" height="30">
	<td bgcolor="#DDDDFF">�귣���</td>
	<td><input type="text" name="BrandName" id="Evt_Title" value="<%=oMainContents.FOneItem.FBrandName%>" ></td>
</tr>
<tr bgcolor="#FFFFFF" height="30">
	<td bgcolor="#DDDDFF">����ī��</td>
	<td><input type="text" name="MainCopy" value="<%=oMainContents.FOneItem.FMainCopy%>" size="50" maxlength="60"></td>
</tr>
<tr bgcolor="#FFFFFF" height="30">
	<td bgcolor="#DDDDFF">�̹���</td>
	<td>
		�����۹�ȣ : <input type="text" name="itemID" value="<%=oMainContents.FOneItem.FitemID%>" size="12">
		<input type="button" name="limg" value="�̹��� ���" onClick="jsSetImg('pcmainnewbrand','<%=oMainContents.FOneItem.FMain_Image%>','Main_Image','newbrandimg')" class="button">
		<div id="newbrandimg" style="padding: 5 5 5 5">
			<%IF oMainContents.FOneItem.FMain_Image <> "" THEN %>
			<a href="javascript:jsImgView('<%=oMainContents.FOneItem.FMain_Image%>')"><img  src="<%=oMainContents.FOneItem.FMain_Image%>" width="400" border="0"></a>
			<a href="javascript:jsDelImg('Main_Image','newbrandimg');"><img src="/images/icon_delete2.gif" border="0"></a>
			<%END IF%>
		</div>
		<%=oMainContents.FOneItem.FMain_Image%>
	</td>
</tr>
<tr bgcolor="#FFFFFF">
  <td width="100" bgcolor="#DDDDFF">������</td>
  <td>
  	<input type="text" name="StartDate" id="startdate" value="<%=oMainContents.FOneItem.FStartDate%>">
	<img src="http://webadmin.10x10.co.kr/images/calicon.gif" id="startdate_trigger" border="0" style="cursor:pointer;" align="absbottom" />
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
  	<input type="text" name="EndDate" id="enddate" value="<%=oMainContents.FOneItem.FEndDate%>">
	<img src="http://webadmin.10x10.co.kr/images/calicon.gif" id="enddate_trigger" border="0" style="cursor:pointer;" align="absbottom" />
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
