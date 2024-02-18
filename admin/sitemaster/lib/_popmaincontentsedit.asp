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
<!-- #include virtual="/lib/classes/sitemasterclass/main_ContentsManageCls.asp" -->
<!-- #include virtual="/admin/eventmanage/common/event_function.asp"-->
<%
dim isusing, fixtype, validdate, prevDate
dim idx, poscode, reload, gubun, edid
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

	if reload="on" then
	    response.write "<script>opener.location.reload(); window.close();</script>"
	    dbget.close()	:	response.End
	end if

	dim oMainContents
		set oMainContents = new CMainContents
		oMainContents.FRectIdx = idx
		oMainContents.GetOneMainContents

dim oposcode, defaultMapStr, defaultXMLMapStr
	set oposcode = new CMainContentsCode
	oposcode.FRectPosCode = poscode
	if poscode<>"" then
	    oposcode.GetOneContentsCode

	    defaultMapStr = "<map name='Map_" +oposcode.FOneItem.FPosvarName + "'>" + VbCrlf
	    defaultMapStr = defaultMapStr + VbCrlf
	    defaultMapStr = defaultMapStr + "</map>"

		defaultXMLMapStr = ""
	    defaultXMLMapStr = defaultXMLMapStr + "<map name='Map_" +oposcode.FOneItem.FPosvarName + "'>"+ VbCrlf
	    defaultXMLMapStr = defaultXMLMapStr + VbCrlf
		defaultXMLMapStr = defaultXMLMapStr + "</map>"
	end if

dim orderidx
	if oMainContents.FOneItem.forderidx = "" then
		orderidx = 99
	else
		orderidx = oMainContents.FOneItem.forderidx
		poscode = oMainContents.FOneItem.fposcode
	end if

	If gubun = "" Then
		gubun = "index"
	End If

	edid = oMainContents.FOneItem.Fworkeruserid
	If edid = "" Then
		If idx <> "" AND idx <> "0" Then
			edid = session("ssBctId")
		End If
	End If

	'// ���Ľ����̼� �ҷ�����
	Dim cultureContents
	Dim cultureEcode ,	cultureEtype ,cultureEname ,cultureEcomment , cultureEimagelist

	If idx <> "" And culturecode = "" Then culturecode = oMainContents.FOneItem.Fecode

	If culturecode <> "" Then
		Dim SqlStr
		sqlStr = "SELECT C.evt_code ,C.evt_type , C.evt_name , C.evt_comment , C.image_list" & vbcrlf
		sqlStr = sqlStr &" FROM db_culture_station.dbo.tbl_culturestation_event as C" & vbcrlf
		sqlStr = sqlStr & "WHERE C.evt_code = "& culturecode

        rsget.Open SqlStr, dbget, 1
		if Not rsget.Eof then
			cultureEcode		= rsget("evt_code")
			cultureEtype		= rsget("evt_type")
			cultureEname		= rsget("evt_name")
			cultureEcomment		= rsget("evt_comment")
			'cultureEimagelist	= webImgUrl &"/culturestation/2009/list/" & rsget("image_list")
			cultureEimagelist	= rsget("image_list")
		end if
        rsget.close
	End If

'// Ư�� �ڵ忡 ��ũ�ؽ�Ʈ �߰�(IMG ALT �� ��)
dim IsLinkTextNeed
	IsLinkTextNeed = (InStr(",630,642,659,673,674,675,687,", ("," & poscode & ",")) > 0)

%>
<script type="text/javascript" src="/js/jquery-1.7.2.min.js"></script>
<script type="text/javascript" src="/js/jsCal/js/jscal2.js"></script>
<script type="text/javascript" src="/js/jsCal/js/lang/ko.js"></script>
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/jscal2.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/border-radius.css" />
<script type="text/javascript">
<%
	'ecode ���Ľ����̼��̺�Ʈid
	'maincopy ��������
	'subcopy �߰� �ڸ�Ʈ����
	'linktext3  ���� (����)
	'xbtncolor 0/1	���м���
	'file1 �̹��� --  �̹��� �� �ְ� �����ؼ� ����ҵ�
%>
	<% if culturecode <> "" then %>
	$(function(){
		var gubuncode = "<%=cultureEtype%>";
		var frm = document.frmcontents;
			frm.ecode.value = "<%=cultureEcode%>";
			frm.maincopy.value = "<%=cultureEname%>";
			frm.subcopy.value = "<%=cultureEcomment%>";
			if (gubuncode == "0"){
				frm.xbtncolor[0].value = "0";
				frm.xbtncolor[0].checked = true;
			}else{
				frm.xbtncolor[1].value = "1";
				frm.xbtncolor[1].checked = true;
			}
			frm.linkurl.value = "/culturestation/culturestation_event.asp?evt_code=<%=cultureEcode%>";
	});
	<% end if %>

	function SaveMainContents(frm){
	    if (frm.poscode.value.length<1){
	        alert('������ ���� ���� �ϼ���.');
	        frm.poscode.focus();
	        return;
	    }

	    if (frm.linkurl.value.length<1){
	        alert('��ũ ���� �Է� �ϼ���.');
	        frm.linkurl.focus();
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
		<% if poscode <> "562" and poscode <> "561" then  %>
		if (!frm.altname.value){
			alert('alt���� �Է� �ϼ���.');
			frm.altname.focus();
			return;
		}
		<% end if %>

	    var vstartdate = new Date(frm.startdate.value.substr(0,4), (1*frm.startdate.value.substr(5,2))-1, frm.startdate.value.substr(8,2));
	    var venddate = new Date(frm.enddate.value.substr(0,4), (1*frm.enddate.value.substr(5,2))-1, frm.enddate.value.substr(8,2));

	    if (vstartdate>venddate){
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

	function putLinkText(key) {
		var frm = document.frmcontents;
		switch(key) {
			case 'search':
				frm.linkurl.value='/search/search_item.asp?rect=�˻���';
				break;
			case 'event':
				frm.linkurl.value='/event/eventmain.asp?eventid=�̺�Ʈ��ȣ';
				break;
			case 'itemid':
				frm.linkurl.value='/shopping/category_prd.asp?itemid=��ǰ�ڵ�';
				break;
			case 'category':
				frm.linkurl.value='/shopping/category_list.asp?disp=ī�װ�';
				break;
			case 'brand':
				frm.linkurl.value='/street/street_brand.asp?makerid=�귣����̵�';
				break;
			case 'showbanner':
				frm.linkurl.value='/showbanner/show_view.asp?showidx=���ʾ��̵�';
				break;
			case 'culture':
				frm.linkurl.value='/culturestation/culturestation_event.asp?evt_code=�̺�Ʈ���̵�';
				break;
			case 'ground':
				frm.linkurl.value='/play/playGround.asp?idx=�׶����ȣ&contentsidx=��������ȣ';
				break;
			case 'styleplus':
				frm.linkurl.value='/play/playStylePlus.asp?idx=��Ÿ���÷�����ȣ&contentsidx=��������ȣ';
				break;
			case 'fingers':
				frm.linkurl.value='/play/playDesignFingers.asp?idx=�ΰŽ���ȣ&contentsidx=��������ȣ';
				break;
			case 'tepisode':
				frm.linkurl.value='/play/playTEpisode.asp?idx=Ƽ���Ǽҵ��ȣ&contentsidx=��������ȣ';
				break;
			case 'gift':
				frm.linkurl.value='/gift/gifttalk/';
				break;
			case 'wish':
				frm.linkurl.value='/wish/index.asp';
				break;
			case 'hitchhiker':
				frm.linkurl.value='/hitchhiker/';
				break;
			case 'giftcard':
				frm.linkurl.value='/giftcard/';
				break;
		}
	}

	function cultureloadpop(){
		winLast = window.open('pop_culturelist.asp?gubun=<%=gubun%>&poscode=<%=poscode%>&pidx=<%=idx%>','pLast','width=1200,height=600, scrollbars=yes')
		winLast.focus();
	}

	function fnSelectBannerType(bannertype){
		if(bannertype==1)
		{
			$("#bnimg2").hide();
			$("#bnalt2").hide();
			$("#bnbg1").hide();
			$("#bnbg2").hide();
			$("#bnlink2").hide();
		}
		else
		{
			$("#bnbg1").show();
			$("#bnbg2").show();
			$("#bnimg2").show();
			$("#bnalt2").show();
			$("#bnlink2").show();
		}
	}
</script>

<table width="100%" cellpadding="2" cellspacing="1" class="a" bgcolor="#3d3d3d">
<form name="frmcontents" method="post" action="<%=uploadUrl%>/linkweb/doMainContentsRegNew.asp" onsubmit="return false;" enctype="multipart/form-data">
<tr bgcolor="#FFFFFF">
    <td width="150" bgcolor="#DDDDFF">Idx</td>
    <td>
        <% if oMainContents.FOneItem.Fidx<>"" then %>
        <%= oMainContents.FOneItem.Fidx %>
        <input type="hidden" name="idx" value="<%= oMainContents.FOneItem.Fidx %>">
        <% else %>

        <% end if %>
    </td>
</tr>
<tr bgcolor="#FFFFFF">
    <td width="150" bgcolor="#DDDDFF">�׷챸��</td>
    <td>
        <% if oMainContents.FOneItem.Fidx<>"" then %>
        <%= oMainContents.FOneItem.Fgubun %>
        <input type="hidden" name="gubun" value="<%= oMainContents.FOneItem.Fgubun %>">
        <% else %>
        <% call DrawGroupGubunCombo("gubun", gubun, "onChange='ChangeGroupGubun(this);'") %>
        <% end if %>
    </td>
</tr>
<tr bgcolor="#FFFFFF">
    <td width="150" bgcolor="#DDDDFF">���и�</td>
    <td>
    	<% IF gubun <> "" Then %>
	        <% if oMainContents.FOneItem.Fidx<>"" then %>
	        <%= oMainContents.FOneItem.Fposname %> (<%= oMainContents.FOneItem.Fposcode %>)
	        <input type="hidden" name="poscode" value="<%= oMainContents.FOneItem.Fposcode %>">
	        <% else %>
	        <% call DrawMainPosCodeCombo("poscode", poscode, "onChange='ChangeGubun(this);'", gubun) %>
	        <% end if %>
	    <% Else %>
	    	<font color="red">�׷챸���� ���� �����ϼ���</font>
	    <% End If %>
		<% If poscode = "714" Then %>
		&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<span><a href="" onclick="cultureloadpop();return false;">�ҷ�����</a></span>
		<% End If %>
    </td>
</tr>
<tr bgcolor="#FFFFFF">
    <td width="150" bgcolor="#DDDDFF">��ũ����</td>
    <td>
    	<% IF gubun <> "" Then %>
	        <% if oMainContents.FOneItem.Fidx<>"" then %>
	        <%= oMainContents.FOneItem.getlinktypeName %>
	        <input type="hidden" name="linktype" value="<%= oMainContents.FOneItem.Flinktype %>">
	        <% else %>
	            <% if poscode<>"" then %>
	            <%= oposcode.FOneItem.getlinktypeName %>
	            <input type="hidden" name="linktype" value="<%= oposcode.FOneItem.Flinktype %>">
	            <% else %>
	            <font color="red">������ ���� �����ϼ���</font>
	            <% end if %>
	        <% end if %>
	    <% Else %>
	    	<font color="red">�׷챸���� ���� �����ϼ���</font>
	    <% End If %>
    </td>
</tr>
<tr bgcolor="#FFFFFF">
    <td width="150" bgcolor="#DDDDFF">���뱸��(�ݿ��ֱ�)</td>
    <td>
		<% IF gubun <> "" Then %>
	        <% if oMainContents.FOneItem.Fidx<>"" then %>
	        <%= oMainContents.FOneItem.getfixtypeName %>
	        <input type="hidden" name="fixtype" value="<%= oMainContents.FOneItem.Ffixtype %>">
	        <% else %>
	            <% if poscode<>"" then %>
	            <%= oposcode.FOneItem.getfixtypeName %>
	            <input type="hidden" name="fixtype" value="<%= oposcode.FOneItem.Ffixtype %>">
	            <% else %>
	            <font color="red">������ ���� �����ϼ���</font>
	            <% end if %>
	        <% end if %>
	    <% Else %>
	    	<font color="red">�׷챸���� ���� �����ϼ���</font>
	    <% End If %>
    </td>
</tr>
<tr bgcolor="#FFFFFF">
    <td width="150" bgcolor="#DDDDFF">�켱����</td>
    <td>
    	<% IF gubun <> "" Then %>
	        <% if oMainContents.FOneItem.Fidx<>"" then %>
	            <input type="text" name="orderidx" size=5 value="<%= orderidx %>" class="text" />
	        <% else %>
	            <% if poscode<>"" then %>
	            	<input type="text" name="orderidx" size=5 value="<%= orderidx %>" class="text" />
	            <% else %>
	            	<font color="red">������ ���� �����ϼ���</font>
	            <% end if %>
	        <% end if %>
	    <% Else %>
	    	<font color="red">�׷챸���� ���� �����ϼ���</font>
	    <% End If %>
    </td>
</tr>
<tr bgcolor="#FFFFFF">
  <td width="150" bgcolor="#DDDDFF">�۾� ��û����</td>
  <td><textarea name="itemDesc" class="textarea" style="width:100%;height:80px;"><%= oMainContents.FOneItem.FitemDesc %></textarea></td>
</tr>
<% If poscode = "706" Then %>
<tr bgcolor="#FFFFFF">
  <td width="150" bgcolor="#DDDDFF">��� Ÿ��</td>
  <td><input type="radio" name="bannertype" value="1"<% If oMainContents.FOneItem.Fbannertype="1" Or oMainContents.FOneItem.Fbannertype="" Then Response.write " checked" %> onclick="fnSelectBannerType(1);">1��&nbsp;&nbsp;<input type="radio" name="bannertype" value="2"<% If oMainContents.FOneItem.Fbannertype="2" Then Response.write " checked" %> onclick="fnSelectBannerType(2);">2��</td>
</tr>
<% End If %>
<%
	'��ũ �ؽ�Ʈ ���� Ȯ��
	dim chkText: chkText="N"
	IF gubun<>"" Then
		if oMainContents.FOneItem.Fidx<>"" then
			if oMainContents.FOneItem.FLinkType="T" then chkText="Y"
		elseif poscode<>"" then
			if oposcode.FOneItem.FLinkType="T" then chkText="Y"
		end if
	end if
	'2013/09/28 ������ �߰� poscode ���
	If oMainContents.FResultCount > 0 Then
		Dim oSQL
		oSQL = " SELECT poscode FROM [db_sitemaster].[dbo].tbl_main_contents where idx = '"&oMainContents.FOneItem.Fidx&"'  "
		rsget.open oSQL, dbget, 1
			poscode = rsget("poscode")
		rsget.close
	End If
%>
<% IF chkText="Y" or (IsLinkTextNeed = True) then %>
<tr bgcolor="#FFFFFF">
  <td width="150" bgcolor="#DDDDFF"><%=chkIIF(poscode="630" or poscode="687","����","��ũ �ؽ�Ʈ")%></td>
  <td><input type="text" name="linkText" value="<%= oMainContents.FOneItem.FlinkText %>" size="32" maxlength="64" class="text" /> </td>
</tr>
<% if poscode="630" or poscode="687" then %>
<tr bgcolor="#FFFFFF">
  <td width="150" bgcolor="#DDDDFF">�ٹ����� �ΰ� ����</td>
  <td>
  	<label><input type="radio" name="linkText2" value="wht" <%=chkIIF(oMainContents.FOneItem.FlinkText2="wht" or oMainContents.FOneItem.FlinkText2="","checked","")%> />ȭ��Ʈ</label>
  	<label><input type="radio" name="linkText2" value="red" <%=chkIIF(oMainContents.FOneItem.FlinkText2="red","checked","")%> />����</label>
  </td>
</tr>
<tr bgcolor="#FFFFFF">
  <td width="150" bgcolor="#DDDDFF">��� ����</td>
  <td>
  	<label><input type="radio" name="linkText4" value="fix" <%=chkIIF(oMainContents.FOneItem.FlinkText4="fix" or oMainContents.FOneItem.FlinkText4="","checked","")%> />FIXED</label>
  	<label><input type="radio" name="linkText4" value="wide" <%=chkIIF(oMainContents.FOneItem.FlinkText4="wide","checked","")%> />WIDE</label>
  </td>
</tr>
<% else %>
<tr bgcolor="#FFFFFF">
  <td width="150" bgcolor="#DDDDFF">�߰� �ؽ�Ʈ #1 (����)</td>
  <td><input type="text" name="linkText2" value="<%= oMainContents.FOneItem.FlinkText2 %>" size="40" maxlength="128" class="text" /> </td>
</tr>
<tr bgcolor="#FFFFFF">
  <td width="150" bgcolor="#DDDDFF">�߰� �ؽ�Ʈ #2 (����)</td>
  <td><textarea name="linkText3" style="width:100%;height:60px;" class="text"><%= oMainContents.FOneItem.FlinkText3 %></textarea></td>
</tr>
<% end if %>
<%
	end if

	if chkText<>"Y" then
%>

	<% If poscode="688" Then %>
		<tr bgcolor="#FFFFFF">
		  <td width="150" bgcolor="#DDDDFF">��� Ÿ��Ʋ(bold)</td>
		  <td><input type="text" name="linkText2" value="<%= oMainContents.FOneItem.FlinkText2 %>" size="40" maxlength="128" class="text" /> </td>
		</tr>
		<tr bgcolor="#FFFFFF">
		  <td width="150" bgcolor="#DDDDFF">�ϴ� ��ǰ����</td>
		  <td><input type="text" name="linkText3" value="<%= oMainContents.FOneItem.FlinkText3 %>" size="40" maxlength="128" class="text" /></td>
		</tr>
		<tr bgcolor="#FFFFFF">
		  <td width="150" bgcolor="#DDDDFF">������</td>
		  <td><input type="text" name="linkText4" value="<%= oMainContents.FOneItem.FlinkText4 %>" size="40" maxlength="128" class="text" />
			<br>�� ������ �ۼ��� �ϴ� ��ǰ������ �������� ����
		</td>
		</tr>
	<% End If %>
	<% If poscode="689" Then %>
		<tr bgcolor="#FFFFFF">
		  <td width="150" bgcolor="#DDDDFF">Ÿ��Ʋ��</td>
		  <td><input type="text" name="linkText2" value="<%= oMainContents.FOneItem.FlinkText2 %>" size="40" maxlength="128" class="text" />
		  <br />�� �Է� ���ϸ� �⺻���� Just1Day�� �ָ�Ư�� ����<br/>�� ����Ư�� �� �Է��ϸ� ������ ������ ����Ư�� ���ڰ� ��µ�.
		  </td>
		</tr>
		<tr bgcolor="#FFFFFF">
		  <td width="150" bgcolor="#DDDDFF">�󼼼���</td>
		  <td><input type="text" name="linkText3" value="<%= oMainContents.FOneItem.FlinkText3 %>" size="40" maxlength="128" class="text" /></td>
		</tr>
	<% End If %>
	<% If poscode="690" Or poscode="691" Or poscode="692" Or poscode="693" Or poscode="699" Then %>
		<tr bgcolor="#FFFFFF">
		  <td width="150" bgcolor="#DDDDFF">��� Ÿ��Ʋ(bold)</td>
		  <td><input type="text" name="linkText2" value="<%= oMainContents.FOneItem.FlinkText2 %>" size="40" maxlength="128" class="text" /> </td>
		</tr>
		<tr bgcolor="#FFFFFF">
		  <td width="150" bgcolor="#DDDDFF">�ϴ� ��ǰ����</td>
		  <td><input type="text" name="linkText3" value="<%= oMainContents.FOneItem.FlinkText3 %>" size="40" maxlength="128" class="text" /></td>
		</tr>
	<% End If %>
<%'2018 ���� �Ѹ� %>
<% If poscode = "710" Then %>
<tr bgcolor="#FFFFFF">
  <td width="150" bgcolor="#DDDDFF">����</td>
  <td>
	�� : # <input type="text" name="linkText" value="<%= oMainContents.FOneItem.FlinkText %>" size="20" maxlength="6" class="text" /><br/>
	�� : # <input type="text" name="bgcode" value="<%=oMainContents.FOneItem.Fbgcode%>" size="20" maxlength="6" class="text">
  </td>
</tr>
<tr bgcolor="#FFFFFF">
  <td width="150" bgcolor="#DDDDFF">��� ����</td>
  <td>
  	<label><input type="radio" name="linkText4" value="fix" <%=chkIIF(oMainContents.FOneItem.FlinkText4="fix" or oMainContents.FOneItem.FlinkText4="","checked","")%> />FIXED</label>
  	<label><input type="radio" name="linkText4" value="wide" <%=chkIIF(oMainContents.FOneItem.FlinkText4="wide","checked","")%> />WIDE</label>
  </td>
</tr>
<tr bgcolor="#FFFFFF">
    <td width="150" bgcolor="#DDDDFF">��Ʈ�÷�����</td>
	<td>
		<input type="radio" name="xbtncolor" value="0" <%=chkiif(oMainContents.FOneItem.Fxbtncolor = "" Or oMainContents.FOneItem.Fxbtncolor = "0" ,"checked","")%>/> : black
		<input type="radio" name="xbtncolor" value="1" <%=chkiif(oMainContents.FOneItem.Fxbtncolor = "1" ,"checked","")%>/> : white
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td width="150" bgcolor="#DDDDFF">����ī��</td>
	<td>
		<input type="text" name="maincopy" value="<%=oMainContents.FOneItem.Fmaincopy%>" size="50" maxlength="25" class="text" /><br/>
		<input type="text" name="maincopy2" value="<%=oMainContents.FOneItem.Fmaincopy2%>" size="80" maxlength="60" class="text" />
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td width="150" bgcolor="#DDDDFF">����ī��</td>
	<td><input type="text" name="subcopy" value="<%=oMainContents.FOneItem.Fsubcopy%>" size="80" maxlength="50" class="text" /></td>
</tr>
<tr bgcolor="#FFFFFF">
	<td width="150" bgcolor="#DDDDFF">�±�</td>
	<td>
		<input type="radio" name="etctag" value="1" <%=chkiif(oMainContents.FOneItem.Fetctag="1" Or oMainContents.FOneItem.Fetctag="","checked","")%>> ���� <input type="radio" name="etctag" value="2" <%=chkiif(oMainContents.FOneItem.Fetctag="2","checked","")%>> ���� <input type="radio" name="etctag" value="3" <%=chkiif(oMainContents.FOneItem.Fetctag="3","checked","")%>> ���� <br/>
		<input type="radio" name="etctag" value="4" <%=chkiif(oMainContents.FOneItem.Fetctag="4","checked","")%>> GIFT <input type="radio" name="etctag" value="5" <%=chkiif(oMainContents.FOneItem.Fetctag="5","checked","")%>> 1+1 <input type="radio" name="etctag" value="6" <%=chkiif(oMainContents.FOneItem.Fetctag="6","checked","")%>> ��Ī <input type="radio" name="etctag" value="7" <%=chkiif(oMainContents.FOneItem.Fetctag="7","checked","")%>> ����<br/>
		<input type="text" name="etctext" value="<%=oMainContents.FOneItem.Fetctext%>" size="20" maxlength="30" class="text" />�� ����,���� �ϰ�츸 �Է� �ϼ���<br/>
		�� �Ѱ����� ���� �ϼ���.
	</td>
</tr>
<% End If %>
<%'2018 ���Ľ����̼�%>
<% If poscode="714" Then %>
<input type="hidden" name="ecode" value=""/><%' cultureidx %>
<tr bgcolor="#FFFFFF">
	<td width="150" bgcolor="#DDDDFF">����ī��</td>
	<td><input type="text" name="maincopy" value="<%=oMainContents.FOneItem.Fmaincopy%>" size="50" maxlength="25" class="text" /></td>
</tr>
<tr bgcolor="#FFFFFF">
	<td width="150" bgcolor="#DDDDFF">����ī��</td>
	<td><input type="text" name="subcopy" value="<%=oMainContents.FOneItem.Fsubcopy%>" size="80" maxlength="60" class="text" /></td>
</tr>
<tr bgcolor="#FFFFFF">
	<td width="150" bgcolor="#DDDDFF">����</td>
	<td><textarea name="linkText3" style="width:100%;height:60px;" class="text"><%= oMainContents.FOneItem.FlinkText3 %></textarea></td>
</tr>
<tr bgcolor="#FFFFFF">
    <td width="150" bgcolor="#DDDDFF">���м���</td>
	<td>
		<input type="radio" name="xbtncolor" value="0" <%=chkiif(oMainContents.FOneItem.Fxbtncolor = "" Or oMainContents.FOneItem.Fxbtncolor = "0" ,"checked","")%>/> : ������
		<input type="radio" name="xbtncolor" value="1" <%=chkiif(oMainContents.FOneItem.Fxbtncolor = "1" ,"checked","")%>/> : �о��
	</td>
</tr>
<% End If %>

<tr bgcolor="#FFFFFF">
  <td width="150" bgcolor="#DDDDFF">�̹���1</td>
  <td>
	<% If poscode <> "714" Then %>
	<input type="file" name="file1" value="" size="32" maxlength="32" class="file">
	<% End If %>
  <% if oMainContents.FOneItem.Fidx<>"" then %>
  <br>
  <img src="<%= oMainContents.FOneItem.GetImageUrl %>" style="max-width:600px;" />
  <br> <%= oMainContents.FOneItem.GetImageUrl %>
  <% end if %>
  <% '���Ľ����̼� %>
  <% If oMainContents.FOneItem.Fidx = "" And poscode = "714" Then %>
  <br>
  <img src="<%= webImgUrl &"/culturestation/2009/list/" & cultureEimagelist %>" style="max-width:600px;" />
  <br> <%= webImgUrl &"/culturestation/2009/list/" & cultureEimagelist %> <br/><br/> �� �̹��� ������ ���Ľ����̼� ���ο��� ���ּ���
  <% ElseIf oMainContents.FOneItem.Fidx <> "" And poscode = "714" Then %>
  <br>
  <img src="<%= webImgUrl &"/culturestation/2009/list/" & cultureEimagelist %>" style="max-width:600px;" />
  <br> <%= webImgUrl &"/culturestation/2009/list/" & cultureEimagelist %> <br/><br/> �� �̹��� ������ ���Ľ����̼� ���ο��� ���ּ���
  <% End If %>
  </td>
</tr>
<tr bgcolor="#FFFFFF">
  <td width="150" bgcolor="#DDDDFF">��Ʈ��1 (�ʼ�)</td>
  <td><input type="text" name="altname" value="<%=oMainContents.FOneItem.Faltname%>" size="20" maxlength="20"> </td>
</tr>
<% If poscode = "706" Then %>
<tr bgcolor="#FFFFFF" id="bnimg2" style="display:none">
  <td width="150" bgcolor="#DDDDFF">�̹���2</td>
  <td>
	<input type="file" name="file2" value="" size="32" maxlength="32" class="file">
  <% if oMainContents.FOneItem.GetImageUrl2<>"" then %>
  <br>
  <img src="<%= oMainContents.FOneItem.GetImageUrl2 %>" style="max-width:600px;" />
  <br> <%= oMainContents.FOneItem.GetImageUrl2 %>
  <% end if %>
  </td>
</tr>
<tr bgcolor="#FFFFFF" id="bnalt2" style="display:none">
  <td width="150" bgcolor="#DDDDFF">��Ʈ��2 (�ʼ�)</td>
  <td><input type="text" name="altname2" value="<%=oMainContents.FOneItem.Faltname2%>" size="20" maxlength="20"> </td>
</tr>
<% End If %>
<% If gubun <> "PCbanner" and gubun <> "MAbanner" And poscode <> "706" Then %>
<tr bgcolor="#FFFFFF">
  <td width="150" bgcolor="#DDDDFF">�߰� �̹��� (����)</td>
  <td><input type="file" name="file2" value="" size="32" maxlength="32" class="file">
  <% if oMainContents.FOneItem.Fidx<>"" then %>
  <br>
  <img src="<%= oMainContents.FOneItem.GetImageUrl2 %>" style="max-width:600px;" />
  <br> <%= oMainContents.FOneItem.GetImageUrl2 %>
  <% end if %>
  </td>
</tr>
<% End If %>
<tr bgcolor="#FFFFFF">
  <td width="150" bgcolor="#DDDDFF">�̹���Width</td>
  <td>
  	<% IF gubun <> "" Then %>
		  <% if oMainContents.FOneItem.Fidx<>"" then %>
		  <input type="text" name="imagewidth" value="<%= oMainContents.FOneItem.Fimagewidth %>" size="8" maxlength="16">
		  <% else %>
		        <% if poscode<>"" then %>
		        <%= oposcode.FOneItem.Fimagewidth %>
		        <% else %>
		        <font color="red">������ ���� �����ϼ���</font>
		        <% end if %>
		  <% end if %>
    <% Else %>
    	<font color="red">�׷챸���� ���� �����ϼ���</font>
    <% End If %>
  </td>
</tr>
<tr bgcolor="#FFFFFF">
  <td width="150" bgcolor="#DDDDFF">�̹���Height</td>
  <td>
  	<% IF gubun <> "" Then %>
		  <% if oMainContents.FOneItem.Fidx<>"" then %>
		  <input type="text" name="imageheight" value="<%= oMainContents.FOneItem.Fimageheight %>" size="8" maxlength="16">
		  <% else %>
		        <% if poscode<>"" then %>
		        <%= oposcode.FOneItem.Fimageheight %>
		        <% else %>
		        <font color="red">������ ���� �����ϼ���</font>
		        <% end if %>
		  <% end if %>
    <% Else %>
    	<font color="red">�׷챸���� ���� �����ϼ���</font>
    <% End If %>
  </td>
</tr>
<% End If %>
<tr bgcolor="#FFFFFF">
    <td width="150" bgcolor="#DDDDFF">��ũ��1</td>
    <td>
    	<% IF gubun <> "" Then %>
	        <% if oMainContents.FOneItem.Fidx<>"" then %>
	            <% if oMainContents.FOneItem.FLinkType="M" then %>
	            <textarea name="linkurl" style="width:100%;height:120px;"><%= oMainContents.FOneItem.Flinkurl %></textarea>
	            <% else %>
	            	<% if oMainContents.FOneItem.Fposcode = 539 Then%>
	            		<textarea name="linkurl" style="width:100%;height:120px;"><%= oMainContents.FOneItem.Flinkurl %></textarea>
            		<% Else%>
            			<input type="text" name="linkurl" value="<%= oMainContents.FOneItem.Flinkurl %>" maxlength="128" style="width:100%;" class="text">
            		<% End If %>
	            <% end if %>
	        <% else %>
	            <% if poscode<>"" then %>
	                <% if oposcode.FOneItem.FLinkType="M" then %>
	                    <textarea name="linkurl" style="width:100%;height:120px;"><%= defaultMapStr %></textarea>
	                    <br>(�̹����� ������ ���� ����)
	            	<% elseif oposcode.FOneItem.FLinkType="B" then %>
	            		<input type="text" class="text_ro" name="linkurl" value="/" maxlength="128" size="40" readonly>
					<% elseif poscode="539" Then %>
	                    <textarea name="linkurl" style="width:100%;height:120px;"><%= defaultXMLMapStr %></textarea>
	                    <br>(�̹����� ������ ���� ����, href���Ͽ� ��ũ�־��ּ���)
                	<% Else %>
	                    <input type="text" name="linkurl" value="" maxlength="128" style="width:100%;" class="text">
	                    <br>ex)<br/>
						- <span style="cursor:pointer" onClick="putLinkText('event');">�̺�Ʈ ��ũ : /event/eventmain.asp?eventid=<span style="color:darkred">�̺�Ʈ�ڵ�</span></span><br/>
						- <span style="cursor:pointer" onClick="putLinkText('itemid');">��ǰ�ڵ� ��ũ : /shopping/category_prd.asp?itemid=<span style="color:darkred">��ǰ�ڵ� (O)</span></span><br/>
						- <span style="cursor:pointer" onClick="putLinkText('category');">ī�װ� ��ũ : /shopping/category_list.asp?disp=<span style="color:darkred">ī�װ�</span></span><br/>
						- <span style="cursor:pointer" onClick="putLinkText('brand');">�귣����̵� ��ũ : /street/street_brand.asp?makerid=<span style="color:darkred">�귣����̵�</span></span><br/>
						- <span style="cursor:pointer" onClick="putLinkText('hitchhiker');">��ġ����Ŀ ��ũ : /hitchhiker/</span><br/>
						- <span style="cursor:pointer" onClick="putLinkText('giftcard');">����Ʈī�� ��ũ : /giftcard/</span><br/>
						- <span style="cursor:pointer" onClick="putLinkText('culture');">���Ľ����̼� ��ũ : /culturestation/culturestation_event.asp?evt_code=<span style="color:darkred">�̺�Ʈ���̵�</span></span><br/>
	                <% end if %>
	            <% else %>
	            <font color="red">������ ���� �����ϼ���</font>
	            <% end if %>
	        <% end if %>
	    <% Else %>
	    	<font color="red">�׷챸���� ���� �����ϼ���</font>
	    <% End If %>
    </td>
</tr>
<% If poscode = "706" Then %>
<tr bgcolor="#FFFFFF" id="bnlink2" style="display:none">
    <td width="150" bgcolor="#DDDDFF">��ũ��2</td>
    <td>
		<input type="text" name="linkurl2" value="<%= oMainContents.FOneItem.Flinkurl2 %>" maxlength="128" style="width:100%;" class="text">
    </td>
</tr>
<tr bgcolor="#FFFFFF">
    <td width="150" bgcolor="#DDDDFF">�¿��� BG�÷��ڵ�</td>
	<td><span  id="bnbg1" style="display:none">�� : </span>#<input type="text" name="bgcode" value="<%=oMainContents.FOneItem.Fbgcode%>" size="20" maxlength="6">
		<div  id="bnbg2" style="display:none">�� : #<input type="text" name="bgcode2" value="<%=oMainContents.FOneItem.Fbgcode2%>" size="20" maxlength="6"></div>
	</td>
</tr>
<tr bgcolor="#FFFFFF">
    <td width="150" bgcolor="#DDDDFF">X��ư����</td>
	<td>
		<input type="radio" name="xbtncolor" value="0" <%=chkiif(oMainContents.FOneItem.Fxbtncolor = "" Or oMainContents.FOneItem.Fxbtncolor = "0" ,"checked","")%>/> : ȭ��Ʈ
		<input type="radio" name="xbtncolor" value="1" <%=chkiif(oMainContents.FOneItem.Fxbtncolor = "1" ,"checked","")%>/> : black
	</td>
</tr>
<% End If %>
<tr bgcolor="#FFFFFF">
    <td width="150" bgcolor="#DDDDFF">�ݿ�������</td>
    <td>
        <input id="startdate" name="startdate" value="<%= Left(oMainContents.FOneItem.Fstartdate,10) %>" class="text" size="10" maxlength="10" />
        <img src="http://webadmin.10x10.co.kr/images/calicon.gif" id="startdate_trigger" border="0" style="cursor:pointer;" align="absbottom" />
        <% if oMainContents.FOneItem.Ffixtype="R" or poscode="687" Or gubun = "PCbanner" Or gubun = "MAbanner" then %> <!-- �ǽð��ΰ�� / �� �ϴ����� ���� (���߿� �ð������� ������ False ����)-->
        <input type="text" name="startdatetime" size="2" maxlength="2" value="<%= Format00(2,Hour(oMainContents.FOneItem.Fstartdate)) %>" />(�� 00~23)
        <input type="text" name="dummy0" value="00:00" size="6" readonly class="text_ro" />
        <% else %>
        <input type="text" name="dummy0" value="00:00:00" size="8" readonly class="text_ro" />
        <% end if %>
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
    <td width="150" bgcolor="#DDDDFF">�ݿ�������</td>
    <td>
        <input id="enddate" name="enddate" value="<%= Left(oMainContents.FOneItem.Fenddate,10) %>" class="text" size="10" maxlength="10" />
        <img src="http://webadmin.10x10.co.kr/images/calicon.gif" id="enddate_trigger" border="0" style="cursor:pointer" align="absbottom" />
        <% if oMainContents.FOneItem.Ffixtype="R"  or poscode="687" Or gubun = "PCbanner" Or gubun = "MAbanner" then %> <!-- �ǽð��ΰ�� -->
        <input type="text" name="enddatetime" size="2" maxlength="2" value="<%= ChkIIF(oMainContents.FOneItem.Fenddate="","23",Format00(2,Hour(oMainContents.FOneItem.Fenddate))) %>">(�� 00~23)
        <input type="text" name="dummy1" value="59:59" size="6" readonly class="text_ro" />
        <% else %>
        <input type="text" name="dummy1" value="23:59:59" size="8" readonly class="text_ro" />
        <% end if %>
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
    <td width="150" bgcolor="#DDDDFF">�����</td>
    <td>
        <%= oMainContents.FOneItem.Fregdate %> (<%= oMainContents.FOneItem.Fregname %>)
    </td>
</tr>
<tr bgcolor="#FFFFFF">
    <td width="150" bgcolor="#DDDDFF">�۾���</td>
    <td>
    	<% If idx <> "" AND idx <> "0" Then %>
    	���� �۾��� : <%=oMainContents.FOneItem.Fworkername%><input type="hidden" name="selDId" value="<%=session("ssBctId")%>">
    	&nbsp;<strong><%=oMainContents.FOneItem.Flastupdate%></strong>
    	<% Else %>
    		<input type="hidden" name="selDId" value="">
    	<% End If %>
    </td>
</tr>
<tr bgcolor="#FFFFFF">
    <td width="150" bgcolor="#DDDDFF">��뿩��</td>
    <td>
        <% if oMainContents.FOneItem.Fisusing="N" then %>
        <input type="radio" name="isusing" value="Y">�����
        <input type="radio" name="isusing" value="N" checked >������
        <% else %>
        <input type="radio" name="isusing" value="Y" checked >�����
        <input type="radio" name="isusing" value="N">������
        <% end if %>
    </td>
</tr>
<tr bgcolor="#FFFFFF">
    <td colspan="2" align="center"><input type="button" value=" �� �� " onClick="SaveMainContents(frmcontents);"></td>
</tr>
</form>
</table>
<%
set oposcode = Nothing
set oMainContents = Nothing
%>

<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
