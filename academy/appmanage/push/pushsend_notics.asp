<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' PageName : pushsend_notics.asp
' Description : �������� Ǫ�� �߼�
' Hieditor : 2016.11.30 ������
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/academy/lib/classes/board/lecturer/lecturer_cls.asp"-->
<!-- #include virtual="/academy/lib/classes/appmanage/fingerpush_msg_cls.asp" -->
<%
    Dim iDoc_Idx '' �������� IDX
    iDoc_Idx = requestCheckvar(request("idoc_idx"),10)
    Dim olect,sDoc_Id,sDoc_Name,sDoc_Status,sDoc_Type,sDoc_Import,sDoc_Subj,sDoc_Content,sDoc_UseYN,sDoc_Regdate,sDoc_part_sn,sDoc_admin_usingyn, IsPushSended
    Dim mode, testlecid
    
    Set olect = New clecturer_list
	olect.FrectDoc_Idx = iDoc_Idx
	olect.FRECTAdmin_UsingNInclude = "on"
	olect.fnGetlecturerView

    if (olect.FREsultCount>0) then
    	sDoc_Id 		= olect.FOneItem.FDoc_Id
    	sDoc_Name		= olect.FOneItem.FDoc_Name
    	sDoc_Status		= olect.FOneItem.FDoc_Status
    	if sDoc_Status = "" then sDoc_Status = "K001"	
    	sDoc_Type		= olect.FOneItem.FDoc_Type
    	sDoc_Import		= olect.FOneItem.FDoc_Import
    	sDoc_Subj		= olect.FOneItem.FDoc_Subj
    	sDoc_Content	= olect.FOneItem.FDoc_Content
    	sDoc_UseYN		= olect.FOneItem.FDoc_UseYN
    	sDoc_Regdate	= olect.FOneItem.FDoc_Regdate
    	sDoc_part_sn	= olect.FOneItem.fpart_sn
        sDoc_admin_usingyn    = olect.FOneItem.fadmin_usingyn
        
        IsPushSended = olect.FOneItem.IsPushSended
    end if
    
    Dim stitle, subtitle
	stitle = "[����] "& replace(TRIM(stripHTML(sDoc_Subj)),"[����]","")
	subtitle = "https://webadmin.10x10.co.kr/apps/academy/notice/noticeView.asp?pmode=pms&idx="&iDoc_Idx
	
%>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript" src="/js/jqueryui/jquery-ui-1.10.2.custom.min.js"></script>
<script type="text/javascript">
    function popviw(icomp){
        var popwin;
        var iurl = icomp.value;
		popwin = window.open(iurl,'popwin11','width=600, height=800');
		popwin.focus();
    }
    
    
	function jsgolist(){
		opener.location.reload();
		self.close();
	}

	function jsPopCal(sName){
		var winCal;

		winCal = window.open('/lib/common_cal.asp?DN='+sName,'pCal','width=250, height=200');
		winCal.focus();
	}

	//�̹��� Ȯ��ȭ�� ��â���� �����ֱ� ------ ���� �߰� �ɶ� �۾�
	function jsImgView(sImgUrl){
	 var wImgView;
	 wImgView = window.open('/admin/sitemaster/play/lib/pop_event_detailImg.asp?sUrl='+sImgUrl,'pImg','width=100,height=100');
	 wImgView.focus();
	}

	function jsDelImg(sName, sSpan){
		if(confirm("�̹����� �����Ͻðڽ��ϱ�?\n\n���� �� �����ư�� ������ ó���Ϸ�˴ϴ�.")){
		   eval("document.all."+sName).value = "";
		   eval("document.all."+sSpan).style.display = "none";
		}
	}

	function jsSetImg(sImg, sName, sSpan){	
		document.domain ="10x10.co.kr";
		
		var winImg;
		winImg = window.open('/admin/mobile/lib/pop_uploadimg.asp?sImg='+sImg+'&sName='+sName+'&sSpan='+sSpan,'popImg','width=370,height=150');
		winImg.focus();
	}

	function subcheck(){
		var frm=document.inputfrm;
        /*
        if ((frm.istargetMsg[1].checked)&&(frm.targetKey.value.length<1)){
            alert('Ÿ���� �������ּ���');
			frm.targetKey.focus();
			return;
        }
        */
		if (!frm.stitle.value){
			alert('Ǫ�� �޽����� ������ּ���');
			frm.stitle.focus();
			return;
		}

		if (!frm.subtitle.value){
			alert('��ũ�� ������ּ���');
			frm.subtitle.focus();
			return;
		}

        /*
		if (!frm.time1.value){
			alert('���� �ð��� ������ּ���');
			frm.reservationdate.focus();
			return;
		}

		if (!frm.time2.value){
			alert('���� ���� ������ּ���');
			frm.time2.focus();
			return;
		}

		if(!frm.state.value){
			alert("���¸� �������ּ���");
			frm.state.focus();
			return;
		}
		*/
		
		//frm.target="_blank";
		if (confirm('�߼��Ͻðڽ��ϱ�?')){
		    frm.submit();
	    }
	}

	function chgstate(v){
		var frm = document.frmstate;
		if ( v == "I" ){
			frm.state.value = 1;
		}else{
			frm.state.value = 0;
		}

		frm.target = "FrameCKP";
        frm.submit();
	}

	function putLinkText(key) {
		var frm = document.inputfrm;
		var urllink = frm.subtitle;
		switch(key) {
			case 'event':
				urllink.value='http://m.10x10.co.kr/apps/appCom/wish/web2014/event/eventmain.asp?eventid=�̺�Ʈ��ȣ';
				break;
			case 'itemid':
				urllink.value='http://m.10x10.co.kr/apps/appCom/wish/web2014/category/category_itemprd.asp?itemid=��ǰ�ڵ�';
				break;
			case 'etc':
				urllink.value='http://m.10x10.co.kr/apps/appCom/wish/web2014/';
				break;
		}
	}

	function chgusing(){
		var frm = document.frmdel

		frm.target = "FrameCKP";
		frm.submit();
	}
	
	function setComp(comp){
	    if (comp.name=="istargetMsg"){
    	    if (comp.value=="1"){
    	        document.getElementById("itargetcmt").style.display="inline";
    	    }else{
    	        document.getElementById("itargetcmt").style.display="none";
    	    }
    	}
    	
    	if (comp.name=="targetKey"){
    	    if (comp.value>1){
    	        document.getElementById("baseIdx").style.display="inline";
    	    }else{
    	        document.getElementById("baseIdx").style.display="none";
    	    }
    	}
	}
	
	function acttarget(){
	    if (!confirm('Ÿ������ �ۼ��Ͻðڽ��ϱ�?')) return;
		var frm = document.frmtarget
        frm.mode.value="target";
        
		frm.target = "FrameCKP";
		frm.submit();
	}
	
	function retarget(){
	    if (!confirm('Ÿ������ (��) �ۼ��Ͻðڽ��ϱ�?')) return;
		var frm = document.frmtarget
        frm.mode.value="retarget";
        
		frm.target = "FrameCKP";
		frm.submit();
	}
	
	function fnChgTestTarget(comp){
	    var lecid = comp.value;
	    if (lecid.length<1) return;
	    $.ajax({
    		url: "getTestPushid.asp?lecid="+lecid+"&selectBoxName=testpushid",
    		cache: false,
    		async: false,
    		success: function(message) {
           		// ���� �ֱ� 
           		if (message.length>0){
           		    message+=("&nbsp;&nbsp;<input type='button' value='�׽�Ʈ �߼�' onClick='sendTestPush()'>");
           	    }
           	    
           		$("#divtestpushid").empty().html(message);
           		
           		
    		}
    	});
	}
	
	function sendTestPush(){
	    var frm = document.inputfrm;
	    if (frm.testlecid.value.length<1){
	        alert('�߼��� �׽�Ʈ ���̵� �����ϼ���.');
	        frm.testlecid.focus();
	        return;   
	    }
	    
	    
	    if (frm.testpushid.value.length<1){
	        alert('�߼��� pushid�� �����ϼ���.');
	        frm.testpushid.focus();
	        return;   
	    }
	    
	    if (!frm.stitle.value){
			alert('Ǫ�� �޽����� ������ּ���');
			frm.stitle.focus();
			return;
		}

		if (!frm.subtitle.value){
			alert('��ũ�� ������ּ���');
			frm.subtitle.focus();
			return;
		}
		
		var testappkey = frm.testpushid.value.split("|")[0];
		var testpushid = frm.testpushid.value.split("|")[1];
		
			
		document.frmtarget.appkey.value=testappkey;
	    document.frmtarget.deviceid.value=testpushid;
	    document.frmtarget.stitle.value=frm.stitle.value;
	    document.frmtarget.subtitle.value=frm.subtitle.value;
	    document.frmtarget.testlecid.value=frm.testlecid.value;
	    document.frmtarget.target="FrameCKP";
	    
	    
	    if (confirm('�׽�Ʈ �߼� �Ͻðڽ��ϱ�?')){
	        document.frmtarget.submit();
	    }
	}
</script>


<form name="frmtarget" method="post" action="pushmsg_proc.asp">
<input type="hidden" name="idoc_idx" value="<%= idoc_idx %>">
<input type="hidden" name="appkey" value="">
<input type="hidden" name="deviceid" value="">
<input type="hidden" name="subtitle" value="">
<input type="hidden" name="stitle" value="">
<input type="hidden" name="testlecid" value="">
<input type="hidden" name="multiPsKey" value="0"><!-- ����1 -->
<input type="hidden" name="mode" value="testsendnoti">
</form>

<iframe name="FrameCKP" src="about:blank" frameborder="0" width="310" height="110" border="1"></iframe>

<form name="inputfrm" method="post" action="pushmsg_proc.asp">
<table width="100%" cellpadding="5" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<input type="hidden" name="idoc_idx" value="<%= idoc_idx %>">
<input type="hidden" name="mode" value="realsendnoti">
<tr height="30">
	<td colspan="4" bgcolor="#FFFFFF">
		<img src="/images/icon_star.gif" align="absmiddle"/>
		<font color="red"><b>Fingers ��ü �������� Ǫ�� �߼�</b></font><br/><br/>
	</td>
</tr>
<% If idoc_idx <> "0" Then %>
<tr>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">�����۹�ȣ</td>
	<td colspan="3" bgcolor="#FFFFFF">
		<b><%=idoc_idx%></b>
	</td>
</tr>
<tr>
	<td width="150" align="center" bgcolor="<%= adminColor("tabletop") %>">��������</td>
	<td colspan="3" bgcolor="#FFFFFF">
	    <%=sDoc_Subj%>
	</td>
</tr>
<tr height="60">
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">Ǫ�� �޽���</td>
	<td colspan="3" bgcolor="#FFFFFF">
		<input type="text" class="text" name="stitle" value="<%=stitle%>" size="90"/>
		<% if (FALSE) then %>
		    <br><font color=gray>�޼����� '(����) '�� ������ �ֽñ� �ٶ��ϴ�. 2014/12</font>
		<% end if %>
	</td>
</tr>

<% End If %>



<tr height="60">
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">��ũ</td>
	<td colspan="3" bgcolor="#FFFFFF">
		<input type="text" class="text" name="subtitle" value="<%=subtitle%>" size="90"/> <input type="button" value="����" onClick="popviw(this.form.subtitle);"><br/>
		Ǯ �ּҷ� �Է�<br/>
		<% if (FALSE) then %>
    		ex)<br>
    		<font color="#707070">
    		- <span style="cursor:pointer" onClick="putLinkText('admnoti')">���� ���� ��ũ : <font color="darkred">���� ���� ��ũ</font></span><br>
    		- <span style="cursor:pointer" onClick="putLinkText('event')">�̺�Ʈ ��ũ : /event/eventmain.asp?eventid=<font color="darkred">�̺�Ʈ�ڵ�</font></span><br>
    		- <span style="cursor:pointer" onClick="putLinkText('itemid')">��ǰ�ڵ� ��ũ : /category/category_itemprd.asp?itemid=<font color="darkred">��ǰ�ڵ� (O)</font></span><br>
    		- <span style="cursor:pointer" onClick="putLinkText('etc')">��Ÿ ��ũ : <font color="darkred">��Ÿ</font></span><br>
	    <% end if %>
	</td>
</tr>
<tr>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">�׽�Ʈ�߼�</td>
	<td colspan="3" bgcolor="#FFFFFF">
		<% call drawSelectBoxTestTarget("testlecid",testlecid,"onchange='fnChgTestTarget(this);'") %>
		<br>
		<div id="divtestpushid" name="divtestpushid" ></div>
	    
	</td>
</tr>
<tr>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">�߼۴��</td>
	<td colspan="3" bgcolor="#FFFFFF">
	    <input type="radio" name="targetgbn" value="D" checked>����̽��� ( <%= getNoticsPushTargetCount(idoc_idx) %> �� )
	    <p>
	    <input type="radio" name="targetgbn" value="U">������̵� ��������̽�(ios,adorid ����) ( <%= getNoticsPushTargetCountLastUser(idoc_idx) %> �� )
		
	</td>
</tr>


<tr bgcolor="#FFFFFF">
	<td colspan="4" align="center">
	    <% if (IsPushSended) then %>
	    
	    <% else %>
		<input type="button" value=" �߼� " class="button" onclick="subcheck();"/> &nbsp;&nbsp;
        <% end if %>
		<input type="button" value=" ���(�ݱ�) " class="button" onclick="self.close();"/>
	</td>
</tr>
</form>
</table>
<%
'	set oSubItemList = Nothing
%>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->
