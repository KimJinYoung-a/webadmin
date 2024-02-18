<%@ language=vbscript %>
<% option explicit 
	Response.Expires = -1440
	Response.CacheControl = "no-cache" 
	Response.AddHeader "Pragma", "no-cache" 
%>
<%
'####################################################
' Description :  ����ǰ ���� ���
' History : 2010.09.27 �ѿ�� ����
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/academy/lib/academy_function.asp"-->
<!-- #include virtual="/academy/lib/classes/gift/giftcls.asp"-->

<%
Dim clsGift, sViewMode, sMode ,arrList, intLoop ,strTxt,strImg,iitemid,igkCode 
	strTxt = requestCheckVar(Request("sGKN"),32) 
	sViewMode  = requestCheckVar(Request("sVM"),32) 	
	IF sViewMode = "" THEN sViewMode = -1
	sMode = "KI"

' �˻��Ϸ��� ����ǰ ���� ���� �� �ش� ����Ʈ �����ش�.
IF strTxt <> "" AND sViewMode < 0 THEN 		
	set clsGift = new CGift
		clsGift.FSearchTxt = strTxt
		arrList = clsGift.fnGetGiftKind
	set clsGift = nothing
END IF

IF sViewMode > 0 THEN
	set clsGift = new CGift
		sMode = "KU"
		igkCode = sViewMode
		clsGift.FGKindCode = igkCode
		clsGift.fnGetGiftKindConts
		strTxt = clsGift.FGKindName	   
		strImg = clsGift.FGKindImg      		
		iitemid= clsGift.FItemid        
	set clsGift = nothing
END IF	

Dim eFolder : eFolder =   igkCode
%> 

<script language="javascript">
	
	// �˻�
	function jsSearch(){
		if(!document.frmSearch.sGKN.value){
			alert("����ǰ�������� �Է����ּ���");
			return;
		}
		
		document.frmSearch.submit();
	}

	
	// ��� �Ǵ� �˻� ȭ������ ����
	function jsChangeMode(sViewMode){
		if (sViewMode ==""){
		document.frmSearch.sGKN.value="";
		}
		document.frmSearch.sVM.value = sViewMode;
		document.frmSearch.submit();
	}
	
	// ����ǰ �������	
	function jsSubmitGiftKind(){
		var frm = document.frmGift;
		if(!frm.sGKN.value){
			 alert("����ǰ�������� �Է����ּ���");
			 frm.sGKN.focus();
			 return false;
		}
			
		return;
	}
	
	//�˻��� ����ǰ���� ����
	function jsSetGiftKind(igk, skn,strImg){	
		opener.document.all.iGK.value = igk;
		opener.document.all.sGKN.value= skn;
		if(strImg !=""){
		opener.document.all.spanImg.innerHTML = "<a href=javascript:jsImgView('"+strImg+"')><img src='"+strImg+"' border=0></a>";		
		}
		window.close();
	}
	
	//-- jsImgView : �̹��� Ȯ��ȭ�� ��â���� �����ֱ� --//
	function jsImgView(sImgUrl){
	 var wImgView;
	 wImgView = window.open('/lib/showimage.asp?img='+sImgUrl,'pImg','width=100,height=100');
	 wImgView.focus();
	}

	function fnAddImage2(strImg,sName,sSpan){
		document.domain ="10x10.co.kr";
		eval("document.frmGift." + sName).value = strImg;		
		eval("document.all." + sSpan ).innerHTML = "<img src='"+strImg+"' border=0 width='60' height='30'>";
	}
    
    function jsSetImg2(sFolder, sImg, sName, sSpan){
    	document.domain ="10x10.co.kr";
    	var winImg;
    	
    	winImg = window.open('popgiftkindupload.asp?sF='+sFolder+'&sImg='+sImg+'&sName='+sName+'&sSpan='+sSpan,'popImg','width=370,height=150');
		winImg.focus();
		
    	//winImg = window.open('/admin/eventmanage/common/pop_event_uploadimg.asp?sF='+sFolder+'&sImg='+sImg+'&sName='+sName+'&sSpan='+sSpan,'popImg','width=370,height=150');
    	//winImg.focus();
    }
    
	function jsSetImg(){		
		document.domain ="10x10.co.kr";
		var winImg;		
		winImg = window.open('popgiftkindupload.asp','popImg','width=370,height=150');
		winImg.focus();
	}
	
 	
	function fnAddImage(strImg){
		document.domain ="10x10.co.kr";
		document.frmGift.sGKImg.value = strImg;		
		document.all.spanImg.innerHTML = "<img src='"+strImg+"' border=0 width='60' height='30'>";
	}
    
    function jsDelImg(sName, sSpan){
    	if(confirm("�̹����� �����Ͻðڽ��ϱ�?\n\n���� �� �����ư�� ������ ó���Ϸ�˴ϴ�.")){
    	   eval("document.all."+sName).value = "";
    	   eval("document.all."+sSpan).style.display = "none";
    	}
    }

</script>

<div style="padding: 0 5 5 5"> <img src="/images/icon_arrow_link.gif" align="absmiddle"> ����ǰ���� ���</div>
<table width="430" border="0" align="left" class="a" cellpadding="3" cellspacing="0" >
<form name="frmSearch" method="get" action="popgiftKindReg.asp" >
<input type="hidden" name="sVM" >
<tr>
	<td>����ǰ������ : <input type="text" name="sGKN" size="30" maxlength="60" value="<%=strTxt%>"> 
		<input type="button" class="button" value="�˻�" onClick="jsSearch();">
		<input type="button" class="button" value="���ε��" onClick="jsChangeMode('0');">
	</td>
</form>	
</tr>
<tr>
	<td><hr wudth="100%"></td>
</tr>
<tr>
	<td> 
		<table width="100%" border="0" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">	
		<%IF isArray(arrList) THEN %>		
			<tr bgcolor="<%= adminColor("tabletop") %>">
			<td align="center">�ڵ��ȣ</td>
			<td align="center">����ǰ������</td>			
			<td align="center">itemid</td>
			<td align="center">�̹���</td>
			<td align="center">�����</td>
			<td align="center">ó��</td>
		</tr>	
		<%	
		For intLoop =0 To UBound(arrList,2)					
		%>
		<tr bgcolor="#FFFFFF">
			<td align="center"><a href="javascript:jsChangeMode('<%=arrList(0,intLoop)%>')" title="����ǰ���� �������"><%=arrList(0,intLoop)%></a></td>
			<td align="center"><a href="javascript:jsChangeMode('<%=arrList(0,intLoop)%>')" title="����ǰ���� �������"><%=arrList(1,intLoop)%></a></td>			
			<td align="center"><a href="javascript:jsChangeMode('<%=arrList(0,intLoop)%>')" title="����ǰ���� �������"><%=arrList(3,intLoop)%></a></td>
			<td align="center"><%IF arrList(2,intLoop) <> "" THEN%><a href="javascript:jsImgView('<%=arrList(2,intLoop)%>')" title="�̹��� Ȯ�뺸��"><img src="<%=arrList(2,intLoop)%>" width="60" height="30" border="0"></a><%END IF%></td>
			<td align="center"><a href="javascript:jsChangeMode('<%=arrList(0,intLoop)%>')" title="����ǰ���� �������"><%=FormatDate(arrList(4,intLoop),"0000.00.00")%></a></td>
			<td align="center"><input type="button" value="���" class="button" onClick="jsSetGiftKind(<%=arrList(0,intLoop)%>,'<%=arrList(1,intLoop)%>','<%=arrList(2,intLoop)%>');"></td>
		</tr>
		<% Next	%>				
		
		<%ELSE%>	

		<%IF sViewMode = -1 AND strTxt <> "" THEN %>
		<tr><td colspan="2"  bgcolor="#FFFFFF"><font color="#E08050"><%=strTxt%></font>�� �ش��ϴ� ����ǰ ������ �����ϴ�. ���� ����� �ּ���</td></tr>		
		<%END IF%>	
		<form name="frmGift" method="post" action="giftProc.asp" onSubmit="return jsSubmitGiftKind();">
		<input type="hidden" name="sM" value="<%=sMode%>">
		<input type="hidden" name="sGKImg" value="<%=strImg%>">
		<input type="hidden" name="iGK" value="<%=igkCode%>">	
		<tr>
			<td align="center" width="100" bgcolor="<%= adminColor("tabletop") %>">����ǰ�ڵ�</td>
			<td bgcolor="#FFFFFF"><%=igkCode%></td>
		</tr>			
		<tr>
			<td align="center" width="100" bgcolor="<%= adminColor("tabletop") %>">����ǰ������</td>
			<td bgcolor="#FFFFFF"><input type="text" name="sGKN" size="40" maxlength="60" value="<%=strTxt%>"></td>
		</tr>		
		<tr>	
			<td align="center" bgcolor="<%= adminColor("tabletop") %>">itemid</td>
			<td bgcolor="#FFFFFF"><input type="text" name="itemid" size="10" value="<%=iitemid%>"></td>
		</tr>
		<tr>	
			<td align="center" bgcolor="<%= adminColor("tabletop") %>">�̹���<br>(�̺�Ʈ�� ����ǰ)</td>
			<td bgcolor="#FFFFFF">
			    <input type="button" class="button" value="�̹������" onClick="jsSetImg2('<%=eFolder%>','<%=strImg%>','sGKImg','spanImg');" >
			    <div id="spanImg">
			    <%IF strImg <> "" THEN%>
			    <a href="javascript:jsImgView('<%=strImg%>');"><img src="<%=strImg%>" width="60" height="30" border="0"></a>
			    <a href="javascript:jsDelImg('sGKImg','spanImg');"><img src="/images/icon_delete2.gif" border="0"></a>
			    <%END IF%>
			    </div>
				
		    </td>
		</tr>
		<tr>
			<td colspan="2" bgcolor="#FFFFFF" align="right"><input type="image" src="/images/icon_confirm.gif">
				<!--<a href="javascript:history.back(0);"><img src="/images/icon_cancel.gif" border="0"></a>-->
			</td>
		</tr>
		</form>		
		<%END IF%>
		</table>	
	</td>
</tr>
</table>

<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->