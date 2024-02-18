<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Page : /admin/eventmanage/event/pop_eventitem_groupImage.asp
' Description :  �̺�Ʈ �׷� �̹��� ����
' History : 2007.02.22 ������ ����
'			2015.02.12 ������ ����
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/event_function.asp"-->
<!-- #include virtual="/lib/classes/event/eventManageCls_V2.asp"--> 
<%
Dim eCode : eCode = requestCheckVar(Request("eC"),10)
Dim eGCode : eGCode = requestCheckVar(Request("eGC"),10) 
Dim vYear : vYear = requestCheckVar(Request("yr"),4)
dim eChannel : eChannel = requestCheckVar(Request("eCh"),1)
Dim cEGroup, arrP,intP,sM
Dim gpcode, gdesc, gsort, gimg,gdepth,gpdesc,glink, gdisp
Dim arrImg, slen, sImgName

 gdisp = True
 
 set cEGroup = new ClsEventGroup
 	cEGroup.FECode = eCode
 	cEGroup.FEChannel = eChannel
 	cEGroup.FGDisp = 1
  	arrP = cEGroup.fnGetRootGroup
  	sM = "I"
  	IF (eGCode <> "" and eGCode <> "0" and not isnull(eGCode)) THEN
	  	cEGroup.FEGCode = eGCode 
	  	cEGroup.fnGetEventItemGroupCont		
	  	gpcode 	= cEGroup.FGPCode
	  	gdesc  	= cEGroup.FGDesc
	  	gsort	= cEGroup.FGSort
	  	gdepth	= cEGroup.FGDepth
	  	gpdesc  = cEGroup.FGPDesc 
		gimg	= cEGroup.FGImg  
	  	glink	= cEGroup.FGlink
		gdisp  = cEGroup.FGDisp 
	  	sM = "U"
	END IF  	
 set cEGroup = nothing
 
 
IF gimg <> "" THEN
	arrImg = split(gimg,"/")
	slen = ubound(arrImg)
	sImgName = arrImg(slen)
END IF	

if gsort = "" then gsort = 0

%>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script> 
<script>
$(function(){
    // â ��������� testarea ���� ����
    $(window).resize(function() { 
        $('#tGMap').css('height', $(window).height()-340); 
    }); 
});
</script>
<div style="padding: 0 5 5 5"> <img src="/images/icon_arrow_link.gif" align="absmiddle"> �̺�Ʈ �׷� �̹��� ���</div>
<form name="frmG" method="post" action="<%= uploadImgUrl %>/linkweb/event_admin/V2/eventgroupImg_process.asp" enctype="MULTIPART/FORM-DATA"   >
<input type="hidden" name="eC" value="<%=eCode%>">
<input type="hidden" name="eGC" value="<%=eGCode%>">
<input type="hidden" name="mode" value="<%=sM%>">
<input type="hidden" name="sOGimg" value="<%=gimg%>">
<input type="hidden" name="yr" value="<%=vYear%>">
<input type="hidden" name="eCh" value="<%=eChannel%>">
<table width="100%" border="0" align="center" class="a" cellpadding="0" cellspacing="0">
<tr>
	<td>
		<table width="100%" border="0" align="left" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
			<tr> 
				<%IF eChannel ="M" then%>
				<td bgcolor="#e3f1fb" align="center"  colspan="2"><b>Mobile / App</b></td>
				<%ELSE%>
				<td bgcolor="#FAECC5" align="center" colspan="2"><b>PC-WEB</b></td>
				<%END IF%>
			</tr>
			<tr>
				<td align="center" bgcolor="<%= adminColor("tabletop") %>">�����׷�</td>
				<td bgcolor="#FFFFFF"> 
			<%IF gdepth = "" THEN%>
				<select name="selPC">
				<option value="0">�ֻ���</option>
				<%IF isArray(arrP) THEN
					For intP =0 To UBound(arrP,2)
					%>
				<option value="<%=arrP(0,intP)%>" <%IF Cstr(gpcode) = CStr(arrP(0,intP)) THEN%>selected<%END IF%>><%=arrP(1,intP)%></option>	
			<%  Next
				END IF%>	
				</select>
				<%ELSE%>
				<input type="hidden" name="selPC" value="<%=gpcode%>">
				<%=gpdesc%>
				<%END IF%>
				</td>
			</tr>
			<tr>
				<td width="100" align="center" bgcolor="<%= adminColor("tabletop") %>">�׷��</td>
				<td bgcolor="#FFFFFF"><input type="text" name="sGD" value="<%=db2html(gdesc)%>" maxlength="32" size="40"></td>
			</tr>		
			<tr>
				<td align="center" bgcolor="<%= adminColor("tabletop") %>">���ļ���</td>
				<td bgcolor="#FFFFFF"><input type="text" size="2" name="sGS"  value="<%=gsort%>"></td>
			</tr> 
			<tr>
				<td align="center" bgcolor="<%= adminColor("tabletop") %>">���ÿ���</td>
				<td bgcolor="#FFFFFF"><input type="radio" name="eIsDisp" value="1" <%if gdisp then%>checked<%end if%>>Y <input type="radio" name="eIsDisp" value="0" <%if not gdisp then%>checked<%end if%>>N </td>
			</tr> 
			<tr>
				<td align="center" bgcolor="<%= adminColor("tabletop") %>">�̹���</td>
				<td bgcolor="#FFFFFF"><input type="file" name="sGimg"><br><%IF gimg <> "" THEN%><%=sImgName%> <input type="checkbox" name="delI">����<%END IF%></td>
			</tr>
			<tr>
				<td align="center" bgcolor="<%= adminColor("tabletop") %>">�ʸ�ũ</td>
				<td bgcolor="#FFFFFF">				
					<font color="red">+ ��! ���ڵ��� �ʸ�Ī�� �� &lt;area shape="rect" ~ �� �Է����ּ���. </font><br>
					<font color="blue">�̺�Ʈ �׷� �������� ��ũ��<br>
					&lt;area shape="rect" coords="0,0,0,0" onclick="TnGotoEventGroupMain('<font color="blue">�̺�Ʈ�ڵ�</font>','<font color="blue">�׷��ڵ�</font>');" onfocus="this.blur();"&gt;<br>		    						
		   			<div style="padding-right:10px;">
			   			<input type="text" value="<map name='mapGroup'>" style="border:0" size="30"><br>
							<textarea id="tGMap" name="tGMap" style="width:100%;height:280px;"><%=db2html(glink)%></textarea>  	
						<input type="text" value="</map>" style="border:0">	
					</div>
				</td>
			</tr>	
		</table>
	</td>
		
</tr>
<tr>
	<td colspan="2" bgcolor="#FFFFFF" align="right" height="40">
		<input type="image" src="/images/icon_confirm.gif">
		<a href="javascript:window.close();"><img src="/images/icon_cancel.gif" border="0"></a>
	</td>
</tr>	
</table>
</form>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->