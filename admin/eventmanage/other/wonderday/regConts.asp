<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/event/eventOtherCls_wonderday.asp"-->
<%
 Dim iIdx, iCurrentpage, sMode
 Dim clsWonderday
 Dim sListImg, sMainImg, blnUse,dOpendate,iVolnum,imaxVol,intLoop
 iIdx 			=  requestCheckVar(Request("idx"),10)
 iCurrentpage 	=  requestCheckVar(Request("iC"),10)
 sMode	= "I"
 imaxVol = 2 '���۾�ó��
 
 IF iIdx <> "" THEN
 	sMode	= "U"
 set clsWonderday = new CWonderday
 	clsWonderday.FIdx = iIdx
 	clsWonderday.fnGetConts
 	sListImg 	= clsWonderday.FListImg
 	sMainImg 	= clsWonderday.FMainImg
 	blnUse		= clsWonderday.FUsing
 	dOpendate	= clsWonderday.FOpendate
 	iVolnum		= clsWonderday.FVolnum
 set clsWonderday = nothing
END IF

IF blnUse = "" THEN blnUse = True
IF dOpendate = "" THEN dOpendate = dateadd("d",date(),1)	
%>
<script language="javascript">
<!--
	function jsSubmit(){
		var frm= document.frmReg;
		
		if(!frm.sLImg.value && !frm.sOrgLImg.value){
			alert("����Ʈ �̹����� ������ּ���");
			return false;
		}
		
		if(!frm.sMImg.value && !frm.sOrgMImg.value){
			alert("���� �̹����� ������ּ���");
			return false;
		}
		
	}
	
		
//-- jsImgView : �̹��� Ȯ��ȭ�� ��â���� �����ֱ� --//
	function jsImgView(sImgUrl){
	 var wImgView;
	 wImgView = window.open('/admin/eventmanage/common/pop_event_detailImg.asp?sUrl='+sImgUrl,'pImg','width=100,height=100');
	 wImgView.focus();
	}
	
	//-- jsPopCal : �޷� �˾� --//
	function jsPopCal(sName){
		var winCal;
		winCal = window.open('/lib/common_cal.asp?DN='+sName,'pCal','width=250, height=200');
		winCal.focus();
	}
	
//-->
</script>
<table width="800" align="left" cellpadding="3" cellspacing="1" class="a">
<form name="frmReg" method="post" action="<%= uploadImgUrl %>/linkweb/event_admin/eventUpload_wonderday.asp" enctype="MULTIPART/FORM-DATA" onSubmit="return jsSubmit();"> 
<input type="hidden" name="sM" value="<%=sMode%>">
<input type="hidden" name="menupos" value="<%=menupos%>">
<input type="hidden" name="sOrgLImg" value="<%=sListImg%>">
<input type="hidden" name="sOrgMImg" value="<%=sMainImg%>">
<input type="hidden" name="idx" value="<%=iIdx%>">
<input type="hidden" name="iC" value="<%=iCurrentpage%>">
<tr>
	<td>
		<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">	
			<tr>
				<td  width="150"  align="center"  bgcolor="<%= adminColor("gray") %>">ID</td>
				<td bgcolor="#FFFFFF"><%=iIdx%></td>
			</tr>
			<tr>
				<td  width="150"  align="center"  bgcolor="<%= adminColor("gray") %>">ȸ��</td>
				<td bgcolor="#FFFFFF">
					<select name="selVol">					
					<%For intLoop = imaxVol To 1 step -1 %>
					<option value="<%=intLoop%>" <%IF iVolnum = intLoop THEN%>selected<%END IF%>><%=intLoop%></option>
					<%Next%>
					</select>
				</td>
			</tr>	
			<tr>
				<td  width="150"  align="center"  bgcolor="<%= adminColor("gray") %>">������</td>
				<td bgcolor="#FFFFFF"> <input type="text" name="dOD" size="10" onClick="jsPopCal('dOD');"  value="<%=dOpendate%>" style="cursor:hand;"> </td>
			</tr>	
			<tr>
				<td width="150" align="center" bgcolor="<%= adminColor("gray") %>">����Ʈ �̹���</td>
				<td bgcolor="#FFFFFF">
					<input type="file" name="sLImg" class="input" style="width:300px;">
					<%IF sListImg <> "" THEN%><div><img src="<%=sListImg%>" onClick="jsImgView('<%=sListImg%>')" style="cursor:hand;"></div><%END IF%>
				</td>
			</tr>
			<tr>
				<td  width="150"  align="center"  bgcolor="<%= adminColor("gray") %>">���� �̹���</td>
				<td bgcolor="#FFFFFF"><input type="file" name="sMImg" class="input" style="width:300px;">
				<%IF sListImg <> "" THEN%>
					<div><img src="<%=sMainImg%>" width="200" height="200" onClick="jsImgView('<%=sMainImg%>')" style="cursor:hand;">
					<br>*�̹����� Ŭ���Ͻø� �̹��� ���� ������� Ȯ�ΰ����մϴ�.
					</div>
				<%END IF%>
				</td>
			</tr>
			<tr>
				<td  width="150"  align="center"  bgcolor="<%= adminColor("gray") %>">���ÿ���</td>
				<td bgcolor="#FFFFFF"><input type="checkbox" name="blnU" value="1" <%IF blnUse THEN%>checked<%END IF%>>������</td>
			</tr>
		</table>
	</td>		
</tr>
<tr>
	<td>	
		<table width="100%">
		<tr>
			<td  align="center"><input type="image" src="/images/icon_save.gif">
				<a href="index.asp?menupos=<%=menupos%>"><img src="/images/icon_cancel.gif" border="0"></a>
			</td>
		</tr>			
		</table>
	</td>		
</tr>
</form>	
</table>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->