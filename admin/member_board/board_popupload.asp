<%@ language=vbscript %>
<% option explicit %>

<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<%
'####################################################
' Description :  �������� ���μ���
' History : 2011.02.23 ������ ����
'####################################################
%>
<%
 Dim sFolder, sImgID, chkIcon,sImgURL, arrImg
 sFolder 	=  requestCheckVar(request("sF"),10)
 sImgID 	=  requestCheckVar(request("sID"),4)
 chkIcon   	=  requestCheckVar(request("chkI"),1)
 sImgURL 	= requestCheckVar(request("sIU"),100)
 
 '//�̹��� �� ����
IF sImgURL <> "" THEN
	arrImg 	= split(sImgURL,"/")
	sImgURL	= arrImg(Ubound(arrImg))
END IF
%>
<script language="javascript">
<!--
	function jsUpload(){
		if(!document.frmImg.sfImg1.value){
			alert("ã�ƺ��� ��ư�� ���� ���ε��� �̹�����\n �ּ� 1�� �̻��� ������ �ּ���.");			
			return false;
		}
		
		var cnt = document.all.div1.rows.length;
		var re = /[,]/gi;

		if (cnt > 1)
		{
			for(var i=0; i<cnt; i++)
			{
				
				if (document.forms["frmImg"].elements["sfImg"+(i+1)+""].value.search(re) != -1) {
				    alert(""+(i+1)+"��° ���ϸ� ��ǥ(,)�� �������ּ���");
				    return false;
				}
			}
		}
		else
		{

				if(document.forms["frmImg"].elements["sfImg1"].value.search(re) != -1){
					alert("1��° ���ϸ� ��ǥ(,)�� �������ּ���.");
					return false;
				}

		}

		document.frmImg.filecnt.value = document.all.div1.rows.length;

		document.frmImg.submit();
		document.all.dvLoad.style.display = "";
	}
	
	function AutoInsert() {
		var f = document.all;
	
		var rowLen = f.div1.rows.length;
		var r  = f.div1.insertRow(rowLen++);
		var c0 = r.insertCell(0);
		
		var Html;

		c0.innerHTML = "&nbsp;";
		var inHtml = "<input type='file' name='sfImg"+rowLen+"'>";
		c0.innerHTML = inHtml;
	}
	
	function uploadOK()
	{
		opener.frm.filename.value = frmImg.filename.value;
		window.close();
	}
//-->
</script>
<div style="padding: 0 5 5 5"> <img src="/images/icon_arrow_link.gif" align="absmiddle"> �̹��� ÷��</div>
<form name="frmImg" method="post" action="<%= uploadImgUrl %>/linkweb/member_board_admin/member_board_upload.asp" enctype="MULTIPART/FORM-DATA">
<input type="hidden" name="mode" value="I">
<input type="hidden" name="sF" value="<%=sFolder%>">
<input type="hidden" name="sID" value="<%=sImgID%>">
<input type="hidden" name="chkI" value ="<%=chkIcon%>">
<input type="hidden" name="filename" value="">
<input type="hidden" name="filecnt" value="0">
<table width="380" border="0" align="left" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>"> 
	<tr>
		<td bgcolor="<%= adminColor("tabletop") %>" align="center" valign="top" style="padding:6 0 0 2">���ϸ�</td>
		<td bgcolor="#FFFFFF">
			<table cellpadding="0" cellspacing="0" border="0" id="div1">
			<tr>
				<td><input type="file" name="sfImg1"></td>
			</tr>
			</table>
		</td>
		<td bgcolor="#FFFFFF" width="50" align="center" valign="top" style="padding:3 0 0 0">
			<input type="button" value="�߰�" onClick="AutoInsert()" class="button">
		</td>
	</tr>	
	<tr>
		<td colspan="3" bgcolor="#FFFFFF" align="right">
			<!--<input type="image" src="/images/icon_confirm.gif">//-->
			<img src="/images/icon_confirm.gif" style="cursor:pointer" onclick="jsUpload();">
			<a href="javascript:window.close();"><img src="/images/icon_cancel.gif" border="0"></a>
		</td>
	</tr> 
<tr>
	<td bgcolor="#FFFFFF" colspan="3">
+ �ִ� ���ϻ����� 20MB(20,480KB)���ϸ� ����<br>		
+ �̹���(gif,jpg) �� zip,xls,pdf,ppt,doc ���ϸ� �����˴ϴ�
</td>
</tr>
</table>
</form>	
<!-- #include virtual="/lib/db/dbclose.asp" -->
<div id="dvLoad" style="display:none;top:50px;left:20;position:absolute;background-color:gray;">
	<table border="0" class="a" cellpadding="5" cellspacing="5">
		<tr>
			<td> <font color="#FFFFFF">���ε� ó�����Դϴ�. ��ø� ��ٷ��ּ���~~</font></td>
		</tr>
	</table>
</div>