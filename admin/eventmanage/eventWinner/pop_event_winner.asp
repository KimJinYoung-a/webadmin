<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Page : /admin/eventmanage/event/pop_event_winner.asp
' Description :  �̺�Ʈ ��÷���
' History : 2007.02.22 ������ ����
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/eventWinner_function.asp"-->
<!-- #include virtual="/lib/function.asp"-->

<script language="javascript">
window.resizeTo(600,460);
<!--
	function jsChType(iVal){
		var frm = document.all;
		if(iVal == "2"){
			frm.div1.style.display = "none";
			frm.div2.style.display = "";
		}else if	(iVal == "3"){
			frm.div1.style.display = "none";
			frm.div2.style.display = "none";
		}else{
			frm.div1.style.display = "";
			frm.div2.style.display = "none";
		}
	}

	//-- jsPopCal : �޷� �˾� --//
	function jsPopCal(sName){
		var winCal;
		winCal = window.open('/lib/common_cal.asp?DN='+sName,'pCal','width=250, height=200');
		winCal.focus();
	}


	function jsWinnerSubmit(frm){
		if(!frm.sR.value){
			alert("����� �Է����ּ���");
			frm.sR.focus();
			return false;
		}

		if(!IsDigit(frm.sR.value)){
			alert("����� ���ڸ� �Է°����մϴ�.");
			frm.sR.focus();
			return false;
		}

		if(!frm.sW.value){
			alert("��÷�ڸ� �Է����ּ���");
			frm.sW.focus();
			return false;
		}

		if(frm.selType.value == "1"){
			if(!frm.sGN.value){
				alert("����ǰ���� �Է����ּ���");
				frm.sGN.focus();
				return false;
			}

			if (frm.reqdeliverdate.value.length<1){
			    frm.reqdeliverdate.focus();
			    alert('��� ��û���� �����ϼ���.');
			    return false;
			}

			if ((!frm.isupchebeasong[0].checked)&&(!frm.isupchebeasong[1].checked)){
        		alert('��� ������ �����ϼ���.');
        		return false;
        	}

            if ((frm.isupchebeasong[1].checked)&&(frm.makerid.value.length<1)){
                alert('��ü ���̵� �����ϼ���.');
        		return false;
            }
		}

		if(frm.selType.value == "2"){
			if(!frm.couponvalue.value){
				alert("�����ݾ� �Ǵ� �������� �Է����ּ���!");
				frm.couponvalue.focus();
				return false;
			}

			if(!frm.minbuyprice.value){
				alert("�ּұݾ��� �Է����ּ���!");
				frm.minbuyprice.focus();
				return false;
			}

			 if(!frm.sDate.value || !frm.eDate.value ){
			  	alert("�Ⱓ�� �Է����ּ���");
			  	frm.sDate.focus();
			  	return false;
			  }

			  if(frm.sDate.value > frm.eDate.value){
			  	alert("�������� �����Ϻ��� �����ϴ�. �ٽ� �Է����ּ���");
			  	frm.sDate.focus();
			  	return false;
			  }
		}

		if(confirm("����Ͻ� ������ ���� �Ǵ� ������ �Ұ����ϸ� ������ �ٷ� ����˴ϴ�.\n\n��� �Ͻðڽ��ϱ�? ")){
			return true;
		}else{
		    return false;
		}
	}

	function disabledBox(comp){
        var frm = comp.form;
        if (comp.value=="Y"){
            frm.makerid.disabled = false;
        }else{
            frm.makerid.selectedIndex = 0;
            frm.makerid.disabled = true;
        }
    }
//-->
</script>
<%
Dim eCode : eCode = Request("eC")
dim arridx : arridx = chkarray(request("arridx"))
%>
<div style="padding: 0 5 5 5"> <img src="/images/icon_arrow_link.gif" align="absmiddle"> ��÷�� ���</div>
<table width="580" border="0" align="center" class="a" cellpadding="0" cellspacing="0">
<form name="frmWin" method="post" action="eventprize_process.asp" onSubmit="return jsWinnerSubmit(this);">
<input type="hidden" name="eC" value="<%=eCode%>">
<input type="hidden" name="mode" value="I">
<tr>
	<td>
		<table width="100%" border="0" align="left" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
			<tr>
				<td width="100" align="center" bgcolor="<%= adminColor("tabletop") %>">����</td>
				<td bgcolor="#FFFFFF">
					<select name="selType" onChange="jsChType(this.value);">
					<option value="1">����ǰ���</option>
					<option value="2">�����߱�</option>
					<option value="3">Only View</option>
					</select>
				</td>
			</tr>
			<tr>
				<td align="center" bgcolor="<%= adminColor("tabletop") %>">���</td>
				<td bgcolor="#FFFFFF"><input type="text" size="2" name="sR"></td>
			</tr>
			<tr>
				<td align="center" bgcolor="<%= adminColor("tabletop") %>">�����Ī</td>
				<td bgcolor="#FFFFFF"><input type="text" name="sRN" size="20"></td>
			</tr>
			<tr>
				<td align="center" bgcolor="<%= adminColor("tabletop") %>">��÷��</td>
				<td bgcolor="#FFFFFF">
					�޸ӷ� ����, ������� (��: aaa,bbb,ccc)<br>
					<textarea name="sW" rows="2" cols="60"><%= arridx %></textarea>
				</td>
			</tr>
		</table>
	</td>

</tr>
<tr>
	<td>
		<div id="div1" style="display:;">
		<table width="100%" border="0" align="left" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
			<tr>
				<td align="center" width="100"  bgcolor="<%= adminColor("tabletop") %>">����� ��ϱ���</td>
				<td bgcolor="#FFFFFF">
					<input type=radio name=rdgubun value="U">User�� ����� �Է�
					<input type=radio name=rdgubun value="F" checked>User �⺻ �ּ� ��� <font color="blue">[������ �⺻ �ּ��� ���]</font>
				</td>
			</tr>
			<tr>
				<td align="center" bgcolor="<%= adminColor("tabletop") %>">����ǰ��</td>
				<td bgcolor="#FFFFFF"><input type="text" name="sGN" size="20"></td>
			</tr>
			<tr>
				<td align="center" bgcolor="<%= adminColor("tabletop") %>">����ǰ ��ǰ��ȣ</td>
				<td bgcolor="#FFFFFF"><input type="text" name="itemid" size="10"></td>
			</tr>
			<!-- ��� ���� �߰� : ������ -->
			<tr>
            	<td align="center" bgcolor="<%= adminColor("tabletop") %>">����û��</td>
            	<td bgcolor="#FFFFFF">
            		<input type="text" name="reqdeliverdate" size="10" maxlength="10"  value="" >
		            <a href="javascript:jsPopCal('reqdeliverdate');"><img src="/images/calicon.gif" border="0" align="absmiddle"></a>
            	</td>
            </tr>
			<tr>
            	<td align="center" bgcolor="<%= adminColor("tabletop") %>">��۱���</td>
            	<td bgcolor="#FFFFFF">
            		<input type=radio name=isupchebeasong value="N" onClick="disabledBox(this);">�ٹ����ٹ��
            		<input type=radio name=isupchebeasong value="Y" onClick="disabledBox(this);">��ü�������
            	</td>
            </tr>
            <tr>
            	<td align="center" bgcolor="<%= adminColor("tabletop") %>">��ü��۽�<br>��üID</td>
            	<td bgcolor="#FFFFFF">
            	    <% drawSelectBoxDesignerwithName "makerid","" %>
            	    <script language='javascript'>
            	    document.frmWin.makerid.disabled=true;
            	    </script>
            	</td>
            </tr>

		</table>
		</div>
		<div id="div2" style="display:;">
		<table width="100%" border="0" align="left" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
			<tr>
				<td align="center" width="100" bgcolor="<%= adminColor("tabletop") %>">����Ÿ��</td>
				<td bgcolor="#FFFFFF">
					<input type=text name=couponvalue maxlength=7 size=10>
					<input type=radio name=coupontype value="1" onclick="alert('% ���� �����Դϴ�.');">%����
					<input type=radio name=coupontype value="2" checked >������
					(�ݾ� �Ǵ� % ����)
				</td>
			</tr>
			<tr>
				<td align="center" bgcolor="<%= adminColor("tabletop") %>">�ּұ��űݾ�</td>
				<td bgcolor="#FFFFFF"><input type=text name=minbuyprice maxlength=7 size=10>�� �̻� ���Ž� ��밡��(����)</td>
			</tr>
			<tr>
				<td align="center" bgcolor="<%= adminColor("tabletop") %>">��ȿ�Ⱓ</td>
				<td bgcolor="#FFFFFF">
					<input type="text" name="sDate" value="<%= left(now(),10) %>"  size="10" maxlength="10" onClick="jsPopCal('sDate');" style="cursor:hand;">
					~<input type="text" name="eDate" size="10"  maxlength="10" onClick="jsPopCal('eDate');" style="cursor:hand;">
				</td>
			</tr>
		</table>
		</div>
	</td>

</tr>
<tr>
	<td colspan="2" bgcolor="#FFFFFF" align="right" height="40">
		<input type="image" src="/images/icon_confirm.gif">
		<a href="javascript:window.close();"><img src="/images/icon_cancel.gif" border="0"></a>
	</td>
</tr>
</form>
</table>

<!-- #include virtual="/lib/db/dbclose.asp" -->