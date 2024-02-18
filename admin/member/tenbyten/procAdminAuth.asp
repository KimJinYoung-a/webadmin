<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  ������ѵ��
' History : ������ ����
'			2022.05.09 �ѿ�� ����(ISMS����������ޱ��� �߰�)
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/checkAllowIPWithLog.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/md5.asp" -->
<%
dim empno, userid, frontid, username, userpass, part_sn, posit_sn, level_sn, job_sn, isusing, userdiv, criticinfouser, userpass64
dim objCmd,returnValue, isdispmember
dim adminid, lv1customerYN, lv2partnerYN, lv3InternalYN
lv1customerYN 	= requestCheckvar(request("lv1customerYN"),1)
lv2partnerYN 	= requestCheckvar(request("lv2partnerYN"),1)
lv3InternalYN 	= requestCheckvar(request("lv3InternalYN"),1)
empno = requestCheckvar(request("sEN"),14)
userid = trim(requestCheckvar(request("sUI"),32))
frontid = requestCheckvar(request("sFUI"),32)
userpass = requestCheckvar(request("sP"),32)

userdiv = requestCheckvar(request("selUD"),10)
level_sn = requestCheckvar(request("selLN"),10)
part_sn = requestCheckvar(request("selPN"),10)
posit_sn = requestCheckvar(request("selPoN"),10)
job_sn = requestCheckvar(request("selJN"),10)
username = requestCheckvar(request("sUN"),32)
userid = Replace(userid, " ", "")
frontid = Replace(frontid, " ", "")
username = Replace(username, " ", "")
adminid = session("ssBctId")
 
''response.write "aaa" & requestCheckvar(request("selPN"),10)
''dbget.close
''response.end

''2014/07/14
criticinfouser = requestCheckvar(request("criticinfouser"),10)
criticinfouser = CHKIIF(criticinfouser="","0",criticinfouser)
if lv1customerYN="" or isnull(lv1customerYN) then lv1customerYN="N"
if lv2partnerYN="" or isnull(lv2partnerYN) then lv2partnerYN="N"
if lv3InternalYN="" or isnull(lv3InternalYN) then lv3InternalYN="N"

IF application("Svr_Info")="Dev" THEN
	isdispmember = true
else
	' ISMS �ɻ�� ���� �������� ���ٱ��� ����/����/���� Ư������� ���̰�(�ѿ��,������,�̹���)	' 2020.10.12 �ѿ��
	if C_privacyadminuser or C_PSMngPart then
		isdispmember = true
	else
		isdispmember = false
	end if
end if

if not(isdispmember) then
	response.write "�λ����̰ų� �������� ���ٱ��� �����ڸ� ���� ������ �Ŵ� �Դϴ�."
	response.end
end if

if (posit_sn = "") then
	posit_sn = "0"
end if
	' �μ��� ������ �ִ� ��쿡��
	if part_sn<>"35" and part_sn<>"" then
		IF frontid <> "" THEN
		Set objCmd = Server.CreateObject("ADODB.COMMAND")
		With objCmd
		.ActiveConnection = dbget
		.CommandType = adCmdText
		.CommandText = "{?= call [db_partner].[dbo].sp_Ten_user_tenbyten_chkFrontId('"&frontid&"','"&username&"')}"
		.Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue)
		.Execute, , adExecuteNoRecords
		End With
	    returnValue = objCmd(0).Value
		Set objCmd = Nothing

		IF  returnValue = 0 THEN
			dbget.Close
		%>
			<script type="text/javascript">
				alert("�ٹ����پ��̵��� �̸��� WEBADMIN ���� �̸��� �������� �ʽ��ϴ�.\n�ٹ����� ���̵� Ȯ�� �� �ٽ� �Է����ּ���.");
				history.back();
			</script>
	<%		response.end
		END IF
		END IF
	END IF

		'//�н����� ��å �˻�
	if userid <> "" and userpass <> "" then
		if chkPasswordComplex(userid,userpass)<>"" then
	    	response.write "<script language='javascript'>" &vbCrLf &_
	    					"	alert('" & chkPasswordComplex(userid,userpass) & "\nSCM �н����带 Ȯ���� �ٽ� �õ����ּ���.');" &vbCrLf &_
	    					" 	history.back();" &vbCrLf &_
	    					"</script>"
	    	dbget.close()	:	response.End
	    end if
	     userpass = md5(trim(userpass))
	     userpass64 = sha256((trim(userpass)))
    end if

    Set objCmd = Server.CreateObject("ADODB.COMMAND")
		With objCmd
		.ActiveConnection = dbget
		.CommandType = adCmdText
		.CommandText = "{?= call [db_partner].[dbo].sp_Ten_user_tenbyten_AdminAuth('"&empno&"','"&userid&"','"&frontid&"','"&username&"','"&userpass&"','"&userpass64&"','"&level_sn&"','"&part_sn&"','"&posit_sn&"','"&job_sn&"','"&userdiv&"',"&criticinfouser&",'"&adminid&"','"&lv1customerYN&"','"&lv2partnerYN&"','"&lv3InternalYN&"')}"

		.Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue)
		.Execute, , adExecuteNoRecords
		End With
	    returnValue = objCmd(0).Value
	Set objCmd = Nothing

''' �߰� ���� ����..

    dim StrSQL, lp1

    '���� �ڷ� ����
    strSQL = "Delete From db_partner.dbo.tbl_partner_AddLevel Where userID='" & userid &"'"
    dbget.Execute(strSQL)

    strSQL = "Insert into db_partner.dbo.tbl_partner_AddLevel"
    strSQL = strSQL & " (userid,part_sn,level_sn,isDefault)"
    strSQL = strSQL & " select p.id, t.part_sn, p.level_sn,'Y'"
    strSQL = strSQL & " from db_partner.dbo.tbl_partner p"
    strSQL = strSQL & " 	Join db_partner.dbo.tbl_user_tenbyten t"
    strSQL = strSQL & " 	on p.id=t.userid"
    strSQL = strSQL & " where p.id='" & userid &"'"
    dbget.Execute(strSQL)


    dim ARRpart_sn : ARRpart_sn		= request.Form("part_sn")
    dim ARRlevel_sn : ARRlevel_sn	= request.Form("level_sn")
    dim splPsn, splLsn

    '�����ڷ� ����
    if ARRpart_sn<>"" then splPsn = Split(ARRpart_sn, ",")
    if ARRlevel_sn<>"" then splLsn = Split(ARRlevel_sn, ",")

    If IsArray(splPsn) Then
        For lp1=0 to Ubound(splPsn)
        	IF Trim(splPsn(lp1))<>"" and Trim(splLsn(lp1))<>"" THEN
        	strSQL =	"Insert into db_partner.dbo.tbl_partner_AddLevel (userid,part_sn,level_sn,isDefault)"
        	strSQL = strSQL & "Values ('" & userid & "', " & splPsn(lp1) & ", " & splLsn(lp1) & ",'N')"

        	dbget.Execute(strSQL)
        	END IF
        Next
    end if


if (returnValue = "1") then
			response.write	"<script  type='text/javascript'>" &_
							"	alert('ó���Ϸ�Ǿ����ϴ�.');" &_
							"	opener.location.reload();" &_
							"	opener.focus();" &_
							"	self.close();" &_
							"</script>"
	Else
		response.write	"<script  type='text/javascript'>" &_
						"	alert('ó���� ������ �߻��߽��ϴ�.');" &_
						"	history.back();" &_
						"</script>"

	End If
%>
