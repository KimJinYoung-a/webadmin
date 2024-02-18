<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : cs���� ��������
' History : �̻󱸻���
'			2023.05.23 �ѿ�� ����(���� ������Ŷ ���� üũ �߰�)
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/cscenter/cs_couponcls.asp" -->
<!-- #include virtual="/lib/classes/cscenter/sp_itemcouponcls.asp" -->
<%
dim mode, submode, coupontype, couponidx, extendday, strSQL
dim userid, orderserial, jukyo, contents_jupsu, adminuserid, BirthdayCouponCnt, OldBirthdayCouponCnt
	mode = requestCheckvar(request("mode"),32)
	submode = requestCheckvar(request("submode"),32)
	coupontype = requestCheckvar(request("coupontype"),32)
	couponidx = requestCheckvar(request("couponidx"),32)
	extendday = requestCheckvar(request("extendday"),32)
	userid = requestCheckvar(request("userid"),32)
	orderserial = requestCheckvar(request("orderserial"),32)
	jukyo = requestCheckvar(request("jukyo"),32)

adminuserid = session("ssBctId")
BirthdayCouponCnt = 0
OldBirthdayCouponCnt = 0

'==============================================================================
'���ʽ�����
dim ocscoupon
set ocscoupon = New CCSCenterCoupon
ocscoupon.FRectBonusCouponIdx = couponidx
'ocscoupon.GetOneCSCenterCoupon

'==============================================================================
dim totay, expireday, baseday, daybeforeonemonth

totay = Left(now(), 10)
daybeforeonemonth = Left(DateAdd("d", -30, totay), 10)

if (mode = "expiredate") then
	'���� �Ⱓ����
	ocscoupon.GetOneCSCenterCoupon

	totay = Left(now(), 10)
	expireday = Left(ocscoupon.FOneItem.Fexpiredate,10)

	baseday = totay
	if (expireday > totay) then
		baseday = expireday
	end if

	baseday = DateAdd("d", Cint(extendday), baseday)
	baseday = Left(baseday,10) & " 23:59:59"
	'response.write baseday
	'response.end

	if (ocscoupon.FOneItem.Fisusing <> "Y") and (ocscoupon.FOneItem.Fdeleteyn <> "Y") and (daybeforeonemonth <= Left(ocscoupon.FOneItem.Fexpiredate,10)) then
		strSQL = "update [db_user].[dbo].tbl_user_coupon " + vbCrlf
		strSQL = strSQL + " set expiredate = '" & baseday & "', reguserid = '" & adminuserid & "' " + vbCrlf
		strSQL = strSQL + " where idx='" + couponidx + "'" + vbCrlf
		rsget.Open strSQL,dbget,1

		response.write "<script type='text/javascript'>alert('���� �Ǿ����ϴ�.');</script>"
		response.write "<script type='text/javascript'>opener.location.reload();</script>"
	else
		response.write "<script type='text/javascript'>alert('����� ������ �Ⱓ���� �� �� �����ϴ�.');</script>"
	end if

	response.write "<script type='text/javascript'>location.replace('/cscenter/coupon/pop_coupon_modify.asp?coupontype=" + coupontype + "&couponidx=" + couponidx + "');</script>"

elseif (mode = "copy") then
	'���� �������
	'�ߺ��������� csorderserial �� ����
	ocscoupon.GetOneCSCenterCoupon

	if (ocscoupon.FOneItem.Fisusing = "Y") and (ocscoupon.FOneItem.Fdeleteyn <> "Y") and (ocscoupon.FOneItem.FprevCopiedCouponCount = 0) and (daybeforeonemonth <= Left(ocscoupon.FOneItem.Fexpiredate,10)) then

		strSQL = "insert into [db_user].[dbo].tbl_user_coupon(reguserid, isusing, masteridx, userid, coupontype, couponvalue, couponname, minbuyprice, targetitemlist, couponimage, startdate, expiredate, deleteyn, exitemid, validsitename, notvalid10x10, couponmeaipprice, ssnkey, scratchcouponidx, evtprize_code, useLevel, csorderserial, targetCpnType  , targetCpnSource, mxCpnDiscount) " + vbCrlf
		strSQL = strSQL + " select top 1 '" & adminuserid & "', 'N', masteridx, userid, coupontype, couponvalue, couponname, minbuyprice, targetitemlist, couponimage, startdate, expiredate, deleteyn, exitemid, validsitename, notvalid10x10, couponmeaipprice, ssnkey, scratchcouponidx, evtprize_code, useLevel, orderserial, targetCpnType  , targetCpnSource, mxCpnDiscount " + vbCrlf
		strSQL = strSQL + " from [db_user].[dbo].tbl_user_coupon " + vbCrlf
		strSQL = strSQL + " where idx = '" + couponidx + "' " + vbCrlf
		rsget.Open strSQL,dbget,1

		response.write "<script type='text/javascript'>alert('���� �Ǿ����ϴ�.');</script>"
		response.write "<script type='text/javascript'>opener.location.reload();</script>"
	elseif (ocscoupon.FOneItem.FprevCopiedCouponCount > 0) then
		response.write "<script type='text/javascript'>alert('����Ұ�!!\n\n������ ���常 ��������� �� �ֽ��ϴ�.');</script>"
	else
		response.write "<script type='text/javascript'>alert('����� ������ ������� �����մϴ�.');</script>"
	end if

	response.write "<script type='text/javascript'>opener.focus(); window.close();</script>"

elseif (mode = "issuecoupon") then
'C_ADMIN_AUTH=false
'C_CSpermanentUser=false
	if (submode = "issuecoupon3000") then
		' �������̰ų� cs�� ������(��� �̻�) �̰�� ���డ��
		if not(C_ADMIN_AUTH or C_CSpermanentUser) then
			response.write "<script type='text/javascript'>"
			response.write "	alert('��������� �����ϴ�.[0]');"
			response.write "</script>"
			dbget.close() : response.end
		end if

		'3000�� ��������
		strSQL = "insert into [db_user].[dbo].tbl_user_coupon"
		strSQL = strSQL + " (reguserid, masteridx,userid,coupontype,couponvalue,couponname,minbuyprice,startdate,expiredate,csorderserial) " & vbCrlf
		strSQL = strSQL + " select '" & adminuserid & "',287,userid,'2','3000','3000�� ��������',100, " & vbCrlf
		strSQL = strSQL + " convert(varchar(10),getdate(),21) , convert(varchar(10),dateadd(m,1,getdate()),21) + ' 23:59:59' " & vbCrlf
		strSQL = strSQL + " ,'" + CStr(orderserial) + "' " & vbCrlf
		strSQL = strSQL + " from [db_user].[dbo].tbl_user_n" & vbCrlf
		strSQL = strSQL + " where userid ='" + userid + "'" & vbCrlf
		rsget.Open strSQL,dbget,1

		contents_jupsu = "3000�� �������� ����"

	elseif (submode = "issuecoupon5per") then
		' �������̰ų� cs�� ������(��� �̻�) �̰�� ���డ��
		if not(C_ADMIN_AUTH or C_CSpermanentUser) then
			response.write "<script type='text/javascript'>"
			response.write "	alert('��������� �����ϴ�.[1]');"
			response.write "</script>"
			dbget.close() : response.end
		end if

		'5% ��������
		strSQL = "insert into [db_user].[dbo].tbl_user_coupon"
		strSQL = strSQL + " (reguserid, masteridx,userid,coupontype,couponvalue,couponname,minbuyprice,startdate,expiredate,csorderserial, mxCpnDiscount) " & vbCrlf
		strSQL = strSQL + " select '" & adminuserid & "',287,userid,'1','5','5% ��������',100, " & vbCrlf
		strSQL = strSQL + " convert(varchar(10),getdate(),21) , convert(varchar(10),dateadd(m,1,getdate()),21) + ' 23:59:59' " & vbCrlf
		strSQL = strSQL + " ,'" + CStr(orderserial) + "', 10000 " & vbCrlf
		strSQL = strSQL + " from [db_user].[dbo].tbl_user_n" & vbCrlf
		strSQL = strSQL + " where userid ='" + userid + "'" & vbCrlf
		rsget.Open strSQL,dbget,1

		contents_jupsu = "5% �������� ����"

	elseif (submode = "issuecoupondeliver") then
		' �������̰ų� cs�� ������(��� �̻�) �̰�� ���డ��
		if not(C_ADMIN_AUTH or C_CSpermanentUser) then
			response.write "<script type='text/javascript'>"
			response.write "	alert('��������� �����ϴ�.[2]');"
			response.write "</script>"
			dbget.close() : response.end
		end if

		'2000�� ��ۺ���������
			strSQL = "insert into [db_user].[dbo].tbl_user_coupon"
			strSQL = strSQL + " (reguserid, masteridx,userid,coupontype,couponvalue,couponname,minbuyprice,startdate,expiredate,csorderserial) " & vbCrlf
			strSQL = strSQL + " select '" & adminuserid & "',287,userid,'3','" + Cstr(getDefaultBeasongPayByDate(now())) + "','��ۺ���������',1000, " & vbCrlf
			strSQL = strSQL + " convert(varchar(10),getdate(),21) , convert(varchar(10),dateadd(d,1,getdate()),21) + ' 23:59:59' " & vbCrlf
			strSQL = strSQL + " ,'" + CStr(orderserial) + "' " & vbCrlf
			strSQL = strSQL + " from [db_user].[dbo].tbl_user_n" & vbCrlf
			strSQL = strSQL + " where userid ='" + userid + "'" & vbCrlf

		'response.write strSQL & "<br>"
		dbget.execute strSQL

		contents_jupsu = Cstr(getDefaultBeasongPayByDate(now())) + "�� ��ۺ��������� ����"

	' �������� ����		' 2018.09.17 �ѿ�� ����
	elseif (submode = "IssueCouponBirthday") then
		' �������̰ų� cs���ϰ�� ���డ��
		if not(C_ADMIN_AUTH or C_CSUser) then
			response.write "<script type='text/javascript'>"
			response.write "	alert('��������� �����ϴ�.[2]');"
			response.write "</script>"
			dbget.close() : response.end
		end if

		BirthdayCouponCnt = 0
		OldBirthdayCouponCnt = 0

		strSQL = " select count(userid) as BirthdayCouponcnt" & vbcrlf
		strSQL = strSQL & " from [db_user].dbo.tbl_user_coupon with (nolock)" & vbcrlf
		strSQL = strSQL & " where masteridx=126 and deleteyn='N'" & vbcrlf
		strSQL = strSQL & " and datediff(year, regdate, getdate()) = 0" & vbcrlf
		strSQL = strSQL & " and userid = '"& userid &"'" & vbcrlf

		'response.write strSQL & "<br>"
		rsget.Open strSQL,dbget,1
		if not rsget.EOF  then
			BirthdayCouponCnt = rsget("BirthdayCouponcnt")
		end if
		rsget.close

		strSQL = " select count(userid) as BirthdayCouponcnt" & vbcrlf
		strSQL = strSQL & " from db_log.dbo.tbl_old_user_coupon with (nolock)" & vbcrlf
		strSQL = strSQL & " where masteridx=126 and deleteyn='N'" & vbcrlf
		strSQL = strSQL & " and datediff(year, regdate, getdate()) = 0" & vbcrlf
		strSQL = strSQL & " and userid = '"& userid &"'" & vbcrlf

		'response.write strSQL & "<br>"
		rsget.Open strSQL,dbget,1
		if not rsget.EOF  then
			OldBirthdayCouponCnt = rsget("BirthdayCouponcnt")
		end if
		rsget.close

		if BirthdayCouponCnt > 0 or OldBirthdayCouponCnt > 0 then
			response.write "<script type='text/javascript'>"
			response.write "	alert('�̹� ���� ����� ���������� �ֽ��ϴ�.'); opener.location.reload(); opener.focus(); window.close();"
			response.write "</script>"
			dbget.close() : response.end
		end if

		strSQL = " insert into [db_user].dbo.tbl_user_coupon(" & vbcrlf
		strSQL = strSQL & " masteridx,userid,coupontype,couponvalue,couponname,minbuyprice,startdate,expiredate,targetitemlist,couponmeaipprice,reguserid)" & vbcrlf
		strSQL = strSQL & " 	select 126,'"& userid &"','2','5000','[��������] ������ ���ϵ����','40000','"& date() &" 00:00:00','"& dateadd("d", +15, date()) &" 23:59:59','',0,'"& adminuserid &"'" & vbcrlf

		'response.write strSQL & "<br>"
		dbget.execute strSQL

		contents_jupsu = "�������� ����"
	end if

	contents_jupsu = contents_jupsu + "(" + CStr(jukyo) + ")"

	'CS�޸�
    strSQL = " insert into [db_cs].[dbo].tbl_cs_memo(orderserial, divcd, userid, mmgubun, qadiv, phoneNumber, writeuser, finishuser, contents_jupsu, finishyn,finishdate,regdate) "
    strSQL = strSQL + " values('" + CStr(orderserial) + "','1','" + CStr(userid) + "','0','20','','" + adminuserid + "','" + adminuserid + "','" + html2db(contents_jupsu) + "','Y',getdate(),getdate()) "
    dbget.Execute strSQL

	response.write "<script type='text/javascript'>alert('���� �Ǿ����ϴ�.');</script>"
	response.write "<script type='text/javascript'>opener.location.reload();</script>"
	response.write "<script type='text/javascript'>opener.focus(); window.close();</script>"
	dbget.close() : response.end
else
	'
end if

%>

<!-- #include virtual="/lib/db/dbclose.asp" -->
