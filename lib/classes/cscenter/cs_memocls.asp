<%
'###########################################################
' Description :  cs �޸� Ŭ����
' History : 2007.10.26 �̻� ����
'           2016.12.07 �ѿ�� ����
'###########################################################

'/������. ���ȭ ��Ŵ. 2016.12.07 �ѿ��
function DrawMemoDivCombo(selName,selVal)
    dim retVal
    retVal = "<select class='select' name='"&selName&"'>"
    retVal = retVal&"    <option value=''>��ü</option>"
    retVal = retVal&"    <option value='00' "& ChkIIF(selVal="00","selected","") &" >��۹���</option>"
    retVal = retVal&"    <option value='01' "& ChkIIF(selVal="01","selected","") &" >�ֹ�����</option>"
    retVal = retVal&"    <option value='02' "& ChkIIF(selVal="02","selected","") &" >��ǰ����</option>"
    retVal = retVal&"    <option value='03' "& ChkIIF(selVal="03","selected","") &" >�����</option>"
    retVal = retVal&"    <option value='04' "& ChkIIF(selVal="04","selected","") &" >��ҹ���</option>"
    retVal = retVal&"    <option value='05' "& ChkIIF(selVal="05","selected","") &" >ȯ�ҹ���</option>"
    retVal = retVal&"    <option value='06' "& ChkIIF(selVal="06","selected","") &" >��ȯ����</option>"
    retVal = retVal&"    <option value='07' "& ChkIIF(selVal="07","selected","") &" >AS����</option>  "
    retVal = retVal&"    <option value='08' "& ChkIIF(selVal="08","selected","") &" >�̺�Ʈ����</option>"
    retVal = retVal&"    <option value='09' "& ChkIIF(selVal="09","selected","") &" >������������</option> "
    retVal = retVal&"    <option value='10' "& ChkIIF(selVal="10","selected","") &" >�ý��۹���</option>"
    retVal = retVal&"    <option value='11' "& ChkIIF(selVal="11","selected","") &" >ȸ����������</option>"
    retVal = retVal&"    <option value='12' "& ChkIIF(selVal="12","selected","") &" >ȸ����������</option>"
    retVal = retVal&"    <option value='13' "& ChkIIF(selVal="13","selected","") &" >��÷����</option>"
    retVal = retVal&"    <option value='14' "& ChkIIF(selVal="14","selected","") &" >��ǰ����</option>"
    retVal = retVal&"    <option value='15' "& ChkIIF(selVal="15","selected","") &" >�Աݹ���</option>"
    retVal = retVal&"    <option value='16' "& ChkIIF(selVal="16","selected","") &" >�������ι���</option>"
    retVal = retVal&"    <option value='17' "& ChkIIF(selVal="17","selected","") &" >����/���ϸ�������</option>"
    retVal = retVal&"    <option value='18' "& ChkIIF(selVal="18","selected","") &" >�����������</option>"
    retVal = retVal&"    <option value='20' "& ChkIIF(selVal="20","selected","") &" >��Ÿ����</option>"
    retVal = retVal&"</select>"

    response.write retVal
end function

'/������. ���ȭ ��Ŵ. 2016.12.07 �ѿ��
function getMemoDivName(iqadiv)
    Select Case iqadiv
        Case "00" : getMemoDivName="��۹���"
        Case "01" : getMemoDivName="�ֹ�����"
        Case "02" : getMemoDivName="��ǰ����"
        Case "03" : getMemoDivName="�����"
        Case "04" : getMemoDivName="��ҹ���"
        Case "05" : getMemoDivName="ȯ�ҹ���"
        Case "06" : getMemoDivName="��ȯ����"
        Case "07" : getMemoDivName="AS����"
        Case "08" : getMemoDivName="�̺�Ʈ����"
        Case "09" : getMemoDivName="������������"
        Case "10" : getMemoDivName="�ý��۹���"
        Case "11" : getMemoDivName="ȸ����������"
        Case "12" : getMemoDivName="ȸ����������"
        Case "13" : getMemoDivName="��÷����"
        Case "14" : getMemoDivName="��ǰ����"
        Case "15" : getMemoDivName="�Աݹ���"
        Case "16" : getMemoDivName="��������"
        Case "17" : getMemoDivName="����/���ϸ�������"
        Case "18" : getMemoDivName="�����������"

        Case "20" : getMemoDivName="��Ÿ����"

        Case "50" : getMemoDivName="�Ϲݸ޸�"
        Case "51" : getMemoDivName="ǰ����ҿ�û"
        Case "52" : getMemoDivName="1:1�Խ���"
        Case "53" : getMemoDivName="D+2��ü����"
        Case "54" : getMemoDivName="D+3��ü����"
        Case "55" : getMemoDivName="��ü�������"
        Case "56" : getMemoDivName="���ǹ�ó��"

        Case "57" : getMemoDivName="��ü�Խ���"
        Case "58" : getMemoDivName="D+4 ��ǰ��ó��"
        Case "59" : getMemoDivName="���޸�"
        Case "60" : getMemoDivName="��ü����"

        CASE Else : getMemoDivName=iqadiv
    End Select
end function

Class CCSMemoItem
    public Fid
    public Forderserial
    public Fdivcd
    public FmmGubun
    public Fuserid
    public FphoneNumber
    public Fwriteuser
    public Ffinishuser
    public Fcontents_jupsu
    public Fregdate
    public Ffinishyn
    public Ffinishdate
	public fcontents_div
	public Fqadiv
	public Fretrydate
	public Fspecialmemo
	public fcomm_cd1
	public fcomm_name1
	public fcomm_cd2
	public fcomm_name2
	Public FupchefinishYN

    public function GetDivCDName()
        if Fdivcd="1" then
			GetDivCDName = "�ܼ�"
        elseif Fdivcd="2" then
			GetDivCDName = "<font color='blue'>��û</font>"
        else
			GetDivCDName = "-"
        end if
    end function

    public function GetmmGubunName()
        if FmmGubun="0" then
			GetmmGubunName = "�Ϲ�"
        elseif FmmGubun="1" then
			GetmmGubunName = "In"
        elseif FmmGubun="2" then
			GetmmGubunName = "Out"
        elseif FmmGubun="3" then
			GetmmGubunName = "��ü"
        elseif FmmGubun="4" then
			GetmmGubunName = "SMS"
        elseif FmmGubun="5" then
			GetmmGubunName = "EMAIL"
        else
			GetmmGubunName = FmmGubun
        end if
    end function

    public function GetQaDivName()
        'GetQaDivName = getMemoDivName(Fqadiv)
        GetQaDivName = fcomm_name2
    end function

    public function GetSiteName()
        if Left(Forderserial,1)="A" then
			GetSiteName = "�ΰŽ�"
        else
			GetSiteName = "���ֹ�"
        end if
    end function

    Private Sub Class_Initialize()
    End Sub
    Private Sub Class_Terminate()
    End Sub
end Class

Class CCSMemo
    public FItemList()
    public FOneItem
    public FCurrPage
    public FTotalPage
    public FPageSize
    public FResultCount
    public FScrollCount
    public FTotalCount

    public FRectUserID
    public FRectOrderserial
    public FRectId
    public FRectIsFinished
    public FRectSiteGubun
    public FRectPhoneNumber
    public FRectDivCD
    public FRectMMGubun
    public FRectWriteUser
    public FRectQadiv
    public FRectContents
	public FRectRegStart
	public FRectRegEnd
	public FRectMMGubunExvlude
	public FRectRetryDateExclude

    public Sub GetCSMemoList()
        dim i,sqlStr, tmpSql, minIdx, tmpStr

		sqlStr = ""
		sqlStr = sqlStr & " SELECT count(*) as cnt, CEILING(CAST(Count(*) AS FLOAT)/" & FPageSize & ") as totPg "
        sqlStr = sqlStr & " from [db_cs].[dbo].tbl_cs_memo c with (nolock)" & vbcrlf
        sqlStr = sqlStr & " left join db_cs.dbo.tbl_cs_comm_code c1 with (nolock)" & vbcrlf
        sqlStr = sqlStr & " 	on c1.comm_group = 'Z030'" & vbcrlf
        sqlStr = sqlStr & " 	and c.mmgubun = right(c1.comm_cd,1)" & vbcrlf
        sqlStr = sqlStr & " left join db_cs.dbo.tbl_cs_comm_code c2 with (nolock)" & vbcrlf
        sqlStr = sqlStr & " 	on c1.comm_cd = c2.comm_group" & vbcrlf
        sqlStr = sqlStr & " 	and c.qadiv = right(c2.comm_cd,2)" & vbcrlf
		If (FRectWriteUser<>"") And (FRectIsFinished = "N") Then
			'// ���� ��üó���Ϸ�
			sqlStr = sqlStr & " 	left join ( "
			sqlStr = sqlStr & " 		select a.orderserial "
			sqlStr = sqlStr & " 		from "
			sqlStr = sqlStr & " 			[db_cs].[dbo].[tbl_cs_memo] m with (nolock) "
			sqlStr = sqlStr & " 			join [db_cs].[dbo].[tbl_new_as_list] a with (nolock) "
			sqlStr = sqlStr & " 			on "
			sqlStr = sqlStr & " 				1 = 1 "
			sqlStr = sqlStr & " 				and m.writeuser = '" & FRectWriteUser & "' "
			sqlStr = sqlStr & " 				and m.finishyn <> 'Y' "
			sqlStr = sqlStr & " 				and m.orderserial = a.orderserial "
			sqlStr = sqlStr & " 				and DateDiff(d, a.finishdate, getdate()) < 1 "
			sqlStr = sqlStr & " 		where "
			sqlStr = sqlStr & " 			1 = 1 "
			sqlStr = sqlStr & " 			and a.requireupche = 'Y' "
			sqlStr = sqlStr & " 			and a.currstate >= 'B006' "
			sqlStr = sqlStr & " 		group by a.orderserial "
			sqlStr = sqlStr & " 	) T "
			sqlStr = sqlStr & " 	on "
			sqlStr = sqlStr & " 		c.orderserial = T.orderserial "
		End If

        sqlStr = sqlStr & " where 1 = 1 " & vbcrlf

		if (FRectRegStart <> "") then
			sqlStr = sqlStr & " and c.regdate>='" & FRectRegStart & "'" & vbcrlf
		end if

		if (FRectRegEnd <> "") then
			sqlStr = sqlStr & " and convert(varchar(10), c.regdate, 21)<='" & FRectRegEnd & "'" & vbcrlf
		end if

		if (FRectContents <> "") then
			sqlStr = sqlStr & " and c.contents_jupsu like '%" & html2db(FRectContents) & "%' " & vbcrlf

			if (FRectUserID = "") and (FRectOrderserial = "") and (FRectPhoneNumber = "") then
				tmpSql = "select (max(id) - 80000) as minIdx from [db_cs].[dbo].tbl_cs_memo with (nolock) "	' �ʹ� ©���� ���´ٰ� �ؼ� 1����->8������ �ø� 2018.03.13 �ѿ��
				rsget.CursorLocation = adUseClient
    			rsget.Open tmpSql, dbget, adOpenForwardOnly, adLockReadOnly
				if  not rsget.EOF  then
					minIdx = rsget("minIdx")
					sqlStr = sqlStr & " and id >= " & minIdx
				end if
				rsget.close
			end if
		end if

		if (FRectMMGubunExvlude <> "") and (FRectContents = "") and (FRectPhoneNumber = "") then
			'// ��ȭ��ȣ �Ǵ� ����˻��� SMS/�̸��� �ȳ� �޸� ���� ����
			'sqlStr = sqlStr & " and c.mmgubun not in (" & FRectMMGubunExvlude & ")" & vbcrlf
			sqlStr = sqlStr & " and c.mmgubun not in ('4','5')" & vbcrlf
		end if

		if (FRectRetryDateExclude <> "") then
			sqlStr = sqlStr & " and not (c.retrydate is not null and c.retrydate > getdate())" & vbcrlf
		end if

        if (FRectUserID <> "") then
            sqlStr = sqlStr & " and c.userid = '" + CStr(FRectUserID) + "' " & vbcrlf
        end if

        if (FRectOrderserial <> "") then
            sqlStr = sqlStr & " and c.orderserial = '" + CStr(FRectOrderserial) + "' " & vbcrlf
        end if

        if (FRectIsFinished <> "") then
            sqlStr = sqlStr & " and c.finishyn = '" + CStr(FRectIsFinished) + "' " & vbcrlf
        end if

	    if (FRectPhoneNumber<> "") then
			tmpStr = "[SMS " + FRectPhoneNumber + "]"
	         ''sqlStr = sqlStr & " and ((c.phonenumber = '" + CStr(FRectPhoneNumber) + "') or (c.contents_jupsu like '%" & html2db(FRectPhoneNumber) & "%')) " & vbcrlf
			 ''sqlStr = sqlStr & " and ((c.phonenumber = '" + CStr(FRectPhoneNumber) + "') or (c.contents_jupsu_varchar like '[[]SMS " & html2db(FRectPhoneNumber) & "%')) " & vbcrlf
			 sqlStr = sqlStr & " and (c.phonenumber = '" + CStr(FRectPhoneNumber) + "') " & vbcrlf
        end if

        if (FRectWriteUser<>"") then
            sqlStr = sqlStr & " and c.writeuser = '" + FRectWriteUser + "' " & vbcrlf
        end if

        if (FRectQadiv<>"") then
            sqlStr = sqlStr & " and c.qadiv = '" + FRectQadiv + "' " & vbcrlf
        end if

        if (FRectDivCD<>"") then
            sqlStr = sqlStr & " and c.divcd = '" + FRectDivCD + "' " & vbcrlf
        end if

        if (FRectMMGubun<>"") then
            sqlStr = sqlStr & " and c.mmgubun = '" + FRectMMGubun + "' " & vbcrlf
        end if

		'response.write sqlStr & "<br>"
		rsget.CursorLocation = adUseClient
    	rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
			FTotalCount = rsget("cnt")
			FTotalPage = rsget("totPg")
		rsget.Close

		'������������ ��ü ���������� Ŭ �� �Լ�����
		If Cint(FCurrPage) > Cint(FTotalPage) Then
			FResultCount = 0
			Exit Sub
		End If

        sqlStr = " select top " & CStr(FPageSize*FCurrPage) & vbcrlf
        sqlStr = sqlStr & " c.id, c.orderserial, c.divcd, c.userid, c.mmgubun, c.qadiv, c.phonenumber, c.writeuser, c.finishuser" & vbcrlf
        sqlStr = sqlStr & " , c.contents_jupsu, c.regdate, c.finishyn, c.finishdate, c.retrydate, c.specialmemo" & vbcrlf
        sqlStr = sqlStr & " , right(c1.comm_cd,1) as comm_cd1, c1.comm_name as comm_name1, right(c2.comm_cd,2) as comm_cd2, c2.comm_name as comm_name2" & vbcrlf

		If (FRectWriteUser<>"") And (FRectIsFinished = "N") Then
			sqlStr = sqlStr & " , (case when T.orderserial is not NULL then 'Y' else 'N' end) as upchefinishYN " & vbcrlf
		Else
			sqlStr = sqlStr & " , '' as upchefinishYN " & vbcrlf
		End If

        sqlStr = sqlStr & " from [db_cs].[dbo].tbl_cs_memo c with (nolock)" & vbcrlf
        sqlStr = sqlStr & " left join db_cs.dbo.tbl_cs_comm_code c1 with (nolock)" & vbcrlf
        sqlStr = sqlStr & " 	on c1.comm_group = 'Z030'" & vbcrlf
        sqlStr = sqlStr & " 	and c.mmgubun = right(c1.comm_cd,1)" & vbcrlf
        sqlStr = sqlStr & " left join db_cs.dbo.tbl_cs_comm_code c2 with (nolock)" & vbcrlf
        sqlStr = sqlStr & " 	on c1.comm_cd = c2.comm_group" & vbcrlf
        sqlStr = sqlStr & " 	and c.qadiv = right(c2.comm_cd,2)" & vbcrlf

		If (FRectWriteUser<>"") And (FRectIsFinished = "N") Then
			'// ���� ��üó���Ϸ�
			sqlStr = sqlStr & " 	left join ( "
			sqlStr = sqlStr & " 		select a.orderserial "
			sqlStr = sqlStr & " 		from "
			sqlStr = sqlStr & " 			[db_cs].[dbo].[tbl_cs_memo] m with (nolock) "
			sqlStr = sqlStr & " 			join [db_cs].[dbo].[tbl_new_as_list] a with (nolock) "
			sqlStr = sqlStr & " 			on "
			sqlStr = sqlStr & " 				1 = 1 "
			sqlStr = sqlStr & " 				and m.writeuser = '" & FRectWriteUser & "' "
			sqlStr = sqlStr & " 				and m.finishyn <> 'Y' "
			sqlStr = sqlStr & " 				and m.orderserial = a.orderserial "
			sqlStr = sqlStr & " 				and DateDiff(d, a.finishdate, getdate()) < 1 "
			sqlStr = sqlStr & " 		where "
			sqlStr = sqlStr & " 			1 = 1 "
			sqlStr = sqlStr & " 			and a.requireupche = 'Y' "
			sqlStr = sqlStr & " 			and a.currstate >= 'B006' "
			sqlStr = sqlStr & " 		group by a.orderserial "
			sqlStr = sqlStr & " 	) T "
			sqlStr = sqlStr & " 	on "
			sqlStr = sqlStr & " 		c.orderserial = T.orderserial "
		End If

        sqlStr = sqlStr & " where 1 = 1 " & vbcrlf

		if (FRectRegStart <> "") then
			sqlStr = sqlStr & " and c.regdate>='" & FRectRegStart & "'" & vbcrlf
		end if

		if (FRectRegEnd <> "") then
			sqlStr = sqlStr & " and convert(varchar(10), c.regdate, 21)<='" & FRectRegEnd & "'" & vbcrlf
		end if

		if (FRectContents <> "") then
			sqlStr = sqlStr & " and c.contents_jupsu like '%" & html2db(FRectContents) & "%' " & vbcrlf

			if (FRectUserID = "") and (FRectOrderserial = "") and (FRectPhoneNumber = "") then
				tmpSql = "select (max(id) - 80000) as minIdx from [db_cs].[dbo].tbl_cs_memo with (nolock) "
				rsget.CursorLocation = adUseClient
    			rsget.Open tmpSql, dbget, adOpenForwardOnly, adLockReadOnly
				if  not rsget.EOF  then
					minIdx = rsget("minIdx")
					'sqlStr = sqlStr & " and id >= " & minIdx
				end if
				rsget.close
			end if
		end if

		if (FRectMMGubunExvlude <> "") and (FRectContents = "") and (FRectPhoneNumber = "") then
			'// ��ȭ��ȣ �Ǵ� ����˻��� SMS/�̸��� �ȳ� �޸� ���� ����
			'sqlStr = sqlStr & " and c.mmgubun not in (" & FRectMMGubunExvlude & ")" & vbcrlf
			sqlStr = sqlStr & " and c.mmgubun not in ('4','5')" & vbcrlf
		end if

		if (FRectRetryDateExclude <> "") then
			sqlStr = sqlStr & " and not (c.retrydate is not null and c.retrydate > getdate())" & vbcrlf
		end if

        if (FRectUserID <> "") then
            sqlStr = sqlStr & " and c.userid = '" + CStr(FRectUserID) + "' " & vbcrlf
        end if

        if (FRectOrderserial <> "") then
            sqlStr = sqlStr & " and c.orderserial = '" + CStr(FRectOrderserial) + "' " & vbcrlf
        end if

        if (FRectIsFinished <> "") then
            sqlStr = sqlStr & " and c.finishyn = '" + CStr(FRectIsFinished) + "' " & vbcrlf
        end if

	    if (FRectPhoneNumber<> "") then
			tmpStr = "[SMS " + FRectPhoneNumber + "]"
	         ''sqlStr = sqlStr & " and ((c.phonenumber = '" + CStr(FRectPhoneNumber) + "') or (c.contents_jupsu like '%" & html2db(FRectPhoneNumber) & "%')) " & vbcrlf
			 ''sqlStr = sqlStr & " and ((c.phonenumber = '" + CStr(FRectPhoneNumber) + "') or (c.contents_jupsu_varchar like '[[]SMS " & html2db(FRectPhoneNumber) & "%')) " & vbcrlf
			 sqlStr = sqlStr & " and (c.phonenumber = '" + CStr(FRectPhoneNumber) + "') " & vbcrlf
        end if

        if (FRectWriteUser<>"") then
            sqlStr = sqlStr & " and c.writeuser = '" + FRectWriteUser + "' " & vbcrlf
        end if

        if (FRectQadiv<>"") then
            sqlStr = sqlStr & " and c.qadiv = '" + FRectQadiv + "' " & vbcrlf
        end if

        if (FRectDivCD<>"") then
            sqlStr = sqlStr & " and c.divcd = '" + FRectDivCD + "' " & vbcrlf
        end if

        if (FRectMMGubun<>"") then
            sqlStr = sqlStr & " and c.mmgubun = '" + FRectMMGubun + "' " & vbcrlf
        end if

        sqlStr = sqlStr & " order by c.id desc " & vbcrlf

		''response.write "<!-- " & sqlStr & " -->" & "<br>"
		''response.end
		rsget.pagesize = FPageSize
		rsget.CursorLocation = adUseClient
    	rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))
		Redim preserve FItemList(FResultCount)
		i = 0
		If not rsget.EOF Then
			rsget.absolutepage = FCurrPage
			Do until rsget.EOF
				Set FItemList(i) = new CCSMemoItem
		            FItemList(i).Fid                = rsget("id")
		            FItemList(i).Forderserial       = rsget("orderserial")
		            FItemList(i).Fdivcd             = rsget("divcd")
		            FItemList(i).FmmGubun           = rsget("mmgubun")
		            FItemList(i).Fuserid            = rsget("userid")
		            FItemList(i).FphoneNumber       = rsget("phonenumber")
		            FItemList(i).Fwriteuser         = rsget("writeuser")
		            FItemList(i).Ffinishuser        = rsget("finishuser")
		            FItemList(i).Fcontents_jupsu    = db2html(rsget("contents_jupsu"))
		            FItemList(i).Fregdate           = rsget("regdate")
		            FItemList(i).Ffinishyn          = rsget("finishyn")
		            FItemList(i).Ffinishdate        = rsget("finishdate")
		            FItemList(i).Fqadiv             = rsget("qadiv")
		            FItemList(i).Fretrydate         = rsget("retrydate")
					FItemList(i).Fspecialmemo       = rsget("specialmemo")
					FItemList(i).fcomm_cd1       	= rsget("comm_cd1")
					FItemList(i).fcomm_name1       	= rsget("comm_name1")
					FItemList(i).fcomm_cd2       	= rsget("comm_cd2")
					FItemList(i).fcomm_name2       	= rsget("comm_name2")
					FItemList(i).FupchefinishYN		= rsget("upchefinishYN")

		            rsget.MoveNext
				i = i + 1
			Loop
        End If
        rsget.close
    end sub

    public Sub GetCSMemoDetail()
        dim i,sqlStr

        sqlStr = " select top 1 id, orderserial, divcd, userid, mmgubun, phonenumber, qadiv, writeuser, finishuser, contents_jupsu,"
        sqlStr = sqlStr & " regdate, finishyn, finishdate, retrydate, specialmemo "
        sqlStr = sqlStr & " from [db_cs].[dbo].tbl_cs_memo with (nolock) "
        sqlStr = sqlStr & " where 1 = 1 "

        if (FRectUserID <> "") then
            sqlStr = sqlStr & " and userid = '" + CStr(FRectUserID) + "' "
        end if

        if (FRectOrderserial <> "") then
            sqlStr = sqlStr & " and orderserial = '" + CStr(FRectOrderserial) + "' "
        end if

		if (FRectId <> "") then
            sqlStr = sqlStr & " and id = '" + CStr(FRectId) + "' "
        end if

        rsget.CursorLocation = adUseClient
    	rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
		'response.write sqlStr

        if  not rsget.EOF  then
			set FOneItem = new CCSMemoItem

			FOneItem.Fid                = rsget("id")
			FOneItem.Forderserial       = rsget("orderserial")
			FOneItem.Fdivcd             = rsget("divcd")
			FOneItem.FmmGubun       	= rsget("mmgubun")
			FOneItem.Fuserid        	= rsget("userid")
			FOneItem.FphoneNumber   	= rsget("phonenumber")
			FOneItem.Fwriteuser         = rsget("writeuser")
			FOneItem.Ffinishuser        = rsget("finishuser")
			FOneItem.Fcontents_jupsu    = db2html(rsget("contents_jupsu"))
			FOneItem.Fregdate           = rsget("regdate")
			FOneItem.Ffinishyn          = rsget("finishyn")
			FOneItem.Ffinishdate        = rsget("finishdate")
			FOneItem.Fqadiv        		= rsget("qadiv")
			FOneItem.Fretrydate        	= rsget("retrydate")
			FOneItem.Fspecialmemo     	= rsget("specialmemo")

        end if
        rsget.close
    end sub

    public Sub GetCSMemoBlankDetail()
        dim i,sqlStr

        set FOneItem = new CCSMemoItem
    end sub

    Private Sub Class_Initialize()
        FCurrPage       = 1
        FPageSize       = 20
        FResultCount    = 0
        FScrollCount    = 10
        FTotalCount     = 0
    End Sub
    Private Sub Class_Terminate()
    End Sub

    public Function HasPreScroll()
		HasPreScroll = StartScrollPage > 1
    end Function
    public Function HasNextScroll()
		HasNextScroll = FTotalPage > StartScrollPage + FScrollCount -1
    end Function
    public Function StartScrollPage()
		StartScrollPage = ((FCurrpage-1)\FScrollCount)*FScrollCount +1
    end Function
end Class

%>
