<%

dim TENBYTEN_SOCNAME : TENBYTEN_SOCNAME = "(��)�ٹ�����"
dim TENBYTEN_SOCNO : TENBYTEN_SOCNO = "211-87-00620"
dim TENBYTEN_SUBSOCNO : TENBYTEN_SUBSOCNO = ""
dim TENBYTEN_CEONAME : TENBYTEN_CEONAME = "������"
dim TENBYTEN_SOCADDR : TENBYTEN_SOCADDR = "����� ���α� ���з� 57 ȫ�ʹ��б� ���з�ķ�۽� ������ 14�� �ٹ�����"
dim TENBYTEN_SOCSTATUS : TENBYTEN_SOCSTATUS = "����,���Ҹſ�"
dim TENBYTEN_SOCEVENT : TENBYTEN_SOCEVENT = "���ڻ�ŷ���"
dim TENBYTEN_MANAGERNAME : TENBYTEN_MANAGERNAME = "��꼭�����"
dim TENBYTEN_MANAGERPHONE : TENBYTEN_MANAGERPHONE = "02-554-2033"
dim TENBYTEN_MANAGERMAIL : TENBYTEN_MANAGERMAIL = "accounts@10x10.co.kr"

public function getMayTaxDate(ipkumdate)
    getMayTaxDate = dateSerial(Year(date),Month(date),1)
    if IsNULL(ipkumdate) then Exit function

    if datediff("m",ipkumdate,date())=0 then
		'�Ա����� ����ް� ������ �Ա��Ϸ�
		getMayTaxDate = dateSerial(Year(ipkumdate),Month(ipkumdate),Day(ipkumdate))
	elseif datediff("m",ipkumdate,date())=1 and datediff("d",date(),dateSerial(year(date),month(date),5))>=0 then
		'�Ա����� �������̸鼭 ��� 5�� �����̶�� �Ա��Ϸ�
		getMayTaxDate = dateSerial(Year(ipkumdate),Month(ipkumdate),Day(ipkumdate))
	elseif datediff("m",ipkumdate,date())>1 and datediff("d",date(),dateSerial(year(date),month(date),5))>=0 then
	    '�Ա����� ������ ���� 5�������̸� ������ 1��
	    getMayTaxDate = DateAdd("m",-1,dateSerial(Year(date),Month(date),1))
	else
		'�׷��� ������ �ݿ� 1�Ϸ� ����
		getMayTaxDate = dateSerial(Year(date),Month(date),1)
	end if
end function

'##### ���ݰ�� ��û�� ���ڵ�¿� Ŭ���� #####
class CTaxItem

	public FtaxIdx
	public ForderIdx

	public Forderserial
	public Fcancelyn
	public Fsubtotalprice
	public FsumPaymentEtc

	public Fcstitle

	public Fuserid
	public Fitemname

	public FtotalPrice
	public FtotalTax
	public Fregdate
	public FisueYn
	public FneoTaxNo
	public FcurUserId
	public Fprintdate

	public FconfirmYn
	public FbusiIdx
	public FbusiNo
	public FbusiSubNo
	public FbusiName
	public FbusiCEOName
	public FbusiAddr
	public FbusiType
	public FbusiItem
	public FrepName
	public FrepEmail
	public FrepTel

	public FsupplyConfirmYn
	public FsupplyBusiIdx
	public FsupplyBusiNo
	public FsupplyBusiSubNo
	public FsupplyBusiName
	public FsupplyBusiCEOName
	public FsupplyBusiAddr
	public FsupplyBusiType
	public FsupplyBusiItem
	public FsupplyRepName
	public FsupplyRepEmail
	public FsupplyRepTel

	public FisueDate
	public Fipkumdate
	public Fbuyname

    public FdelYn

    public Fbilldiv

    public Ftaxtype


    public Freforderserial

	public Ftaxissuetype
	public FsellBizCd
	public Fselltype
	public FsellBizNm
	public FselltypeNm


	public Fminmultiorderidx
	public Fmultiordercnt

	public FsupplyGroupid
	public FsupplyGroupidCnt

    public Fgroupid
    public FgroupidCnt

	public Ftplcompanyid

	public FconsignYN
	public FissueMethod

	public function GetMultiOrderIdxSUM()
		dim strSql

		GetMultiOrderIdxSUM = ""

		if (Fmultiordercnt > 0) then
			GetMultiOrderIdxSUM = Fminmultiorderidx
			if (Fmultiordercnt > 1) then
				GetMultiOrderIdxSUM = GetMultiOrderIdxSUM & " �� " & (Fmultiordercnt - 1) & " ��"
			end if
		end if

	end function

	public function GetMultiOrderIdxList()
		dim strSql

		GetMultiOrderIdxList = ""

		strSql = "select matchlinkkey from db_order.[dbo].tbl_taxSheet_Match where taxidx = " & FtaxIdx & " order by matchlinkkey "
		rsget.Open strSql, dbget, 1

		if Not(rsget.EOF or rsget.BOF) then

			do until rsget.eof
				if (GetMultiOrderIdxList = "") then
					GetMultiOrderIdxList = rsget("matchlinkkey")
				else
					GetMultiOrderIdxList = GetMultiOrderIdxList & "," & rsget("matchlinkkey")
				end if

				rsget.moveNext
			loop
		end if
		rsget.close

	end function

	public function BillDivString()
		if Fbilldiv="01" then
			BillDivString ="�Һ���"
		elseif Fbilldiv="11" then
			BillDivString ="�Һ���(��ü��)"
		elseif Fbilldiv="02" then
			BillDivString ="������"
		elseif Fbilldiv="03" then
			BillDivString ="���θ��"
		elseif Fbilldiv="51" then
			BillDivString ="��Ÿ����"
		elseif Fbilldiv="52" then
			BillDivString ="���Ʒ���"
		elseif Fbilldiv="53" then
			BillDivString ="���̶��"
		elseif Fbilldiv="54" then
			BillDivString ="�ٹ����� ����"
		elseif Fbilldiv="55" then
			BillDivString ="�����÷�����"
		elseif Fbilldiv="99" then
			BillDivString ="��Ÿ����(3PL)"
		else
			BillDivString ="��Ÿ"
		end if
	end function

	public function BillDivCompany()
		if (Fbilldiv="52") then
			BillDivCompany ="���ش�����"
		elseif (Fbilldiv="53") then
			BillDivCompany ="���̶��"
		elseif (Fbilldiv="55") then
			BillDivCompany ="�����÷�����"
		elseif (Fbilldiv="99") then
			BillDivCompany ="3PL"
		else
			BillDivCompany ="�ٹ�����"
		end if
	end function

	public function TaxTypeString()
		if (Ftaxtype="Y") then
			TaxTypeString ="����"
		elseif (Ftaxtype="N") then
			TaxTypeString ="�鼼"
		elseif (Ftaxtype="0") then
			TaxTypeString ="����"
		else
			if ((FtotalTax <> "") and (CStr(FtotalTax) <> "0")) then
				TaxTypeString ="����"
			else
				TaxTypeString = Ftaxtype
			end if
		end if
	end function

	public function GetTaxIssueTypeName()
		if (Ftaxissuetype = "C") then
			GetTaxIssueTypeName ="�Һ��ڸ���(�ֹ�����)"
		elseif (Ftaxissuetype="E") then
			GetTaxIssueTypeName ="��Ÿ����(���곻��)"
		elseif (Ftaxissuetype="S") then
			GetTaxIssueTypeName ="��Ÿ����(�����)"
		elseif (Ftaxissuetype="X") then
			GetTaxIssueTypeName ="������"
		else
			GetTaxIssueTypeName = Ftaxissuetype
		end if
	end function

	public function GetConsignmentYN()
		if (FconsignYN = "Y") then
			GetConsignmentYN ="����Ź"
		elseif (FconsignYN = "N") or (FconsignYN = " ") then
			GetConsignmentYN ="����"
		else
			GetConsignmentYN = "aaa" & FconsignYN & "aaa"
		end if
	end function

	Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub

end Class


'##### ���ݰ�� ��û�� Ŭ���� #####
Class CTax
	public FTaxList()
	public FOneItem

	public FTotalCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FResultCount
	public FScrollCount
	public FRecttaxIdx
	public FRectsearchDiv
	public FRectsearchBilldiv
	public FRectsearchKey
	public FRectsearchString
	public FRectSdate
	public FRectEdate
	public FRectchkTerm

    public FRectDelYn
	public FRectConsignYN

	public FRectSupplyGroupID

	'// �⺻ ������ ����
	Private Sub Class_Initialize()
		redim preserve FTaxList(0)

		FCurrPage         = 1
		FPageSize         = 10
		FResultCount      = 0
		FScrollCount      = 10
		FTotalCount       = 0
	End Sub

	Private Sub Class_Terminate()

	End Sub


	'// ���ݰ�� ��û�� ��� ���
	public Sub GetTaxList()
		dim strSql, AddSQL, i

		'�˻� �߰� ����
		if FRectSearchKey <> "" and FRectSearchString <> "" then
			if FRectSearchKey="c.busiName" then
				AddSQL = AddSQL & " and " & FRectSearchKey & " like '%" & FRectSearchString & "%' "
			elseif FRectSearchKey="t.orderserial" then
				'// �������ݰ�꼭 ���԰˻�
				AddSQL = AddSQL & " and ((t.orderserial = '" & FRectSearchString & "') or (t.reforderserial = '" & FRectSearchString & "')) "
			else
				AddSQL = AddSQL & " and " & FRectSearchKey & " = '" & FRectSearchString & "' "
			end if
		end if

		if FRectsearchDiv<>"" then
			AddSQL = AddSQL & " and t.isueYn='" & FRectsearchDiv & "' "
		end if

		if FRectsearchBilldiv<>"" then
			AddSQL = AddSQL & " and t.billdiv='" & FRectsearchBilldiv & "' "
		end if

		if FRectchkTerm="Y" then
			AddSQL = AddSQL & " and t.isueDate between '" & FRectSdate & "' and '" & FRectEdate & "' "
		end if

		if (FRectDelYn<>"") then
			AddSQL = AddSQL & " and t.delYn='"&FRectDelYn&"'"
		end if

		if (FRectSupplyGroupID <> "") then
			AddSQL = AddSQL & " and IsNull(s.busiNo, '') in ( "
			AddSQL = AddSQL & " 	select company_no "
			AddSQL = AddSQL & " 	from db_partner.dbo.tbl_partner "
			AddSQL = AddSQL & " 	where IsNull(groupid, '') <> '' and IsNull(company_no, '') <> '' and IsNull(groupid, '') = '" + CStr(FRectSupplyGroupID) + "' "
			AddSQL = AddSQL & " ) "
		end if

		if (FRectConsignYN <> "") then
			AddSQL = AddSQL & " and IsNull(t.consignYN, 'N') = '" + CStr(FRectConsignYN) + "' "
		end if

		'@ �ѵ����ͼ�
		strSql = " Select count(t.taxIdx) as cnt "
		strSql = strSql + " from "
		strSql = strSql + " 	db_order.[dbo].tbl_taxSheet as t with (nolock)"
		strSql = strSql + " 	left Join db_order.[dbo].tbl_busiinfo as s with (nolock)"
		strSql = strSql + " 	on "
		strSql = strSql + " 		t.supplyBusiIdx=s.busiIdx "
		strSql = strSql + " 	Left Join db_order.[dbo].tbl_busiinfo as c with (nolock)"
		strSql = strSql + " 	on "
		strSql = strSql + " 		t.busiIdx=c.busiIdx "
		strSql = strSql + " 	left Join db_order.[dbo].tbl_order_master as o with (nolock)"
		strSql = strSql + " 	on "
		strSql = strSql + " 		t.orderIdx=o.idx "
		strSql = strSql + " where "
		strSql = strSql + " 	1 = 1 "
		strSql = strSql + AddSQL

		'response.write strSql & "<br>"
		rsget.CursorLocation = adUseClient
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
			FTotalCount = rsget("cnt")
		rsget.close

		strSql = " select  top " + CStr(CStr(FPageSize*FCurrPage)) + " "
		strSql = strSql + " 	t.taxIdx, t.orderIdx, t.orderserial, t.userid, t.itemname "
		strSql = strSql + " 	, t.totalPrice, t.totalTax, t.regdate, t.isueYn "
		strSql = strSql + " 	, t.neoTaxNo, t.curUserId, t.printdate, IsNull(t.taxtype, '') as taxtype, t.isueDate, t.delYn "
		strSql = strSql + " 	, t.supplyBusiIdx, t.busiIdx, IsNull(t.billdiv, '01') as billdiv "
		strSql = strSql + " 	, t.taxissuetype, t.reforderserial, isnull(t.sellBizCd,'') as sellBizCd, t.selltype, t.tplcompanyid "
		strSql = strSql + " 	, s.busiNo as supplyBusiNo, s.busiSubNo as supplyBusiSubNo, s.busiName as supplyBusiName, s.busiCEOName as supplyBusiCEOName, s.busiAddr as supplyBusiAddr, s.busiType as supplyBusiType, s.busiItem as supplyBusiItem, s.confirmYn as supplyConfirmYn "
		strSql = strSql + " 	, c.busiNo, c.busiSubNo, c.busiName, c.busiCEOName, c.busiAddr, c.busiType, c.busiItem, c.confirmYn "
		strSql = strSql + " 	, s.repName as supplyRepName, s.repEmail as supplyRepEmail, s.repTel as supplyRepTel "
		strSql = strSql + " 	, t.repName, t.repEmail, t.repTel "
		strSql = strSql + " 	, o.ipkumdate "
		strSql = strSql + " 	, (select min(matchlinkkey) from db_order.[dbo].tbl_taxSheet_Match with (nolock) where taxIdx = t.taxIdx) as minmultiorderidx "
		strSql = strSql + " 	, (select count(matchlinkkey) from db_order.[dbo].tbl_taxSheet_Match with (nolock) where taxIdx = t.taxIdx) as multiordercnt "
		strSql = strSql + " 	,( "
		strSql = strSql + " 		SELECT TOP 1 (case when g.company_no = '211-87-00620' then 'G00456' else g.groupid end) "
		strSql = strSql + " 		FROM db_partner.dbo.tbl_partner_group g with (nolock)"
		strSql = strSql + " 		WHERE g.company_no = s.busino "
		strSql = strSql + " 		) AS supplyGroupid "
		strSql = strSql + " 	,( "
		strSql = strSql + " 		SELECT (case when g.company_no = '211-87-00620' then 1 else count(*) end) "
		strSql = strSql + " 		FROM db_partner.dbo.tbl_partner_group g with (nolock)"
		strSql = strSql + " 		WHERE g.company_no = s.busino "
		strSql = strSql + " 		GROUP BY g.company_no "
		strSql = strSql + " 		) AS supplyGroupidCnt "
		strSql = strSql + " 	,( "
		strSql = strSql + " 		SELECT TOP 1 (case when g.company_no = '211-87-00620' then 'G00456' else g.groupid end) "
		strSql = strSql + " 		FROM db_partner.dbo.tbl_partner_group g with (nolock)"
		strSql = strSql + " 		WHERE g.company_no = c.busino "
		strSql = strSql + " 		) AS Groupid "
		strSql = strSql + " 	,( "
		strSql = strSql + " 		SELECT (case when g.company_no = '211-87-00620' then 1 else count(*) end) "
		strSql = strSql + " 		FROM db_partner.dbo.tbl_partner_group g with (nolock)"
		strSql = strSql + " 		WHERE g.company_no = c.busino "
		strSql = strSql + " 		GROUP BY g.company_no "
		strSql = strSql + " 		) AS GroupidCnt "
		strSql = strSql + " 	, b.bizsection_nm, p.pcomm_name, IsNull(t.consignYN, 'N') as consignYN, t.issueMethod "
		strSql = strSql + " from "
		strSql = strSql + " 	db_order.[dbo].tbl_taxSheet as t with (nolock)"
		strSql = strSql + " 	left Join db_order.[dbo].tbl_busiinfo as s with (nolock)"
		strSql = strSql + " 	on "
		strSql = strSql + " 		t.supplyBusiIdx=s.busiIdx "
		strSql = strSql + " 	Left Join db_order.[dbo].tbl_busiinfo as c with (nolock)"
		strSql = strSql + " 	on "
		strSql = strSql + " 		t.busiIdx=c.busiIdx "
		strSql = strSql + " 	left Join db_order.[dbo].tbl_order_master as o with (nolock)"
		strSql = strSql + " 	on "
		strSql = strSql + " 		t.orderIdx=o.idx "
		strSql = strSql + " 	left join db_partner.dbo.tbl_TMS_BA_BIZSECTION b with (nolock)"
		strSql = strSql + " 	on "
		strSql = strSql + " 		t.sellBizCd = b.bizsection_cd "
		strSql = strSql + " 	left join [db_partner].[dbo].tbl_partner_comm_code p with (nolock)"
		strSql = strSql + " 	on "
		strSql = strSql + " 		p.pcomm_group = 'sellacccd' and p.pcomm_cd = t.selltype "
		strSql = strSql + " where "
		strSql = strSql + " 	1 = 1 "

		strSql = strSql + AddSQL

		strSql = strSql + " order by "
		strSql = strSql + " 	t.taxidx desc "

		'response.write strSql & "<br>"
		rsget.pagesize = FPageSize
		rsget.CursorLocation = adUseClient
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly

		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

		redim FTaxList(FResultCount)

		if Not(rsget.EOF or rsget.BOF) then

		    i = 0
			rsget.AbsolutePage = FCurrPage
			do until rsget.eof
				set FTaxList(i) = new CTaxItem

				FTaxList(i).FtaxIdx				= rsget("taxIdx")
				FTaxList(i).ForderIdx			= rsget("orderIdx")
				FTaxList(i).Forderserial		= rsget("orderserial")
				FTaxList(i).Fuserid				= rsget("userid")
				FTaxList(i).Fitemname			= rsget("itemname")
				FTaxList(i).FtotalPrice			= rsget("totalPrice")
				FTaxList(i).FtotalTax			= rsget("totalTax")
				FTaxList(i).Fregdate			= rsget("regdate")
				FTaxList(i).FisueYn				= rsget("isueYn")
				FTaxList(i).FneoTaxNo			= rsget("neoTaxNo")
				FTaxList(i).FcurUserId			= rsget("curUserId")
				FTaxList(i).Fprintdate			= rsget("printdate")

				FTaxList(i).Ftaxtype			= rsget("taxtype")			'// ��������

				'// ������
				FTaxList(i).FsupplyConfirmYn	= rsget("supplyConfirmYn")
				FTaxList(i).FsupplyBusiIdx		= rsget("supplyBusiIdx")
				FTaxList(i).FsupplyBusiNo		= rsget("supplyBusiNo")
				FTaxList(i).FsupplyBusiSubNo	= rsget("supplyBusiSubNo")
				FTaxList(i).FsupplyBusiName		= rsget("supplyBusiName")
				FTaxList(i).FsupplyBusiCEOName	= rsget("supplyBusiCEOName")
				FTaxList(i).FsupplyBusiAddr		= rsget("supplyBusiAddr")
				FTaxList(i).FsupplyBusiType		= db2html(rsget("supplyBusiType"))
				FTaxList(i).FsupplyBusiItem		= db2html(rsget("supplyBusiItem"))
				FTaxList(i).FsupplyRepName		= rsget("supplyRepName")
				FTaxList(i).FsupplyRepEmail		= rsget("supplyRepEmail")
				FTaxList(i).FsupplyRepTel		= rsget("supplyRepTel")

				'// ���޹޴���
				FTaxList(i).FconfirmYn			= rsget("confirmYn")				'// ����ڵ������ �ѽ��� Ȯ���Ҷ� ����ߴ� ���(����� ����û ����ڹ�ȣ ��ȸ���񽺸� �̿��ϹǷ� ������)
				FTaxList(i).FbusiIdx			= rsget("busiIdx")
				FTaxList(i).FbusiNo				= rsget("busiNo")
				FTaxList(i).FbusiSubNo			= rsget("busiSubNo")
				FTaxList(i).FbusiName			= rsget("busiName")
				FTaxList(i).FbusiCEOName		= rsget("busiCEOName")
				FTaxList(i).FbusiAddr			= rsget("busiAddr")
				FTaxList(i).FbusiType			= db2html(rsget("busiType"))
				FTaxList(i).FbusiItem			= db2html(rsget("busiItem"))
				FTaxList(i).FrepName			= rsget("repName")
				FTaxList(i).FrepEmail			= rsget("repEmail")
				FTaxList(i).FrepTel				= rsget("repTel")

				FTaxList(i).Fipkumdate			= rsget("ipkumdate")
				FTaxList(i).FisueDate			= rsget("isueDate")
				FTaxList(i).FdelYn         		= rsget("delYn")
				FTaxList(i).Fbilldiv       		= rsget("billdiv")
				FTaxList(i).Freforderserial		= rsget("reforderserial")

				FTaxList(i).Ftaxissuetype  		= rsget("taxissuetype")
				FTaxList(i).FsellBizCd   		= rsget("sellBizCd")
				FTaxList(i).Fselltype   		= rsget("selltype")

				FTaxList(i).Fminmultiorderidx  	= rsget("minmultiorderidx")
				FTaxList(i).Fmultiordercnt   	= rsget("multiordercnt")

				FTaxList(i).FsellBizNm 			= rsget("bizsection_nm")
				FTaxList(i).FselltypeNm			= rsget("pcomm_name")

				'// ������
                FTaxList(i).FsupplyGroupid       = rsget("supplyGroupid")
                FTaxList(i).FsupplyGroupidCnt    = rsget("supplyGroupidCnt")

				'// ���޹޴���
                FTaxList(i).Fgroupid       		= rsget("groupid")
                FTaxList(i).FgroupidCnt     	= rsget("groupidCnt")

				FTaxList(i).Ftplcompanyid   	= rsget("tplcompanyid")

				FTaxList(i).FconsignYN   		= rsget("consignYN")
				FTaxList(i).FissueMethod   		= rsget("issueMethod")
				if ucase(FTaxList(i).FissueMethod)=ucase("bill36524") then FTaxList(i).FissueMethod="WEHAGO"

				i = i + 1
				rsget.moveNext
			loop
		end if
		rsget.close

	end Sub

	'// ���ݰ�� ��û�� ���� ����
	public Sub GetTaxRead()
		dim strSql

		strSql = " select  top 1 "
		strSql = strSql + " 	t.taxIdx, t.orderIdx, t.orderserial, t.userid, t.itemname "
		strSql = strSql + " 	, t.totalPrice, t.totalTax, t.regdate, t.isueYn "
		strSql = strSql + " 	, t.neoTaxNo, t.curUserId, t.printdate, IsNull(t.taxtype, '') as taxtype, t.isueDate, t.delYn "
		strSql = strSql + " 	, t.supplyBusiIdx, t.busiIdx, IsNull(t.billdiv, '01') as billdiv "
		strSql = strSql + " 	, t.taxissuetype, t.reforderserial, t.sellBizCd, t.selltype, t.tplcompanyid "
		strSql = strSql + " 	, s.busiNo as supplyBusiNo, s.busiSubNo as supplyBusiSubNo, s.busiName as supplyBusiName, s.busiCEOName as supplyBusiCEOName, s.busiAddr as supplyBusiAddr, s.busiType as supplyBusiType, s.busiItem as supplyBusiItem, s.confirmYn as supplyConfirmYn "
		strSql = strSql + " 	, c.busiNo, c.busiSubNo, c.busiName, c.busiCEOName, c.busiAddr, c.busiType, c.busiItem, c.confirmYn "
		strSql = strSql + " 	, s.repName as supplyRepName, s.repEmail as supplyRepEmail, s.repTel as supplyRepTel "
		strSql = strSql + " 	, t.repName, t.repEmail, t.repTel, IsNull(t.consignYN, 'N') as consignYN, IsNull(t.issueMethod, 'WEHAGO') as issueMethod "
		strSql = strSql + " 	, o.ipkumdate "
		strSql = strSql + " 	, (select min(matchlinkkey) from db_order.[dbo].tbl_taxSheet_Match with (nolock) where taxIdx = t.taxIdx) as minmultiorderidx "
		strSql = strSql + " 	, (select count(matchlinkkey) from db_order.[dbo].tbl_taxSheet_Match with (nolock) where taxIdx = t.taxIdx) as multiordercnt "
		strSql = strSql + " from "
		strSql = strSql + " 	db_order.[dbo].tbl_taxSheet as t with (nolock)"
		strSql = strSql + " 	left Join db_order.[dbo].tbl_busiinfo as s with (nolock) "
		strSql = strSql + " 	on "
		strSql = strSql + " 		t.supplyBusiIdx=s.busiIdx "
		strSql = strSql + " 	Join db_order.[dbo].tbl_busiinfo as c with (nolock) "
		strSql = strSql + " 	on "
		strSql = strSql + " 		t.busiIdx=c.busiIdx "
		strSql = strSql + " 	left Join db_order.[dbo].tbl_order_master as o with (nolock) "
		strSql = strSql + " 	on "
		strSql = strSql + " 		1 = 1 "
		strSql = strSql + " 		and ( "
		strSql = strSql + " 			(IsNull(t.orderIdx, 0) <> 0 and IsNull(t.orderIdx, 0) = o.idx) "
		strSql = strSql + " 			or "
		strSql = strSql + " 			(IsNull(t.orderserial, '') = o.orderserial) "
		strSql = strSql + " 		) "
		strSql = strSql + " where "
		strSql = strSql + " 	t.taxIdx = " + CStr(FRectTaxIdx) + " "

		'response.write strSql & "<br>"
		rsget.CursorLocation = adUseClient
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly

		if Not(rsget.EOF or rsget.BOF) then

			set FOneItem = new CTaxItem

			FOneItem.FtaxIdx		= rsget("taxIdx")
			FOneItem.ForderIdx		= rsget("orderIdx")
			FOneItem.Forderserial	= rsget("orderserial")
			FOneItem.Fuserid		= rsget("userid")
			FOneItem.Fitemname		= rsget("itemname")
			FOneItem.FtotalPrice	= rsget("totalPrice")
			FOneItem.FtotalTax		= rsget("totalTax")
			FOneItem.Fregdate		= rsget("regdate")
			FOneItem.FisueYn		= rsget("isueYn")
			FOneItem.FneoTaxNo		= rsget("neoTaxNo")
			FOneItem.FcurUserId		= rsget("curUserId")
			FOneItem.Fprintdate		= rsget("printdate")

			FOneItem.Ftaxtype		= rsget("taxtype")			'// ��������

			'// ������
			FOneItem.FsupplyConfirmYn		= rsget("supplyConfirmYn")
			FOneItem.FsupplyBusiIdx			= rsget("supplyBusiIdx")
			FOneItem.FsupplyBusiNo			= rsget("supplyBusiNo")
			FOneItem.FsupplyBusiSubNo		= rsget("supplyBusiSubNo")
			FOneItem.FsupplyBusiName		= rsget("supplyBusiName")
			FOneItem.FsupplyBusiCEOName		= rsget("supplyBusiCEOName")
			FOneItem.FsupplyBusiAddr		= rsget("supplyBusiAddr")
			FOneItem.FsupplyBusiType		= db2html(rsget("supplyBusiType"))
			FOneItem.FsupplyBusiItem		= db2html(rsget("supplyBusiItem"))
			FOneItem.FsupplyRepName			= rsget("supplyRepName")
			FOneItem.FsupplyRepEmail		= rsget("supplyRepEmail")
			FOneItem.FsupplyRepTel			= rsget("supplyRepTel")

			'// ���޹޴���
			FOneItem.FconfirmYn		= rsget("confirmYn")				'// ����ڵ������ �ѽ��� Ȯ���Ҷ� ����ߴ� ���(����� ����û ����ڹ�ȣ ��ȸ���񽺸� �̿��ϹǷ� ������)
			FOneItem.FbusiIdx		= rsget("busiIdx")
			FOneItem.FbusiNo		= rsget("busiNo")
			FOneItem.FbusiSubNo		= rsget("busiSubNo")
			FOneItem.FbusiName		= rsget("busiName")
			FOneItem.FbusiCEOName	= rsget("busiCEOName")
			FOneItem.FbusiAddr		= rsget("busiAddr")
			FOneItem.FbusiType		= db2html(rsget("busiType"))
			FOneItem.FbusiItem		= db2html(rsget("busiItem"))
			FOneItem.FrepName		= rsget("repName")
			FOneItem.FrepEmail		= rsget("repEmail")
			FOneItem.FrepTel		= rsget("repTel")

			FOneItem.Fipkumdate			= rsget("ipkumdate")
			FOneItem.FisueDate			= rsget("isueDate")
            FOneItem.FdelYn         	= rsget("delYn")
            FOneItem.Fbilldiv       	= rsget("billdiv")
            FOneItem.Freforderserial	= rsget("reforderserial")

            FOneItem.Ftaxissuetype  	= rsget("taxissuetype")
            FOneItem.FsellBizCd   		= rsget("sellBizCd")
            FOneItem.Fselltype   		= rsget("selltype")

            FOneItem.Fminmultiorderidx  = rsget("minmultiorderidx")
            FOneItem.Fmultiordercnt   	= rsget("multiordercnt")

			FOneItem.Ftplcompanyid   	= rsget("tplcompanyid")

			FOneItem.FconsignYN   		= rsget("consignYN")
			FOneItem.FissueMethod   	= rsget("issueMethod")
			if ucase(FOneItem.FissueMethod)=ucase("bill36524") then FOneItem.FissueMethod="WEHAGO"
		end if
		rsget.close

	end sub

	'// �� ���ݰ�꼭 ������(��ü��)
	public Sub GetTaxListUpche()
		dim strSql, AddSQL, i

		AddSQL = AddSQL & " and IsNull(t.isueYn, '') = 'Y' "
		AddSQL = AddSQL & " and t.billdiv='11' "
		AddSQL = AddSQL & " and t.delYn='N' "
		AddSQL = AddSQL & " and IsNull(s.busiNo, '') in ( "
		AddSQL = AddSQL & " 	select company_no "
		AddSQL = AddSQL & " 	from db_partner.dbo.tbl_partner "
		AddSQL = AddSQL & " 	where IsNull(groupid, '') <> '' and IsNull(company_no, '') <> '' and IsNull(groupid, '') = '" + CStr(FRectSupplyGroupID) + "' "
		AddSQL = AddSQL & " ) "
		AddSQL = AddSQL & " and t.isueDate >= '" & FRectSdate & "' and t.isueDate < '" & FRectEdate & "' "


		'// ====================================================================
		'@ �ѵ����ͼ�

		strSql = " Select count(t.taxIdx) as cnt "
		strSql = strSql + " from "
		strSql = strSql + " 	db_order.[dbo].tbl_taxSheet as t "
		strSql = strSql + " 	left Join db_order.[dbo].tbl_busiinfo as s "
		strSql = strSql + " 	on "
		strSql = strSql + " 		t.supplyBusiIdx=s.busiIdx "
		strSql = strSql + " 	Left Join db_order.[dbo].tbl_busiinfo as c "
		strSql = strSql + " 	on "
		strSql = strSql + " 		t.busiIdx=c.busiIdx "
		strSql = strSql + " 	left Join db_order.[dbo].tbl_order_master as o "
		strSql = strSql + " 	on "
		strSql = strSql + " 		t.orderIdx=o.idx "
		strSql = strSql + " where "
		strSql = strSql + " 	1 = 1 "
		strSql = strSql + AddSQL

		''response.write strSql
		rsget.Open strSql, dbget, 1
			FTotalCount = rsget("cnt")
		rsget.close


		'// ====================================================================
		strSql = " select  top " + CStr(CStr(FPageSize*FCurrPage)) + " "
		strSql = strSql + " 	t.taxIdx, t.orderIdx, t.orderserial, t.userid, t.itemname "
		strSql = strSql + " 	, t.totalPrice, t.totalTax, t.regdate, t.isueYn "
		strSql = strSql + " 	, t.neoTaxNo, t.curUserId, t.printdate, IsNull(t.taxtype, '') as taxtype, t.isueDate, t.delYn "
		strSql = strSql + " 	, t.supplyBusiIdx, t.busiIdx, IsNull(t.billdiv, '01') as billdiv "
		strSql = strSql + " 	, t.taxissuetype, t.reforderserial, t.sellBizCd, t.selltype, t.tplcompanyid "
		strSql = strSql + " 	, s.busiNo as supplyBusiNo, s.busiSubNo as supplyBusiSubNo, s.busiName as supplyBusiName, s.busiCEOName as supplyBusiCEOName, s.busiAddr as supplyBusiAddr, s.busiType as supplyBusiType, s.busiItem as supplyBusiItem, s.confirmYn as supplyConfirmYn "
		strSql = strSql + " 	, c.busiNo, c.busiSubNo, c.busiName, c.busiCEOName, c.busiAddr, c.busiType, c.busiItem, c.confirmYn "
		strSql = strSql + " 	, s.repName as supplyRepName, s.repEmail as supplyRepEmail, s.repTel as supplyRepTel "
		strSql = strSql + " 	, t.repName, t.repEmail, t.repTel "
		strSql = strSql + " 	, o.ipkumdate "
		strSql = strSql + " 	, b.bizsection_nm, p.pcomm_name, IsNull(t.consignYN, 'N') as consignYN, t.issueMethod "
		strSql = strSql + " from "
		strSql = strSql + " 	db_order.[dbo].tbl_taxSheet as t "
		strSql = strSql + " 	left Join db_order.[dbo].tbl_busiinfo as s "
		strSql = strSql + " 	on "
		strSql = strSql + " 		t.supplyBusiIdx=s.busiIdx "
		strSql = strSql + " 	Left Join db_order.[dbo].tbl_busiinfo as c "
		strSql = strSql + " 	on "
		strSql = strSql + " 		t.busiIdx=c.busiIdx "
		strSql = strSql + " 	left Join db_order.[dbo].tbl_order_master as o "
		strSql = strSql + " 	on "
		strSql = strSql + " 		t.orderIdx=o.idx "
		strSql = strSql + " 	left join db_partner.dbo.tbl_TMS_BA_BIZSECTION b "
		strSql = strSql + " 	on "
		strSql = strSql + " 		t.sellBizCd = b.bizsection_cd "
		strSql = strSql + " 	left join [db_partner].[dbo].tbl_partner_comm_code p "
		strSql = strSql + " 	on "
		strSql = strSql + " 		p.pcomm_group = 'sellacccd' and p.pcomm_cd = t.selltype "
		strSql = strSql + " where "
		strSql = strSql + " 	1 = 1 "

		strSql = strSql + AddSQL

		strSql = strSql + " order by "
		strSql = strSql + " 	t.taxidx desc "

		''response.write strSql
		rsget.pagesize = FPageSize
		rsget.Open strSql, dbget, 1

		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

		redim FTaxList(FResultCount)

		if Not(rsget.EOF or rsget.BOF) then

		    i = 0
			rsget.AbsolutePage = FCurrPage
			do until rsget.eof
				set FTaxList(i) = new CTaxItem

				FTaxList(i).FtaxIdx				= rsget("taxIdx")
				FTaxList(i).ForderIdx			= rsget("orderIdx")
				FTaxList(i).Forderserial		= rsget("orderserial")
				FTaxList(i).Fuserid				= rsget("userid")
				FTaxList(i).Fitemname			= rsget("itemname")
				FTaxList(i).FtotalPrice			= rsget("totalPrice")
				FTaxList(i).FtotalTax			= rsget("totalTax")
				FTaxList(i).Fregdate			= rsget("regdate")
				FTaxList(i).FisueYn				= rsget("isueYn")
				FTaxList(i).FneoTaxNo			= rsget("neoTaxNo")
				FTaxList(i).FcurUserId			= rsget("curUserId")
				FTaxList(i).Fprintdate			= rsget("printdate")

				FTaxList(i).Ftaxtype			= rsget("taxtype")			'// ��������

				'// ������
				FTaxList(i).FsupplyConfirmYn	= rsget("supplyConfirmYn")
				FTaxList(i).FsupplyBusiIdx		= rsget("supplyBusiIdx")
				FTaxList(i).FsupplyBusiNo		= rsget("supplyBusiNo")
				FTaxList(i).FsupplyBusiSubNo	= rsget("supplyBusiSubNo")
				FTaxList(i).FsupplyBusiName		= rsget("supplyBusiName")
				FTaxList(i).FsupplyBusiCEOName	= rsget("supplyBusiCEOName")
				FTaxList(i).FsupplyBusiAddr		= rsget("supplyBusiAddr")
				FTaxList(i).FsupplyBusiType		= db2html(rsget("supplyBusiType"))
				FTaxList(i).FsupplyBusiItem		= db2html(rsget("supplyBusiItem"))
				FTaxList(i).FsupplyRepName		= rsget("supplyRepName")
				FTaxList(i).FsupplyRepEmail		= rsget("supplyRepEmail")
				FTaxList(i).FsupplyRepTel		= rsget("supplyRepTel")

				'// ���޹޴���
				FTaxList(i).FconfirmYn			= rsget("confirmYn")				'// ����ڵ������ �ѽ��� Ȯ���Ҷ� ����ߴ� ���(����� ����û ����ڹ�ȣ ��ȸ���񽺸� �̿��ϹǷ� ������)
				FTaxList(i).FbusiIdx			= rsget("busiIdx")
				FTaxList(i).FbusiNo				= rsget("busiNo")
				FTaxList(i).FbusiSubNo			= rsget("busiSubNo")
				FTaxList(i).FbusiName			= rsget("busiName")
				FTaxList(i).FbusiCEOName		= rsget("busiCEOName")
				FTaxList(i).FbusiAddr			= rsget("busiAddr")
				FTaxList(i).FbusiType			= db2html(rsget("busiType"))
				FTaxList(i).FbusiItem			= db2html(rsget("busiItem"))
				FTaxList(i).FrepName			= rsget("repName")
				FTaxList(i).FrepEmail			= rsget("repEmail")
				FTaxList(i).FrepTel				= rsget("repTel")

				FTaxList(i).Fipkumdate			= rsget("ipkumdate")
				FTaxList(i).FisueDate			= rsget("isueDate")
				FTaxList(i).FdelYn         		= rsget("delYn")
				FTaxList(i).Fbilldiv       		= rsget("billdiv")
				FTaxList(i).Freforderserial		= rsget("reforderserial")

				FTaxList(i).Ftaxissuetype  		= rsget("taxissuetype")
				FTaxList(i).FsellBizCd   		= rsget("sellBizCd")
				FTaxList(i).Fselltype   		= rsget("selltype")

				''FTaxList(i).Fminmultiorderidx  	= rsget("minmultiorderidx")
				''FTaxList(i).Fmultiordercnt   	= rsget("multiordercnt")

				FTaxList(i).FsellBizNm 			= rsget("bizsection_nm")
				FTaxList(i).FselltypeNm			= rsget("pcomm_name")

				'// ������
                ''FTaxList(i).FsupplyGroupid       = rsget("supplyGroupid")
                ''FTaxList(i).FsupplyGroupidCnt    = rsget("supplyGroupidCnt")

				'// ���޹޴���
                ''FTaxList(i).Fgroupid       		= rsget("groupid")
                ''FTaxList(i).FgroupidCnt     	= rsget("groupidCnt")

				FTaxList(i).Ftplcompanyid   	= rsget("tplcompanyid")

				FTaxList(i).FconsignYN   		= rsget("consignYN")
				FTaxList(i).FissueMethod   		= rsget("issueMethod")

				i = i + 1
				rsget.moveNext
			loop
		end if
		rsget.close

	end Sub

	'// �������ݰ�꼭 ���� ��� ���
	public Sub GetAmendedTaxList()
		dim strSql, fromWhereSql, i

		fromWhereSql = " from "
		fromWhereSql = fromWhereSql + " 	db_order.dbo.tbl_order_master m "
		fromWhereSql = fromWhereSql + " 	join db_order.dbo.tbl_taxSheet t "
		fromWhereSql = fromWhereSql + " 	on "
		fromWhereSql = fromWhereSql + " 		m.orderserial = t.orderserial "
		fromWhereSql = fromWhereSql + " 	join db_order.[dbo].tbl_busiinfo b "
		fromWhereSql = fromWhereSql + " 	on "
		fromWhereSql = fromWhereSql + " 		t.busiIdx=b.busiIdx "
		fromWhereSql = fromWhereSql + " 	left join db_cs.dbo.tbl_new_as_list c "
		fromWhereSql = fromWhereSql + " 	on "
		fromWhereSql = fromWhereSql + " 		1 = 1 "
		fromWhereSql = fromWhereSql + " 		and m.orderserial = c.orderserial "
		fromWhereSql = fromWhereSql + " 		and t.delYn <> 'Y' "
		fromWhereSql = fromWhereSql + " where "
		fromWhereSql = fromWhereSql + " 	1 = 1 "
		fromWhereSql = fromWhereSql + " 	and m.cashreceiptreq in ('T', 'U') "
		fromWhereSql = fromWhereSql + " 	and c.divcd not in ('A900', 'A006', 'A000', 'A002', 'A008', 'A004', 'A011', 'A010', 'A700', 'A001') "
		fromWhereSql = fromWhereSql + " 	and c.currstate = 'B007' "

		if FRectSearchDiv <> "" then
			fromWhereSql = fromWhereSql + " 	and t.isueYn = '" + CStr(FRectSearchDiv) + "' "
		end if

		if FRectSearchBilldiv <> "" then
			fromWhereSql = fromWhereSql + " 	and t.billdiv = '" + CStr(FRectSearchBilldiv) + "' "
		end if

		if FRectSearchString<>"" then

			if FRectSearchKey="b.busiName" then
				fromWhereSql = fromWhereSql & " and " & FRectSearchKey & " like '%" & FRectSearchString & "%' "
			elseif FRectSearchKey="t.orderserial" then
				'// �������ݰ�꼭 ���԰˻�
				fromWhereSql = fromWhereSql & " and ((t.orderserial = '" & FRectSearchString & "') or (t.reforderserial = '" & FRectSearchString & "')) "
			else
				fromWhereSql = fromWhereSql & " and " & FRectSearchKey & " = '" & FRectSearchString & "' "
			end if

		end if

		if FRectChkTerm="Y" then
			fromWhereSql = fromWhereSql & " and t.isueDate between '" & FRectSdate & "' and '" & FRectEdate & "' "
		end if

        if (FRectDelYn<>"") then
			fromWhereSql = fromWhereSql & " and t.delYn='"&FRectDelYn&"'"
		end if

		'// ===================================================================
		'// �ѵ����ͼ�
		strSql = " select count(m.orderserial) as cnt "

		strSql = strSql + fromWhereSql

		rsget.Open strSql, dbget, 1
			FTotalCount = rsget("cnt")
		rsget.close

		'// ===================================================================
		'@ ������
		strSql = " select top " & CStr(FPageSize*FCurrPage) & " "
		strSql = strSql + " m.cancelyn, m.subtotalprice, m.sumPaymentEtc "

		strSql = strSql + " , c.title as cstitle "

		strSql = strSql + " , t.taxIdx, t.orderIdx, t.orderserial, t.userid "
		strSql = strSql + " , t.itemname "
		strSql = strSql + " , t.totalPrice, t.totalTax, t.regdate, t.isueYn, t.billdiv, b.confirmYn "
		strSql = strSql + " , t.isueDate, t.delYn, b.busiName, b.busiNo "
		strSql = strSql + " , t.repName, t.repEmail, t.repTel "
		strSql = strSql + " , b.busiCEOName, b.busiAddr, b.busiType, b.busiItem "

		strSql = strSql + fromWhereSql

		strSql = strSql + " order by m.orderserial desc, m.idx desc "

		'response.write strSql
		rsget.pagesize = FPageSize
		rsget.Open strSql, dbget, 1

		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

		redim FTaxList(FResultCount)

		if Not(rsget.EOF or rsget.BOF) then

		    i = 0
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FTaxList(i) = new CTaxItem

				FTaxList(i).FtaxIdx			= rsget("taxIdx")
				FTaxList(i).ForderIdx		= rsget("orderIdx")
				FTaxList(i).Forderserial	= rsget("orderserial")

				FTaxList(i).Fcancelyn		= rsget("cancelyn")
				FTaxList(i).Fsubtotalprice	= rsget("subtotalprice")
				FTaxList(i).FsumPaymentEtc	= rsget("sumPaymentEtc")

				FTaxList(i).Fcstitle		= rsget("cstitle")

				FTaxList(i).Fuserid			= rsget("userid")
				FTaxList(i).Fitemname		= rsget("itemname")
				FTaxList(i).FtotalPrice		= rsget("totalPrice")
				FTaxList(i).FtotalTax		= rsget("totalTax")
				FTaxList(i).Fregdate		= rsget("regdate")
				FTaxList(i).FisueYn			= rsget("isueYn")
				FTaxList(i).FconfirmYn		= rsget("confirmYn")
				FTaxList(i).FisueDate		= rsget("isueDate")

                FTaxList(i).FbusiNo        	= rsget("busiNo")
                FTaxList(i).FbusiName      	= rsget("busiName")
                FTaxList(i).FdelYn         	= rsget("delYn")

                FTaxList(i).Fbilldiv        = rsget("billdiv")

				FTaxList(i).FrepName		= rsget("repName")
				FTaxList(i).FrepEmail		= rsget("repEmail")
				FTaxList(i).FrepTel			= rsget("repTel")

				FTaxList(i).FbusiCEOName	= rsget("busiCEOName")
				FTaxList(i).FbusiAddr		= rsget("busiAddr")
				FTaxList(i).FbusiType		= db2html(rsget("busiType"))
				FTaxList(i).FbusiItem		= db2html(rsget("busiItem"))

				i = i + 1
				rsget.moveNext
			loop
		end if
		rsget.close

	end Sub

	public Sub GetTaxEmptyOne()
		set FOneItem = new CTaxItem
	end sub

	public FPrevID
	public FNextID

	'// ���� ������ �˻�
	public Function HasPreScroll()
		HasPreScroll = StartScrollPage > 1
	end Function

	'// ���� ������ �˻�
	public Function HasNextScroll()
		HasNextScroll = FTotalPage > StartScrollPage + FScrollCount -1
	end Function

	'// ù������ ����
	public Function StartScrollPage()
		StartScrollPage = ((FCurrpage-1)\FScrollCount)*FScrollCount +1
	end Function
end Class


'##### ��û�� ����Ʈ Ŭ���� #####
Class CTaxPrint
	public FTaxList()
	public FTotalCount
	public FRectChkPrint

	'// �⺻ ������ ����
	Private Sub Class_Initialize()
		redim preserve FTaxList(0)
		FTotalCount       = 0
	End Sub

	Private Sub Class_Terminate()

	End Sub


	'// ��û�� ����Ʈ ��� ���
	public Sub GetTaxPrint()
		Dim SQL, lp

		SQL =	" Select " &_
				"	t1.printdate, t1.curUserId " &_
				"	, t3.ipkumdate " &_
				"	, t2.busiName, t2.busiNo " &_
				"	, t1.repName, t1.repEmail, t1.repTel, t1.totalPrice, t1.itemname, t1.billdiv " &_
				"	, t3.buyname, t3.orderserial " &_
				"	, t2.busiAddr " &_
				" From db_order.[dbo].tbl_taxSheet as t1 " &_
				"	Join db_order.[dbo].tbl_busiInfo as t2 on t1.busiIdx=t2.busiIdx " &_
				"	Join db_order.[dbo].tbl_order_master as t3 on t1.orderIdx=t3.idx " &_
				" Where t1.taxIdx in (" & FRectChkPrint & ")"
		rsget.Open sql, dbget, 1

		'���ڵ� ��
		FTotalCount = rsget.RecordCount

		redim FTaxList(FTotalCount)

		if Not(rsget.EOF or rsget.BOF) then
		    lp = 0
			do until rsget.eof
				set FTaxList(lp) = new CTaxItem

				FTaxList(lp).Fprintdate		= rsget("printdate")
				FTaxList(lp).FcurUserId		= rsget("curUserId")
				FTaxList(lp).Fipkumdate		= rsget("ipkumdate")
				FTaxList(lp).FbusiName		= rsget("busiName")
				FTaxList(lp).FbusiNo		= rsget("busiNo")
				FTaxList(lp).FrepName		= rsget("repName")
				FTaxList(lp).FrepEmail		= rsget("repEmail")
				FTaxList(lp).FrepTel		= rsget("repTel")
				FTaxList(lp).FtotalPrice	= rsget("totalPrice")
				FTaxList(lp).Fitemname		= rsget("itemname")
				FTaxList(lp).Fbuyname		= rsget("buyname")
				FTaxList(lp).Forderserial	= rsget("orderserial")
				FTaxList(lp).FbusiAddr		= rsget("busiAddr")

				FTaxList(lp).Fbilldiv		= rsget("billdiv")

				lp=lp+1
				rsget.moveNext
			loop
		end if
		rsget.Close
	end sub
end Class


'// �߱� ��û �ִ���  Ȯ��
Function chkRegTax(ordSn)
	Dim SQL

	SQL = 	"Select isueYn " &_
			"From db_order.[dbo].tbl_taxSheet " &_
			"Where orderserial='" & ordSn & "'" &_
			"	and delYn='N'"
	rsget.Open sql, dbget, 1
		if rsget.EOF or rsget.BOF then
			chkRegTax = "none"
		else
			chkRegTax = rsget(0)
		end if
	rsget.Close

End Function

Function getOrderSerialPK(iorderserial)
    Dim sqlStr
    sqlStr = " IF (select count(*) from db_order.dbo.tbl_taxsheet"&VbCRLF
	sqlStr = sqlStr & " where orderserial='"&iorderserial&"')=1 "&VbCRLF
    sqlStr = sqlStr & " BEGIN "&VbCRLF
    sqlStr = sqlStr & " 	select '"&iorderserial&"' as ipk "&VbCRLF
    sqlStr = sqlStr & " END"
    sqlStr = sqlStr & " ELSE IF (select count(*) from db_order.dbo.tbl_taxsheet "&VbCRLF
    sqlStr = sqlStr & " 		where delyn='N'"&VbCRLF
    sqlStr = sqlStr & " 		and orderserial='"&iorderserial&"')=1 "&VbCRLF
    sqlStr = sqlStr & " BEGIN "&VbCRLF
    sqlStr = sqlStr & " 	select '"&iorderserial&"'+'_'+convert(varchar(10),taxidx)  as ipk "&VbCRLF
    sqlStr = sqlStr & " 	from db_order.dbo.tbl_taxsheet "&VbCRLF
    sqlStr = sqlStr & " 	where delyn='N' "&VbCRLF
    sqlStr = sqlStr & " 	and orderserial='"&iorderserial&"' "&VbCRLF
    sqlStr = sqlStr & " END "&VbCRLF
    sqlStr = sqlStr & " ELSE "&VbCRLF
    sqlStr = sqlStr & " BEGIN "&VbCRLF
    sqlStr = sqlStr & " 	select '' as ipk "&VbCRLF
    sqlStr = sqlStr & " END"

    rsget.Open sqlStr, dbget, 1
		if rsget.EOF or rsget.BOF then
			getOrderSerialPK = ""
		else
			getOrderSerialPK = rsget(0)
		end if
	rsget.Close

end function

%>
