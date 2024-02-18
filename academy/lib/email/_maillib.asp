<!-- #include virtual="/lib/util/htmllib.asp" -->
<%
sub sendmail(mailfrom, mailto, mailtitle, mailcontent)
        dim mailobject

        set mailobject=server.createobject("CDONTS.NewMail")
        mailobject.from = mailfrom
        mailobject.to = mailto
        mailobject.subject = mailtitle

        'html style
        mailobject.bodyformat = 0
        mailobject.mailformat = 0

        mailobject.body = mailcontent
        mailobject.send
        set mailobject = nothing
end sub

sub dsendmail(mailfrom, mailto, mailtitle, mailcontent)
        dim mailobject

        set mailobject=server.createobject("CDONTS.NewMail")
        mailobject.from = mailfrom
        mailobject.to = mailto
        mailobject.subject = mailtitle

        'html style
        mailobject.bodyformat = 0
        mailobject.mailformat = 0

        mailobject.body = mailcontent
        mailobject.send
        set mailobject = nothing
end sub

function sendmailnewuser2(mailto,userName) ' ���Ը��������� �о���̴� ������� ��ȯ
        dim mailfrom, mailtitle, mailcontent,dirPath,fileName
        dim fs,objFile

        mailfrom = "customer@10x10.co.kr"
        mailtitle = "10x10 ����Ʈ ������ ���� �帳�ϴ�."

        Set fs = Server.CreateObject("Scripting.FileSystemObject")
        dirPath = server.mappath("/lib/email")
        fileName = dirPath&"\\email_join.htm"
        Set objFile = fs.OpenTextFile(fileName,1)
        mailcontent = objFile.readall
        mailcontent = replace(mailcontent,":USERNAME:",userName)

        call sendmail(mailfrom, mailto, mailtitle, mailcontent)
        sendmailnewuser2 = mailcontent
end function


function sendmailorder2(orderserial)
        dim sql,discountrate,paymethod
        dim mailfrom, mailto, mailtitle, mailcontent,itemHtml,itemHtmlOri
        dim fs,objFile,dirPath,fileName,beforeItemHtml,afterItemHtml,itemHtmlTotal

        mailfrom = "customer@10x10.co.kr"
        mailtitle = "�ֹ��� ���������� �����Ǿ����ϴ�!"

        '�ֹ��� �����ּ� Ȯ��,�ֹ��ŷ����� ����
        sql = "select buyemail,accountdiv from [db_order].[10x10].tbl_order_master where orderserial = '" + orderserial + "'"
        rsget.Open sql,dbget,1
        if  not rsget.EOF  then
                rsget.Movefirst
                mailto = rsget("buyemail")
                paymethod = trim(rsget("accountdiv"))
        else
                exit function
        end if
        rsget.close

        ' ������ �ҷ��ͼ�
        Set fs = Server.CreateObject("Scripting.FileSystemObject")
        dirPath = server.mappath("/lib/email")
        if paymethod = "7" then    ' ������
            fileName = dirPath&"\\email_bank1.htm"
        elseif paymethod = "100" then   ' �ſ�ī��
            fileName = dirPath&"\\email_card1.htm"
        end if

        Set objFile = fs.OpenTextFile(fileName,1)
        mailcontent = objFile.readall
'        mailcontent = replace(mailcontent,":USERNAME:",userName)



        '�ֹ����� Ȯ��.
        sql = "select buyname,regdate, reqzipcode, (a.reqzipaddr + ' ' + a.reqaddress) as reqalladdress, a.totalcost, a.totalmileage, c.itemcost,a.discountrate,a.subtotalprice from [db_order].[10x10].tbl_order_master a, [db_order].[10x10].tbl_order_detail c"
        sql = sql + " where a.orderserial = '" + orderserial + "' and c.orderserial = '" + orderserial + "' and c.itemid = '0'"
        rsget.Open sql,dbget,1
        if  not rsget.EOF  then
                discountrate = rsget("discountrate")
                rsget.Movefirst
                mailcontent = replace(mailcontent,":SUBTOTALPRICE:", FormatCurrency(cstr(rsget("subtotalprice")))) ' �ֹ��Ѿ�
                mailcontent = replace(mailcontent,":TOTALITEMPRICE:",  FormatCurrency(cstr(rsget("subtotalprice") - rsget("itemcost"))) ) ' �ֹ��� ��item  ����
                mailcontent = replace(mailcontent,":DELIVERYFEE:",  FormatCurrency(cstr(rsget("itemcost"))) ) ' ��۱ݾ�
                mailcontent = replace(mailcontent,":BUYNAME:", rsget("buyname")) ' �ֹ��� �̸�
                mailcontent = replace(mailcontent,":ORDERSERIAL:", orderserial) ' �ֹ���ȣ
                mailcontent = replace(mailcontent,":REQZIPCODE:", rsget("reqzipcode")) ' ��ۿ����ȣ
                mailcontent = replace(mailcontent,":REQALLADDRESS:", rsget("reqalladdress")) ' ����ּ�
        else
                exit function
        end if
        rsget.close

        'item ���� �յںκ� ¥����
        beforeItemHtml = Left(mailcontent,InStr(mailcontent,":ITEMSTART:")-1)
        afterItemHtml = Mid(mailcontent,InStr(mailcontent,":ITEMEND:")+11)

        'item ������ �����κ� �ڸ���
        itemHtmlOri = Left(mailcontent,InStr(mailcontent,":ITEMEND:")-1)
        itemHtmlOri = Mid(itemHtmlOri,InStr(itemHtmlOri,":ITEMSTART:")+11)

        '�ֹ������� ���� Ȯ��.
        dim itemserial,inx
        sql = " select a.itemid, a.itemoptionname, c.listimage, c.itemname, (c.itemserial_large + c.itemserial_mid + c.itemserial_small) as itemserial, c.sellcash, a.itemno from [db_order].[10x10].tbl_order_detail a, [db_item].[10x10].tbl_item c"
        sql = sql + " where a.orderserial = '" + orderserial + "' and a.itemid <> '0' and c.itemid = a.itemid"
        inx = 1
        rsget.Open sql,dbget,1
        if  not rsget.EOF  then
                rsget.Movefirst
                do until rsget.eof
                        itemserial = rsget("itemserial") + "-" + FormatCode(rsget("itemid"))
                        itemHtml = replace(itemHtmlOri,":ITEMSERIAL:", itemserial) ' ��ǰ�ڵ�
                        itemHtml = replace(itemHtml,":ITEMNAME:", rsget("itemname")) ' ��ǰ�̸�
                        itemHtml = replace(itemHtml,":ITEMOPTIONNAME:", rsget("itemoptionname")) ' �ɼǸ�
                        if discountrate=1 then
                        	itemHtml = replace(itemHtml,":ITEMPRICE:",  CStr(rsget("sellcash"))) ' ��ǰ����
                        else
                        	itemHtml = replace(itemHtml,":ITEMPRICE:",  cstr(round(rsget("sellcash")*cdbl(discountrate)/100)*100) ) ' ��ǰ����
                    	end if
                        itemHtml = replace(itemHtml,":ITEMNO:", cstr(rsget("itemno"))) ' ����
                        itemHtml = replace(itemHtml,":IMGLIST:", cstr(  "0" + CStr(Clng(rsget("itemid")\10000)) + "/" + rsget("listimage"))) ' �̹���
                        if  inx mod 3 = 0 then
                            itemHtml = itemHtml + vbcr + "<tr></tr>"
                        end if
                        itemHtmlTotal = itemHtmlTotal & itemHtml

                inx = inx + 1
                rsget.movenext
                loop
        else
                exit function
        end if
        rsget.close

        mailcontent = beforeItemHtml & itemHtmlTotal & afterItemHtml



        call sendmail(mailfrom, mailto, mailtitle, mailcontent)
        sendmailorder2 = mailcontent
end function


function sendmailorder3(orderserial,mailfrom)
        dim sql,discountrate,paymethod
        dim mailto, mailtitle, mailcontent,itemHtml,itemHtmlOri
        dim fs,objFile,dirPath,fileName,beforeItemHtml,afterItemHtml,itemHtmlTotal


        mailtitle = "�ֹ��� ���������� �����Ǿ����ϴ�!"

        '�ֹ��� �����ּ� Ȯ��,�ֹ��ŷ����� ����---------------------------------------------------------------------------
        sql = "select buyemail,accountdiv from [db_order].[10x10].tbl_order_master where orderserial = '" + orderserial + "'"
        rsget.Open sql,dbget,1
        if  not rsget.EOF  then
                rsget.Movefirst
                mailto = rsget("buyemail")
                paymethod = trim(rsget("accountdiv"))
        else
                exit function
        end if
        rsget.close

        ' ������ �ҷ��ͼ� ---------------------------------------------------------------------------
        Set fs = Server.CreateObject("Scripting.FileSystemObject")
        dirPath = server.mappath("/lib/email")
        if paymethod = "7" then    ' ������
            fileName = dirPath&"\\email_bank1.htm"
        elseif paymethod = "100" then   ' �ſ�ī��
            fileName = dirPath&"\\email_card1.htm"
        elseif paymethod = "90" then   ' ��Ƽ�̵�
            fileName = dirPath&"\\email_multi1.htm"
        else
        	fileName = dirPath&"\\email_default.htm"
        end if

        Set objFile = fs.OpenTextFile(fileName,1)
        mailcontent = objFile.readall
'        mailcontent = replace(mailcontent,":USERNAME:",userName)


		dim SpendMile, tencardspend
        '�ֹ����� Ȯ��.---------------------------------------------------------------------------
        sql = "select buyname,regdate, reqname, reqzipcode, (a.reqzipaddr + ' ' + a.reqaddress) as reqalladdress, a.reqphone, a.totalcost, a.totalmileage, c.itemcost,a.discountrate,a.subtotalprice, a.miletotalprice ,a.tencardspend, a.comment from [db_order].[10x10].tbl_order_master a, [db_order].[10x10].tbl_order_detail c"
        sql = sql + " where a.orderserial = '" + orderserial + "' and c.orderserial = '" + orderserial + "' and c.itemid = '0'"

		rsget.Open sql,dbget,1
        if  not rsget.EOF  then
                discountrate = rsget("discountrate")
                tencardspend = rsget("tencardspend")
                rsget.Movefirst
                mailcontent = replace(mailcontent,":SUBTOTALPRICE:", formatNumber(FormatCurrency(cstr(rsget("subtotalprice"))),0) ) ' �ֹ��Ѿ�
                'mailcontent = replace(mailcontent,":TOTALITEMPRICE:",  FormatCurrency(cstr(Titemcost - rsget("itemcost"))) ) ' �ֹ��� ��item  ����
                mailcontent = replace(mailcontent,":DELIVERYFEE:",  formatNumber(FormatCurrency(cstr(rsget("itemcost"))),0) ) ' ��۱ݾ�
                mailcontent = replace(mailcontent,":BUYNAME:", rsget("buyname")) ' �ֹ��� �̸�
                mailcontent = replace(mailcontent,":ORDERSERIAL:", orderserial) ' �ֹ���ȣ
                mailcontent = replace(mailcontent,":REQNAME:", rsget("reqname")) ' ������ �̸�
                mailcontent = replace(mailcontent,":REQZIPCODE:", rsget("reqzipcode")) ' ��ۿ����ȣ
                mailcontent = replace(mailcontent,":REQALLADDRESS:", rsget("reqalladdress")) ' ����ּ�
                mailcontent = replace(mailcontent,":REQPHONE:", rsget("reqphone")) ' �ֹ��� ��ȭ��ȣ
                mailcontent = replace(mailcontent,":BEASONGMEMO:", rsget("comment")) ' ��۸޸�
                if IsNull(rsget("miletotalprice")) then
                	SpendMile =""
                else
                	SpendMile = rsget("miletotalprice") + tencardspend
                	SpendMile = formatNumber(FormatCurrency(SpendMile),0)
            	end if
            	mailcontent = replace(mailcontent,":SPENDMILEAGE:", SpendMile) ' ���ϸ���
		else
                exit function
        end if
        rsget.close


		'�ֹ������� ���� Ȯ��.-----------------------------------------------------------------------------
        dim itemserial,inx,sinx,einx
        dim Titemcost,BufCost

        Titemcost = 0

        sql = " select a.itemid, a.itemoptionname, c.smallimage, c.itemname," + vbcrlf
        sql = sql + " (c.itemserial_large + c.itemserial_mid + c.itemserial_small) as itemserial," + vbcrlf
        sql = sql + " a.itemcost as sellcash, a.itemno, a.isupchebeasong" + vbcrlf
        sql = sql + " from [db_order].[10x10].tbl_order_detail a," + vbcrlf
        sql = sql + " [db_item].[10x10].tbl_item c" + vbcrlf
        sql = sql + " where a.orderserial = '" + orderserial + "'" + vbcrlf
        sql = sql + " and a.itemid <> '0'" + vbcrlf
        sql = sql + " and c.itemid = a.itemid" + vbcrlf
        sql = sql + " and (a.cancelyn<>'Y')" + vbcrlf
        sql = sql + " order by a.isupchebeasong asc" + vbcrlf

        rsget.Open sql,dbget,1

        inx = 0
		  sinx = 1
		  einx = 0
itemHtml = "<table border='0' cellpadding='0' cellspacing='0'>"

        if  not rsget.EOF  then
                rsget.Movefirst
                do until rsget.eof

						  if inx = 0 then
								if rsget("isupchebeasong") = "N" then
									sinx = 0' �ٹ����ٹ���� ó������ɶ�
									einx = 1
								elseif rsget("isupchebeasong") = "Y" then
									sinx = 0'��ü����� ó������ɶ�
								end if
						  elseif einx = 1 and (rsget("isupchebeasong") = "Y") then
									einx = 0
									sinx = 0'�ٹ����ٹ�� �ѷ����� ��ü��� ó�� �ѷ��ٶ�
						  end if
if sinx = 0 then
itemHtml = itemHtml + "<tr>"
itemHtml = itemHtml + "<td>"
itemHtml = itemHtml + "<table width='550' border='0' cellspacing='1' cellpadding='0'>"
itemHtml = itemHtml + "<tr>"
if rsget("isupchebeasong") = "N" then
itemHtml = itemHtml + "<td align='left' valign='top'><img src='http://www.10x10.co.kr/lib/email/images/deliver_ten_t.gif' width='121' height='30'></td>"
else
itemHtml = itemHtml + "<td align='left' valign='top'><img src='http://www.10x10.co.kr/lib/email/images/deliver_upche_t.gif' width='121' height='30'></td>"
end if
itemHtml = itemHtml + "<td>&nbsp;</td>"
itemHtml = itemHtml + "</tr>"
itemHtml = itemHtml + "</table>"
itemHtml = itemHtml + "</td>"
itemHtml = itemHtml + "</tr>"
itemHtml = itemHtml + "<tr>"
itemHtml = itemHtml + "<td>"
itemHtml = itemHtml + "<table style='border-top: 1px solid #aaaaaa' border='0' cellpadding='0' cellspacing='0' height='4' bgcolor='ECECEC'width='550'>"
itemHtml = itemHtml + "<tr>"
itemHtml = itemHtml + "<td><img src='http://www.10x10.co.kr/lib/email/images/spacer.gif' width='550' height='4' align='center'></td>"
itemHtml = itemHtml + "</tr>"
itemHtml = itemHtml + "</table>"
itemHtml = itemHtml + "<table style='border-bottom: 1px solid #555555;'width='550' border='0' height='23' cellpadding='0' cellspacing='0'>"
itemHtml = itemHtml + "<tr>"
itemHtml = itemHtml + "<td width='50' class='p11' align='center'>��ǰ</td>"
itemHtml = itemHtml + "<td style='padding:3 6 3 6' class='p11' align='center'>��ǰ��</td>"
itemHtml = itemHtml + "<td style='padding:3 6 3 6' width='100' class='p11' align='center'>��ǰ�ڵ�</td>"
itemHtml = itemHtml + "<td style='padding:3 6 3 6' width='80' class='p11' align='center'>�ɼ�</td>"
itemHtml = itemHtml + "<td style='padding:3 6 3 6' width='30' class='p11' align='center'>����</td>"
itemHtml = itemHtml + "<td style='padding:3 6 3 6' width='80' class='p11' align='center'>����</td>"
itemHtml = itemHtml + "</tr>"
itemHtml = itemHtml + "</table>"
itemHtml = itemHtml + "</td>"
itemHtml = itemHtml + "</tr>"
end if


                        itemserial = rsget("itemserial") + "-" + FormatCode(rsget("itemid")) ' �����۹�ȣ

								if CDbl(discountrate)=1 then
                        	BufCost = rsget("sellcash") * rsget("itemno")
                        else
                        	BufCost = round(rsget("sellcash")*cdbl(discountrate)/100)*100 * rsget("itemno")
                    		end if
								Titemcost = Titemcost + BufCost '���ֹ���

itemHtml = itemHtml + "<tr>"
itemHtml = itemHtml + "<td>"
itemHtml = itemHtml + "<table style='border-bottom: 1px solid #c8c8c8' width='550' border='0' height='57' cellpadding='0' cellspacing='0'>"
itemHtml = itemHtml + "<tr>"
itemHtml = itemHtml + "<td width='50'><img src='http://webimage.10x10.co.kr/image/small/" + cstr( "0" + CStr(Clng(rsget("itemid")\10000)) + "/" + rsget("smallimage")) + "' width='50' height='50'></td>"
itemHtml = itemHtml + "<td style='padding:3 6 3 6'>" + db2html(rsget("itemname")) + "</td>"
itemHtml = itemHtml + "<td style='padding:3 6 3 6' width='100' align='center'>" + itemserial + "</td>"
itemHtml = itemHtml + "<td style='padding:3 6 3 6' width='80' align='center'>" + rsget("itemoptionname") + "</td>"
itemHtml = itemHtml + "<td style='padding:3 6 3 6' width='30' align='center'>" + Cstr(rsget("itemno")) + "ea</td>"
itemHtml = itemHtml + "<td style='padding:3 6 3 6' width='80' align='center'>" + CStr(BufCost) + "won</td>"
itemHtml = itemHtml + "</tr>"
itemHtml = itemHtml + "</table>"
itemHtml = itemHtml + "</td>"
itemHtml = itemHtml + "</tr>"

                inx = inx + 1
                sinx = sinx + 1
                rsget.movenext
                loop
        else
                exit function
        end if
        rsget.close

itemHtml = itemHtml + "</table>"

		itemHtmlTotal = replace(mailcontent,":INNERORDERTABLE:", itemHtml) ' �ֹ��������̺� �ֱ�

      mailcontent = itemHtmlTotal

		mailcontent = replace(mailcontent,":TOTALITEMPRICE:",  formatNumber(FormatCurrency(cstr(Titemcost)),0) ) ' �ֹ��� ��item  ����


        call sendmail(mailfrom, mailto, mailtitle, mailcontent)
        sendmailorder3 = mailcontent
end function

function ReSendmailorder(orderserial,mailfrom)
        sendmailorder3 orderserial,mailfrom
end function

function sendmailcome(orderserial) ' �������ɽ� ���� ������
        dim sql,discountrate,paymethod
        dim mailfrom, mailto, mailtitle, mailcontent,itemHtml,itemHtmlOri
        dim fs,objFile,dirPath,fileName,beforeItemHtml,afterItemHtml,itemHtmlTotal

        mailfrom = "customer@10x10.co.kr"
        mailtitle = "10X10 ���� �ȳ� �����Դϴ�!"

        '�ֹ��� �����ּ� Ȯ��,�ֹ��ŷ����� ����
        sql = "select buyemail,accountdiv from [db_order].[10x10].tbl_order_master where orderserial = '" + orderserial + "'"
        rsget.Open sql,dbget,1
        if  not rsget.EOF  then
                rsget.Movefirst
                mailto = rsget("buyemail")
        else
                exit function
        end if
        rsget.close

        ' ������ �ҷ��ͼ�
        Set fs = Server.CreateObject("Scripting.FileSystemObject")
        dirPath = server.mappath("/lib/email")
        fileName = dirPath&"\\email_come.htm"

        Set objFile = fs.OpenTextFile(fileName,1)
        mailcontent = objFile.readall

        '�ֹ����� Ȯ��.
        sql = "select buyname,regdate, reqzipcode, (a.reqzipaddr + ' ' + a.reqaddress) as reqalladdress, a.totalcost, a.totalmileage, c.itemcost,a.discountrate,a.subtotalprice from [db_order].[10x10].tbl_order_master a, [db_order].[10x10].tbl_order_detail c"
        sql = sql + " where a.orderserial = '" + orderserial + "' and c.orderserial = '" + orderserial + "' and c.itemid = '0'"
        rsget.Open sql,dbget,1
        if  not rsget.EOF  then
                discountrate = rsget("discountrate")
                rsget.Movefirst
                mailcontent = replace(mailcontent,":SUBTOTALPRICE:", FormatCurrency(cstr(rsget("subtotalprice")))) ' �ֹ��Ѿ�
                mailcontent = replace(mailcontent,":TOTALITEMPRICE:",  FormatCurrency(cstr(rsget("subtotalprice") - rsget("itemcost"))) ) ' �ֹ��� ��item  ����
                mailcontent = replace(mailcontent,":DELIVERYFEE:",  FormatCurrency(cstr(rsget("itemcost"))) ) ' ��۱ݾ�
                mailcontent = replace(mailcontent,":BUYNAME:", rsget("buyname")) ' �ֹ��� �̸�
                mailcontent = replace(mailcontent,":ORDERSERIAL:", orderserial) ' �ֹ���ȣ
                mailcontent = replace(mailcontent,":REQZIPCODE:", rsget("reqzipcode")) ' ��ۿ����ȣ
                mailcontent = replace(mailcontent,":REQALLADDRESS:", rsget("reqalladdress")) ' ����ּ�
        else
                exit function
        end if
        rsget.close

        'item ���� �յںκ� ¥����
        beforeItemHtml = Left(mailcontent,InStr(mailcontent,":ITEMSTART:")-1)
        afterItemHtml = Mid(mailcontent,InStr(mailcontent,":ITEMEND:")+11)

        'item ������ �����κ� �ڸ���
        itemHtmlOri = Left(mailcontent,InStr(mailcontent,":ITEMEND:")-1)
        itemHtmlOri = Mid(itemHtmlOri,InStr(itemHtmlOri,":ITEMSTART:")+11)

        '�ֹ������� ���� Ȯ��.
        dim itemserial,inx
        sql = " select a.itemid, a.itemoptionname, c.listimage, c.itemname,"
		sql = sql + " (c.itemserial_large + c.itemserial_mid + c.itemserial_small) as itemserial, c.sellcash, c.makerid, a.itemno"
		sql = sql + " from [db_order].[10x10].tbl_order_detail a, [db_item].[10x10].tbl_item c"
        sql = sql + " where a.orderserial = '" + orderserial + "' and a.itemid <> '0' and c.itemid = a.itemid"
        sql = sql + " and (a.cancelyn='N' or a.cancelyn='A')"
        inx = 1
        rsget.Open sql,dbget,1
        if  not rsget.EOF  then
                rsget.Movefirst
                do until rsget.eof
                        itemserial = rsget("itemserial") + "-" + FormatCode(rsget("itemid"))
                        itemHtml = replace(itemHtmlOri,":ITEMSERIAL:", itemserial) ' ��ǰ�ڵ�
                        itemHtml = replace(itemHtml,":ITEMNAME:", rsget("itemname")) ' ��ǰ�̸�
                        itemHtml = replace(itemHtml,":ITEMPRICE:",  cstr(rsget("sellcash")*cdbl(discountrate)) ) ' ��ǰ����
                        itemHtml = replace(itemHtml,":ITEMNO:", cstr(rsget("itemno"))) ' ����

						if rsget("itemoptionname") <> "" then
                        itemHtml = replace(itemHtml,":ITEMOPTIONNAME:", rsget("itemoptionname")) ' �ɼǸ�
						else
                        itemHtml = replace(itemHtml,":ITEMOPTIONNAME:", "-") ' �ɼǸ�
						end if

                        itemHtml = replace(itemHtml,":IMGLIST:", cstr( "0" + CStr(Clng(rsget("itemid")\10000)) + "/" + rsget("listimage"))) ' ��ǰ�̹���
                        itemHtml = replace(itemHtml,":MAKERID:", cstr(rsget("makerid"))) ' ����Ŀ���̵�
                        itemHtmlTotal = itemHtmlTotal & itemHtml

                inx = inx + 1
                rsget.movenext
                loop
        else
                exit function
        end if
        rsget.close

        mailcontent = beforeItemHtml & itemHtmlTotal & afterItemHtml

        call sendmail(mailfrom, mailto, mailtitle, mailcontent)
        sendmailcome = mailcontent
end function

function sendmailbankok(mailto,userName,orderserial) ' ��ۿϷ����
        dim sql,discountrate
        dim mailfrom, mailtitle, mailcontent
        dim fs,objFile,dirPath,fileName

        mailfrom = "customer@10x10.co.kr"
        mailtitle = "������ �Ա��� ���������� ó�� �Ǿ����ϴ�!"

        ' ������ �ҷ��ͼ�
        Set fs = Server.CreateObject("Scripting.FileSystemObject")
        dirPath = server.mappath("/lib/email")
        fileName = dirPath&"\\email_bank2.htm"
        Set objFile = fs.OpenTextFile(fileName,1)
        mailcontent = objFile.readall
        mailcontent = replace(mailcontent,":USERNAME:",userName)
        mailcontent = replace(mailcontent,":ORDERSERIAL:",orderserial)

        call sendmail(mailfrom, mailto, mailtitle, mailcontent)
'        sendmailbankok = mailcontent
end function

function sendmailfinish(orderserial,deliverno)
        dim sql,discountrate,paymethod
        dim mailfrom, mailto, mailtitle, mailcontent,itemHtml,itemHtmlOri
        dim fs,objFile,dirPath,fileName,beforeItemHtml,afterItemHtml,itemHtmlTotal
        dim subtotalprice
        mailfrom = "customer@10x10.co.kr"
        mailtitle = "�ֹ��Ͻ� ��ǰ�� ���� �ٹ����� ��۾ȳ��Դϴ�!"


        ' ������ �ҷ��ͼ�
        Set fs = Server.CreateObject("Scripting.FileSystemObject")
        dirPath = server.mappath("/lib/email")
        fileName = dirPath&"\\email_finish.htm"

        Set objFile = fs.OpenTextFile(fileName,1)
        mailcontent = objFile.readall

		sql = "select top 1 buyname,buyemail,subtotalprice from [db_order].[10x10].tbl_order_master"
		sql = sql + " where orderserial = '" + orderserial + "'"
		rsget.Open sql,dbget,1
		if  not rsget.EOF  then
			mailto = rsget("buyemail")
			subtotalprice = rsget("subtotalprice")
			mailcontent = replace(mailcontent,":BUYNAME:", db2html(rsget("buyname"))) ' �ֹ��� �̸�
			'if Left(deliverno,1)="6" then
			'	mailcontent = replace(mailcontent,":DELIVERNOWITHSRC:",  "http://www.cjgls.co.kr/contents/gls/gls004/gls004_06_01.asp?slipno=" + CStr(deliverno) ) ' ������ȣ
			'else
				mailcontent = replace(mailcontent,":DELIVERNOWITHSRC:",  "http://www.hydex.net/ehydex/jsp/home/distribution/tracking/trackingViewCus.jsp?InvNo=" + CStr(deliverno) ) ' ������ȣ
			'end if

			mailcontent = replace(mailcontent,":ORDERSERIAL:", orderserial) ' �ֹ���ȣ
		else
			exit function
		end if
		rsget.close



        'item ���� �յںκ� ¥����
'        beforeItemHtml = Left(mailcontent,InStr(mailcontent,":ITEMSTART:")-1)
'        afterItemHtml = Mid(mailcontent,InStr(mailcontent,":ITEMEND:")+11)

        'item ������ �����κ� �ڸ���
'        itemHtmlOri = Left(mailcontent,InStr(mailcontent,":ITEMEND:")-1)
'        itemHtmlOri = Mid(itemHtmlOri,InStr(itemHtmlOri,":ITEMSTART:")+11)

        '�ֹ������� ���� Ȯ��.
        dim itemserial,inx,sinx,einx
		  dim BaesongState
		  dim transco,transurl,songjangstr
'        sql = " select d.itemid, d.itemoptionname, m.imglist, d.itemname,"
'		   sql = sql + " d.itemcost, d.makerid, d.itemno"
'		   sql = sql + " from [db_order].[10x10].tbl_order_detail d"
'		   sql = sql + " left join [db_item].[10x10].tbl_item_image m on d.itemid=m.itemid"
'        sql = sql + " where d.orderserial = '" + orderserial + "'"
'        sql = sql + " and d.itemid <>0"
'        sql = sql + " and d.cancelyn<>'Y'"

        sql = " select a.itemid, a.itemoptionname, c.smallimage, c.itemname," + vbcrlf
        sql = sql + " (c.itemserial_large + c.itemserial_mid + c.itemserial_small) as itemserial," + vbcrlf
        sql = sql + " a.itemcost as sellcash, a.itemno, a.isupchebeasong, a.songjangdiv, isnull(a.songjangno,'') as songjangno, a.currstate" + vbcrlf
        sql = sql + " from [db_order].[10x10].tbl_order_detail a," + vbcrlf
        sql = sql + " [db_item].[10x10].tbl_item c" + vbcrlf
        sql = sql + " where a.orderserial = '" + orderserial + "'" + vbcrlf
        sql = sql + " and a.itemid <> '0'" + vbcrlf
        sql = sql + " and c.itemid = a.itemid" + vbcrlf
        sql = sql + " and (a.cancelyn<>'Y')" + vbcrlf
        sql = sql + " order by a.isupchebeasong asc" + vbcrlf

        inx = 0
		  sinx = 1
		  einx = 0
itemHtml = "<table border='0' cellpadding='0' cellspacing='0'>"

        rsget.Open sql,dbget,1
        if  not rsget.EOF  then
                rsget.Movefirst
                do until rsget.eof

						  if inx = 0 then
								if rsget("isupchebeasong") = "N" then
									sinx = 0' �ٹ����ٹ���� ó������ɶ�
									einx = 1
								elseif rsget("isupchebeasong") = "Y" then
									sinx = 0'��ü����� ó������ɶ�
								end if
						  elseif einx = 1 and (rsget("isupchebeasong") = "Y") then
									einx = 0
									sinx = 0'�ٹ����ٹ�� �ѷ����� ��ü��� ó�� �ѷ��ٶ�
						  end if
'response.write sinx & "<br>"
'response.write einx
'dbget.close()	:	response.End
if sinx = 0 then
itemHtml = itemHtml + "<tr>"
itemHtml = itemHtml + "<td>"
itemHtml = itemHtml + "<table width='550' border='0' cellspacing='1' cellpadding='0'>"
itemHtml = itemHtml + "<tr>"
if rsget("isupchebeasong") = "N" then
itemHtml = itemHtml + "<td align='left' valign='top'><img src='http://www.10x10.co.kr/lib/email/images/deliver_ten_t.gif' width='121' height='30'></td>"
else
itemHtml = itemHtml + "<td align='left' valign='top'><img src='http://www.10x10.co.kr/lib/email/images/deliver_upche_t.gif' width='121' height='30'></td>"
end if
itemHtml = itemHtml + "<td>&nbsp;</td>"
itemHtml = itemHtml + "</tr>"
itemHtml = itemHtml + "</table>"
itemHtml = itemHtml + "</td>"
itemHtml = itemHtml + "</tr>"
itemHtml = itemHtml + "<tr>"
itemHtml = itemHtml + "<td>"
itemHtml = itemHtml + "<table style='border-top: 1px solid #aaaaaa' border='0' cellpadding='0' cellspacing='0' height='4' bgcolor='ECECEC'width='550'>"
itemHtml = itemHtml + "<tr>"
itemHtml = itemHtml + "<td><img src='http://www.10x10.co.kr/lib/email/images/spacer.gif' width='550' height='4' align='center'></td>"
itemHtml = itemHtml + "</tr>"
itemHtml = itemHtml + "</table>"
itemHtml = itemHtml + "<table style='border-bottom: 1px solid #555555;'width='550' border='0' height='23' cellpadding='0' cellspacing='0'>"
itemHtml = itemHtml + "<tr>"
itemHtml = itemHtml + "<td width='50' class='p11' align='center'>��ǰ</td>"
itemHtml = itemHtml + "<td style='padding:3 6 3 6' class='p11' align='center'>��ǰ��</td>"
itemHtml = itemHtml + "<td style='padding:3 6 3 6' width='80' class='p11' align='center'>�ɼ�</td>"
itemHtml = itemHtml + "<td style='padding:3 6 3 6' width='30' class='p11' align='center'>����</td>"
itemHtml = itemHtml + "<td style='padding:3 6 3 6' width='80' class='p11' align='center'>�����Ȳ</td>"
itemHtml = itemHtml + "<td style='padding:3 6 3 6' width='100' class='p11' align='center'>�ù�/����</td>"
itemHtml = itemHtml + "</tr>"
itemHtml = itemHtml + "</table>"
itemHtml = itemHtml + "</td>"
itemHtml = itemHtml + "</tr>"
end if

'��ۻ��� ����
if rsget("isupchebeasong") = "N" then
		 BaesongState = "<font color='red'>���Ϸ�</font>"
else
	 if rsget("currstate") = 7 then
		 BaesongState = "<font color='red'>���Ϸ�</font>"
	 else
		 BaesongState = "<font color='#004080'>��ǰ�غ���</font>"
	 end if
end if

'�ù�� ����
if rsget("songjangdiv") = "1" then
transco = "�����ù�"
transurl = "http://www.hanjin.co.kr/transmission/main.htm"
elseif rsget("songjangdiv") = "2" then
transco = "�����ù�"
transurl = "http://www.hydex.net/ehydex/jsp/home/distribution/tracking/trackingViewCus.jsp?InvNo="
elseif rsget("songjangdiv") = "3" then
transco = "�������"
transurl = "http://doortodoor.korex.co.kr/jsp/cmn/index.jsp"
elseif rsget("songjangdiv") = "4" then
transco = "CJ GLS"
transurl = "http://www.cjgls.co.kr"
elseif rsget("songjangdiv") = "5" then
transco = "��Ŭ����"
transurl = "http://www.ecline.net/tracking/customer02.html#t01"
elseif rsget("songjangdiv") = "6" then
transco = "HTH"
transurl = "https://samsunghth.com/homepage/searchTraceGoods/SearchTraceResult.jhtml?dtdShtno="
elseif rsget("songjangdiv") = "7" then
transco = "�ѹ̸��ù�"
transurl = "http://www.e-family.co.kr/"
elseif rsget("songjangdiv") = "8" then
transco = "��ü��"
transurl = "http://parcel.epost.go.kr"
elseif rsget("songjangdiv") = "9" then
transco = "KGB"
transurl = "http://www.kgbl.co.kr/"
elseif rsget("songjangdiv") = "10" then
transco = "�����ù�"
transurl = "http://www.ajulogis.co.kr/"
elseif rsget("songjangdiv") = "11" then
transco = "�������ù�"
transurl = "http://www.10x10.co.kr/cscenter/csmain.asp"
elseif rsget("songjangdiv") = "12" then
transco = "�ѱ��ù�"
transurl = "http://www.kls.co.kr/"
elseif rsget("songjangdiv") = "13" then
transco = "���ο�ĸ"
transurl = "http://www.yellowcap.co.kr/"
elseif rsget("songjangdiv") = "14" then
transco = "���̽��ù�"
transurl = "http://www.10x10.co.kr/cscenter/csmain.asp"
elseif rsget("songjangdiv") = "15" then
transco = "�߾��ù�"
transurl = "http://www.10x10.co.kr/cscenter/csmain.asp"
elseif rsget("songjangdiv") = "16" then
transco = "�����ù�"
transurl = "http://www.10x10.co.kr/cscenter/csmain.asp"
elseif rsget("songjangdiv") = "17" then
transco = "Ʈ����ù�"
transurl = "http://www.etranet.co.kr/"
elseif rsget("songjangdiv") = "18" then
transco = "�����ù�"
transurl = "http://www.ilogen.com/"
elseif rsget("songjangdiv") = "19" then
transco = "KGBƯ���ù�"
transurl = "http://www.ikgb.co.kr/"
elseif rsget("songjangdiv") = "20" then
transco = "KT������"
transurl = "http://www.kls.co.kr/customer/cus_trace_01.asp"
elseif rsget("songjangdiv") = "21" then
transco = "�浿�ù�"
transurl = "http://www.kdexp.com"
else
transco = "��Ÿ"
transurl = "http://www.10x10.co.kr/cscenter/csmain.asp"
end if

'�ù�/���� ����
if rsget("isupchebeasong") = "N" then
	songjangstr =  "�����ù�<br>(<a href='http://www.hydex.net/ehydex/jsp/home/distribution/tracking/trackingViewCus.jsp?InvNo=" + Cstr(deliverno) + "' target='_blank'>" + Cstr(deliverno) + "</a>)"
else
	 If rsget("songjangdiv") = "2" Then
		  if rsget("songjangno")<>"" or isnull(rsget("songjangno")) then
			  songjangstr =  "�����ù�<br>(<a href='http://www.hydex.net/ehydex/jsp/home/distribution/tracking/trackingViewCus.jsp?InvNo=" + Cstr(rsget("songjangno")) + "' target='_blank'>" + rsget("songjangno") + "</a>)"
		  else
			  songjangstr="-"
		  end if
	 Else
		  if rsget("songjangno")<>"" or isnull(rsget("songjangno")) then
			  songjangstr = transco + "<br>(<a href='" + transurl + "' target='_blank'>" + rsget("songjangno") + "</a>)"
		  else
			  songjangstr="-"
		  end If
	 End If
end if

itemHtml = itemHtml + "<tr>"
itemHtml = itemHtml + "<td>"
itemHtml = itemHtml + "<table style='border-bottom: 1px solid #c8c8c8' width='550' border='0' height='57' cellpadding='0' cellspacing='0'>"
itemHtml = itemHtml + "<tr>"
itemHtml = itemHtml + "<td width='50'><img src='http://webimage.10x10.co.kr/image/small/" + cstr( "0" + CStr(Clng(rsget("itemid")\10000)) + "/" + rsget("smallimage")) + "' width='50' height='50'></td>"
itemHtml = itemHtml + "<td style='padding:3 6 3 6'>" + db2html(rsget("itemname")) + "</td>"
itemHtml = itemHtml + "<td style='padding:3 6 3 6' width='80' align='center'>" + rsget("itemoptionname") + "</td>"
itemHtml = itemHtml + "<td style='padding:3 6 3 6' width='30' align='center'>" + Cstr(rsget("itemno")) + "ea</td>"
itemHtml = itemHtml + "<td style='padding:3 6 3 6' width='80' align='center'>" + BaesongState + "</td>"
itemHtml = itemHtml + "<td style='padding:3 6 3 6' width='100' align='center'>" + songjangstr + "</td>"
itemHtml = itemHtml + "</tr>"
itemHtml = itemHtml + "</table>"
itemHtml = itemHtml + "</td>"
itemHtml = itemHtml + "</tr>"


                inx = inx + 1
                sinx = sinx + 1
                rsget.movenext
                loop
        else
                exit function
        end if
        rsget.close

		itemHtml = itemHtml + "</table>"

		itemHtmlTotal = replace(mailcontent,":INNERORDERTABLE:", itemHtml) ' �ֹ��������̺� �ֱ�

      mailcontent = itemHtmlTotal

        call sendmail(mailfrom, mailto, mailtitle, mailcontent)
        sendmailfinish = mailcontent
end function


function sendmailfinish_old(orderserial,deliverno)
        dim sql,discountrate,paymethod
        dim mailfrom, mailto, mailtitle, mailcontent,itemHtml,itemHtmlOri
        dim fs,objFile,dirPath,fileName,beforeItemHtml,afterItemHtml,itemHtmlTotal
        dim subtotalprice
        mailfrom = "customer@10x10.co.kr"
        mailtitle = "�ֹ��Ͻ� ��ǰ�� ���� �ٹ����� ��۾ȳ��Դϴ�!"
        '�ֹ��� �����ּ� Ȯ��,�ֹ��ŷ����� ����
        sql = "select buyemail,discountrate,subtotalprice from [db_order].[10x10].tbl_order_master where orderserial = '" + orderserial + "'"
        rsget.Open sql,dbget,1
        if  not rsget.EOF  then
                rsget.Movefirst
                mailto = rsget("buyemail")
                discountrate = rsget("discountrate")
                subtotalprice = rsget("subtotalprice")
        else
                exit function
        end if
        rsget.close

        ' ������ �ҷ��ͼ�
        Set fs = Server.CreateObject("Scripting.FileSystemObject")
        dirPath = server.mappath("/lib/email")
        fileName = dirPath&"\\email_finish.htm"

        Set objFile = fs.OpenTextFile(fileName,1)
        mailcontent = objFile.readall


        '�ֹ����� Ȯ��.
        sql = "select buyname,regdate, reqzipcode, "
        sql = sql + " (a.reqzipaddr + ' ' + a.reqaddress) as reqalladdress, a.totalcost, a.totalmileage, c.itemcost,a.discountrate,a.subtotalprice "
        sql = sql + " from [db_order].[10x10].tbl_order_master a,  [db_order].[10x10].tbl_order_detail c"
        sql = sql + " where a.orderserial = '" + orderserial + "' and c.orderserial = '" + orderserial + "' and c.itemid = '0'"
        rsget.Open sql,dbget,1
        if  not rsget.EOF  then
                discountrate = rsget("discountrate")
                rsget.Movefirst
                mailcontent = replace(mailcontent,":SUBTOTALPRICE:", FormatCurrency(cstr(rsget("subtotalprice")))) ' �ֹ��Ѿ�
                mailcontent = replace(mailcontent,":TOTALITEMPRICE:",  FormatCurrency(cstr(rsget("subtotalprice") - rsget("itemcost"))) ) ' �ֹ��� ��item  ����
                mailcontent = replace(mailcontent,":DELIVERYFEE:",  FormatCurrency(cstr(rsget("itemcost"))) ) ' ��۱ݾ�

                'if (Left(deliverno,1)="6") then
                	mailcontent = replace(mailcontent,":DELIVERNOWITHSRC:",  "http://www.hydex.net/ehydex/jsp/home/distribution/tracking/trackingViewCus.jsp?InvNo=" + CStr(deliverno) ) ' ������ȣ
                'else
                '	mailcontent = replace(mailcontent,":DELIVERNOWITHSRC:",  "http://www.doortodoor.co.kr/html/parcels/Tracking/TrackingResult.asp?TDNUM=" + CStr(deliverno) ) ' ������ȣ
                'end if

                mailcontent = replace(mailcontent,":DELIVERNO:",  deliverno ) ' ������ȣ
                mailcontent = replace(mailcontent,":BUYNAME:", rsget("buyname")) ' �ֹ��� �̸�
                mailcontent = replace(mailcontent,":ORDERSERIAL:", orderserial) ' �ֹ���ȣ
                mailcontent = replace(mailcontent,":REQZIPCODE:", rsget("reqzipcode")) ' ��ۿ����ȣ
                mailcontent = replace(mailcontent,":REQALLADDRESS:", rsget("reqalladdress")) ' ����ּ�


        else
                exit function
        end if
        rsget.close

        'item ���� �յںκ� ¥����
        beforeItemHtml = Left(mailcontent,InStr(mailcontent,":ITEMSTART:")-1)
        afterItemHtml = Mid(mailcontent,InStr(mailcontent,":ITEMEND:")+11)

        'item ������ �����κ� �ڸ���
        itemHtmlOri = Left(mailcontent,InStr(mailcontent,":ITEMEND:")-1)
        itemHtmlOri = Mid(itemHtmlOri,InStr(itemHtmlOri,":ITEMSTART:")+11)

        '�ֹ������� ���� Ȯ��.
        dim itemserial,inx
        sql = " select a.itemid, a.itemoptionname, c.listimage, c.itemname,"
		sql = sql + " (c.itemserial_large + c.itemserial_mid + c.itemserial_small) as itemserial, c.sellcash, c.makerid, a.itemno"
		sql = sql + " from [db_order].[10x10].tbl_order_detail a, [db_item].[10x10].tbl_item c"
        sql = sql + " where a.orderserial = '" + orderserial + "' and a.itemid <> '0' and c.itemid = a.itemid"
        sql = sql + " and (a.cancelyn='N' or a.cancelyn='A')"
        inx = 1
        rsget.Open sql,dbget,1
        if  not rsget.EOF  then
                rsget.Movefirst
                do until rsget.eof
                        itemserial = rsget("itemserial") + "-" + FormatCode(rsget("itemid"))
                        itemHtml = replace(itemHtmlOri,":ITEMSERIAL:", itemserial) ' ��ǰ�ڵ�
                        itemHtml = replace(itemHtml,":ITEMNAME:", rsget("itemname")) ' ��ǰ�̸�
                        itemHtml = replace(itemHtml,":ITEMPRICE:",  cstr(rsget("sellcash")*cdbl(discountrate)) ) ' ��ǰ����
                        itemHtml = replace(itemHtml,":ITEMNO:", cstr(rsget("itemno"))) ' ����

						if rsget("itemoptionname") <> "" then
                        itemHtml = replace(itemHtml,":ITEMOPTIONNAME:", rsget("itemoptionname")) ' �ɼǸ�
						else
                        itemHtml = replace(itemHtml,":ITEMOPTIONNAME:", "-") ' �ɼǸ�
						end if

                        itemHtml = replace(itemHtml,":IMGLIST:", cstr( "0" + CStr(Clng(rsget("itemid")\10000)) + "/" + rsget("listimage"))) ' ��ǰ�̹���
                        itemHtml = replace(itemHtml,":MAKERID:", cstr(rsget("makerid"))) ' ��ǰ�̹���

                        itemHtmlTotal = itemHtmlTotal & itemHtml

                inx = inx + 1
                rsget.movenext
                loop
        else
                exit function
        end if
        rsget.close

        mailcontent = beforeItemHtml & itemHtmlTotal & afterItemHtml

        call sendmail(mailfrom, mailto, mailtitle, mailcontent)
        sendmailfinish_old = mailcontent
end function

function sendmailfinish_ting(orderserial,deliverno)
		on error resume next

        dim sql
        dim mailfrom, mailto, mailtitle, mailcontent
        dim itemid, imglist
        dim buyname,itemname, itemoption

        dim fs,objFile,dirPath,fileName,beforeItemHtml,afterItemHtml,itemHtmlTotal

        mailfrom = "tingmart@011ting.com"
        mailtitle = "�ֹ��Ͻ� ��ǰ�� ���� �ø�Ʈ ��۾ȳ��Դϴ�!"

        '�ֹ��� �����ּ� Ȯ��,�ֹ��ŷ����� ����
        sql = "select * from [db_ting].[dbo].tbl_new_ting_orderhistory where orderserial = '" + orderserial + "'"
        rsget.Open sql,dbget,1
        if  not rsget.EOF  then
        		rsget.movefirst

                mailto = rsget("buyemail")
                buyname = rsget("buyname")
                itemname = rsget("itemname")
                itemid = rsget("itemid")
                itemoption = rsget("itemoption")

                rsget.Close

                sql = "select m.imglist , IsNull(o.codeview,'-') as optname, m.itemid"
                sql = sql + " from [db_item].[10x10].tbl_item_image m"
                sql = sql + " left join [db_item].[10x10].vw_all_option o on o.optioncode='" + CStr(itemoption) + "'"
                sql = sql + " where m.itemid=" + CStr(itemid)

                rsget.Open sql,dbget,1
                if Not rsget.Eof then
                	imglist = "http://image.10x10.co.kr/image/list/" + GetImageSubFolderByItemid(itemid) + "/" + rsget("listimage")
                	itemoption = rsget("optname")
                end if
                rsget.Close
        else
        	rsget.Close
                exit function
        end if


        ' ������ �ҷ��ͼ�
        Set fs = Server.CreateObject("Scripting.FileSystemObject")
        dirPath = server.mappath("/ext/ting/mail")
        fileName = dirPath & "\\email_finish_ting.htm"

        Set objFile = fs.OpenTextFile(fileName,1)
        mailcontent = objFile.readall

        mailcontent = replace(mailcontent,"[IBUYNAME]",buyname)
        mailcontent = replace(mailcontent,"[ILISTIMAGE]",imglist)
        mailcontent = replace(mailcontent,"[IITEMNAME]",itemname)
        mailcontent = replace(mailcontent,"[IOPTION]",itemoption)
        mailcontent = replace(mailcontent,"[ISONGJANG]",deliverno)

		if (Left(deliverno,1)="6") then
			mailcontent = replace(mailcontent,"[ISONGJANGWITSRC]","http://www.cjgls.co.kr/contents/gls/gls004/gls004_06_01.asp?slipno=" + deliverno)
		else
			mailcontent = replace(mailcontent,"[ISONGJANGWITSRC]","http://www.doortodoor.co.kr/html/parcels/Tracking/TrackingResult.asp?TDNUM=" + deliverno)
		end if

        call sendmail(mailfrom, mailto, mailtitle, mailcontent)
        sendmailfinish_ting = mailcontent

        if err then
        	response.write err.description
        end if
end function

function sendmailsearchpass(mailto,userName,imsipass)
        dim mailfrom, mailtitle, mailcontent,dirPath,fileName
        dim fs,objFile

        mailfrom = "customer@10x10.co.kr"
        mailtitle = "[10x10] " + userName + "���� �ӽú�й�ȣ �Դϴ�."

        Set fs = Server.CreateObject("Scripting.FileSystemObject")
        dirPath = server.mappath("/lib/email")
        fileName = dirPath&"\\email_searchpass.htm"
        Set objFile = fs.OpenTextFile(fileName,1)
        mailcontent = objFile.readall
        mailcontent = replace(mailcontent,":USERNAME:",userName)
        mailcontent = replace(mailcontent,":IMSIPASS:",imsipass)

        call sendmail(mailfrom, mailto, mailtitle, mailcontent)
        sendmailsearchpass = mailcontent
end function

%>
