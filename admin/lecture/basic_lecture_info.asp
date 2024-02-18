<% option Explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
%>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<link rel="stylesheet" href="/bct.css" type="text/css">
<script language="JavaScript">
<!--
function TnCheckForm(){
	if(document.itemfrm.itemid.value == ""){
		alert("상품번호를 넣어주세요");
		document.itemfrm.itemid.focus();
	}
	else{
	document.itemfrm.submit();
	}
}
//-->
</script>
<%
dim itemid
dim sql,Tcnt
dim	Fitemserial_large,Fitemserial_mid,Fitemserial_small
dim Fitemname,Fmakerid,Fitemsource
dim Fitemsize,Fkeywords,Fmakername
dim Fsourcearea,Fdeliverytype,Fmwdiv,Fdeliverarea
dim Fsellcash,Fbuycash,Fitemcontent,Fusinghtml
dim Fsellvat,Fbuyvat
	dim FResultCount

	dim Fidx
	dim Flinkitemid
	dim Flecturerid
	dim Flecturer
	dim Flectitle
	dim Flecsum
	dim Flecspace
	dim Fmatinclude
	dim Fmatsum
	dim Fleccount
	dim Flectime
	dim Ftottime
	dim Fmatdesc
	dim Flecperiod
	dim Fproperperson
	dim Fminperson
	dim Freservestart
	dim Freserveend
	dim Flecdate01
	dim Flecdate02
	dim Flecdate03
	dim Flecdate04
	dim Flecdate05
	dim Flecdate06
	dim Flecdate07
	dim Flecdate08

	dim Flecdate01_end
	dim Flecdate02_end
	dim Flecdate03_end
	dim Flecdate04_end
	dim Flecdate05_end
	dim Flecdate06_end
	dim Flecdate07_end
	dim Flecdate08_end
	dim Fleccontents
	dim Fleccurry
	dim Flecetc
	dim Fisusing
	dim FYyyymm

	dim FRegFinish
itemid = request("itemid")

if itemid <> "" then
    '상품정보 가져오기

	sql = " select top 1 i.itemserial_large, i.itemserial_mid, i.itemserial_small" + vbCrlf
	sql = sql + ", i.itemname, i.makerid, i.itemsource, i.itemsize, i.keywords, i.makername, i.usinghtml" + vbCrlf
	sql = sql + ", i.sellcash, i.sellvat, i.buycash, i.buyvat, i.sourcearea, i.deliverytype, i.deliverarea, i.mwdiv, i.itemcontent" + vbCrlf
	sql = sql + ", lec.idx, lec.linkitemid, lec.lecturerid, lec.lecturer" + vbCrlf
	sql = sql + ", lec.lectitle, lec.lecsum, lec.lecspace, lec.matinclude, lec.matsum" + vbCrlf
	sql = sql + ", lec.leccount, lec.lecperiod, lec.lectime, lec.tottime, lec.matdesc, lec.properperson" + vbCrlf
	sql = sql + ", lec.minperson, lec.reservestart, lec.reserveend" + vbCrlf
	sql = sql + ", convert(varchar(19),lec.lecdate01,21) as lecdate01" + vbCrlf
	sql = sql + ", convert(varchar(19),lec.lecdate02,21) as lecdate02" + vbCrlf
	sql = sql + ", convert(varchar(19),lec.lecdate03,21) as lecdate03" + vbCrlf
	sql = sql + ", convert(varchar(19),lec.lecdate04,21) as lecdate04" + vbCrlf
	sql = sql + ", convert(varchar(19),lec.lecdate05,21) as lecdate05" + vbCrlf
	sql = sql + ", convert(varchar(19),lec.lecdate06,21) as lecdate06" + vbCrlf
	sql = sql + ", convert(varchar(19),lec.lecdate07,21) as lecdate07" + vbCrlf
	sql = sql + ", convert(varchar(19),lec.lecdate08,21) as lecdate08" + vbCrlf
	sql = sql + ", convert(varchar(19),lec.lecdate01_end,21) as lecdate01_end" + vbCrlf
	sql = sql + ", convert(varchar(19),lec.lecdate02_end,21) as lecdate02_end" + vbCrlf
	sql = sql + ", convert(varchar(19),lec.lecdate03_end,21) as lecdate03_end" + vbCrlf
	sql = sql + ", convert(varchar(19),lec.lecdate04_end,21) as lecdate04_end" + vbCrlf
	sql = sql + ", convert(varchar(19),lec.lecdate05_end,21) as lecdate05_end" + vbCrlf
	sql = sql + ", convert(varchar(19),lec.lecdate06_end,21) as lecdate06_end" + vbCrlf
	sql = sql + ", convert(varchar(19),lec.lecdate07_end,21) as lecdate07_end" + vbCrlf
	sql = sql + ", convert(varchar(19),lec.lecdate08_end,21) as lecdate08_end" + vbCrlf
	sql = sql + ", lec.leccontents, lec.leccurry, lec.lecetc, lec.isusing, lec.mastercode, lec.regfinish " + vbCrlf
	sql = sql + " from [db_contents].[dbo].tbl_lecture_item lec" + vbCrlf
	sql = sql + ", [db_item].[dbo].tbl_item i" + vbCrlf
	sql = sql + "where i.itemid=lec.linkitemid " + vbCrlf
	sql = sql + "and i.itemid='" + Cstr(itemid) + "' " + vbCrlf

	'sql = "select i.itemserial_large,i.itemserial_mid,i.itemserial_small," + vbCrlf
	'sql = "i.itemname,i.makerid,i.itemsource,i.itemsize,i.keywords,i.makername," + vbCrlf
	'sql = sql + ",i.sellcash,i.sellvat,i.buycash,i.buyvat" + vbCrlf
	'sql = sql + " from  [db_item].[dbo].tbl_item" + vbCrlf
	'sql = sql + " where itemid = '" + Cstr(itemid) + "' "

	rsget.Open sql, dbget, 1

	Tcnt = rsget.RecordCount

	if  not rsget.EOF  then

		Fitemserial_large = rsget("itemserial_large")
		Fitemserial_mid = rsget("itemserial_mid")
		Fitemserial_small = rsget("itemserial_small")
		Fitemname = db2html(rsget("itemname"))
		Fmakerid = rsget("makerid")
		Fitemsource = db2html(rsget("itemsource"))
		Fitemsize = db2html(rsget("itemsize"))
		Fkeywords = db2html(rsget("keywords"))
		Fmakername = db2html(rsget("makername"))
		Fsourcearea = db2html(rsget("sourcearea"))
		Fdeliverytype = rsget("deliverytype")
		Fdeliverarea = rsget("deliverarea")
		Fmwdiv = rsget("mwdiv")
		Fsellcash = rsget("sellcash")
		Fsellvat = rsget("sellvat")
		Fbuycash = rsget("buycash")
		Fbuyvat = rsget("buyvat")
		Fitemcontent = replace(db2html(rsget("itemcontent")),vbcrlf,"<br>")
		Fusinghtml = rsget("usinghtml")

		Flinkitemid   = rsget("linkitemid")
		Flecturerid     = rsget("lecturerid")
		Flecturer     = rsget("lecturer")
		Flectitle     = db2html(rsget("lectitle"))
		Flecsum       = rsget("lecsum")
		Fmatinclude   = rsget("matinclude")
		Fmatsum       = rsget("matsum")
		Flecspace       = rsget("lecspace")
		Fleccount     = rsget("leccount")
		Flecperiod      = rsget("lecperiod")
		Flectime      = rsget("lectime")
		Ftottime      = rsget("tottime")
		Fmatdesc      = db2html(rsget("matdesc"))
		Fproperperson = rsget("properperson")
		Fminperson    = rsget("minperson")
		Freservestart = rsget("reservestart")
		Freserveend   = rsget("reserveend")
		Flecdate01    = rsget("lecdate01")
		Flecdate02    = rsget("lecdate02")
		Flecdate03    = rsget("lecdate03")
		Flecdate04    = rsget("lecdate04")
		Flecdate05    = rsget("lecdate05")
		Flecdate06    = rsget("lecdate06")
		Flecdate07    = rsget("lecdate07")
		Flecdate08    = rsget("lecdate08")
		Flecdate01_end = rsget("lecdate01_end")
		Flecdate02_end = rsget("lecdate02_end")
		Flecdate03_end = rsget("lecdate03_end")
		Flecdate04_end = rsget("lecdate04_end")
		Flecdate05_end = rsget("lecdate05_end")
		Flecdate06_end = rsget("lecdate06_end")
		Flecdate07_end = rsget("lecdate07_end")
		Flecdate08_end = rsget("lecdate08_end")
		Fleccontents  = db2html(rsget("leccontents"))
		Fleccurry     = db2html(rsget("leccurry"))
		Flecetc       = db2html(rsget("lecetc"))
		Fisusing      = rsget("isusing")
		FYyyymm      = rsget("mastercode")
		FRegFinish = rsget("regfinish")

	end if
	rsget.close
if Tcnt > 0 then
%>
<script language="JavaScript">
<!--
	var frm = opener.itemreg;
	var source,convert,temp;

source = "<br>";
convert = "\n";
temp = "<% = replace(Fitemcontent,chr(34),"") %>";

while (temp.indexOf(source)>-1) {
	 pos= temp.indexOf(source);
	 temp = "" + (temp.substring(0, pos) + convert +
	 temp.substring((pos + source.length), temp.length));
}
	frm.yyyymm.value				='<%= FYyyymm %>';  //강좌 월구분

	frm.itemname.value			= '<% =Fitemname  %>';	//강좌명
	frm.designerid.value			= '<% = Fmakerid %>';		//소속아이디
	frm.tempid.value				= '<% = Fmakerid %>';			//소속아이디

//	frm.itemsource.value		= '<% = Fitemsource %>';		//상품재질
//	frm.itemsize.value				= '<% = Fitemsize %>';		//상품사이즈

	frm.keywords.value			= '<% = Fkeywords %>';	//키워드
	frm.makename.value		= "<% = Fmakername %>";	//제조사
	frm.sourcearea.value		= '<% = Fsourcearea %>';	//원산지
	frm.deliverytype.value		= '<% = Fdeliverytype %>';	//배송구분함

	frm.sellcash.value			= '<% = Fsellcash %>';		//판매가
	frm.sellvat.value				= '<% = Fsellvat %>';				//판매부가세소
	frm.buycash.value			= '<% = Fbuycash %>';		//매입가
	frm.buyvat.value				= '<% = Fbuyvat %>';				//매입부가세
	frm.mileage.value				='11';								//마일리지
	frm.lecturerid.value			=	'<%= Flecturerid %>';	//소속아이디
	frm.lecturer.value				='<%= Flecturer %>';			//강사명
	frm.lecsum.value				='<%= Flecsum %>';			//강좌비
	frm.matinclude.value		='<%= Fmatinclude %>;'	//재료비포함유
	frm.matsum.value				='<%= Fmatsum %>';		//재료비
	frm.lecspace.value 			='<%= Flecspace %>';		//장소
	frm.leccount.value			='<%= Fleccount %>';		//강좌횟수
	frm.lecperiod.value			='<%= Flecperiod %>';		//강의기간(주기)
	frm.lectime.value				='<%= Flectime %>';			//강의시간
	frm.tottime.value				='<%= Ftottime %>'; 			//총강의시간
	frm.matdesc.value			='<%= Fmatdesc %>';		//재료비설명
	frm.properperson.value	='<%= Fproperperson %>';	//적정인원
	frm.minperson.value    		='<%= Fminperson %>';	//최소인원
	frm.reservestart.value 		='<%= Freservestart %>';		//예약등록일
	frm.reserveend.value		='<%= Freserveend %>';	//예약마감일
	frm.lecdate01.value			='<%= Flecdate01 %>';			//강좌내용(커리큘럼)
	frm.lecdate02.value			='<%= Flecdate02 %>';
	frm.lecdate03.value			='<%= Flecdate03 %>';
	frm.lecdate04.value			='<%= Flecdate04 %>';
	frm.lecdate05.value			='<%= Flecdate05 %>';
	frm.lecdate06.value			='<%= Flecdate06 %>';
	frm.lecdate07.value			='<%= Flecdate07 %>';
	frm.lecdate08.value			='<%= Flecdate08 %>';
	frm.lecdate01_end.value 	='<%= Flecdate01_end %>';
	frm.lecdate02_end.value 	='<%= Flecdate02_end %>';
	frm.lecdate03_end.value 	='<%= Flecdate03_end %>';
	frm.lecdate04_end.value 	='<%= Flecdate04_end %>';
	frm.lecdate05_end.value 	='<%= Flecdate05_end %>';
	frm.lecdate06_end.value 	='<%= Flecdate06_end %>';
	frm.lecdate07_end.value 	='<%= Flecdate07_end %>';
	frm.lecdate08_end.value 	='<%= Flecdate08_end %>';

	<%
	Fleccontents = replace(Fleccontents,chr(34),"&#34;")
	Fleccontents = replace(Fleccontents,chr(39),"&#39;")
	Fleccontents = replace(nl2br(Fleccontents),"<br>","\n")
	%>

	var leccontents='<%= Fleccontents %>';

	leccontents= leccontents.replace(/&#34;/gi,"\"");
	leccontents= leccontents.replace(/&#39;/gi,"'");
	frm.leccontents.innerText=leccontents;

	<%
	Fleccurry = replace(Fleccurry,chr(34),"&#34;")
	Fleccurry = replace(Fleccurry,chr(39),"&#39;")
	Fleccurry = replace(nl2br(Fleccurry),"<br>","\n")
	%>

	var leccurry='<%= Fleccurry %>';

	leccurry= leccurry.replace(/&#34;/gi,"\"");
	leccurry= leccurry.replace(/&#39;/gi,"'");
	frm.leccurry.innerText=leccurry;


	<%
	Flecetc = replace(Flecetc,chr(34),"&#34;")
	Flecetc = replace(Flecetc,chr(39),"&#39;")
	Flecetc = replace(nl2br(Flecetc),"<br>","\n")
	%>

	var lecetc='<%= Flecetc %>';

	lecetc= lecetc.replace(/&#34;/gi,"\"");
	lecetc= lecetc.replace(/&#39;/gi,"'");
	frm.lecetc.innerText=lecetc;

//	frm.regfinish.value			='<%= Fregfinish %>';	//접수종료

	self.close();

//-->
</script>
<% else %>
<script language="JavaScript">
<!--

	self.close();

//-->
</script>
<% end if %>
<% else %>
<table border="0" cellpadding="0" cellspacing="0" class="a" width="100%" height="100%">
<form method="post" name="itemfrm">
<tr>
	<td align="center">상품기본틀생성하기</td>
</tr>
<tr>
	<td align="center">상품번호<input type="text" name="itemid" size="6"><input type="button" value="전송하기" onclick="TnCheckForm();"></td>
</tr>
</form>
</table>
<% end if %>
<!-- #include virtual="/lib/db/dbclose.asp" -->