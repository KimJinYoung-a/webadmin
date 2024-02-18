<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->


<%
dim lec_idx
dim target
lec_idx=RequestCheckvar(request("lec_idx"),10)
'target= request("target")

dim sql
	sql = "select top 1 l.*,convert(varchar(19),lec_startday1,121) as lec_startday1 ,convert(varchar(19),lec_endday1,121) as lec_endday1 ,CL.code_nm as large_name , CM.code_nm as mid_name"
	sql = sql + " from [db_academy].[dbo].tbl_lec_item l inner join [db_academy].[dbo].tbl_lec_cate_large CL on l.newCate_large = CL.code_large inner join [db_academy].[dbo].tbl_lec_cate_mid CM on l.newCate_large = CM.code_large and l.newCate_mid = CM.code_mid" + vbcrlf
	sql = sql + " where idx=" + CStr(lec_idx)

	''rw sql

	rsAcademyget.open sql,dbAcademyget,1

	if not rsAcademyget.eof and not rsAcademyget.bof then

%>
	<script>

	var frm = parent.opener;

			//frm.lecfrm.oldidx.value		='<%= rsACADEMYget("idx") %>';
			//frm.lecfrm.cate_large.value		='<%= rsACADEMYget("cate_large") %>';

			frm.lecfrm.code_large.value				='<%= rsACADEMYget("newCate_large") %>'; //대카테고리
			frm.lecfrm.code_mid.value				='<%= rsACADEMYget("newCate_mid") %>'; //중카테고리
			frm.lecfrm.large_name.value				='<%= rsACADEMYget("large_name") %>'; //대카테고리 이름
			frm.lecfrm.mid_name.value				='<%= rsACADEMYget("mid_name") %>'; //중카테고리 이름

			// frm.lecfrm.lec_date.value			='<%= rsACADEMYget("lec_date") %>';   //복사않함.
			frm.lecfrm.lec_title.value			='<%= db2html(rsACADEMYget("lec_title")) %>';
			frm.lecfrm.lecturer_id.value		='<%= rsACADEMYget("lecturer_id") %>';
			frm.lecfrm.lecturer_name.value	='<%= db2html(rsACADEMYget("lecturer_name")) %>';
			frm.lecfrm.lec_cost.value			='<%= rsACADEMYget("lec_cost") %>';
			frm.lecfrm.buying_cost.value		='<%= rsACADEMYget("buying_cost") %>';
			frm.lecfrm.mileage.value				='<%= rsACADEMYget("mileage") %>';
			frm.lecfrm.margin.value				='<%= rsACADEMYget("margin") %>';
			frm.lecfrm.mat_cost.value			='<%= rsACADEMYget("mat_cost") %>';
            
			frm.lecfrm.temp_lec_id.value		='<%= rsACADEMYget("lecturer_id") %>,<%= db2html(rsACADEMYget("lecturer_name")) %>,<%= rsACADEMYget("margin") %>';
            
            //2010-추가
            frm.lecfrm.mat_margin.value			='<%= rsACADEMYget("mat_margin") %>';           
            frm.lecfrm.buying_cost.value			='<%= rsACADEMYget("buying_cost") %>';
            
			
            frm.lecfrm.matinclude_yn.value='<%= rsACADEMYget("matinclude_yn") %>';
            frm.lecfrm.lecjgubun.value='<%= rsACADEMYget("lecjgubun") %>';
            frm.lecfrm.CateCD1.value='<%= rsACADEMYget("CateCD1") %>';
            frm.lecfrm.CateCD3.value='<%= rsACADEMYget("CateCD3") %>';
            frm.lecfrm.classlevel.value='<%= rsACADEMYget("classlev") %>';

			<%
					dim mat_contents,lec_outline,lec_contents,lec_etccontents

					mat_contents = replace(rsACADEMYget("mat_contents"),chr(34),"#&34;")
					mat_contents = replace(mat_contents,chr(39),"#&39;")
					mat_contents = replace(nl2br(mat_contents),"<br>","\r\n")
					mat_contents = replace(nl2br(mat_contents),"<br />","\r\n")   ''추가 2016/12/13
				    mat_contents = replace(mat_contents,VbCR,"")                ''추가 2016/12/13
				    mat_contents = replace(mat_contents,VbLf,"")                ''추가 2016/12/13
			%>

			var mat_contents='<%= mat_contents %>';
			mat_contents= mat_contents.replace(/#&34;/gi,"\"");
			mat_contents= mat_contents.replace(/#&39;/gi,"'");
			frm.lecfrm.mat_contents.value=mat_contents;

			frm.lecfrm.keyword.value			='<%= db2html(rsACADEMYget("keyword")) %>';


			frm.lecfrm.limit_count.value			='<%= rsACADEMYget("limit_count") %>';
			frm.lecfrm.min_count.value			='<%= rsACADEMYget("min_count") %>';

			//frm.lecfrm.limit_sold.value			='<%= rsACADEMYget("limit_sold") %>';

			frm.lecfrm.reg_startday.value		='<%= rsACADEMYget("reg_startday") %>';
			frm.lecfrm.reg_endday.value		='<%= rsACADEMYget("reg_endday") %>';

			frm.lecfrm.lec_count.value			='<%= rsACADEMYget("lec_count") %>';
			frm.lecfrm.lec_time.value			='<%= rsACADEMYget("lec_time") %>';
			frm.lecfrm.lec_period.value		='<%= rsACADEMYget("lec_period") %>';
			frm.lecfrm.lec_space.value		='<%= rsACADEMYget("lec_space") %>';
			frm.lecfrm.lec_mapimg.value		='<%= db2html(rsACADEMYget("lec_mapimg")) %>';

			//frm.lecfrm.lec_startday.value	='<%= rsACADEMYget("lec_startday1") %>';
			//frm.lecfrm.lec_endday.value		='<%= rsACADEMYget("lec_endday1") %>';
            frm.lecfrm.map_idx.value		='<%= rsACADEMYget("map_idx") %>';

			<%
				lec_outline = replace(rsACADEMYget("lec_outline"),chr(34),"#&34;")
				lec_outline = replace(lec_outline,chr(39),"#&39;")
				lec_outline = replace(nl2br(lec_outline),"<br>","\r\n")
				lec_outline = replace(nl2br(lec_outline),"<br />","\r\n")   ''추가 2016/12/13
				lec_outline = replace(lec_outline,VbCR,"")                ''추가 2016/12/13
				lec_outline = replace(lec_outline,VbLf,"")                ''추가 2016/12/13
			%>

			var lec_outline='<%= lec_outline %>';
			lec_outline= lec_outline.replace(/#&34;/gi,"\"");
			lec_outline= lec_outline.replace(/#&39;/gi,"'");
			//frm.lecfrm.lec_outline.innerText=lec_outline;
			frm.lecfrm.lec_outline.value=lec_outline;

			<%
				lec_contents = replace(rsACADEMYget("lec_contents"),chr(34),"#&34;")
				lec_contents = replace(lec_contents,chr(39),"#&39;")
				lec_contents = replace(nl2br(lec_contents),"<br>","\r\n")
				lec_contents = replace(nl2br(lec_contents),"<br />","\r\n")   ''추가 2016/12/13
				lec_contents = replace(lec_contents,VbCR,"")                ''추가 2016/12/13
				lec_contents = replace(lec_contents,VbLf,"")                ''추가 2016/12/13
			%>

			var lec_contents='<%= trim(lec_contents) %>';
			lec_contents= lec_contents.replace(/#&34;/gi,"\"");
			lec_contents= lec_contents.replace(/#&39;/gi,"'");
			//frm.lecfrm.lec_contents.innerText=lec_contents;
			frm.lecfrm.lec_contents.value=lec_contents;

			<%
				lec_etccontents = replace(db2html(rsACADEMYget("lec_etccontents")),chr(34),"#&34;")
				lec_etccontents = replace(lec_etccontents,chr(39),"#&39;")
				lec_etccontents = replace(nl2br(lec_etccontents),"<br>","\r\n")
				lec_etccontents = replace(nl2br(lec_etccontents),"<br />","\r\n")   ''추가 2016/12/13
				lec_etccontents = replace(lec_etccontents,VbCR,"")                ''추가 2016/12/13
				lec_etccontents = replace(lec_etccontents,VbLf,"")                ''추가 2016/12/13
			%>

			var lec_etccontents='<%= (lec_etccontents) %>';
			lec_etccontents= lec_etccontents.replace(/#&34;/gi,"\"");
			lec_etccontents= lec_etccontents.replace(/#&39;/gi,"'");
			//frm.lecfrm.lec_etccontents.innerText=lec_etccontents;
			frm.lecfrm.lec_etccontents.value=lec_etccontents;


			//frm.lecfrm.isusing.value				=' <%= rsACADEMYget("isusing") %>';
			frm.lecfrm.reg_yn.value				=' <%= rsACADEMYget("reg_yn") %>';
			frm.lecfrm.disp_yn.value				='<%= rsACADEMYget("disp_yn") %>';


			//frm.lecfrm.regdate.value				='<%= rsACADEMYget("regdate") %>';



			frm.showimgyn();
    
            parent.close();
		</script>
<%
	rsACADEMYget.close
	end if

%>
<script>
//self.close();
</script>
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->
