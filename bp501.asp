<!--#include file="nop.inc"-->
<% Response.Buffer = True%>
<% response.cookies("amenu") = "bp501.asp" %>
   <FORM action="bp501.asp" method=post name="bp501"> 
   <!--'#include file="bpmenu.asp"-->
 
   
 <script language="javascript">
 function invalid_cetak(a)
 { alert(a+" Kemasukan Data Belum Selesai !!! ");
   return(true);
 }
</script>
   
    <% 	 
   
        bulan = request.form("bulan")
   		tahun = request.form("tahun")	
		hantar = request.form("a")
		reset = request.form("c")
		j = request.form("j")
		m = request.form("m")
		l = request.form("l")
		s = request.form("s")
		pek = request.cookies("gnop")
		nopek = request.form("nopek")
		cagar = ucase(request.form("cagar"))
		butir = replace(ucase(request.form("butir")),"'","''")
		simpan = request.form("d")
		ctr = request.form("ctr")
		freset = request.form("f")
		pro = request.form("pro")
		proses = Request.Form("cetak") 
		counter = Request.Form("counter") 
		'response.write("COUNTER:"&counter)

		

		if j="" then j=m	
		if j <> "" then pro = j		
		if s = "Hantar" then j=pro			'j = "Tuntutan Perjalanan"	

		if reset = "Reset" then			
			j = ""
			pro = ""
			hantar = ""
			ctr = ""
			bulan = ""
			tahun = ""
		end if
		
		
			  
  		if pek <> "" then
			hfz =" select a.keterangan,b.jabatan from iabs.jabatan a,majlis.pengguna b "
			hfz = hfz & " where a.kod = b.jabatan and b.no_pekerja ='"&pek&"' "
			set hfz = objconn.execute(hfz)
			
			if not hfz.eof then 
				hjabatan = hfz("jabatan")
				hketer = hfz("keterangan")
			end if
		end if
		
		if pro ="Tuntutan Perjalanan" or j ="Tuntutan Perjalanan" or pro="Lebihmasa" then	
					if bulan = "" or tahun = "" then
			
					y = "select to_char(add_months(sysdate,-1),'mmyyyy') bln from dual "
						Set rsy = objConn.Execute(y)	
			  
						bln = rsy("bln")
						bulan = mid(bln,1,2)
						tahun = mid(bln,3,4)
			
					else
						y = "select lpad('"& bulan &"',2,'0') bln from dual "
						Set rsy = objConn.Execute(y)	
			  
						bln = rsy("bln")
						bulan = mid(bln,1,2)
					end if		
		else
		
			y = "select to_char(sysdate,'mmyyyy') bln from dual "
    		Set rsy = objConn.Execute(y)	
  
     		bln = rsy("bln")
    		bulan = mid(bln,1,2)
     		tahun = mid(bln,3,4)
		
		end if
		
		mula
				
		if j <> "Tuntutan Perjalanan"  and j <> "" and (pro <> "Tuntutan Perjalanan" and pro<>"Lebihmasa")  then input
		
		if s = "Hantar" then		
			'j = "Tuntutan Perjalanan"
			j=pro
			'l = "Lebihmasa"
				input
		end if
		
		
		'if proses = "Cetak Kelompok2" then
		
		
		'w = "select rowid,no_dokumen nopek, null nama,butir "
		'w = w & " from iabs.kelompok_mpsp "
		'w = w & " where jenis='A' and ptj = '"&hjabatan&"' "
		'w = w & " and (post <> 'Y' or post is null) and no_pekerja is null"		
		'set sw = objconn.execute(w)
		'response.write (w)
		
		'if sw.eof then
        'mula
       
        'response.write "<script language = ""javascript"">"
        'response.write "var timeID = setTimeout('invalid_cetak("" "");',1)"
        'response.write "</'script>"
        'response.end
         
		'else
		'response.redirect "bp501c.asp?notis=bp501.asp&butir="&bulan&""&tahun&""&hjabatan&"&perihal="&pro&"&ctr="&ctr&"&hprint="&hprint&""
	
        'end if
		'end if			
		'************************************************  PROSES SIMPAN *****************************
		if simpan = "Simpan" then
		rowid = request.form("rowid")
		
		if rowid <> "" then
				de = " delete iabs.kelompok_mpsp where rowid = '"& rowid &"' "
			   	Set sde = objConn.Execute(de)
		end if
		
		if pro = "Tuntutan Perjalanan" or pro="Lebihmasa" then
		if nopek = "" then
			input
			response.write "<script language = ""vbscript"">"
		    response.write " MsgBox ""Sila masukkan nombor pekerja!!."", vbInformation, ""Perhatian!"" "
			response.write "</script>"
			response.end
		end if
		end if
		if pro = "Cagaran" then
		if cagar = "" then
		input
			response.write "<script language = ""vbscript"">"
		    response.write " MsgBox ""Sila masukkan No Akaun!!"", vbInformation, ""Perhatian!"" "
			response.write "</script>"
			response.end
		else
			checkakaun="false"
			'semak no akaun
			vb="select kod from iabs.cagaran_2002 where kod='"& cagar &"'"
			set qvb = objConn.Execute(vb)
			if not qvb.bof and not qvb.eof then
				checkakaun="true"
			end if
			if checkakaun="false" then
				input
				response.write "<script language = ""vbscript"">"
		    	response.write " MsgBox ""No Akaun Salah!!  Sila masukkan semula"", vbInformation, ""Perhatian!"" "
				response.write "</script>"
				response.end
			end if
		end if
		end if
		
		if butir = "" and pro <> "Cagaran" then
		response.write "<script language = ""vbscript"">"
		response.write " MsgBox ""Sila masukkan butir tuntutan!!."", vbInformation, ""Perhatian!"" "
		response.write "</script>"
		else
		
		if pro = "Tuntutan Perjalanan" then

			q = "select no_pekerja from payroll.paymas where no_pekerja = '"&nopek&"' "
			q = q &"union select no_pekerja from payroll.paymas_sambilan where no_pekerja = '"&nopek&"' "

			set sq = objconn.execute(q)
		
			if sq.eof then
				response.write "<script language = ""vbscript"">"
				response.write " MsgBox ""No Pekerja ini tidak wujud!!"", vbInformation, ""Perhatian!"" "
				response.write "</script>"
			else
				
				af = "insert into iabs.kelompok_mpsp(ptj,bulan,tahun,no_pekerja,butir,tkh_kelompok,jenis) "
				af = af & " values('"&hjabatan&"','"&bulan&"','"&tahun&"','"&nopek&"','"&butir&"',null,'T') "
				set saf =  objconn.execute(af)
			
			end if
		end if
		
		if pro = "Lebihmasa" then

			q = "select no_pekerja from payroll.paymas where no_pekerja = '"&nopek&"' "
			'13052013 : jun tutup, ot utk pkj tetap saja.
			'q = q &"union select no_pekerja from payroll.paymas_sambilan where no_pekerja = '"&nopek&"' "
			set sq = objconn.execute(q)
		
			if sq.eof then
				response.write "<script language = ""vbscript"">"
				response.write " MsgBox ""No Pekerja ini tidak wujud!!"", vbInformation, ""Perhatian!"" "
				response.write "</script>"
			else
				
				af = "insert into iabs.kelompok_mpsp(ptj,bulan,tahun,no_pekerja,butir,tkh_kelompok,jenis,keyin_person) "
				af = af & " values('"&hjabatan&"','"&bulan&"','"&tahun&"','"&nopek&"','"&butir&"',null,'O','"&pek&"') "
				set saf =  objconn.execute(af)
			
			end if
		end if	
		
		
		if pro = "Cagaran" then
		'20092013:asma modify checking, utk allow insert hantar semula kelompok ke kewangan wpun rekod dah ada
		q1 = "select no_dokumen from iabs.kelompok_mpsp where no_dokumen = '"&cagar&"' "
		q1 = q1 & " and KELOMPOKQ is null"
		'q1 = q1 & " and bulan = '"&bulan&"' and tahun = '"&tahun&"' "
		set sq1 = objconn.execute(q1)
		
		if not sq1.eof then
		response.write "<script language = ""vbscript"">"
		response.write " MsgBox ""No Akaun ini sudah diinput !!."", vbInformation, ""Perhatian!"" "
		response.write "</script>"
		else
		
		p = " select nama from iabs.cagaran_2002 where kod='"&cagar&"' "
		set sp = objconn.execute(p)
		
		if not sp.eof then
			if butir = "" then butir = replace(sp("nama"),"'","''")
		end if
			
 		af = "insert into iabs.kelompok_mpsp(ptj,no_dokumen,butir,tkh_kelompok,bulan,tahun,jenis) "
		af = af & " values('"&hjabatan&"','"&cagar&"','"&butir&"',null,'"&bulan&"','"&tahun&"','C') "
		set saf =  objconn.execute(af)
		nopek = ""
		butir = ""
		
		end if
		end if
		
		if pro = "Lain-lain" then
		'20092013 : asma check dokumen jenis cagaran. jika ya, x allow utk buat kelompok melalui Lain2		
		s1="select * from dual where '"& cagar &"' like '9%'"
		set qs1 = objConn.Execute(s1)
		if not qs1.bof and not qs1.eof then %>
		<script>
			alert("Akaun cagaran tidak diterima.\nMohon daftar di bahagian cagaran");
		</script>
		<%	response.redirect "bp501.asp"
		end if
		
		af = "insert into iabs.kelompok_mpsp(ptj,no_dokumen,butir,tkh_kelompok,bulan,tahun,jenis) "
		af = af & " values('"&hjabatan&"','"&cagar&"','"&butir&"',null,'"&bulan&"','"&tahun&"','A') "
		set saf =  objconn.execute(af)
		'response.write(af)
		end if
		
		nopek = ""
		butir = ""
		cagar = ""
		end if
		j=pro

		input
		end if 				
		'***************************************** END OF PROCESS SIMPAN ********************************
		
		if ctr <> "" then
		for i=0 to ctr
			frowid="frowid"+cStr(i)
			fnopek="fnopek"+cStr(i)
			fbutir="fbutir"+cStr(i)
			g = "g"+cstr(i)
			h = "h"+cstr(i)
			
			frowid=request.form(""&frowid&"")
			fnopek=request.form(""&fnopek&"")
			fbutir=request.form(""&fbutir&"")
			g=request.form(""&g&"")
			h=request.form(""&h&"")
									
			if g="Edit" then			
			if pro <> "Tuntutan Perjalanan" and pro <> "Lebihmasa"  then						
				cagar = fnopek
			else
				nopek = fnopek
			end if
			butir = fbutir
			rowid = frowid
			j=pro
			input						
			elseif h="Hapus" then
	   			del = " delete iabs.kelompok_mpsp where rowid = '"& frowid &"' "
			   	Set sdel = objConn.Execute(del)
				'response.write (del)
			j = pro
			input	
			end if						
		next
	end if
	
		if freset = "Reset" then
			nopek = ""
			butir = ""
			frowid = ""
			cagar = ""
			j = pro
			input			
		end if
		
		sub mula		%>
  <table width="100%" align="center" cellpadding="1" cellspacing="1" style="font-family: Trebuchet MS; font-size: 10pt; ">
    <tr>
      <td bgcolor="<%=color1%>">Jenis Kelompok</td>
      <td bgcolor="<%=color2%>">
	  <input name="j" type="submit" id="j" style="font-family: Trebuchet MS; font-size: 9pt; height:17pt;" value="Tuntutan Perjalanan">	 
	  <input name="m" type="submit" id="m" style="font-family: Trebuchet MS; font-size: 9pt; height:17pt;" value="Lebihmasa"> 
        <input name="j" type="submit" id="j" style="font-family: Trebuchet Ms; font-size: 9pt; height:17pt;" value="Cagaran">
        <input name="j" type="submit" id="j" style="font-family: Trebuchet Ms; font-size: 9pt; height:17pt;" value="Lain-lain"> 
		<input name="c" type="submit" value="Reset" style="font-family: Trebuchet Ms; font-size: 9pt; height:17pt; color:red;">
      </td>
    </tr>
    <tr> 
	<% if j = "Tuntutan Perjalanan" or pro = "Tuntutan Perjalanan" or j = "Lebihmasa" or pro = "Lebihmasa" then %>
      <td width="19%" bgcolor="<%=color1%>">Bulan/Tahun</td>
      <td width="81%" bgcolor="<%=color2%>"><input name="bulan" type="text" value="<%=bulan%>" size="2" maxlength="2">
        /  <input name="tahun" type="text" value="<%=tahun%>" size="4" maxlength="4">
        <input name="s" type="submit" id="s" style="font-family: Trebuchet Ms; font-size: 9pt; height:17pt;" value="Hantar"></td>
    </tr>
	<%end if%>
  </table>
  <%	end sub		%>
<input type="hidden" name="pro" value="<%=j%>">	
	<%	sub input 
	
	    w = "select count(*) bil "
		w = w & " from iabs.kelompok_mpsp "
		w = w & " where jenis='A' and ptj = '"&hjabatan&"' "
		w = w & " and (post <> 'Y' or post is null) and no_pekerja is null"		
		set sw = objconn.execute(w)
		'response.write (w)
		if not sw.eof then 
		   counter = sw("bil")
		end if
	
	
	
	%> <br>
  <table width="95%" align="center" cellpadding="1" cellspacing="1" style="font-family: Trebuchet MS; font-size: 10pt;">
    <tr align="center"> 
      <td colspan="4" bgcolor="<%=color1%>">
	  <% if j="Lebihmasa" then%>
	  <input type="submit" name="f2" value="Cetak Kelompok" style="font-family: Trebuchet MS; font-size: 9pt; height:17pt;" onclick="this.form.action='bp501d.asp?notis=bp501.asp&butir=<%=bulan%><%=tahun%><%=hjabatan%>&perihal=<%=pro%>';">
	 
	 <% elseif j="Lain-lain"  and counter > "0" then %> 
    <input type="submit" name="f2" value="Cetak Kelompok" style="font-family: Trebuchet MS; font-size: 9pt; height:17pt;" onclick="this.form.action='bp501c.asp?notis=bp501.asp&butir=<%=bulan%><%=tahun%><%=hjabatan%>&perihal=<%=pro%>';">	 
	 
	 <% elseif  j<>"Lain-lain" and j<>"Lebihmasa" then 
	 
	 rekod="tiada"
	 if pro = "Cagaran" or j = "Cagaran" then
		bx = " select no_dokumen num ,butir,rowid, null nama "
		bx = bx & " from iabs.kelompok_mpsp where no_dokumen like '9%' and ptj = '"&hjabatan&"' "
		bx = bx & " and (post <> 'Y' or post is null) "
		set qbx = objConn.Execute(bx)
		if not qbx.bof and not qbx.eof then rekod="ada"
	end  if
	
		if rekod="ada" or j <> "Cagaran" then
	 %> 
	 
			<input type="submit" name="f2" value="Cetak Kelompok" style="font-family: Trebuchet MS; font-size: 9pt; height:17pt;" onclick="this.form.action='bp501c.asp?notis=bp501.asp&butir=<%=bulan%><%=tahun%><%=hjabatan%>&perihal=<%=pro%>';">	 
	 <% 
	 	end if
	 end if%>
	  </td>
    </tr>
    <tr align="center"> 
      <td width="21%" bgcolor="<%=color1%>">
        <% if j = "Tuntutan Perjalanan" or pro = "Tuntutan Perjalanan" or j = "Lebihmasa" or pro = "Lebihmasa" then %>
        No Pekerja 
        <%elseif j ="Cagaran"  or pro ="Cagaran"  then %>
        No Akaun 
        <%elseif j ="Lain-lain" or pro ="Lain-lain" then %>
        No Dokumen 
        <%end if%>
      </td>
      <td width="52%" bgcolor="<%=color1%>">Nama Syarikat &amp; Keterangan BP </td>
      <td width="9%" bgcolor="<%=color1%>">Cetak Kelompok</td>
      <td width="18%" bgcolor="<%=color1%>">Proses</td>
    </tr>
    <tr align="center" valign="top" bgcolor="<%=color3%>"> 
      <td> 
        <% 
	if j = "Tuntutan Perjalanan" or pro="Tuntutan Perjalanan" or j = "Lebihmasa" or pro="Lebihmasa" then %>
        <input name="nopek" type="text" id="nopek" value="<%=nopek%>" size="5" maxlength="5"  style="font-family: Trebuchet MS; font-size: 10pt;" onKeyDown="if(event.keyCode==13) event.keyCode=9;"> 
        <a href="javascript:void(0)" onClick="open_staf('bp501.nopek');" onMouseOver="window.status='Senarai Nombor Pekerja';return true;" onMouseOut="window.status='';return true;" > 
        <input name="h" type="button" style="font-family: Trebuchet Ms; font-size: 9pt; font-weight:bold; color:red;" value="?">
        </a> <script>document.bp501.nopek.focus()</script> 
        <%elseif j ="Cagaran" or j ="Lain-lain" or pro ="Cagaran" or pro ="Lain-lain"then %>
        <input name="cagar" type="text" onKeyDown="if(event.keyCode==13) event.keyCode=9;" id="cagar" value="<%=cagar%>" size="11" maxlength="11"  style="font-family: Trebuchet MS; font-size: 10pt;"> 
        <%if j ="Cagaran" or pro ="Cagaran" then%>
        <a href="javascript:void(0)" onClick="open_cagar('bp501.cagar','bp501.butir');" onMouseOver="window.status='Senarai No Akaun Cagaran';return true;" onMouseOut="window.status='';return true;"> 
        <input name="h2" type="button" id="h2" style="font-family: Trebuchet Ms; font-size: 9pt; font-weight:bold; color:red;" value="?">
        </a> 
        <% end if%>
        <%end if%>
      </td>
      <td> <textarea  name="butir" rows="5" onKeyDown="if(event.keyCode==13) event.keyCode=9;" id="butir" style="font-family: Trebuchet MS; font-size: 10pt;"cols="50"><%=butir%></textarea></td>
      <td>&nbsp;</td>
      <td align="left"> <input name="d" type="submit" id="d" onFocus="nextfield='done';" style="font-family: Trebuchet Ms; font-size: 9pt; height:17pt; font-weight:bold;" value="Simpan"> 
        <input name="f" type="submit" id="f" style="font-family: Trebuchet Ms; font-size: 9pt; font-weight:bold; color:red; height:17pt;" onFocus="nextfield='done';" value="Reset"></td>
    </tr>
    <%  if j <> "" then pro = j	
		if pro ="Tuntutan Perjalanan" or j ="Tuntutan Perjalanan" then	
		
		'15032012 : jun tutup, yang lama ni baca paymas saja, tak baca dari paymas_sambilan
		'b = b & " select upper(a.nama)nama,b.rowid,lpad(b.no_pekerja,5,0) num,b.butir from payroll.paymas a,iabs.kelompok_mpsp b "
	    'b = b & " where a.no_pekerja = b.no_pekerja and b.ptj = '"&hjabatan&"' and b.bulan = '"&bulan&"' and b.tahun = '"&tahun&"' "
		'b = b & " and b.jenis='T' and (b.post <> 'Y' or b.post is null) "
		
		'13052013 : jun create view untuk select paymas/paymas_sambilan dlm 1 view, tak payah nak guna union		
		b = b & " select upper(a.nama)nama,b.rowid,lpad(b.no_pekerja,5,0) num,b.butir from payroll.paymas_all a,iabs.kelompok_mpsp b "
	    b = b & " where a.no_pekerja = b.no_pekerja and b.ptj = '"&hjabatan&"' and b.bulan = '"&bulan&"' and b.tahun = '"&tahun&"' "
		b = b & " and b.jenis='T' and (b.post <> 'Y' or b.post is null) "
		b = b & " order by b.no_pekerja "		
	
		
		elseif 	pro ="Lebihmasa" or j ="Lebihmasa" then	
		
		b = " select upper(a.nama)nama,b.rowid,lpad(b.no_pekerja,5,0) num,b.butir from payroll.paymas a,iabs.kelompok_mpsp b "
	    b = b & " where a.no_pekerja = b.no_pekerja and b.ptj = '"&hjabatan&"' and b.bulan = '"&bulan&"' and b.tahun = '"&tahun&"' "
		b = b & " and (b.post <> 'Y' or b.post is null) and jenis='O' and keyin_person='"&pek&"' "
		b = b & " order by a.no_pekerja "
		'13052013 : jun tutup select pkj sambilan, lbhmasa utk tetap saja
		'b = b & " union "
		'b = b & " select upper(a.nama)nama,b.rowid,lpad(b.no_pekerja,5,0) num,b.butir from payroll.paymas_sambilan a,iabs.kelompok_mpsp b "
        'b = b & " where a.no_pekerja = b.no_pekerja and b.ptj = '"&hjabatan&"' and b.bulan = '"&bulan&"' and b.tahun = '"&tahun&"' "
        'b = b & " and (b.post <> 'Y' or b.post is null) and jenis='O' and keyin_person='"&pek&"'"
		
		elseif pro = "Cagaran" or j = "Cagaran" then
		b = " select no_dokumen num ,butir,rowid, null nama "
		b = b & " from iabs.kelompok_mpsp where no_dokumen like '9%' and ptj = '"&hjabatan&"' "
		b = b & " and (post <> 'Y' or post is null) "
		
		elseif pro = "Lain-lain" or j = "Lain-lain" then

		b = " select no_dokumen num,butir,rowid,null,no_pekerja,null nama from iabs.kelompok_mpsp "
		b = b & " where jenis ='A' and ptj = '"&hjabatan&"' "
		b = b & " and (post <> 'Y' or post is null) and no_pekerja is null"		
		end if
		set sb = objconn.execute(b)		
		ctr = 0	
		
'		if 	pro ="Lebihmasa" or j ="Lebihmasa" then
'			if sb.eof then
'				b = " select upper(a.nama)nama,b.rowid,lpad(b.no_pekerja,5,0) num,b.butir from payroll.paymas_sambilan a,iabs.kelompok_mpsp b "
'				b = b & " where a.no_pekerja = b.no_pekerja and b.ptj = '"&hjabatan&"' and b.bulan = '"&bulan&"' and b.tahun = '"&tahun&"' "
'				b = b & " and (b.post <> 'Y' or b.post is null) and jenis='O' and keyin_person='"&pek&"'"
'			end if
'			set sb = objconn.execute(b)	
'		end if		
		
       'response.Write(b)
		
		do while not sb.eof 
		ctr = ctr + 1
		warna = ctr mod 2
		hrowid = sb("rowid")		
		hnopek = sb("num")
		hbutir = sb("butir")
		hnama = sb("nama")
		
				
	%>
    <tr valign="top" <%if warna ="1" then %>bgcolor="#CCCCCC" <%else%>bgcolor="<%=color2%>" <%end if%> > 
      <td><%=hnopek%> 
        <%if hnama <> "" then %>
        -<%=hnama%> 
        <%end if%>
      </td>
      <td> 
        <%if hbutir <> "" then
			response.write(replace(hbutir,vbcrlf,"<br>"))
end if %>
    </td>
      <td align="center"><input type="checkbox" name="hprint<%=ctr%>" value="Y" <%if hprint="" or hprint="Y" then%> checked <%end if%>></td>
      <td align="center"> <input name="g<%=ctr%>" type="submit" id="g<%=ctr%>" onFocus="nextfield='done';" style="font-family: Trebuchet Ms; font-size: 9pt; font-weight:bold; height:17pt;" value="Edit"> 
        <input name="h<%=ctr%>" type="submit" id="h<%=ctr%>" onClick="return confirm('Hapus Rekod Ini ?')" style="font-family: Trebuchet Ms; font-size: 9pt; font-weight:bold; color:red; height:17pt;" onFocus="nextfield='done';" value="Hapus"></td>
      <input type="hidden" name="frowid<%=ctr%>" value="<%=hrowid%>" >
      <input type="hidden" name="fnopek<%=ctr%>" value="<%=hnopek%>" >
      <input type="hidden" name="fbutir<%=ctr%>" value="<%=hbutir%>" >
	  <input type="hidden" name="npek<%=ctr%>" value="<%=hnopek%>" >
       
    </tr>
    <%	sb.movenext
		loop	%>
  </table>
  <input type="hidden" name="rowid" value="<%=rowid%>" >
  <input type="hidden" name="ctr" value="<%=ctr%>" >
  <%	if pro <> "" and j = "" then %><input type="hidden" name="pro" value="<%=pro%>"><%'=pro%>
<%	end if
	end sub %>  
    	</form>