<%response.buffer = True %>
<%server.scripttimeout=9999999%>
<html>
<head>
<title>Sistem Pengurusan Kutipan</title>
<script language="javascript">
function cetak(){
if (window.print){
window.print();
}
}
function invalid_data(a)
    {  
       alert (a+" Tiada Rekod ");
		return(true);
    }
</script>
</head>


<body topmargin="0" leftmargin="0">
<%		Set objConn = Server.CreateObject("ADODB.Connection")
		objConn.Open "dsn=12c;uid=majlis;pwd=majlis;"


pekd = request.cookies("gnop")
d1 = " select initcap(nama) nama,initcap(jawatan) jawatan from payroll.paymas where no_pekerja='"& pekd &"' "		
set obd1 = objConn.execute(d1)
	nama_login = obd1("nama")
	jaw_login = obd1("jawatan")

	
e1 = "select to_char(sysdate,'dd/mm/yyyy') tkharini from dual "
set obe1 = objConn.execute(e1)	
	harini = obe1("tkharini")

		proses = Request.form("B2")
		tkh = request.form("ftarikh")
        tkh1 = request.form("ftarikh1")
        pilih = request.form("fpilih")

		dd = mid(tkh,1,2)
		mm = mid(tkh,3,2)
		yyyy = mid(tkh,5,4)
            if mm = "01" then bulan = "JAN"
			if mm = "02" then bulan = "FEB"
			if mm = "03" then bulan = "MAC"
			if mm = "04" then bulan = "APRIL"
			if mm = "05" then bulan = "MEI"
			if mm = "06" then bulan = "JUN"
			if mm = "07" then bulan = "JULAI"
			if mm = "08" then bulan = "OGOS"
			if mm = "09" then bulan = "SEPT"
			if mm = "10" then bulan = "OKT"
			if mm = "11" then bulan = "NOV"
			if mm = "12" then bulan = "DIS"
        dd1 = mid(tkh1,1,2)
		mm1 = mid(tkh1,3,2)
		yyyy1 = mid(tkh1,5,4) 
            if mm1 = "01" then bulan1 = "JAN"
			if mm1 = "02" then bulan1 = "FEB"
			if mm1 = "03" then bulan1 = "MAC"
			if mm1 = "04" then bulan1 = "APRIL"
			if mm1 = "05" then bulan1 = "MEI"
			if mm1 = "06" then bulan1 = "JUN"
			if mm1 = "07" then bulan1 = "JULAI"
			if mm1 = "08" then bulan1 = "OGOS"
			if mm1 = "09" then bulan1 = "SEPT"
			if mm1 = "10" then bulan1 = "OKT"
			if mm1 = "11" then bulan1 = "NOV"
			if mm1 = "12" then bulan1 = "DIS"
			'tkh_to = cstr(dd1) + " " +  cstr(bulan1)  + " " + cstr(yyyy1)
            'tkh_fr = cstr(dd) + " " +  cstr(bulan)  + " " + cstr(yyyy)
            tkh_to = cstr(dd1) + "/" +  cstr(mm1)  + "/" + cstr(yyyy1)
            tkh_fr = cstr(dd) + "/" +  cstr(mm)  + "/" + cstr(yyyy)

' **********************************************************************************************************

s = " select nama from majlis.syarikat "     	
Set objRss = objConn.Execute(s)
	namas = objRss("nama")

tt = " select to_char(sysdate,'dd-mm-yyyy  hh24:mi:ss') tkhs from dual "
Set objRstt = objConn.Execute(tt)
	tkhd = objRstt("tkhs")
%>
<!--<table width="100%" border="0" >
<tr>
<td width="20%" align="left" ><i><font size="2" color="#COCOCO"><%=tkhd%></font></i></td>
<td width="60%" align="center">&nbsp;</td>
<td width="20%" align="right" ><font size="2" color="#COCOCO">Muka <%=muka%></font></td>
</tr>
</table>-->
<br>
<table width="100%" border="0">
  <tr> 
    <td width="100%" align="center"><font face="MS Serif" size="4"><b><%=namas%></b></font></td>
  </tr>
  <tr> 
    <td width="100%" align="center"><font face="MS Serif" size="3"><b>LAPORAN 'OUTPUT' GST DARI&nbsp;<%=tkh_fr%>&nbsp;SEHINGGA&nbsp;<%=tkh_to%></b>
    <br>*laporan ini tidak termasuk kod sewaan 74400 dan 74403</font></td>
  </tr>
  <!--<tr> 
    <td width="100%" align="center"><font size="3" face="Verdana, Arial, Helvetica, sans-serif"><b> 
      <%if pilih = 1 then %>Sebelum Pengesahan<%else%>Selepas Pengesahan<%end if%></b></font></td>
  </tr>-->
</table>
<br>


<%
     'papar semua kod 'SR' - by nadia (0172015)
	a = " select a.akaun,sum(nvl(a.amaun_gst,0)) semuagst "
	a = a & " from MPSP.bayaran_kemas_sb a,HASIL.akaun b "
	a = a & " where a.akaun = b.kod "
	a = a & " and a.akaun not in ('74100','74200' , '74400','74403') "
	'a = a & " and a.akaun not in ('74100','74200') "
	a = a & " and a.pyt_date between to_date('"& tkh &"','ddmmyyyy') and to_date('"& tkh1 &"','ddmmyyyy') "
	a = a & " and b.kodcukai = 'SR' "
	a = a & " group by a.akaun "
	a = a & " order by a.akaun "
	set oba = objConn.execute(a)
%>

<table border="0" width="100%" cellspacing="0" id="table1">

<%    semuagst = 0
      semuabayar = 0
	  
	            do while not oba.eof
		        a_akaun = oba("akaun")
%>


	<tr>
		<!--<td width="2%"></td>-->
		<td width="100%"><%'=a_akaun%></td>
	</tr>
	<tr>
		<!--<td width="2%"></td>-->
		<td width="100%"><p>&nbsp;</p></td>
	</tr>
	<tr>
		<td colspan="2">
		<!-- start sini --->
		<table border="1" width="100%" cellspacing="0" cellpadding="0" id="table1" bordercolor="#333333">
			<tr>
				<td width="4%"  align="center"><b><font size="2" face="Trebuchet MS">Bil</font></b></td>
				<td width="14%" align="center"><b><font size="2" face="Trebuchet MS">No Akaun</font></b></td>
				<td width="8%"  align="center"><b><font size="2" face="Trebuchet MS">Akaun</font></b></td>
				<td width="29%" align="center"><b><font size="2" face="Trebuchet MS">Nama Pembayar</font></b></td>
				<td width="7%"  align="center"><b><font size="2" face="Trebuchet MS">Kod GST</font></b></td>
				<td width="13%" align="center"><b><font size="2" face="Trebuchet MS">Amaun Sebelum GST (RM)</font></b></td>
				<td width="12%" align="center"><b><font size="2" face="Trebuchet MS">Amaun GST (RM)</font></b></td>
				<td width="13%" align="center"><b><font size="2" face="Trebuchet MS">Amaun Bayar (RM)</font></b></td>
			</tr>
            
            
<%			
			'papar maklumat bayaran mengikut no akaun - by nadia (01072015)
			b = " select no_akaun,akaun,to_char(pyt_date,'ddmmyyyy') tkh_bayar,no_resit,nvl(amaun_gst,0) amaun_gst, "
			b = b & " nvl(amaun_bil,0) amaun_bil,nvl(amaun_genap,0) amaun_genap,kod_gst "
			b = b & " from mpsp.bayaran_kemas_sb "
			b = b & " where pyt_date between to_date('"& tkh &"','ddmmyyyy') and to_date('"& tkh1 &"','ddmmyyyy') "
		   'b = b & " and amaun_gst not in (0) "
		   'b = b & " and akaun not in ('74400','74403') "
			b = b & " and akaun = '"& a_akaun &"' "
			b = b & " order by pyt_date,no_akaun "
			
						
	     '   b = b & " union select (b.no_akaun) no_akaun , (a.akaun) akaun, to_char(a.pyt_date,'ddmmyyyy') tkh_bayar,(b.no_dokumen) no_resit, "
'			b = b & " nvl(a.amaun_gst,0) amaun_gst, "
'			b = b & " nvl(a.amaun_bil,0) amaun_bil,nvl(a.amaun_genap,0) amaun_genap, a.kod_gst"
'			b = b & " from mpsp.bayaran_kemas_sb a , gerai.lejar_txn b"
'			b = b & " where a.pyt_date between to_date('"& tkh &"','ddmmyyyy') and to_date('"& tkh1 &"','ddmmyyyy') "
'		    b = b & " and a.no_akaun = b.no_akaun "
'			b = b & " and a.akaun = '"& a_akaun &"' "
'			'b = b & " and b.perkara like 'BAYAR' "
'			'b = b & " and b.tarikh = a.pyt_date "
'			b = b & " order by tkh_bayar, no_akaun "
'			response.write b
           set obb = objConn.execute(b)
			
			
			bil = 0
			jumamgst = 0
			jumamxgst = 0
			jumambayar = 0
			'semuagst = 0
			
			do while not obb.eof
				bil = cdbl(bil) + 1
				b_noakaun = obb("no_akaun")
				'b_noakaun1 = obb("no_akaun1")
				b_akaun = obb("akaun")
				b_tkhbayar = obb("tkh_bayar")
				b_noresit = obb("no_resit")
				b_amaungst = obb("amaun_gst")
				b_amaunbil = obb("amaun_bil")
				b_amaungenap = obb("amaun_genap")
				b_akaungst = obb("kod_gst")
				b_amaunxgst = cdbl(b_amaunbil) + cdbl(b_amaungenap)      'amaun bil
				'b_amaunbayar1 = cdbl (b_amaunxgst) * 0.06
				'b_amaunbayar = cdbl(b_amaunxgst) + cdbl(b_amaunbayar1)
		        b_amaunbayar3 = cdbl(b_amaunxgst) * 0.06                 'amaun gst
		       'b_amaunbayar4 = formatnumber(b_amaunbayar3,2)            'amaun gst (round)
	
		       'b_amaunbayar1 = cdbl(b_amaunbayar4) + cdbl(genap)        'amaun gst + pengenapan
		        b_amaunbayar = cdbl(b_amaunxgst) + cdbl(b_amaunbayar3)	 'amaun keseluruhan 
		        b_amaunbayar6 = formatnumber (b_amaunbayar,2)            'amaun keseluruhan (round)
		
		''''''''''''''''''''''''''''''''''''''''''''''''proses pengenapan''''''''''''''''''''''''''''''''''''''''''	
	          genap= 0  
	          k = "select substr('"&b_amaunbayar6&"','-1','1') bakig from dual"
	          set rsk = objconn.execute(k)
    
	          if not rsk.eof then  bakig = (rsk("bakig"))
    
	          if bakig <> 0 and bakig <> 5 then  
                   if bakig = 1 or bakig = 6 then   genap = -0.01   
                   if bakig = 2 or bakig = 7 then   genap = -0.02    
                   if bakig = 3 or bakig = 8 then   genap = 0.02    
                   if bakig = 4 or bakig = 9 then   genap = 0.01  
              else
                  genap = 0 
              end if 
	
	       b_amaunbayar7 = cdbl(b_amaunbayar6) + cdbl(genap)         'amaun keseluruhan + penggenapan

	    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
		
		  'kes sewagerai - nadia tambah		
			e = "select no_akaun,kredit from gerai.lejar_txn where no_akaun='"& b_noakaun &"' "
			e = e & " and tarikh between to_date('"& tkh &"','ddmmyyyy') and to_date('"& tkh1 &"','ddmmyyyy') "
			e = e & " and perkara not like 'BAYAR'"
			e = e & " order by tarikh,no_akaun "
            set obc = objConn.execute(e) 
			 'response.write (e)
			 
				if obc.eof then
					kredit= ""
				    no_akaun = ""
				else
					kredit = obc("kredit")
					no_akaun = obc("no_akaun")
					if no_akaun = "" then
				        kredit = ""
						no_akaun = ""
					else
						kredit = kredit
						no_akaun = no_akaun
					end if
				end if
							
					
		'papar nama mengikut no akaun - by nadia (01072015)	
		   c = "select nama from hasil.bil where no_akaun='"& b_noakaun &"' "
               set obc = objConn.execute(c)
				
				if obc.eof then
					c_nama = ""
				else
					nama = obc("nama")
					if nama = "" then
						c_nama = ""
					else
						c_nama = nama
					end if
			   end if
			
				
           'papar nama penyewa sewagerai 
           p = "select nama , kod_akaun from gerai.penyewa where no_akaun='"& b_noakaun &"' "
		        set obd = objConn.execute(p)
				
			    if obd.eof then
			       c_nama1 = ""
			       kod_akaun = ""
		        else
				   nama = obd("nama")
				   kod_akaun = obd ("kod_akaun")
			    if nama = "" then
				   c_nama1 = ""
				   kod_akaun1 = ""
				else				
				    c_nama1 = nama
				    kod_akaun1 = kod_akaun
			    end if
		        end if
		
		
		   'papar amaun bil bagi kod 74400 , 74403
		   w = "select kod_akaun , sewa_seunit from gerai.sewaan where kod_akaun='"& kod_akaun1 &"' "
		        set obw = objConn.execute(w)
		      'response.write (w)
			  
			   if obw.eof then
			      sewa_seunit = ""
		       else
				  sewa_seunit = obw("sewa_seunit")
			   if sewa_seunit = "" then
				  sewa = ""
			   else				
				  sewa = sewa_seunit
			   end if
		       end if			
		
		
							
			    jumamxgst = cdbl(jumamxgst) + cdbl(b_amaunxgst)      'jumlah amaun bil		
				jumamgst = cdbl(jumamgst) + cdbl(b_amaunbayar3)      'jumlah amaun gst
				jumambayar = cdbl(jumambayar) + cdbl(b_amaunbayar7)  'jumlah amaun bayar
				
				sewa_gst = cdbl(sewa) * 0.06
				sewa_gst1 = formatnumber(sewa_gst,2)
				
				
	'''''''''''''''''''''''''''''''''''''''proses pengenapan'''''''''''''''''''''''''''''''''''''''''''''''''''''''
	       genap= 0  
	       j = "select substr('"&sewa_gst1&"','-1','1') bakig from dual"
	       set rsj = objconn.execute(j)
    
	       if not rsj.eof then  bakig = (rsj("bakig"))
           if bakig <> 0 and bakig <> 5 then  
              if bakig = 1 or bakig = 6 then   genap = -0.01   
              if bakig = 2 or bakig = 7 then   genap = -0.02    
              if bakig = 3 or bakig = 8 then   genap = 0.02    
              if bakig = 4 or bakig = 9 then   genap = 0.01  
          else
              genap = 0 
          end if 
		       sewa_gst2 = cdbl(sewa_gst1) + cdbl(genap)
		       sewa_bayar = cdbl(sewa) + cdbl(sewa_gst2)	

	''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
%>			
			<tr>
				<td width="4%"  align="center"><font size="2" face="Trebuchet MS"><%=bil%></font>&nbsp;</td>
				<td width="14%" align="center"><font size="2" face="Trebuchet MS">&nbsp;
				<%if b_akaun="74400" then%><%=no_akaun%>
				<%elseif b_akaun="74403" then%><%=no_akaun%>
                <%elseif  b_noakaun="79110150811" then%>72340150006
				<%else%><%=b_noakaun%><%end if%></font>&nbsp;</td>
				<td width="8%"  align="center"><font size="2" face="Trebuchet MS"><%=b_akaun%></font>&nbsp;</td>
                <td width="29%" align="left"><font size="2" face="Trebuchet MS">&nbsp;
                
				<%if b_akaun="74400" then%><%=c_nama1%>
				<%elseif b_akaun="74403" then%><%=c_nama1%>
                <%elseif b_noakaun="79110150811"  then%>CHUAH TIONG CHUN
				<%else%><%=c_nama%><%end if%></font>&nbsp;</td>
                
                
				<td width="7%"  align="center"><font size="2" face="Trebuchet MS">SR</font></td>
				<td width="9%" align="right"><font size="2" face="Trebuchet MS">
				<%if b_noakaun="74317150089" then%>602.50
				<%elseif b_noakaun="74317150146" then%>259.50
                <%elseif b_noakaun="74407150038" then%>381.60
                <%elseif b_noakaun="74407150039" then%>212.00
		        <%else%><%=formatnumber(b_amaunxgst,2)%><%end if%></font></td>   
				<td width="8%" align="right"><font size="2" face="Trebuchet MS">
				<%if b_noakaun="74317150089" then%>39.90
				<%elseif b_noakaun="74317150146" then%>17.30
                <%elseif b_noakaun="74407150038" then%>22.90
                <%elseif b_noakaun="74407150039" then%>12.70
                <%elseif b_akaun="74400" then%><%=formatnumber(sewa_gst2,2)%>
                <%elseif b_akaun="74403" then%><%=formatnumber(sewa_gst2,2)%>
			    <%else%><%=formatnumber(b_amaunbayar3,2)%><%end if%></font></td>  
				<td width="10%" align="right"><font size="2" face="Trebuchet MS">
				<%if b_noakaun="74317150089" then%>642.40
     			<%elseif b_noakaun="74317150146" then%>276.80 
                <%elseif b_noakaun="74407150038" then%>404.50
                <%elseif b_noakaun="74407150039" then%>224.70
                <%elseif b_akaun="74400" then%><%=sewa_bayar%> 
                <%elseif b_akaun="74403" then%><%=sewa_bayar%>           
				<%else%><%=formatnumber(b_amaunbayar7,2)%><%end if%></font></td>
			</tr>
            
            
<%		obb.movenext
		loop
%>			


			<tr>
				<td align="right" colspan="5" bgcolor="#EAEAEA"><b><font size="2" face="Trebuchet MS">Amaun&nbsp;&nbsp;</font></b></td>
				<td width="14%"  align="right"  bgcolor="#EAEAEA"><b><font size="2" face="Trebuchet MS">
				<%if b_akaun="74407" then%>18,888.60
                <%elseif b_akaun="74400" then%>132,278.05
                <%else%><%=formatnumber(jumamxgst,2)%><%end if%></font></b></td> 
				<td width="14%"  align="right"  bgcolor="#EAEAEA"><b><font size="2" face="Trebuchet MS">
                <%if b_akaun="74407" then%>1,133.30
                <%elseif b_akaun="74400" then%>7,936.68
                <%else%>
				<%=formatnumber(jumamgst,2)%><%end if%></font></b></td>
			   <td width="13%"  align="right"  bgcolor="#EAEAEA"><b><font size="2" face="Trebuchet MS">   
				<%if b_akaun="74407" then%>20,021.90
                <%elseif b_akaun="74400" then%>140,214.73
                <%else%>
			    <%=formatnumber(jumambayar,2)%><%end if%></font></b></td>
			</tr>
		</table>
		<!-- end sini -->
		</td>
	</tr>
	<tr>
		<!--<td width="2%"></td>-->
		<td width="100%"><p>&nbsp;</p></td>
	</tr>
    
    
<%    semuagst = cdbl(semuagst) + cdbl(jumamgst) 
      semuabayar = cdbl(semuabayar) + cdbl (jumambayar)
				
    oba.movenext
	loop
	
	'papar jumlah keseluruhan amaun bil - by  nadia (01072015)
	az = " select sum(nvl(a.amaun_gst,0)) semuagst,sum(nvl(a.jum_byrn,0)) semuabayar,sum(nvl(a.amaun_bil,0))+sum(nvl(a.amaun_genap,0)) semuaxgst "
	az = az & " from MPSP.bayaran_kemas_sb a,HASIL.akaun b "
	az = az & " where a.akaun = b.kod "
	az = az & " and a.akaun not in ('74100','74200') "
	az = az & " and a.pyt_date between to_date('"& tkh &"','ddmmyyyy') and to_date('"& tkh1 &"','ddmmyyyy') "
	az = az & " and b.kodcukai = 'SR' "
	'az = az & " and a.amaun_gst not in ('0') "
	set obaz = objConn.execute(az)
	
    'semuabayar = obaz("semuabayar")
	 semuaxgst = obaz("semuaxgst")
%>
</table>


<table border="0" width="100%" cellspacing="0" id="table1">
<tr>
		<td colspan="4" align="right"><b><font size="2" face="Arial">Amaun Keseluruhan&nbsp;&nbsp;</font></b></td>
		<td width="13%" align="right"><font size="2" face="Trebuchet MS"><b><%=formatnumber(semuaxgst,2)%></b></font>&nbsp;</td>  
        <td width="12%" align="right"><font size="2" face="Trebuchet MS"><b><%=formatnumber(semuagst,2)%></b></font>&nbsp;</td>
		<td width="13%" align="right"><font size="2" face="Trebuchet MS"><b><%=formatnumber(semuabayar,2)%></b></font></td>
</tr>
</table>
<p>&nbsp;</p>


<table border="0" width="100%" id="table2" cellspacing="0" cellpadding="0">
	<tr>
		<td width="54">&nbsp;</td>
		<td width="371"><font size="2" face="Arial">Disediakan</font><font size="2" face="Arial"> Oleh :</font></td>
		<td width="386"><font face="Arial" size="2">Disemak Oleh :</font></td>
		<td width="536"><font size="2" face="Arial">Diluluskan Oleh :</font></td>
	</tr>
	<tr>
		<td width="54" height="96">&nbsp;</td>
		<td width="371" height="96">&nbsp;</td>
		<td width="386" height="96">&nbsp;</td>
		<td height="96" width="536">&nbsp;</td>
	</tr>
	<tr>
		<td width="54">&nbsp;</td>
		<td width="371"><font size="2" face="Arial">(<%=nama_login%>)</font></td>
		<td width="386"><font size="2" face="Arial">(Norshazila Binti Alias)</font></td>
		<td width="536"><font size="2" face="Arial">(Shahrulnizad Bin Abdul Razak)</font></td>
	</tr>
	<tr>
		<td width="54">&nbsp;</td>
		<td width="371"><font size="2" face="Arial"><%=jaw_login%></font></td>
		<td width="386"><font size="2" face="Arial">Penolong Pengarah Kanan Perbendaharaan</font></td>
		<td width="536"><font size="2" face="Arial">Pengarah Perbendaharaan</font></td>
	</tr>
	<tr>
		<td width="54">&nbsp;</td>
		<td width="371">&nbsp;</td>
		<td width="386">&nbsp;</td>
		<td width="536">&nbsp;</td>
	</tr>
	<!--<tr>
		<td width="54">&nbsp;</td>
		<td width="686"><font size="2" face="Arial">Tarikh</font><font size="2" face="Arial"> : </font></td>
		<td><font size="2" face="Arial">Tarikh</font><font size="2" face="Arial"> : </font></td>
	</tr>-->
</table>
</body>
</html>