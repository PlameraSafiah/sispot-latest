<!--#include file="connection.asp" -->
<%response.buffer=true%>
<%response.cookies("amenu") = "sp27a.asp"%>
<!--'#include file="spmenu.asp"-->


<html>
<head>
<title>Cetak Senarai Pekerja</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link type="text/css" href="css/menu1.css" REL="stylesheet">
<BODY leftMargin=0 onLoad="" topMargin=0 bgcolor="<%=color4%>">

<script type="text/javascript" src="menu1.js"></script>
<SCRIPT LANGUAGE="JavaScript">
<!-- Begin
function check(f){
	if (f.blnthn.value=="")
	{
		alert("Sila masukkan bulan dan tahun arahan kerja");
		f.blnthn.focus();
		return false
	}
}
//  End -->
</script>
<script language="javascript">
function submitForm1()
{
	document.test.submit();
}
</script>
</head>


<BODY leftMargin=0 onLoad="" topMargin=0 bgcolor="<%'=color4%>">
<% 
proses = request.form("proses")
pemohon = request.form("pemohon")
blnthn = request.form("blnthn")
tkh_ot = request.form("tkh_ot")
id_ak = request.form("id_ak")
tkh_ot2 = request.form("tkh_ot2")
ptj=plokasi
p7 = request.form("b7")	'cetak permohonan


papar
if proses <> "" then	
			borang
end if



sub papar ()
if blnthn = "" then
Dim TodayYearNumber, TodayMonthName
TodayYearNumber = DatePart("yyyy",Date)
TodayMonth = DatePart("m",Date)
if len(TodayMonth)="1" then TodayMonth="0"& TodayMonth
blnthn=TodayMonth&TodayYearNumber
end if


   'check jabatan by id pekerja
    pekd = request.cookies("gnop")
    a =     "select jabatan from majlis.pengguna where no_pekerja = '"& pekd &"' "
    Set rsa = objConn.Execute(a)


    'set bulan/tahun - arahan kerja 
    q4="select distinct to_char(tkh_ot,'mm/yyyy') as tkh_ot"
    q4 = q4 & " from payroll.jadual_ot order by tkh_ot"  
    set oq4 = objConn.Execute(q4)
    tkh_ot2 = oq4 ("tkh_ot")


    idd = "select a.nama, a.lokasi, b.keterangan, to_char(sysdate,'dd/mm/yyyy') tkh  from payroll.paymas a, iabs.jabatan b "
    idd = idd & " where a.no_pekerja = '"&pekd&"' and b.kod = a.lokasi "
    set idd2 = objConn.Execute(idd)
    'response.write idd
    lokasi = idd2 ("lokasi")


%>
<br>


<br>
<form name="test" method="post" action="sp27a.asp" onSubmit="return check(this)">
<TABLE border=1 borderColor=black cellPadding=1 cellSpacing=0 rules=all class="hd">
   <tr class="hd">
      <td colspan="2">Laporan Arahan Kerja - Mengikut Senarai Pekerja</td>
    </tr>
    
  <tr>
  <td width="15%"><b>Bulan/Tahun (mm/yyyy) : </b></td>
  <td><!--<input type="text" name="blnthn" size="8" value="<%=tkh_ot%>">-->
     <select size="1" name="tkh_ot" onChange="submitForm1(this)">
          <option value="">Pilih Bulan/Tahun</option>
          <%     	  
    'semak bulan/tahun OT 
   q2="select distinct to_char(a.tkh_ot,'mm/yyyy') as tkh_ot "
   q2 = q2 & " from payroll.pekerja_ot a , payroll.arahan_kerja b where a.id_ak=b.id_ak and b.ptj = '"&ptj&"' "
   q2 = q2 & " order by tkh_ot desc"  
   set oq2 = objConn.Execute(q2)
  ' response.write q2
   if not oq2.bof and oq2.eof then
   tkh_ot = oq2 ("tkh_ot")
   end if
		  	  
		  do while not oq2.EOF  %>
          <option <%if tkh_ot = oq2("tkh_ot") then%> selected <%end if%>value="<%=oq2("tkh_ot")%>"><%=oq2("tkh_ot")%></option>
          <%
            oq2.MoveNext
            loop
'------
oq2.close
'------
           %>
    </select>
    </td>
    </tr> 
    
 
  <tr>
  <td><b>Nama Pemohon : </b></td>
  <td><!--<input type="text" name="pekerja" size="8" value="<%=pemohon%>">-->
  <select size="1" name="pemohon" onChange="submitForm1(this)">
  <option value="">Pilih Pemohon</option>
          <% 
		  
   'papar senarai pemohon by blnthn
   q2b="select distinct a.pemohon, to_char(c.tkh_ot,'mmyyyy') , b.nama "
   q2b = q2b & " from payroll.arahan_kerja a, payroll.pekerja_ot c , payroll.paymas b where a.ptj= "&ptj&" and to_char(c.tkh_ot,'mm/yyyy') = '"& tkh_ot &"' "
   q2b = q2b & "and a.id_ak = c.id_ak and a.pemohon = b.no_pekerja order by  a.pemohon asc "
   set oq2b = objConn.Execute(q2b)
   'response.write  q2b
   if not oq2b.bof and oq2b.eof then
   pemohon = oq2b ("pemohon")
   end if
		  	  
		  do while not oq2b.EOF %>
          <option <%if cstr(pemohon) =  cstr(oq2b("pemohon")) then%> selected <%end if%>value="<%=oq2b("pemohon")%>"><%=oq2b("pemohon")%> - <%=oq2b("nama")%></option>
          <%
            oq2b.MoveNext
            loop
'------
oq2b.close
'------
           %>
    </select>
  </td>
</tr>


  <tr>
  <td width="15%"><b>Arahan Kerja  : </b></td>
  <td><!--<input type="text" name="blnthn" size="8" value="<%=id_ak%>">-->
     <select size="1" name="id_ak" onChange="submitForm1(this)">
          <option value="">Pilih Bulan/Tahun</option>
          <%     
		  
    'papar id arahan kerja by jabatan
    qq2="select distinct a.id_ak , a.ptj"
    qq2 = qq2 & " from payroll.arahan_kerja a , payroll.pekerja_ot b where  a.id_ak=b.id_ak  and  a.ptj = '"&ptj&"' and a.pemohon = '"& pemohon & "' "
    qq2 = qq2 & " and to_char(b.tkh_ot,'mm/yyyy') = '"& tkh_ot &"' order by a.id_ak desc"  
    set oqq2 = objConn.Execute(qq2)
	if not oqq2.bof and oqq2.eof then
    id_ak = oqq2 ("id_ak")
	end if
		  		  		  
		  do while not oqq2.EOF  %>
          <option <%if id_ak = oqq2("id_ak") then%> selected <%end if%>value="<%=oqq2("id_ak")%>"><%=oqq2("id_ak")%></option>
          <%
            oqq2.MoveNext
            loop
'------
oqq2.close
'------
           %>
    </select>
    </td>
    </tr> 

<tr align="center"><td colspan="2"><input type="submit" name="proses" value="Hantar"></td></tr>
</table>
</form>
<%
end sub %>


<%
sub borang()
blnthn=request.form("blnthn")
ptj=plokasi
p7 = request.form("b7")	
id_ak = request.form("id_ak")
tkh_ot = request.form("tkh_ot")
tkh_ot2 = request.form("tkh_ot2")
pemohon = request.form("pemohon")


  ss = " select a.id_ak , a.no_pekerja , to_char(a.masa_tamat_angg,'HH24:MI') as masa_tamat , b.ptj , b.pemohon , b.pengesahan ,   "
  ss = ss & " to_char(a.tkh_ot,'dd/mm/yyyy') as tkh_ot ,  to_char(a.masa_mula_angg,'HH24:MI') as masa_mula , to_char(a.tkh_ot,'mm/yyyy') as tkh_ot2 "
  ss = ss & " from payroll.pekerja_ot a , payroll.arahan_kerja b where a.id_ak = b.id_ak "
  ss = ss & "  and to_char(a.tkh_ot,'mm/yyyy') ='"& tkh_ot &"' and b.pemohon = '"&pemohon&"' "
  ss = ss & " and a.id_ak = '"&id_ak&"' order by tkh_ot , masa_mula ,masa_tamat asc  "
  
  Set rsss = objConn.Execute (ss)
    'response.write ss
    id = rsss ("id_ak")
	tkh_ot = rsss ("tkh_ot")
	tkh_ot2 = rsss ("tkh_ot2")
	pemohon = rsss ("no_pekerja")
	pemohon3 = rsss ("pemohon")
	mula = rsss ("masa_mula")
	tamat = rsss ("masa_tamat")
	ptj = rsss ("ptj")
	pengesahan = rsss ("pengesahan")
	
   
   
    dd = " select nama from payroll.paymas where no_pekerja='"&pemohon3&"' " 
    Set rsdd = objConn.Execute (dd)
	nama = rsdd ("nama")
   
		
  m = " select kod , initcap(keterangan) keterangan from iabs.jabatan "
  m = m & " where kod = '"&ptj&"' "
  Set rsm = objConn.Execute(m)
  ket = rsm ("keterangan")	
	
%>
<br>
<br>




<TABLE border=1 borderColor=black cellPadding=1 cellSpacing=0 rules=all class="hd" height="20%">
  <TBODY>
    <tr class="hd"> 
      <td colspan="10" >Senarai Pekerja Mengikut Arahan Kerja<br> 
      Jabatan <%=ket%><br> 
      Nama Pemohon : <%=nama%> (<%=pemohon3%>)<br>
      Arahan Kerja : <%=id_ak%><br>
      Status Pengesahan : <%if pengesahan="Y" then%><font color="#0000ff"><b>Telah Disahkan</b></font><%end if%><%if pengesahan="T" then%><font color="#b30000"><b>Belum Disahkan</b></font><%end if%><br>
      Bulan <%=tkh_ot2%></td>
    </tr> 
    <tr style="BACKGROUND-COLOR: #ffffff"><td colspan="10">&nbsp;</td></tr>
    <tr class="hd"> 
    <td width="2%"  align="center">Bil</td>
    <td width="8%"  align="center">No Rujukan Arahan Kerja</td>
    <td width="8%"  align="center">Tarikh Dilaksana</td>
    <td width="4%"  align="center">Jabatan</td>
    <td width="6%"  align="center">Masa Mula</td>
    <td width="7%"  align="center">Masa Tamat</td>
    <td width="6%"  align="center">No Pekerja</td>
    <td width="19%" align="center">Nama Pekerja</td>
    <td width="23%" align="center">Keterangan Kerja</td>
    </tr>
  
  
<%  
bil=0
Do While Not rsss.eof 
bil=bil + 1
tkh_ot = rsss("tkh_ot")
id = rsss ("id_ak")	
pemohon = rsss ("no_pekerja")
mula = rsss ("masa_mula")
tamat = rsss ("masa_tamat")
	
	
   se = " select nvl(angg_ot_S,0) + nvl(angg_ot_M,0) anggaran "
   se = se & " from payroll.pekerja_ot"
   se = se & " where id_ak='"&id&"' and no_pekerja='"&pemohon&"' and  to_char(tkh_ot,'dd/mm/yyyy')='"&tkh_ot&"' and to_char(masa_mula_angg,'HH24:MI')= '"&mula&"' " 
   Set rse= objConn.Execute (se)
   'response.write se
   anggaran = rse("anggaran")
  

  p = "  select id_ak, keterangan, nota , ptj "
  p = p & "  from payroll.arahan_kerja "
  p = p & " where id_ak = '"&id&"' "
  Set rsp = objConn.Execute(p)
   Do While Not rsp.eof 
  keterangan1 = rsp ("keterangan")
  nota = rsp ("nota")
  ptj = rsp ("ptj")
  

  n = " select nama , jawatan from payroll.paymas " 
  n = n & " where no_pekerja = '"&pemohon&"' " 
  Set rsn = objConn.Execute (n)
  
   Do While Not rsn.eof 
	nama1 = rsn ("nama")
	jaw = rsn ("jawatan")
%>

<form name="test<%=bil%>"  method="post" action="sp27a.asp" target="_new">
  <tr bgcolor="#EFF5F5">
    <td rowspan="2" align="center"><%=bil%></td>
    <td rowspan="2" align="center"><%=rsss ("id_ak")%></td>
    <td rowspan="2" align="center"><%=tkh_ot%></td>
    <td rowspan="2" align="center"><%=ptj%></td>
    <td rowspan="2" align="center"><%=rsss ("masa_mula")%></td>
    <td rowspan="2" align="center"><%=rsss ("masa_tamat")%></td>
    <td rowspan="2" align="center"><%=rsss("no_pekerja")%></td>
    <td rowspan="2"><%=nama1%></td>
    <td rowspan="2"><%=keterangan1%></td>
  </tr>
 </form>
 <form name="test<%=bil%>"  method="post" action="sp27a.asp" target="_new">
  <tr bgcolor="#EFF5F5">
  </tr>
 </form>  
  
<% 
rsn.movenext
loop
rsp.movenext
loop 
rsss.movenext
loop 

'------
rsss.close
'------
%>  


 <% 
   anggaran = rse("anggaran")	
   sss = " select sum(nvl(a.angg_ot_s,0) + nvl(a.angg_ot_m,0)) semuaxgst1 "
   sss = sss & " from payroll.pekerja_ot a , payroll.arahan_kerja c "
   sss = sss & " where a.id_ak = c.id_ak and to_char(a.tkh_ot,'mm/yyyy') ='"& tkh_ot2 &"' " 
   sss = sss & " and c.id_ak = '"&id_ak&"'   "
   Set rssss = objConn.Execute (sss)
   'response.write sss
   semuaxgst1 = rssss("semuaxgst1")
   'Do While Not rssss.eof 
   'jum_ot_semua = jum_ot_semua + rssss("semuaxgst1")
  '------
rssss.close
objConn.close
'------
  %>
   
  

<% 
'rssss.movenext
'loop 
%> 


  </tbody>
</table>
<br>
  

  <TABLE border=1 borderColor=black cellPadding=1 cellSpacing=0 rules=all class="hd">
    <tr align="center">
    <td colspan="2">
   <form name="test<%=bil%>"  method="post" action="sp27a1.asp?tkh_ot=<%=tkh_ot2%>&id_ak=<%=id_ak%>&amenu='sp27a1'" target="_new">
   <input type="submit" name="b7" value="Cetak Laporan" class="button" onFocus="nextfield='done';">
    </form>
    </td>
     </tr>
  </table>
	   
<% end sub %>
<input type="hidden" name="bilrec" value="<%=ctrz%>" >
</body>
</html>