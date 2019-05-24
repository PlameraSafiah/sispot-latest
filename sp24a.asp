<%response.buffer=true%>
<!--#include file="connection.asp" -->
<%Set objConn = Server.CreateObject("ADODB.Connection")
  objConn.Open "dsn=12c;uid=majlis;pwd=majlis;" %>
<%' response.cookies("amenu") = "ajsip.asp" %>
<%response.cookies("amenu") = "a812.asp" %>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link type="text/css" href="css/menu1.css" REL="stylesheet">
<script type="text/javascript" src="menu1.js"></script>
<FORM name="" action="sp24a.asp" method="post" >
<!--#include file="spmenu.asp"-->
 
<%'if b = "" then b = request.querystring("B1")
  b1 = request.form("b1") 'button lain
  b = request.form("b")  'hantar
 
  rowid = request.form(""& rowid &"")
  jenis_bank = request.form("jenis_bank")
  jenis_bank = replace(jenis_bank,"'","''")
  bank = ucase(request.form("bank"))
  bank = replace(bank,"'","''")
     d = request.form(""& d &"")
     e = request.form(""& e &"")
	 c = request.form(""& c &"")
  nama_bank = request.form("nama_bank")
  nama_bank = replace(nama_bank,"'","''")
  alamat1 = request.form("alamat1")
  alamat1 = replace(alamat1,"'","''")
  alamat2 = request.form("alamat2")
  alamat2 = replace(alamat2,"'","''")
  alamat3 = request.form("alamat3")
  alamat3 = replace(alamat3,"'","''")
  tel = request.form("tel")
  tel = replace(tel,"'","''")
  faks = request.form("faks")
  faks = replace(faks,"'","''")
  pengurus = ucase(request.form("pengurus"))
  pengurus = replace(pengurus,"'","''")
  pegawai1 = ucase(request.form("pegawai1"))
  pegawai1 = replace(pegawai1,"'","''")
  tel1 = request.form("tel1")
  tel1 = replace(tel1,"'","''")
  pegawai2 = ucase(request.form("pegawai2"))
  pegawai2 = replace(pegawai2,"'","''") 
  tel2 = request.form("tel2")
  tel2 = replace(tel2,"'","''")
  pegawai3 = ucase(request.form("pegawai3"))
  pegawai3 = replace(pegawai3,"'","''")
  tel3 = request.form("tel3")
  tel3 = replace(tel3,"'","''")
  hari_tahun = request.form("hari_tahun")
  hari_tahun = replace(hari_tahun,"'","''")
  

  
 if id_ak = "" then id_ak = request.querystring("id_ak")

if b = "" then b = request.querystring("B1")

 if b = "" then
  end if 
 
 

	  '*********************************kemaskini*********************************************** 
 
   if b1 = "Simpan" then
   

  
    u = " update iabs.nama_bank "
    u = u & " set  jenis_bank = '"& jenis_bank &"',bank = '"& bank &"', nama_bank = '"& nama_bank &"', "
    u = u & " alamat1 = '"& alamat1 &"',alamat2 = '"& alamat2 &"',alamat3 = '"& alamat3 &"', tel= '"&  tel &"',faks = '"& faks &"',pengurus = '"& pengurus &"',"
    u = u & " pegawai1 = '"& pegawai1 &"',tel1 = '"& tel1 &"', pegawai2= '"& pegawai2 &"', "
 	u = u & " tel2 = '"&tel2 &"',pegawai3 = '"& pegawai3 &"', "
    u = u & " tel3= '"& tel3 &"',"
    u = u & " hari_tahun = '"& hari_tahun &"' "
    u = u & " where bank = '"& bank &"' "
   'response.write(u)
	set su = objconn.execute(u)
	
	auditLog "Pelaburan >> Input Bank","a812e.asp","Edit","iabs.nama_bank","Update","Bank="&bank&", Jenis Bank = "&jenis_bank&", Nama Bank = "&nama_bank&", Alamat = "&alamat1&" "&alamat2&" "&alamat3&", Telefon = "&tel&", No Faks = "&faks&", Pengurus = "&pengurus&", Pegawai 1 = "&pegawai1&", Telefon Pegawai 1 = "&tel1&", Pegawai 2 = "&pegawai2&", Telefon Pegawai 2 = "&tel2&" , Pegawai 3 = "&pegawai3&", Telefon Pegawai 3  = "&tel3&", Bil Hari Setahun  = "& hari_tahun&". "

	  response.write "<script language = ""vbscript"">"
	  response.write " MsgBox ""Data Di Kemaskini!"", vbInformation, ""Perhatian!"" "
	  response.write "</script>"
   ' end if
 end if


 

 '**********************************************keluarkan data apabila dipanggil***********************************
 
    a =     " select rowid,jenis_bank,bank,nama_bank,alamat1,alamat2,alamat3,tel,faks,pengurus,pegawai1,tel1,pegawai2,tel2,pegawai3,tel3,hari_tahun"
	a = a & " from iabs.nama_bank  "
	a = a & " where bank ='"& bank  &"'"
 
 
  Set Rsa = objConn.Execute(a)
  
  if not RSa.eof then 
    ctrz       = cdbl(ctrz) + 1
    rowid      = rsa("rowid")
	jenis_bank = rsa("jenis_bank")
	bank       = rsa("bank")
    nama_bank  = rsa("nama_bank")
    alamat1    = rsa("alamat1")
    alamat2    = rsa("alamat2")
	alamat3    = rsa("alamat3")
    tel        = rsa("tel")
    faks       = rsa("faks")
	pengurus   = rsa("pengurus")
	pegawai1   = rsa("pegawai1") 
	tel1       = rsa("tel1")
    pegawai2   = rsa("pegawai2")
	tel2       = rsa("tel2")
    pegawai3   = rsa("pegawai3")
    tel3       = rsa("tel3")
	hari_tahun = rsa("hari_tahun")
	warna = ctrz mod 2
	
	end if%>

	 <TABLE border=1 borderColor=black cellPadding=1 cellSpacing=0 rules=all class="hd">
  <TBODY>
    <tr class="hd"> 
      <td colspan="10" >Senarai Arahan Kerja</td>
    </tr>
    <tr style="BACKGROUND-COLOR: #ffffff"> 
      <td colspan="10">&nbsp;</td>
    </tr>
    <tr class="hd"> 
      <td align="center">Bil</td>
      <td align="center">No Rujukan Arahan Kerja</td>
      <td align="center">Bln/Thn</td>
      <td align="center">Pemohon</td>
      <td align="center">Keterangan</td>
      <td align="center">Nota Tambahan</td>
      <td align="center">Bil Pekerja</td>
      <td align="center">Anggaran OT (RM) </td>
      <td align="center">Proses</td>
    </tr>
	 <%
 '******************************************display data form ************************************************
 

q1="select id_ak,no_pekerja,tkh_ot  from payroll.pekerja_ot where id_ak ="& id_ak &" "



set oq1=objConn.Execute(q1)

bil=0
tot_angg_semua=0
Do While Not oq1.eof
bil=bil + 1

q2="select upper(nama) as nama from payroll.paymas where no_pekerja="& oq1("no_pekerja") &""
set oq2 = objConn.Execute(q2)
If not oq2.bof and not oq2.eof then
	nama = oq2("nama")
end if

q3="select count(distinct no_pekerja) as bil,nvl(sum(angg_ot_S),0) + nvl(sum(angg_ot_M),0) as tot_angg from payroll.pekerja_ot where id_ak='"& oq1("id_ak") &"'"
set oq3 = objConn.Execute(q3)
if not oq3.bof and not oq3.eof then
	tot_angg_semua = cdbl(tot_angg_semua) + cdbl(oq3("tot_angg"))
end if
  %>
<p align="center"><strong>EDIT BANK</strong></p>
	
   
<form name="test" method="post" action="sp24.asp" onSubmit="return check(this)">

      <form name="test<%=bil%>"  method="post" action="sp24.asp">
    <input type="hidden" name="pemohon" value="<%=pemohon%>">
    <input type="hidden" name="blnthn" value="<%=blnthn%>">
    <tr> 
      <td><%=bil%></td>
      <td> <a href="sp24a.asp?id_ak=<%=oq1("id_ak")%>"><font face="Verdana" size="2" color="Blue"><b><%=oq1("id_ak")%></b></font></a></td>
      <td ><%=oq1("tkh_ot")%></td>
      <td ><%=oq2("nama")%>&nbsp;(<%=oq1("no_pekerja")%>)</td>
      <td><%'=oq1("keterangan")%></td>
      <td><%'=oq1("nota")%></td>
      <td><div align="center"><%=oq3("bil")%></div></td>
      <td><div align="right"><%=oq3("tot_angg")%></div></td>
      <td> <input type="hidden" name="tkh_post_ak" value="<%=tkh_post_ak%>"> <input type="hidden" name="id_ak" value="<%=oq1("id_ak")%>"> 
        <div align="center">
          <%if cdbl(oq3("bil")) > 0 then %>
          <input type="submit" name="proses" value="Sah">
          <%end if%>
        </div></td>
    </tr>
  </form>
  <%
  oq1.movenext
  loop %>
  <tr> 
    <td colspan="7"><div align="right"><b>Jumlah Keseluruhan Anggaran OT</b></div></td>
    <td><div align="right"><%=formatnumber(tot_angg_semua,2)%></div></td>
    <td>&nbsp;</td><input type="submit" value="Simpan" name="b1" style="font-family: Tahoma; font-size: 9pt; font-weight: bold"> 
  </tr></tbody>

       

 

</body></table>




