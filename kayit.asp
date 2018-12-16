<center>
<% @language="javascript" %>
<h2>Ziyaretçi Bilgileri<hr></h2>
<p>Adý....:<%=Request.Form("ad")%>
<p>Soyadý....:<%=Request.Form("soyad")%>
<p>E-Posta....:<%=Request.Form("eposta")%>
<p>Düþünceler:<%=Request.Form("mesaj")%>
<%
fso=Server.CreateObject("scripting.FileSystemObject")
if(fso.FileExists(Server.Mappath("uye.txt"))==true)
dosya=fso.OpenTextFile(Server.Mappath("uye.txt"),8)
else
dosya=fso.CreateTextFile(Server.Mappath("uye.txt"),2)
dosya.WriteLine(Request.Form("ad"))
dosya.WriteLine(Request.Form("soyad"))
dosya.WriteLine(Request.Form("eposta"))
dosya.WriteLine(Request.Form("mesaj"))
dosya.Close()
%>
<h3>Ziyaretçi Defterine Kaydedildiniz.</h3>
<a href="liste.asp">Ziyaretçi Defteri...</a>
<a href="index.asp">Ana Sayfa...</a>
</center>