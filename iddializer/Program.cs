using System;
using System.Collections.Generic;
using System.Linq;
using RestSharp;
using Newtonsoft.Json;
using OfficeOpenXml;
using System.IO;
using System.Drawing;

namespace iddializer
{
    class Program
    {
        static void Main(string[] args)
        {
            var data = new FileInfo(@"data.xlsx"); //Data excel dosyasını oku
            using (var p = new ExcelPackage(data))
            {

                var ws = p.Workbook.Worksheets["BULTEN"]; //Excel dosyasında sayfa seç
                DateTime bugun = DateTime.Now; //Bu günün tarihini al

                logla("=====Transaction Is Starting=====");
                logla("-**-Islem baslangici : " + bugun.ToString());

                try
                {
                    if (ws.Cells["B" + 3].Value != null) // Eğer veri varsa son tarihi ekrana yazdır
                        logla("-**-Son Alınan Tarih " + ws.Cells["B" + Convert.ToString(GetLastUsedRow(ws))].Value); //
                }
                catch
                {
                    logla("-**-\'Data Dosyası Bulunamadı. Program Kapanıyor.\'");
                    Console.ReadKey();
                    Environment.Exit(0);
                }

                Console.WriteLine("Lütfen Verinin Alınmaya Başlayacağı Tarihi G/A/YYYY Şeklinde Girin");
                Console.WriteLine("Boş Bırakmanız durumunda tablo kontrol edilecektir. Tablo boşsa 17/04/2004 tarihinden itibaren bütün veriler istenecektir");
                Console.WriteLine("Örneğin : 9/7/2017");
                string[] baslangicData;
                

                string baslangicKontrol = Console.ReadLine();

                if (baslangicKontrol == "") //Veri Girilmemişse
                {
                    if (ws.Cells["B" + 3].Value != null) //Tabloda Veri Varsa
                    {

                        //Bu Gün Alınmış Mı?
                        string sonBitisData = bugun.ToString(("dd'/'MM'/'yyyy"));
                        if (Convert.ToString(ws.Cells["B" + Convert.ToString(GetLastUsedRow(ws))].Value) == sonBitisData)
                        {
                            logla("-**-Baslanacak tarih : " + bugun.ToString());
                            logla("-**-Bu Gün Zaten Senkoronize Edilmiş. Program Kapatılıyor...");
                            Console.ReadKey();
                            Environment.Exit(0);
                        }


                        baslangicData = Convert.ToString(ws.Cells["B" + GetLastUsedRow(ws)].Value).Split('/'); //Son eklenen verinin tarihini al

                        //Tarihi 1 artır
                        DateTime eskiTarih = new DateTime(Convert.ToInt32(baslangicData[2]), Convert.ToInt32(baslangicData[1]), Convert.ToInt32(baslangicData[0]));
                        DateTime yeniTarih = eskiTarih.AddDays(1);
                        string[] yeniGun = yeniTarih.ToString(("dd'/'MM'/'yyyy")).Split('/');
                        baslangicData = yeniGun;
                        logla("-**-Baslanacak tarih (1 gün eklenmiştir): " + yeniTarih.ToString());

                    }
                    else//Tabloda veri yoksa
                    {
                        baslangicData = "17/04/2004".Split('/'); //İlk veriyi çekmeye başla
                        logla("-**-Baslanacak tarih : 17/04/2004");
                    }
                        

                }

                else //Veri girilmişse
                {
                    baslangicData = baslangicKontrol.Split('/'); //Girilen veriyi diziye at

                    logla("-**-Baslanacak tarih : " + baslangicKontrol);

                    //O tarih daha önce alınmış mı?
                    bool durum = true;
                    DateTime kontrol = new DateTime(Convert.ToInt32(baslangicKontrol.Split('/')[2]), Convert.ToInt32(baslangicKontrol.Split('/')[1]), Convert.ToInt32(baslangicKontrol.Split('/')[0]));
                    for (int i = 3; i < GetLastUsedRow(ws) - 4; i++)
                    {
                        if (Convert.ToString(ws.Cells["B" + i].Value) == kontrol.ToString("dd\\/MM\\/yyyy"))
                        {
                            durum = false;
                            break;
                        }
                    }

                    if (durum != true) //Daha önce alınmış bir veri girildiyse programı kapat
                    {
                        logla("-**-Bu Tarihteki Data Daha Önce Alınmış. Program Kapatmak için bir tuşa basın.");
                        Console.ReadKey();
                        Environment.Exit(0);
                    }
                }

                int baslangicGun = Convert.ToInt32(baslangicData[0]);
                int baslangicAy = Convert.ToInt32(baslangicData[1]);
                int baslangicYil = Convert.ToInt32(baslangicData[2]);


                //Bitiş Tarihi
                Console.WriteLine("Lütfen Verinin Alınmaya Durdurulacağı Tarihi G/A/YYYY Şeklinde Girin");
                Console.WriteLine("Boş Bırakmanız durumunda tablo kontrol edilecektir. Tablo boşsa belirttiğiniz başlangıç tarihinden dün'e kadar olan bütün veriler istenecektir.");
                Console.WriteLine("Örneğin : 5/11/2018");
                string[] bitisData;

                string bitisKontrol = Console.ReadLine();

                if (bitisKontrol == "") //Bitiş Tarihi Girilmemişse bugünkü veriye kadar al
                {
                    bitisData = bugun.AddDays(-1).ToString(("dd'/'MM'/'yyyy")).Split('/'); //TODO Bu gün istenmeyebilir.

                    logla("-**-Bitirilecek tarih : " + bugun.ToString());
                }
                else //Bitiş Tarihi girilmişse girilen tarih dahil tüm verileri al
                {
                    bitisData = bitisKontrol.Split('/');
                    bitisData[0] = Convert.ToString(Convert.ToInt32(bitisData[0]));
                    logla("-**-Bitirlecek tarih : " + bitisKontrol);
                }

                int bitisGun = Convert.ToInt32(bitisData[0]);
                int bitisAy = Convert.ToInt32(bitisData[1]);
                int bitisYil = Convert.ToInt32(bitisData[2]);


                DateTime start = new DateTime(baslangicYil, baslangicAy, baslangicGun);
                DateTime end = new DateTime(bitisYil, bitisAy, bitisGun);
                int days = (end - start).Days;

                if (start > bugun)
                {
                    logla("-**-Girilen Başlangıç Tarihi, Bu Günün Tarihinden Büyüktü. Program Kapatılıyor...");
                    Console.ReadKey();
                    Environment.Exit(0);
                }
                if (end > bugun)
                {
                    logla("-**-Girilen Bitiş Tarihi, Bu Günün Tarihinden Büyüktü. Program Kapatılıyor...");
                    Console.ReadKey();
                    Environment.Exit(0);
                }

                int satir;

                int lastUsedRow = GetLastUsedRow(ws);
                if (lastUsedRow > 3)
                    satir = lastUsedRow + 1;
                else
                    satir = 3;

                Enumerable
                    .Range(0, days + 1)
                    .Select(x => start.AddDays(x))
                    .ToList()
                    .ForEach(d =>
                    {
                        //Console.WriteLine(d.ToString("dd\\/MM\\/yyyy"));

                        string tarih = d.ToString("dd\\/MM\\/yyyy");
                        string json = getData(tarih);
                        if (json != "error" && json != "null")
                        {
                            int macSayaci = 0;

                            Maclar datalist = JsonConvert.DeserializeObject<Maclar>(json);
                            for (int i = 0; i < datalist.m.Count; i++)
                            {
                                bool failState = false;

                                if (Convert.ToString(datalist.m[i][14]) != "0" && Convert.ToString(datalist.m[i][6]) != "ERT")
                                {
                                    string details = null;
                                    Detaylar detaylar = null;
                                    try
                                    {
                                        details = getDetails(Convert.ToString(datalist.m[i][0]), Convert.ToString(datalist.m[i][14]));
                                        detaylar = JsonConvert.DeserializeObject<Detaylar>(details);
                                    }
                                    catch
                                    {
                                        logla("-**-Detaylar Servisine Erişilemedi. 10 Saniye Sonra Tekrar Denenecek");
                                        System.Threading.Thread.Sleep(10000);
                                        try
                                        {
                                            details = getDetails(Convert.ToString(datalist.m[i][0]), Convert.ToString(datalist.m[i][14]));
                                            detaylar = JsonConvert.DeserializeObject<Detaylar>(details);
                                        }
                                        catch
                                        {
                                            logla("-**-Detaylar Patladı.");
                                            //break;
                                            failState = true;
                                        }

                                    }

                                    try
                                    {
                                        if (details != "" && details != null && failState != true) //else Basketbol Datası
                                        {
                                            for (int j = 0; j < detaylar.ARR.Count; j++)
                                            {
                                                //Console.WriteLine("{0} : {1} - {2}", datalist.m[i][35], detaylar.ARR[0].T1, detaylar.ARR[0].T2);
                                                logla(datalist.m[i][35] + " : " + detaylar.ARR[0].T1 + " - " + detaylar.ARR[0].T2);

                                                string[] ulkeler = Convert.ToString(datalist.m[i][36]).Split(',');
                                                string ulke = ulkeler[9].Trim();
                                                ulke = ulke.Replace("\"", "");

                                                

                                                ws.Cells["A" + satir].Value = ulke;//[36][9] Ülke
                                                ws.Cells["B" + satir].Value = datalist.m[i][35]; //tarih VE SAAT

                                                try
                                                {
                                                    HANDİKAP handikap = JsonConvert.DeserializeObject<HANDİKAP>(Convert.ToString(datalist.m[i][15]));

                                                    if (handikap.h1 > 0) //Eğer 1.Takımın handikap değeri 0'dan büyükse
                                                    {
                                                        ws.Cells["C" + satir].Value = "(h:" + handikap.h1 + ") " + detaylar.ARR[j].T1; //Takım 1 Handikap Değeri + Takım Adı
                                                        ws.Cells["C" + satir].Style.Font.Color.SetColor(Color.Red);
                                                    }
                                                    else
                                                        ws.Cells["C" + satir].Value = detaylar.ARR[j].T1; //Takım 1

                                                    if (handikap.h2 > 0)//Eğer 2.Takımın handikap değeri 0'dan büyükse
                                                    {
                                                        ws.Cells["D" + satir].Value = detaylar.ARR[j].T2 + " (h:" + handikap.h2 + ")"; //Takım 2 Handikap Değeri + Takım Adı
                                                        ws.Cells["D" + satir].Style.Font.Color.SetColor(Color.Red);
                                                    }
                                                    else
                                                        ws.Cells["D" + satir].Value = detaylar.ARR[j].T2; //Takım 2
                                                }
                                                catch
                                                {
                                                    logla("-**-Handikaplar Patladı");
                                                }


                                                ws.Cells["E" + satir].Value = datalist.m[i][12] + "-" + datalist.m[i][13]; //Maç Sonucu
                                                ws.Cells["F" + satir].Value = datalist.m[i][7]; //İlk Yarı
                                                ws.Cells["G" + satir].Value = detaylar.ARR[j].MS1; //Maç Sonucu 1
                                                ws.Cells["H" + satir].Value = detaylar.ARR[j].MS0; //Maç Sonucu x
                                                ws.Cells["I" + satir].Value = detaylar.ARR[j].MS2; //Maç Sonucu 2
                                                ws.Cells["J" + satir].Value = detaylar.ARR[j].IY1; //İlk Yarı Sonucu 1
                                                ws.Cells["K" + satir].Value = detaylar.ARR[j].IY0; //İlk Yarı Sonucu x
                                                ws.Cells["L" + satir].Value = detaylar.ARR[j].IY2; //İlk Yarı Sonucu 2
                                                ws.Cells["M" + satir].Value = detaylar.ARR[j].IYA15; //İlk Yarı 1.5 Alt
                                                ws.Cells["N" + satir].Value = detaylar.ARR[j].IYU15; //İlk Yarı 1.5 Üst
                                                ws.Cells["O" + satir].Value = detaylar.ARR[j].A15; //Maç Sonucu Alt Üst 1.5 Alt
                                                ws.Cells["P" + satir].Value = detaylar.ARR[j].U15; //Maç Sonucu Alt Üst 1.5 Üst
                                                ws.Cells["Q" + satir].Value = detaylar.ARR[j].A; //Maç Sonucu Alt Üst 2.5 Alt
                                                ws.Cells["R" + satir].Value = detaylar.ARR[j].U; //Maç Sonucu Alt Üst 2.5 Üst
                                                ws.Cells["S" + satir].Value = detaylar.ARR[j].A35; //Maç Sonucu Alt Üst 3.5 Alt
                                                ws.Cells["T" + satir].Value = detaylar.ARR[j].U35; //Maç Sonucu Alt Üst 3.5 Üst
                                                ws.Cells["U" + satir].Value = detaylar.ARR[j].KGVAR; //Karşılıklı Gol Var
                                                ws.Cells["V" + satir].Value = detaylar.ARR[j].KGYOK; //Karşılıklı Gol Yok
                                                ws.Cells["W" + satir].Value = detaylar.ARR[j].CS10; //Çifte Şans 1-0
                                                ws.Cells["X" + satir].Value = detaylar.ARR[j].CS12; //Çİfte ŞAns 1-2
                                                ws.Cells["Y" + satir].Value = detaylar.ARR[j].CS02; //Çifte Şans 0-2

                                                //Handikap
                                                ws.Cells["Z" + satir].Value = detaylar.ARR[j].HMS1; //Handikaplı Maç Sonucu 1
                                                ws.Cells["AA" + satir].Value = detaylar.ARR[j].HMS0; //Handikaplı Maç Sonucu x //Handikaplı Takım Gözükecek mi?
                                                ws.Cells["AB" + satir].Value = detaylar.ARR[j].HMS2; // Handikaplı Maç Sonucu 2

                                                if (detaylar.ARR[j].H1 == 1)
                                                    ws.Cells["Z" + satir].Style.Font.Color.SetColor(Color.Red);

                                                if (detaylar.ARR[j].H2 == 1)
                                                    ws.Cells["AB" + satir].Style.Font.Color.SetColor(Color.Red);

                                                //Toplam Gol
                                                ws.Cells["AC" + satir].Value = detaylar.ARR[j].TG01;
                                                ws.Cells["AD" + satir].Value = detaylar.ARR[j].TG23;
                                                ws.Cells["AE" + satir].Value = detaylar.ARR[j].TG46;
                                                ws.Cells["AF" + satir].Value = detaylar.ARR[j].TG7;


                                                //İlk Yarı Maç Sonucu
                                                ws.Cells["AG" + satir].Value = detaylar.ARR[j].IYMS11;
                                                ws.Cells["AH" + satir].Value = detaylar.ARR[j].IYMS10;
                                                ws.Cells["AI" + satir].Value = detaylar.ARR[j].IYMS12;
                                                ws.Cells["AJ" + satir].Value = detaylar.ARR[j].IYMS01;
                                                ws.Cells["AK" + satir].Value = detaylar.ARR[j].IYMS00;
                                                ws.Cells["AL" + satir].Value = detaylar.ARR[j].IYMS02;
                                                ws.Cells["AM" + satir].Value = detaylar.ARR[j].IYMS21;
                                                ws.Cells["AN" + satir].Value = detaylar.ARR[j].IYMS20;
                                                ws.Cells["AO" + satir].Value = detaylar.ARR[j].IYMS22;


                                                //Maç Skoru
                                                ws.Cells["AP" + satir].Value = detaylar.ARR[j].SK10;
                                                ws.Cells["AQ" + satir].Value = detaylar.ARR[j].SK20;
                                                ws.Cells["AR" + satir].Value = detaylar.ARR[j].SK21;
                                                ws.Cells["AS" + satir].Value = detaylar.ARR[j].SK30;
                                                ws.Cells["AT" + satir].Value = detaylar.ARR[j].SK31;
                                                ws.Cells["AU" + satir].Value = detaylar.ARR[j].SK32;
                                                ws.Cells["AV" + satir].Value = detaylar.ARR[j].SK40;
                                                ws.Cells["AW" + satir].Value = detaylar.ARR[j].SK41;
                                                ws.Cells["AX" + satir].Value = detaylar.ARR[j].SK42;
                                                ws.Cells["AY" + satir].Value = detaylar.ARR[j].SK43;
                                                ws.Cells["AZ" + satir].Value = detaylar.ARR[j].SK50;
                                                ws.Cells["BA" + satir].Value = detaylar.ARR[j].SK51;
                                                ws.Cells["BB" + satir].Value = detaylar.ARR[j].SK52;
                                                ws.Cells["BC" + satir].Value = detaylar.ARR[j].SK53;
                                                ws.Cells["BD" + satir].Value = detaylar.ARR[j].SK54;
                                                ws.Cells["BE" + satir].Value = detaylar.ARR[j].SK00;
                                                ws.Cells["BF" + satir].Value = detaylar.ARR[j].SK11;
                                                ws.Cells["BG" + satir].Value = detaylar.ARR[j].SK22;
                                                ws.Cells["BH" + satir].Value = detaylar.ARR[j].SK33;
                                                ws.Cells["BI" + satir].Value = detaylar.ARR[j].SK44;
                                                ws.Cells["BJ" + satir].Value = detaylar.ARR[j].SK55;
                                                ws.Cells["BK" + satir].Value = detaylar.ARR[j].SK01;
                                                ws.Cells["BL" + satir].Value = detaylar.ARR[j].SK02;
                                                ws.Cells["BM" + satir].Value = detaylar.ARR[j].SK12;
                                                ws.Cells["BN" + satir].Value = detaylar.ARR[j].SK03;
                                                ws.Cells["BO" + satir].Value = detaylar.ARR[j].SK13;
                                                ws.Cells["BP" + satir].Value = detaylar.ARR[j].SK23;
                                                ws.Cells["BQ" + satir].Value = detaylar.ARR[j].SK04;
                                                ws.Cells["BR" + satir].Value = detaylar.ARR[j].SK14;
                                                ws.Cells["BS" + satir].Value = detaylar.ARR[j].SK24;
                                                ws.Cells["BT" + satir].Value = detaylar.ARR[j].SK34;
                                                ws.Cells["BU" + satir].Value = detaylar.ARR[j].SK05;
                                                ws.Cells["BV" + satir].Value = detaylar.ARR[j].SK15;
                                                ws.Cells["BW" + satir].Value = detaylar.ARR[j].SK25;
                                                ws.Cells["BX" + satir].Value = detaylar.ARR[j].SK35;
                                                ws.Cells["BY" + satir].Value = detaylar.ARR[j].SK45;

                                                //sayac++;
                                                satir++;
                                                macSayaci++;

                                            }
                                        }
                                    }
                                    catch
                                    {

                                        //Console.WriteLine(e.Message);
                                        logla("-**-Excell ya da Parser Hatası");

                                    }
                                }
                            }
                            logla("-**-Bu Gün Kaydedilen Toplam Maç : " + macSayaci.ToString());
                        }
                        if (d.ToString("dd\\/MM\\/yyyy").Split('/')[0] == "28")
                        {
                            try
                            {
                                p.Save();
                                logla("Ay Sonu Kaydı!");
                            }
                            catch
                            {
                                logla("Ay Sonu Kaydı Başarısız");
                            }
                        }
                    });
                logla("Veriler Kaydediliyor. Lütfen Bekleyiniz...");
                try
                {
                    p.Save();
                }
                catch (Exception e)
                {
                    logla(e.Message);
                }
                logla("Kayıt Tamamlandı!");
                int kacmac = GetLastUsedRow(ws) - 2;
                logla("Toplam Kayıtlı Maç Sayısı : " + Convert.ToString(kacmac));

                DateTime bitis = DateTime.Now; //Bu günün tarihini al
                logla("-**-Islem bitisi : " + bitis.ToString());
                logla("=====Transaction Is Finish=====");


            }

            Console.ReadKey();

            string getData(string date)
            {
                var client = new RestClient("http://goapi.mackolik.com/livedata?date=" + date);
                var request = new RestRequest(Method.GET);
                request.AddHeader("accept", "*/*");
                request.AddHeader("origin", "http://arsiv.mackolik.com");
                request.AddHeader("referer", "http://arsiv.mackolik.com/Canli-Sonuclar");
                request.AddHeader("user-agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/75.0.3770.142 Safari/537.36");
                IRestResponse response = client.Execute(request);

                if (response.Content == "")
                {
                    return "error";
                }
                else
                    return response.Content;
            }

            string getDetails(string macid, string iddaaid)
            {

                var client = new RestClient("http://arsiv.mackolik.com/AjaxHandlers/IddaaHandler.aspx?command=morebets&mac=" + macid + "&iddaaId=" + iddaaid);
                var request = new RestRequest(Method.GET);
                request.AddHeader("accept", "text/plain, */*; q=0.01");
                request.AddHeader("accept-encoding", "gzip, deflate");
                request.AddHeader("accept-language", "tr-TR,tr;q=0.9,en-US;q=0.8,en;q=0.7");
                request.AddHeader("cache-control", "no-cache");
                request.AddHeader("connection", "keep-alive");
                request.AddHeader("cookie", "__auc=e4f0755616c527e06e560742536; _ga=GA1.2.213032902.1564752087; _gid=GA1.2.559745995.1564752087; __gfp_64b=RRC2vDL0usNm5BspB69n79N39F.4dDanXJsjK0lsg8z.X7; _hjid=d42a1d6b-5e2e-4e60-a823-3d9f4e0ca775; cookieconsent_dismissed=yes; gig_hasGmid=login; SOUND=false; duello=false; _gat_UA-241588-3=1; am_cookie_test=true; __asc=4cbed33f16c5d24cdc41caa52c2; _gat=1");
                request.AddHeader("host", "arsiv.mackolik.com");
                request.AddHeader("pragma", "no-cache");
                request.AddHeader("referer", "http://arsiv.mackolik.com/Canli-Sonuclar");
                request.AddHeader("user-agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/75.0.3770.142 Safari/537.36");
                request.AddHeader("x-requested-with", "XMLHttpRequest");
                IRestResponse response = client.Execute(request);

                return response.Content;

            }

            int GetLastUsedRow(ExcelWorksheet sheet)
            {
                var row = sheet.Dimension.End.Row;
                while (row >= 1)
                {
                    var range = sheet.Cells[row, 1, row, sheet.Dimension.End.Column];
                    if (range.Any(c => !string.IsNullOrEmpty(c.Text)))
                    {
                        break;
                    }
                    row--;
                }
                return row;
            }

            void logla(string log)
            {
                try
                {
                    //using (StreamWriter writer = new StreamWriter("output.txt"))
                    using (StreamWriter writer = File.AppendText("log.txt"))
                    {
                        //writer.WriteLine("=====Transaction Is Starting=====");
                        writer.WriteLine(log);
                        Console.WriteLine(log);
                    }
                }
                catch
                {
                    Console.WriteLine("Logla Fonksiyonu Patladı");
                }
            }
        }


    }

    public class ARR
    {
        public string Type { get; set; }
        public int MID { get; set; }
        public int ID { get; set; }
        public int H1 { get; set; }
        public int H2 { get; set; }
        public string T1 { get; set; }
        public string T2 { get; set; }
        public int T1I { get; set; }
        public int T2I { get; set; }
        public int MB { get; set; }
        public int MD { get; set; }
        public string MS1 { get; set; }
        public string MS0 { get; set; }
        public string MS2 { get; set; }
        public string CS10 { get; set; }
        public string CS12 { get; set; }
        public string CS02 { get; set; }
        public string IY1 { get; set; }
        public string IY0 { get; set; }
        public string IY2 { get; set; }
        public string A { get; set; }
        public string U { get; set; }
        public string IYMS11 { get; set; }
        public string IYMS10 { get; set; }
        public string IYMS12 { get; set; }
        public string IYMS01 { get; set; }
        public string IYMS00 { get; set; }
        public string IYMS02 { get; set; }
        public string IYMS21 { get; set; }
        public string IYMS20 { get; set; }
        public string IYMS22 { get; set; }
        public string TG01 { get; set; }
        public string TG23 { get; set; }
        public string TG46 { get; set; }
        public string TG7 { get; set; }
        public string HMS1 { get; set; }
        public string HMS0 { get; set; }
        public string HMS2 { get; set; }
        public string KGVAR { get; set; }
        public string KGYOK { get; set; }
        public string SK00 { get; set; }
        public string SK01 { get; set; }
        public string SK02 { get; set; }
        public string SK03 { get; set; }
        public string SK04 { get; set; }
        public string SK05 { get; set; }
        public string SK10 { get; set; }
        public string SK11 { get; set; }
        public string SK12 { get; set; }
        public string SK13 { get; set; }
        public string SK14 { get; set; }
        public string SK15 { get; set; }
        public string SK20 { get; set; }
        public string SK21 { get; set; }
        public string SK22 { get; set; }
        public string SK23 { get; set; }
        public string SK24 { get; set; }
        public string SK25 { get; set; }
        public string SK30 { get; set; }
        public string SK31 { get; set; }
        public string SK32 { get; set; }
        public string SK33 { get; set; }
        public string SK34 { get; set; }
        public string SK35 { get; set; }
        public string SK40 { get; set; }
        public string SK41 { get; set; }
        public string SK42 { get; set; }
        public string SK43 { get; set; }
        public string SK44 { get; set; }
        public string SK45 { get; set; }
        public string SK50 { get; set; }
        public string SK51 { get; set; }
        public string SK52 { get; set; }
        public string SK53 { get; set; }
        public string SK54 { get; set; }
        public string SK55 { get; set; }
        public int FT1 { get; set; }
        public int FT2 { get; set; }
        public int HT1 { get; set; }
        public int HT2 { get; set; }
        public int MOH { get; set; }
        public string ISD { get; set; }
        public string A15 { get; set; }
        public string U15 { get; set; }
        public string A35 { get; set; }
        public string U35 { get; set; }
        public string IYA15 { get; set; }
        public string IYU15 { get; set; }
    }

    public class Detaylar
    {
        public int ID { get; set; }
        public List<ARR> ARR { get; set; }
    }

    public class Maclar
    {
        public List<List<object>> e { get; set; }
        public int eId { get; set; }
        public List<List<object>> m { get; set; }
        public string t { get; set; }
    }

    public class HANDİKAP
    {
        public int e { get; set; }
        public string goal { get; set; }
        public int h2 { get; set; }
        public int h1 { get; set; }
        public int k2 { get; set; }
        public int k1 { get; set; }
        public int aeleme { get; set; }
        public int tId { get; set; }
        public int ogd { get; set; }
    }
}
