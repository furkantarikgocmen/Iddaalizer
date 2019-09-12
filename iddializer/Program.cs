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
                    logla("-!!-\'Data Dosyası Bulunamadı. Program Kapanıyor.\'");
                    Console.ReadKey();
                    Environment.Exit(0);
                }

                Console.WriteLine("Lütfen Verinin Alınmaya Başlayacağı Tarihi G/A/YYYY Şeklinde Girin");
                Console.WriteLine("Boş Bırakmanız durumunda tablo kontrol edilecektir. Tablo boşsa 29/08/2019 tarihinden itibaren bütün veriler istenecektir");
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
                            logla("-!!-Bu Gün Zaten Senkoronize Edilmiş. Program Kapatılıyor...");
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
                        baslangicData = "29/08/2019".Split('/'); //İlk veriyi çekmeye başla
                        logla("-**-Baslanacak tarih : 29/08/2019");
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
                        logla("-!!-Bu Tarihteki Data Daha Önce Alınmış. Program Kapatmak için bir tuşa basın.");
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

                    logla("-**-Bitirilecek tarih : " + bugun.AddDays(-1).ToString("dd'/'MM'/'yyyy")); //bitisdata alınabilir bugun.ToString()
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
                    logla("-!!-Girilen Başlangıç Tarihi, Bu Günün Tarihinden Büyüktü. Program Kapatılıyor...");
                    Console.ReadKey();
                    Environment.Exit(0);
                }
                if (end > bugun)
                {
                    logla("-!!-Girilen Bitiş Tarihi, Bu Günün Tarihinden Büyüktü. Program Kapatılıyor...");
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

                            Maclar datalist = JsonConvert.DeserializeObject<Maclar>(json); //gun datası convert to obj
                            for (int i = 0; i < datalist.m.Count; i++)
                            {

                                bool failState = false;

                                if (Convert.ToString(datalist.m[i][14]) != "0" && Convert.ToString(datalist.m[i][6]) != "ERT" && Convert.ToString(datalist.m[i][23]) != "2") //iddia ve ertelenmemiş
                                {

                                    string details = null;
                                    Detaylar detaylar = null;
                                    try
                                    {
                                        details = getDetails(Convert.ToString(datalist.m[i][0]));
                                        detaylar = JsonConvert.DeserializeObject<Detaylar>(details);
                                    }
                                    catch
                                    {
                                        logla("-!!-Detaylar Servisine Erişilemedi. 10 Saniye Sonra Tekrar Denenecek");
                                        System.Threading.Thread.Sleep(10000);
                                        try
                                        {
                                            details = getDetails(Convert.ToString(datalist.m[i][0]));
                                            detaylar = JsonConvert.DeserializeObject<Detaylar>(details);
                                        }
                                        catch
                                        {
                                            logla("-!!-Detaylar Alınamadı.");
                                            //break;
                                            failState = true;
                                        }

                                    }

                                    try
                                    {
                                        if (details != "" && details != null && failState != true) //else Basketbol Datası or error
                                        {
                                            //Console.WriteLine("{0} : {1} - {2} === {3}", datalist.m[i][35], datalist.m[i][2], datalist.m[i][4], datalist.m[i][0]);
                                            logla(datalist.m[i][35] + " : " + datalist.m[i][2] + " - " + datalist.m[i][4]);

                                            for (int j = 0; j < detaylar.Event.Markets.Count; j++)
                                            {

                                                try
                                                {
                                                    for (int w = 0; w < detaylar.Event.Markets[j].Outcomes.Count; w++)
                                                    {
                                                        string column = getColumn(detaylar.Event.Markets[j].Name);

                                                        if ( column == "unkown")
                                                            break;
                                                        string[] kolonlar = column.Split(',');

                                                        string kolon = kolonlar[w];

                                                        string outcome = Convert.ToString(detaylar.Event.Markets[j].Outcomes[w].Odd);

                                                        ws.Cells[kolon + satir].Value = outcome;
                                                    }
                                                }
                                                catch
                                                {
                                                    logla("---!!GetColumn fonksiyonundan gelen dizi uzunluğu, Outcomes uzunluğudnan düşük.Dizi Belirlenen Aralığın Dışındaydı");
                                                }
                                            }
                                            string[] ulkeler = Convert.ToString(datalist.m[i][36]).Split(',');
                                            string ulke = ulkeler[9].Trim();
                                            ulke = ulke.Replace("\"", "");

                                            ws.Cells["A" + satir].Value = ulke;//[36][9] Ülke
                                            ws.Cells["B" + satir].Value = datalist.m[i][35]; //tarih VE SAAT
                                            ws.Cells["C" + satir].Value = datalist.m[i][2];
                                            ws.Cells["D" + satir].Value = datalist.m[i][4];

                                            ws.Cells["E" + satir].Value = datalist.m[i][12] + "-" + datalist.m[i][13]; //Maç Sonucu
                                            ws.Cells["F" + satir].Value = datalist.m[i][7]; //İlk Yarı

                                            satir++;
                                            macSayaci++;

                                        }
                                    }
                                    catch (Exception e)
                                    {
                                        logla("-!!-Excell ya da Parser Hatası");
                                        logla(e.Message);
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
                                logla("--**-Ay Sonu Kaydı!");
                            }
                            catch
                            {
                                logla("--!!--Ay Sonu Kaydı Başarısız");
                            }
                        }
                    });
                logla("--**--Veriler Kaydediliyor. Lütfen Bekleyiniz...");
                try
                {
                    //p.Save();
                }
                catch (Exception e)
                {
                    logla(e.Message);
                }
                logla("--**--Kayıt Tamamlandı");
                int kacmac = GetLastUsedRow(ws) - 2;
                logla("--**--Toplam Kayıtlı Maç Sayısı : " + Convert.ToString(kacmac));

                DateTime bitis = DateTime.Now; //Bu günün tarihini al
                logla("-**-Islem bitisi : " + bitis.ToString());
                logla("=====Transaction Is Finish=====");

            }

            Console.ReadKey(); //System.enviroment.exit(0) olabilir?

            string getColumn(string name)
            {
                string column = "";
                switch (name)
                {
                    case "Maç Sonucu":
                        column = "G,H,I";
                        break;
                    case "Çifte Şans"://3
                        column = "AP,AQ,AR";
                        break;
                    case "Handikaplı Maç Sonucu (0:2)"://3 //yok
                        column = "unkown";
                        break;
                    case "Handikaplı Maç Sonucu (0:1)"://3 //yok
                        column = "unkown";
                        break;
                    case "Handikaplı Maç Sonucu (1:0)"://3 //yok
                        column = "unkown";
                        break;
                    case "Handikaplı Maç Sonucu (2:0)"://3 //yok
                        column = "unkown";
                        break;
                    case "İlk Yarı/Maç Sonucu"://9
                        column = "AY,AZ,BA,BB,BC,BD,BE,BF,BG";
                        break;
                    case "Maç Sonucu ve (1,5) Alt/Üst"://6 //yok
                        column = "unkown";
                        break;
                    case "Maç Sonucu ve (2,5) Alt/Üst"://6 //bunlardan birini incele //yok
                        column = "unkown";
                        break;
                    case "Maç Sonucu ve (3,5) Alt/Üst"://6 //yok
                        column = "unkown";
                        break;
                    case "Maç Sonucu ve (4,5) Alt/Üst"://6 //yok
                        column = "unkown";
                        break;
                    case "İlk Gol"://3 //yok //bak buna
                        column = "unkown";
                        break;
                    case "0,5 Alt/Üst"://2
                        column = "Z,AA";
                        break;
                    case "1,5 Alt/Üst"://2
                        column = "AB,AC";
                        break;
                    case "2,5 Alt/Üst"://2
                        column = "AD,AE";
                        break;
                    case "3,5 Alt/Üst"://2
                        column = "AF,AG";
                        break;
                    case "4,5 Alt/Üst"://2
                        column = "AH,AI";
                        break;
                    case "5,5 Alt/Üst"://2 
                        column = "AJ,AK";
                        break;
                    case "6,5 Alt/Üst"://2
                        column = "AL,AM";
                        break;
                    case "Karşılıklı Gol"://2
                        column = "AN,AO";
                        break;
                    case "Toplam Gol Aralığı"://4
                        column = "P,Q,R,S";
                        break;
                    case "1. Yarı Sonucu"://3
                        column = "J,K,L";
                        break;
                    case "1. Yarı Çifte Şans"://3
                        column = "AS,AT,AU";
                        break;
                    case "2. Yarı Sonucu"://3
                        column = "M,N,O";
                        break;
                    case "1. Yarı 0,5 Alt/Üst"://2
                        column = "T,U";
                        break;
                    case "1. Yarı 1,5 Alt/Üst"://2
                        column = "V,W";
                        break;
                    case "1. Yarı 2,5 Alt/Üst"://2
                        column = "X,Y";
                        break;
                    case "Evsahibi 0,5 Alt/Üst"://2
                        column = "unkown";
                        break;
                    case "Evsahibi 1,5 Alt/Üst"://2
                        column = "unkown";
                        break;
                    case "Evsahibi 2,5 Alt/Üst"://2
                        column = "unkown";
                        break;
                    case "Deplasman 0,5 Alt/Üst"://2
                        column = "unkown";
                        break;
                    case "Deplasman 1,5 Alt/Üst"://2
                        column = "unkown";
                        break;
                    case "Deplasman 2,5 Alt/Üst"://2
                        column = "unkown";
                        break;
                    case "Tek/Çift"://2 //bunu da vermiyorum
                        column = "unkown";
                        break;
                    case "Maç Skoru"://29 //bunu vermiyorum
                        column = "unkown";
                        break;
                    case "Evsahibi Gol Yemeden Kazanır"://2
                        column = "unkown";
                        break;
                    case "Deplasman Gol Yemeden Kazanır"://2
                        column = "unkown";
                        break;
                    case "Gol Atacak Takımlar"://4
                        column = "unkown";
                        break;
                    case "Daha Çok Gol Olacak Yarı"://3
                        column = "unkown";
                        break;
                    case "Evsahibi Gol Yemez"://2
                        column = "unkown";
                        break;
                    case "Deplasman Gol Yemez"://2
                        column = "unkown";
                        break;
                    case "(8,5) Korner Alt/Üst"://2
                        column = "unkown";
                        break;
                    case "(9,5) Korner Alt/Üst"://2
                        column = "unkown";
                        break;
                    case "(10,5) Korner Alt/Üst"://2
                        column = "unkown";
                        break;
                    case "(11,5) Korner Alt/Üst"://2
                        column = "unkown";
                        break;
                    case "(12,5) Korner Alt/Üst"://2
                        column = "unkown";
                        break;
                    case "1.Yarı (4,5) Korner Alt/Üst"://2
                        column = "unkown";
                        break;
                    case "1.Yarı (5,5) Korner Alt/Üst"://2
                        column = "unkown";
                        break;
                    case "1.Yarı (6,5) Korner Alt/Üst"://2
                        column = "unkown";
                        break;
                    case "En Çok Korner"://3
                        column = "unkown";
                        break;
                    case "1. Yarı En Çok Korner"://3 //bundan iki tane var
                        column = "unkown";
                        break;
                    case "İlk Korner"://3
                        column = "unkown";
                        break;
                    case "(1,5) Kart Alt/Üst"://2
                        column = "unkown";
                        break;
                    case "(2,5) Kart Alt/Üst"://2
                        column = "unkown";
                        break;
                    case "(3,5) Kart Alt/Üst"://2
                        column = "unkown";
                        break;
                    case "(4,5) Kart Alt/Üst"://2
                        column = "unkown";
                        break;
                    case "(5,5) Kart Alt/Üst"://2
                        column = "unkown";
                        break;
                    case "(6,5) Kart Alt/Üst"://2
                        column = "unkown";
                        break;
                    case "(7,5) Kart Alt/Üst"://2
                        column = "unkown";
                        break;
                    case "Kırmızı Kart"://2
                        column = "unkown";
                        break;
                    case "Uzatma Oynanır":// bilmiyorum bak buna
                        column = "unkown";
                        break;
                    case "1. Yarı Toplam Gol Sayısı"://3
                        column = "BH,BI,BJ";
                        break;
                    default:
                        column = "unkown";
                        break;
                }
                return column;
            }

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

            string getDetails(string macid)
            {

                var client = new RestClient("http://arsiv.mackolik.com/AjaxHandlers/IddaaHandler.aspx?command=morebets&mac=" + macid);
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
                    Console.WriteLine("--!!--Logla Fonksiyonu Hatası");
                }
            }
        }

    }

    public class Maclar
    {
        public List<List<object>> e { get; set; }
        public int eId { get; set; }
        public List<List<object>> m { get; set; }
        public string t { get; set; }
    }

    public class MarketType
    {
        public int Id { get; set; }
        public string Name { get; set; }
        public string Title { get; set; }
    }

    public class Outcome
    {
        public int EventId { get; set; }
        public int MarketId { get; set; }
        public int MarketTypeId { get; set; }
        public int OutcomeNo { get; set; }
        public string OutcomeName { get; set; }
        public string Odd { get; set; } //burası doble'idi string yaptım
    }

    public class Market
    {
        public int MarketId { get; set; }
        public int MarketNo { get; set; }
        public int EventId { get; set; }
        public MarketType MarketType { get; set; }
        public int MBS { get; set; }
        public double SOV { get; set; }
        public int MarketStatus { get; set; }
        public List<Outcome> Outcomes { get; set; }
        public string Title { get; set; }
        public string Name { get; set; }
    }

    public class Event
    {
        public int EventId { get; set; }
        public int SportId { get; set; }
        public DateTime StartDate { get; set; }
        public int LeagueCode { get; set; }
        public bool HasLive { get; set; }
        public bool IsLive { get; set; }
        public List<Market> Markets { get; set; }
    }

    public class Detaylar
    {
        public string Match { get; set; }
        public Event Event { get; set; }
    }

}
