using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using RestSharp;
using Newtonsoft.Json;
using OfficeOpenXml;
using System.IO;
using Newtonsoft.Json.Linq;

namespace iddializer
{
    class Program
    {
        static void Main(string[] args)
        {
            //string mac = "3201596";
            //string iddaa = "585";
            //string date = "01/08/2019";

            string[] gun = { "01", "02", "03", "04", "05", "06", "07","08","09","10","11","12","13","14","15","16","17","18","19","20","21","22","23","24","25","26","27","28","29","30","31"};
            string[] ay = { "01", "02", "03", "04", "05", "06", "07", "08", "09", "10", "11", "12"};
            int sayac = 3;
            var data = new FileInfo(@"data.xlsx");
            using (var p = new ExcelPackage(data))
            {
                var ws = p.Workbook.Worksheets["BULTEN"];
                for (int y = 2019; y <= 2019 + 1; y++) //2000
                {
                    for (int a = 7; a < 8; a++) // < ay.length
                    {
                        for (int g = 0; g < 5; g++) // < gun.length-1
                        {
                            string tarih = gun[g] + "/" + ay[a] + "/" + y;
                            string json = getData(tarih);
                            if (json != "error" && json != "null")
                            {
                                Maclar datalist = JsonConvert.DeserializeObject<Maclar>(json);
                                for (int i = 0; i < datalist.m.Count; i++)
                                {
                                    if (Convert.ToString(datalist.m[i][14]) != "0" && Convert.ToString(datalist.m[i][6]) != "ERT")
                                    {
                                        string details = getDetails(Convert.ToString(datalist.m[i][0]), Convert.ToString(datalist.m[i][14]));
                                        Detaylar detaylar = JsonConvert.DeserializeObject<Detaylar>(details);

                                        try
                                        {
                                            for (int j = 0; j < detaylar.ARR.Count; j++)
                                            {
                                                Console.WriteLine("{0} : {1} - {2}", datalist.m[i][35], detaylar.ARR[0].T1, detaylar.ARR[0].T2);

                                                string[] ulkeler = Convert.ToString(datalist.m[i][36]).Split(',');
                                                string ulke = ulkeler[9].Trim();
                                                ulke = ulke.Replace("\"", "");

                                                ws.Cells["A" + sayac].Value = ulke;//[36][9] Ülke
                                                ws.Cells["B" + sayac].Value = datalist.m[i][35]; //tarih VE SAAT
                                                ws.Cells["C" + sayac].Value = detaylar.ARR[j].T1; //Takım 1
                                                ws.Cells["D" + sayac].Value = detaylar.ARR[j].T2; //Takım 2
                                                ws.Cells["E" + sayac].Value = datalist.m[i][12] + "-" + datalist.m[i][13]; //Maç Sonucu
                                                ws.Cells["F" + sayac].Value = datalist.m[i][7]; //İlk Yarı
                                                ws.Cells["G" + sayac].Value = detaylar.ARR[j].MS1; //Maç Sonucu 1
                                                ws.Cells["H" + sayac].Value = detaylar.ARR[j].MS0; //Maç Sonucu x
                                                ws.Cells["I" + sayac].Value = detaylar.ARR[j].MS2; //Maç Sonucu 2
                                                ws.Cells["J" + sayac].Value = detaylar.ARR[j].IY1; //İlk Yarı Sonucu 1
                                                ws.Cells["K" + sayac].Value = detaylar.ARR[j].IY0; //İlk Yarı Sonucu x
                                                ws.Cells["L" + sayac].Value = detaylar.ARR[j].IY2; //İlk Yarı Sonucu 2
                                                ws.Cells["M" + sayac].Value = detaylar.ARR[j].IYA15; //İlk Yarı 1.5 Alt
                                                ws.Cells["N" + sayac].Value = detaylar.ARR[j].IYU15; //İlk Yarı 1.5 Üst
                                                ws.Cells["O" + sayac].Value = detaylar.ARR[j].A15; //Maç Sonucu Alt Üst 1.5 Alt
                                                ws.Cells["P" + sayac].Value = detaylar.ARR[j].U15; //Maç Sonucu Alt Üst 1.5 Üst
                                                ws.Cells["Q" + sayac].Value = detaylar.ARR[j].A; //Maç Sonucu Alt Üst 2.5 Alt
                                                ws.Cells["R" + sayac].Value = detaylar.ARR[j].U; //Maç Sonucu Alt Üst 2.5 Üst
                                                ws.Cells["S" + sayac].Value = detaylar.ARR[j].A35; //Maç Sonucu Alt Üst 3.5 Alt
                                                ws.Cells["T" + sayac].Value = detaylar.ARR[j].U35; //Maç Sonucu Alt Üst 3.5 Üst
                                                ws.Cells["U" + sayac].Value = detaylar.ARR[j].KGVAR; //Karşılıklı Gol Var
                                                ws.Cells["V" + sayac].Value = detaylar.ARR[j].KGYOK; //Karşılıklı Gol Yok
                                                ws.Cells["W" + sayac].Value = detaylar.ARR[j].CS10; //Çifte Şans 1-0
                                                ws.Cells["X" + sayac].Value = detaylar.ARR[j].CS12; //Çİfte ŞAns 1-2
                                                ws.Cells["Y" + sayac].Value = detaylar.ARR[j].CS02; //Çifte Şans 0-2


                                                ws.Cells["Z" + sayac].Value = detaylar.ARR[j].HMS1; //Handikaplı Maç Sonucu 1
                                                ws.Cells["AA" + sayac].Value = detaylar.ARR[j].HMS0; //Handikaplı Maç Sonucu x //Handikaplı Takım Gözükecek mi?
                                                ws.Cells["AB" + sayac].Value = detaylar.ARR[j].HMS2; // Handikaplı Maç Sonucu 2


                                                //İlk Yarı Maç Sonucu
                                                ws.Cells["AC" + sayac].Value = detaylar.ARR[j].IYMS11;
                                                ws.Cells["AD" + sayac].Value = detaylar.ARR[j].IYMS10;
                                                ws.Cells["AE" + sayac].Value = detaylar.ARR[j].IYMS12;
                                                ws.Cells["AF" + sayac].Value = detaylar.ARR[j].IYMS01;
                                                ws.Cells["AG" + sayac].Value = detaylar.ARR[j].IYMS00;
                                                ws.Cells["AH" + sayac].Value = detaylar.ARR[j].IYMS02;
                                                ws.Cells["AI" + sayac].Value = detaylar.ARR[j].IYMS21;
                                                ws.Cells["AJ" + sayac].Value = detaylar.ARR[j].IYMS20;
                                                ws.Cells["AK" + sayac].Value = detaylar.ARR[j].IYMS22;


                                                //p.Save();
                                                sayac++;
                                            }
                                        }
                                        catch(Exception e)
                                        {
                                            /*
                                            Console.WriteLine(e.Message);
                                            Console.WriteLine("Excell'e Aktarılırken Bir Hata Oldu");
                                            */
                                        }
                                    }
                                }
                            }

                        }
                    }
                }
                p.Save();
            }

            



            /*
            string json = getData("01/08/2019");
            Maclar datalist = JsonConvert.DeserializeObject<Maclar>(json);

            for (int i = 0; i < datalist.m.Count; i++)
            {
                Console.WriteLine(datalist.m[i]);
                if (Convert.ToString(datalist.m[i][14]) != "0")
                {
                    try
                    {
                        string details = getDetails(Convert.ToString(datalist.m[i][0]), Convert.ToString(datalist.m[i][14]));
                        Detaylar detaylar = JsonConvert.DeserializeObject<Detaylar>(details);
                        for (int j = 0; j < detaylar.ARR.Count; j++)
                        {
                            //Console.WriteLine("{0} - {1}", detaylar.ARR[0].T1, detaylar.ARR[0].T2);
                        }
                    }
                    catch
                    {
                        //basketbol ya da hata
                    }
                }
            }*/
            Console.WriteLine("End of the data");
            

            string getData(string date)
            {
                var client = new RestClient("http://goapi.mackolik.com/livedata?date=" + date);
                var request = new RestRequest(Method.GET);
                request.AddHeader("accept", "*/*");
                request.AddHeader("origin", "http://arsiv.mackolik.com");
                request.AddHeader("referer", "http://arsiv.mackolik.com/Canli-Sonuclar");
                request.AddHeader("user-agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/75.0.3770.142 Safari/537.36");
                IRestResponse response = client.Execute(request);

                if(response.Content == "")
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

            Console.ReadKey();
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
}
