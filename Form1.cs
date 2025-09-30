using Microsoft.Extensions.Configuration;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using Primary;
using Primary.Data;
using System.Configuration;
using System.Globalization;
using System.Media;
using System.Net;
using System.Text;

namespace BOTArbitradorPorPlazo
{
    public partial class frmBOT : Form
    {
        const string sURL = "https://api.invertironline.com";
        const string SURLOper = "https://www.invertironline.com";
        const string sURLVETA = "https://api.veta.xoms.com.ar";
        const string prefijoPrimary = "MERV - XMEV - ";
        const string sufijoCI = " - CI";
        const string sufijo24 = " - 24hs";
        const int HorarioApertura = 1035;
        const int HorarioCierre = 1655;

        string tokenVETA;
        string bearer;
        string refresh;
        double umbral;
        double timeOffset;
        double umbralBonos;
        double umbralAcciones;
        DateTime expires;

        List<string> tickersIOL;
        List<Ticker> tickers;
        List<string> tickersCI;
        List<string> tickers24;

        public frmBOT()
        {
            InitializeComponent();
        }

        private void frmBOT_Load(object sender, EventArgs e)
        {
            this.Top = 10;
            this.Text = "BOT Arbitrador - Sion Capital";

            DoubleBuffered = true;
            CheckForIllegalCrossThreadCalls = false;
            umbralAcciones = 0.50;  // Establecer los umbrales de acuerdo a la comisión de cada uno.
            umbralBonos = 0.50;

            tickersIOL = new List<string>();
            tickers = new List<Ticker>();

            FillTickersIOL();
            ConfigureGrid();

            for (double j = 0; j <= 4; j = j + 0.1)
                cboUmbral.Items.Add(Math.Round(j, 2));

            var configuracion = new ConfigurationBuilder()
                .SetBasePath(AppContext.BaseDirectory)
                .AddJsonFile("appsettings.json", optional: true, reloadOnChange: true)
                .Build();
            try
            {
                cboUmbral.SelectedIndex = int.Parse(configuracion.GetSection("MiConfiguracion:Umbral").Value);
                txtPresupuesto.Text = configuracion.GetSection("MiConfiguracion:Presupuesto").Value;
                txtUsuarioIOL.Text = configuracion.GetSection("MiConfiguracion:UserIOL").Value;
                txtClaveIOL.Text = configuracion.GetSection("MiConfiguracion:ClaveIOL").Value;
                txtUsuarioVETA.Text = configuracion.GetSection("MiConfiguracion:UserVETA").Value;
                txtClaveVETA.Text = configuracion.GetSection("MiConfiguracion:ClaveVETA").Value;
            }
            catch { }

            if (cboUmbral.Items.Count > 0 && cboUmbral.SelectedIndex < 0)
                cboUmbral.SelectedIndex = 0;
        }

        private void FillTickersIOL()
        {
            // ETFs
            tickersIOL.Add("-ETFs-");
            tickersIOL.Add("ARKK");
            tickersIOL.Add("DIA");
            tickersIOL.Add("EEM");
            tickersIOL.Add("EWZ");
            tickersIOL.Add("IWM");
            tickersIOL.Add("QQQ");
            tickersIOL.Add("SPY");
            tickersIOL.Add("XLE");
            tickersIOL.Add("XLF");

            // Bonos
            tickersIOL.Add("-BONDs-");
            tickersIOL.Add("AL29");
            tickersIOL.Add("AL35");
            tickersIOL.Add("AE38");
            tickersIOL.Add("AL41");
            tickersIOL.Add("DICP");
            tickersIOL.Add("PARP");
            tickersIOL.Add("GD29");
            tickersIOL.Add("GD35");
            tickersIOL.Add("GD38");
            tickersIOL.Add("GD41");
            tickersIOL.Add("GD46");
            tickersIOL.Add("PBA25");
            tickersIOL.Add("TX26");
            tickersIOL.Add("TX28");
            tickersIOL.Add("TO26");
            tickersIOL.Add("TDG24");

            // Acciones
            tickersIOL.Add("-ACCs-");
            tickersIOL.Add("ALUA");
            tickersIOL.Add("BBAR");
            tickersIOL.Add("BMA");
            tickersIOL.Add("BYMA");
            tickersIOL.Add("GGAL");
            tickersIOL.Add("PAMP");
            tickersIOL.Add("SUPV");
            tickersIOL.Add("TXAR");
            tickersIOL.Add("YPFD");

            // CEDEARs nuevos
            tickersIOL.Add("-NCEDs-");
            tickersIOL.Add("AAL");
            tickersIOL.Add("AKO.B");
            tickersIOL.Add("COIN");
            tickersIOL.Add("DOW");
            tickersIOL.Add("EA");
            tickersIOL.Add("GM");
            tickersIOL.Add("LRCX");
            tickersIOL.Add("NIO");
            tickersIOL.Add("SE");
            tickersIOL.Add("SPGI");
            tickersIOL.Add("TWLO");
            tickersIOL.Add("XP");
            tickersIOL.Add("ABBV");
            tickersIOL.Add("AVGO");
            tickersIOL.Add("BIOX");
            tickersIOL.Add("BRKB");
            tickersIOL.Add("CAAP");
            tickersIOL.Add("DOCU");
            tickersIOL.Add("ETSY");
            tickersIOL.Add("MA");
            tickersIOL.Add("PAAS");
            tickersIOL.Add("PSX");
            tickersIOL.Add("SHOP");
            tickersIOL.Add("SNOW");
            tickersIOL.Add("SPOT");
            tickersIOL.Add("SQ");
            tickersIOL.Add("UNH");
            tickersIOL.Add("UNP");
            tickersIOL.Add("WBA");
            tickersIOL.Add("ZM");

            tickersIOL.Add("-NCEDs2-");
            tickersIOL.Add("ABNB");
            tickersIOL.Add("BITF");
            tickersIOL.Add("F");
            tickersIOL.Add("HUT");
            tickersIOL.Add("JMIA");
            tickersIOL.Add("MOS");
            tickersIOL.Add("MSTR");
            tickersIOL.Add("MU");
            tickersIOL.Add("OXY");
            tickersIOL.Add("PANW");
            tickersIOL.Add("RBLX");
            tickersIOL.Add("SATL");
            tickersIOL.Add("UAL");
            tickersIOL.Add("UBER");
            tickersIOL.Add("UPST");

            tickersIOL.Add("-NCEDs3-");
            tickersIOL.Add("CCL");
            tickersIOL.Add("BKNG");
            tickersIOL.Add("CVS");
            tickersIOL.Add("DAL");
            tickersIOL.Add("MDLZ");
            tickersIOL.Add("MRNA");
            tickersIOL.Add("PINS");
            tickersIOL.Add("PM");
            tickersIOL.Add("RACE");
            tickersIOL.Add("RIOT");
            tickersIOL.Add("ROKU");
            tickersIOL.Add("SCHW");
            tickersIOL.Add("SPCE");
            tickersIOL.Add("STLA");
            tickersIOL.Add("SWKS");
            tickersIOL.Add("TMUS");

            // CEDEARs viejos
            tickersIOL.Add("-OCEDs-");
            tickersIOL.Add("AAPL");
            tickersIOL.Add("ABEV");
            tickersIOL.Add("ABT");
            tickersIOL.Add("ADBE");
            tickersIOL.Add("ADGO");
            tickersIOL.Add("AIG");
            tickersIOL.Add("AMD");
            tickersIOL.Add("AMGN");
            tickersIOL.Add("AMX");
            tickersIOL.Add("AMZN");
            tickersIOL.Add("ARCO");
            tickersIOL.Add("AXP");
            tickersIOL.Add("AZN");
            tickersIOL.Add("BA");
            tickersIOL.Add("BA.C");
            tickersIOL.Add("BABA");
            tickersIOL.Add("BB");
            tickersIOL.Add("BBD");
            tickersIOL.Add("BBV");
            tickersIOL.Add("BCS");
            tickersIOL.Add("BHP");
            tickersIOL.Add("BIDU");
            tickersIOL.Add("BIIB");
            tickersIOL.Add("BNG");
            tickersIOL.Add("BP");
            tickersIOL.Add("BRFS");
            tickersIOL.Add("BSBR");
            tickersIOL.Add("C");
            tickersIOL.Add("CAH");
            tickersIOL.Add("CAT");
            tickersIOL.Add("CL");
            tickersIOL.Add("COST");
            tickersIOL.Add("CRM");
            tickersIOL.Add("CSCO");
            tickersIOL.Add("CVX");
            tickersIOL.Add("CX");
            tickersIOL.Add("DE");
            tickersIOL.Add("DESP");
            tickersIOL.Add("DISN");
            tickersIOL.Add("EBAY");
            tickersIOL.Add("ERIC");
            tickersIOL.Add("ERJ");
            tickersIOL.Add("FCX");
            tickersIOL.Add("FDX");
            tickersIOL.Add("FMX");
            tickersIOL.Add("FSLR");
            tickersIOL.Add("GE");
            tickersIOL.Add("GFI");
            tickersIOL.Add("GGB");
            tickersIOL.Add("GILD");
            tickersIOL.Add("GLOB");
            tickersIOL.Add("GOLD");
            tickersIOL.Add("GOOGL");
            tickersIOL.Add("GS");
            tickersIOL.Add("HD");
            tickersIOL.Add("HMY");
            tickersIOL.Add("HPQ");
            tickersIOL.Add("HSBC");
            tickersIOL.Add("IBM");
            tickersIOL.Add("INTC");
            tickersIOL.Add("ITUB");
            tickersIOL.Add("JD");
            tickersIOL.Add("JNJ");
            tickersIOL.Add("JPM");
            tickersIOL.Add("KO");
            tickersIOL.Add("MCD");
            tickersIOL.Add("MELI");
            tickersIOL.Add("META");
            tickersIOL.Add("MMM");
            tickersIOL.Add("MO");
            tickersIOL.Add("MRK");
            tickersIOL.Add("MSFT");
            tickersIOL.Add("NFLX");
            tickersIOL.Add("NKE");
            tickersIOL.Add("NOKA");
            tickersIOL.Add("NVDA");
            tickersIOL.Add("ORAN");
            tickersIOL.Add("ORCL");
            tickersIOL.Add("PBR");
            tickersIOL.Add("PEP");
            tickersIOL.Add("PFE");
            tickersIOL.Add("PG");
            tickersIOL.Add("PYPL");
            tickersIOL.Add("QCOM");
            tickersIOL.Add("RIO");
            tickersIOL.Add("RTX");
            tickersIOL.Add("SAP");
            tickersIOL.Add("SBUX");
            tickersIOL.Add("SID");
            tickersIOL.Add("SLB");
            tickersIOL.Add("SNAP");
            tickersIOL.Add("SONY");
            tickersIOL.Add("T");
            tickersIOL.Add("TEN");
            tickersIOL.Add("TRIP");
            tickersIOL.Add("TSM");
            tickersIOL.Add("TSLA");
            tickersIOL.Add("TWTR");
            tickersIOL.Add("TXN");
            tickersIOL.Add("TXR");
            tickersIOL.Add("UGP");
            tickersIOL.Add("UL");
            tickersIOL.Add("V");
            tickersIOL.Add("VALE");
            tickersIOL.Add("VIV");
            tickersIOL.Add("VOD");
            tickersIOL.Add("VRSN");
            tickersIOL.Add("VZ");
            tickersIOL.Add("WFC");
            tickersIOL.Add("WMT");
            tickersIOL.Add("X");
            tickersIOL.Add("XOM");

            // Galpones
            tickersIOL.Add("-GALPs-");
            tickersIOL.Add("CEPU");
            tickersIOL.Add("COME");
            tickersIOL.Add("CRES");
            tickersIOL.Add("CTIO");
            tickersIOL.Add("EDN");
            tickersIOL.Add("GAMI");
            tickersIOL.Add("LOMA");
            tickersIOL.Add("MIRG");
            tickersIOL.Add("TECO2");
            tickersIOL.Add("TGNO4");
            tickersIOL.Add("TGSU2");
            tickersIOL.Add("TRAN");
            tickersIOL.Add("VALO");
        }

        private void btnLogin_Click(object sender, EventArgs e)
        {
            Inicio();
        }

        private async Task Inicio()
        {
            var api = new Api(new Uri(sURLVETA));
            await api.Login(txtUsuarioVETA.Text, txtClaveVETA.Text);
            ToLog("Login VETA Ok");

            var allInstruments = await api.GetAllInstruments();

            var entries = new[] { Entry.Bids, Entry.Offers };

            FillListaTickers();

            var instrumentos = allInstruments.Where(c => tickersCI.Contains(c.Symbol))
                .Concat(allInstruments.Where(c => tickers24.Contains(c.Symbol)));

            using var socket = api.CreateMarketDataSocket(instrumentos, entries, 1, 1);
            socket.OnData = OnMarketData;
            var socketTask = await socket.Start();
            socketTask.Wait(1000);

            LoginIOL();
            tmr.Start();
            ToLog("Login IOL Ok");
            await socketTask;
        }

        // ====== HTTP HELPERS (UTF-8 + Content-Type correcto) ======
        private string PostForm(string url, string formBody) // /token
        {
            var req = (HttpWebRequest)WebRequest.Create(url);
            byte[] data = Encoding.UTF8.GetBytes(formBody);
            req.Method = "POST";
            req.ContentType = "application/x-www-form-urlencoded";
            req.ContentLength = data.Length;
            if (!string.IsNullOrEmpty(bearer)) req.Headers["Authorization"] = bearer;

            using (var s = req.GetRequestStream()) s.Write(data, 0, data.Length);

            try
            {
                using var resp = (HttpWebResponse)req.GetResponse();
                using var sr = new StreamReader(resp.GetResponseStream(), Encoding.UTF8);
                return sr.ReadToEnd();
            }
            catch (WebException ex)
            {
                var er = ex.Response as HttpWebResponse;
                using var sr = new StreamReader(er?.GetResponseStream() ?? Stream.Null, Encoding.UTF8);
                return $"HTTP {(int)(er?.StatusCode ?? 0)} {er?.StatusCode}: {sr.ReadToEnd()}";
            }
        }

        private string PostJson(string url, string jsonBody) // /api/v2/operar/*
        {
            var req = (HttpWebRequest)WebRequest.Create(url);
            byte[] data = Encoding.UTF8.GetBytes(jsonBody);
            req.Method = "POST";
            req.ContentType = "application/json";
            req.ContentLength = data.Length;
            if (!string.IsNullOrEmpty(bearer)) req.Headers["Authorization"] = bearer;

            using (var s = req.GetRequestStream()) s.Write(data, 0, data.Length);

            try
            {
                using var resp = (HttpWebResponse)req.GetResponse();
                using var sr = new StreamReader(resp.GetResponseStream(), Encoding.UTF8);
                return sr.ReadToEnd();
            }
            catch (WebException ex)
            {
                var er = ex.Response as HttpWebResponse;
                using var sr = new StreamReader(er?.GetResponseStream() ?? Stream.Null, Encoding.UTF8);
                return $"HTTP {(int)(er?.StatusCode ?? 0)} {er?.StatusCode}: {sr.ReadToEnd()}";
            }
        }

        private string GetJson(string url) // GET con bearer
        {
            var req = (HttpWebRequest)WebRequest.Create(url);
            req.Method = "GET";
            req.ContentType = "application/json";
            if (!string.IsNullOrEmpty(bearer)) req.Headers["Authorization"] = bearer;

            try
            {
                using var resp = (HttpWebResponse)req.GetResponse();
                using var sr = new StreamReader(resp.GetResponseStream(), Encoding.UTF8);
                return sr.ReadToEnd();
            }
            catch (WebException ex)
            {
                var er = ex.Response as HttpWebResponse;
                using var sr = new StreamReader(er?.GetResponseStream() ?? Stream.Null, Encoding.UTF8);
                return $"HTTP {(int)(er?.StatusCode ?? 0)} {er?.StatusCode}: {sr.ReadToEnd()}";
            }
        }
        // ==========================================================

        private async void LoginIOL()
        {
            try
            {
                // Login inicial
                if (expires == DateTime.MinValue)
                {
                    string postData = $"username={txtUsuarioIOL.Text}&password={txtClaveIOL.Text}&grant_type=password";
                    string response = PostForm(sURL + "/token", postData);
                    dynamic json = JObject.Parse(response);
                    bearer = "Bearer " + json.access_token;
                    expires = DateTime.Now.AddSeconds((double)json.expires_in - 300);
                    refresh = json.refresh_token;
                    ToLog("Token IOL OK");
                }
                // Refresh si está por vencer o vencido
                else if (DateTime.Now >= expires)
                {
                    string postData = $"refresh_token={refresh}&grant_type=refresh_token";
                    string response = PostForm(sURL + "/token", postData);
                    if (response.StartsWith("HTTP 401"))
                    {
                        ToLog("Refresh 401: " + response);
                    }
                    else
                    {
                        dynamic json = JObject.Parse(response);
                        bearer = "Bearer " + json.access_token;
                        expires = DateTime.Now.AddSeconds((double)json.expires_in - 100);
                        refresh = json.refresh_token;
                        ToLog("Token IOL Refrescado");
                    }
                }
            }
            catch (Exception e)
            {
                ToLog("LoginIOL error: " + e.Message);
            }
        }

        private async Task<string> Comprar(string simbolo, string cantidad, string precio)
        {
            if (!int.TryParse(cantidad, out var qty) || qty <= 0)
            {
                ToLog("Error de cantidad: " + cantidad);
                return "Error";
            }
            if (!double.TryParse(precio, NumberStyles.Any, CultureInfo.InvariantCulture, out var px))
            {
                ToLog("Error de precio: " + precio);
                return "Error";
            }

            ToLog($"Comprando {qty} {simbolo} a {px.ToString(CultureInfo.InvariantCulture)}");
            await Task.Run(() => Application.DoEvents());

            string validez = DateTime.Today.ToString("yyyy-MM-dd") + "T17:59:59.000Z";

            var order = new
            {
                mercado = "bCBA",
                simbolo = simbolo,
                cantidad = qty,
                precio = px,
                plazo = "t0",
                validez = validez
            };

            string postDataJson = JsonConvert.SerializeObject(order);
            string response = PostJson(sURL + "/api/v2/operar/Comprar", postDataJson);

            if (response.StartsWith("HTTP 401"))
            {
                ToLog("401 en Comprar → renovando token IOL y reintentando...");
                LoginIOL();
                response = PostJson(sURL + "/api/v2/operar/Comprar", postDataJson);
            }

            try
            {
                dynamic json = JObject.Parse(response);
                string okStr = json.ok != null ? json.ok.ToString() : "true";
                if (okStr.Equals("false", StringComparison.OrdinalIgnoreCase))
                {
                    ToLog("Comprar NOK: " + response);
                    return "Error";
                }
                string operacion = json.numeroOperacion;
                return operacion ?? "Error";
            }
            catch
            {
                ToLog("Comprar parse error: " + response);
                return "Error";
            }
        }

        private string Vender(string simbolo, string cantidad, string precio)
        {
            if (!int.TryParse(cantidad, out var qty) || qty <= 0)
            {
                ToLog("Error de cantidad: " + cantidad);
                return "Error";
            }
            if (!double.TryParse(precio, NumberStyles.Any, CultureInfo.InvariantCulture, out var px))
            {
                ToLog("Error de precio: " + precio);
                return "Error";
            }

            ToLog($"Vendiendo {qty} {simbolo} a {px.ToString(CultureInfo.InvariantCulture)}");
            Application.DoEvents();

            string validez = DateTime.Today.ToString("yyyy-MM-dd") + "T17:59:59.000Z";

            var order = new
            {
                mercado = "bCBA",
                simbolo = simbolo,
                cantidad = qty,
                precio = px,
                plazo = "t1",
                validez = validez
            };

            string postDataJson = JsonConvert.SerializeObject(order);
            string response = PostJson(sURL + "/api/v2/operar/Vender", postDataJson);

            if (response.StartsWith("HTTP 401"))
            {
                ToLog("401 en Vender → renovando token IOL y reintentando...");
                LoginIOL();
                response = PostJson(sURL + "/api/v2/operar/Vender", postDataJson);
            }

            try
            {
                dynamic json = JObject.Parse(response);
                string okStr = json.ok != null ? json.ok.ToString() : "true";
                if (okStr.Equals("false", StringComparison.OrdinalIgnoreCase))
                {
                    ToLog("Vender NOK: " + response);
                    return "Error";
                }
                string operacion = json.numeroOperacion;
                return operacion ?? "Error";
            }
            catch
            {
                ToLog("Vender parse error: " + response);
                return "Error";
            }
        }

        private string GetEstadoOperacion(string idoperacion)
        {
            string url = sURL + "/api/v2/operaciones/" + idoperacion;
            string response = GetJson(url);

            if (response.StartsWith("HTTP 401"))
            {
                ToLog("401 en GetEstadoOperacion → renovando token IOL y reintentando...");
                LoginIOL();
                response = GetJson(url);
            }

            try
            {
                dynamic json = JObject.Parse(response);
                string estado =
                    (string)(json.estadoActual != null ? json.estadoActual :
                             json.estado != null ? json.estado : "");

                if (string.IsNullOrWhiteSpace(estado))
                    return "Error";

                return estado.Trim().ToLowerInvariant(); // "terminada", "pendiente", etc.
            }
            catch
            {
                if (response.IndexOf("HTTP 404", StringComparison.OrdinalIgnoreCase) >= 0 ||
                    response.IndexOf("no existe", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    return "noencontrada";
                }
                ToLog("Estado parse error: " + response);
                return "Error";
            }
        }

        private void FillListaTickers()
        {
            foreach (string tickerIOL in tickersIOL)
            {
                tickers.Add(new Ticker(tickerIOL, prefijoPrimary + tickerIOL + sufijoCI, prefijoPrimary + tickerIOL + sufijo24));
                tickersCI = tickers.Select(t => t.PrimaryCI).ToList();
                tickers24 = tickers.Select(t => t.Primary24).ToList();
            }
        }

        private async void OnMarketData(Api api, MarketData marketData)
        {
            var ticker = marketData.InstrumentId.Symbol;
            var bid = default(decimal);
            var offer = default(decimal);
            var bidSize = default(decimal);
            var offerSize = default(decimal);

            if (marketData.Data.Bids != null)
            {
                foreach (var trade in marketData.Data.Bids)
                {
                    bid = trade.Price;
                    bidSize = trade.Size;
                }
            }

            if (marketData.Data.Offers != null)
            {
                foreach (var trade in marketData.Data.Offers)
                {
                    offer = trade.Price;
                    offerSize = trade.Size;
                }
            }

            if (ticker.EndsWith("24hs"))
            {
                for (int j = 0; j < grdPanel.Rows.Count; j++)
                {
                    string left = grdPanel.Rows[j].Cells[0].Value.ToString();
                    string right = tickers.Where(t => t.Primary24 == ticker).Select(t => t.IOL).First().ToString();
                    if (left == right)
                    {
                        if (bidSize == 0)
                        {
                            grdPanel.Rows[j].Cells[4].Value = "";
                            grdPanel.Rows[j].Cells[5].Value = "";
                        }
                        else
                        {
                            grdPanel.Rows[j].Cells[4].Value = bid;
                            grdPanel.Rows[j].Cells[5].Value = bidSize;
                        }
                        refreshRatio(j);
                    }
                }
            }
            if (ticker.EndsWith("CI"))
            {
                for (int j = 0; j < grdPanel.Rows.Count; j++)
                {
                    string left = grdPanel.Rows[j].Cells[0].Value.ToString();
                    string right = tickers.Where(t => t.PrimaryCI == ticker).Select(t => t.IOL).First().ToString();

                    if (left == right)
                    {
                        if (offerSize == 0)
                        {
                            grdPanel.Rows[j].Cells[2].Value = "";
                            grdPanel.Rows[j].Cells[3].Value = "";
                        }
                        else
                        {
                            grdPanel.Rows[j].Cells[2].Value = offerSize;
                            grdPanel.Rows[j].Cells[3].Value = offer;
                        }
                        refreshRatio(j);
                    }
                }
            }
            Application.DoEvents();
        }

        private void ConfigureGrid()
        {
            grdPanel.Columns.Clear();
            grdPanel.Rows.Clear();
            grdPanel.Columns.Add("Ticker", "Ticker");
            grdPanel.Columns[0].Width = 70;
            grdPanel.Columns[0].CellTemplate.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;

            grdPanel.Columns.Add("Momento", "Momento");
            grdPanel.Columns[1].Width = 70;
            grdPanel.Columns[1].CellTemplate.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;

            grdPanel.Columns.Add("QVCI", "QVCI");
            grdPanel.Columns[2].Width = 70;
            grdPanel.Columns[2].CellTemplate.Style.Alignment = DataGridViewContentAlignment.MiddleRight;

            grdPanel.Columns.Add("PVCI", "PVCI");
            grdPanel.Columns[3].Width = 70;
            grdPanel.Columns[3].CellTemplate.Style.Alignment = DataGridViewContentAlignment.MiddleRight;

            grdPanel.Columns.Add("PC24", "PC24");
            grdPanel.Columns[4].Width = 70;
            grdPanel.Columns[4].CellTemplate.Style.Alignment = DataGridViewContentAlignment.MiddleRight;

            grdPanel.Columns.Add("QC24", "QC24");
            grdPanel.Columns[5].Width = 70;
            grdPanel.Columns[5].CellTemplate.Style.Alignment = DataGridViewContentAlignment.MiddleRight;

            grdPanel.Columns.Add("Ratio", "Ratio");
            grdPanel.Columns[6].Width = 70;
            grdPanel.Columns[6].CellTemplate.Style.Alignment = DataGridViewContentAlignment.MiddleRight;

            grdPanel.RowHeadersWidth = 4;

            foreach (var ticker in tickersIOL)
                grdPanel.Rows.Add(ticker);
        }

        private async void refreshRatio(int i)
        {
            bool esBono = false;
            string PIV = "";
            string P24C = "";
            string QIV = "";
            string Q24C = "";
            int Q;

            string simbolo = grdPanel.Rows[i].Cells[0].Value?.ToString() ?? "";
            grdPanel.ClearSelection();
            grdPanel.Rows[i].Cells[1].Value = DateTime.Now.ToLongTimeString();

            if (chkFollow.Checked)
                grdPanel.CurrentCell = grdPanel.Rows[i].Cells[1];

            grdPanel.Rows[i].Selected = true;
            Application.DoEvents();

            if (grdPanel.Rows[i].Cells[3].Value != null)
            {
                PIV = grdPanel.Rows[i].Cells[3].Value.ToString();
                QIV = grdPanel.Rows[i].Cells[2].Value?.ToString() ?? "";
            }
            if (grdPanel.Rows[i].Cells[4].Value != null)
            {
                P24C = grdPanel.Rows[i].Cells[4].Value.ToString();
                Q24C = grdPanel.Rows[i].Cells[5].Value?.ToString() ?? "";
            }

            if (string.IsNullOrWhiteSpace(PIV) || string.IsNullOrWhiteSpace(P24C))
            {
                grdPanel.Rows[i].Cells[6].Value = "";
                return;
            }

            // Marcar si es bono
            if (simbolo == "AL30" || simbolo == "AL29" || simbolo == "AL35" || simbolo == "AE38" ||
                simbolo == "AL41" || simbolo == "TC23" || simbolo == "TC24" || simbolo == "CO26" ||
                simbolo == "CUAP" || simbolo == "DICP" || simbolo == "GD29" || simbolo == "GD30" ||
                simbolo == "GD35" || simbolo == "GD38" || simbolo == "GD41" || simbolo == "GD46" ||
                simbolo == "PARP" || simbolo == "PR13" || simbolo == "PR15" || simbolo == "TO23" ||
                simbolo == "TO26" || simbolo == "T2X2" || simbolo == "T2X3" || simbolo == "T2X4" ||
                simbolo == "TX22" || simbolo == "TX23" || simbolo == "TX24" || simbolo == "TX26" ||
                simbolo == "TX28" || simbolo == "TDJ23" || simbolo == "TDL23" || simbolo == "TDS23" ||
                simbolo == "TDF24")
            {
                esBono = true;
            }

            double piv = Convert.ToDouble(PIV);
            double p24 = Convert.ToDouble(P24C);

            double porcentual = Math.Round(100 - ((piv / p24) * 100), 4);
            grdPanel.Rows[i].Cells[6].Value = Math.Round(porcentual, 2);

            grdPanel.Rows[i].Cells[6].Style.ForeColor = porcentual > 0 ? Color.DarkGreen : Color.Red;

            double umbralUI = (cboUmbral.SelectedItem is double d) ? d : 0d;
            double limiteBonos = umbralBonos + umbralUI;
            double limiteAcciones = umbralAcciones + umbralUI;

            bool superaUmbral = (esBono && (porcentual > limiteBonos)) || (!esBono && (porcentual > limiteAcciones));

            if (!superaUmbral)
                return;

            if (chkBeep.Checked)
                SystemSounds.Beep.Play();

            grdPanel.Rows[i].Cells[6].Style.ForeColor = Color.DarkSlateBlue;
            ToLog("Arbitraje en " + simbolo + " con ratio " + porcentual.ToString());

            if (!chkAuto.Checked)
                return;

            // Calcular cantidad a operar
            if (!int.TryParse(Q24C, out var q24)) q24 = 0;
            if (!int.TryParse(QIV, out var qiv)) qiv = 0;

            Q = (q24 > 0 && qiv > 0) ? Math.Min(q24, qiv) : 0;

            double presupuesto = double.TryParse(txtPresupuesto.Text, out var p) ? p : 0;

            string cant;
            if (Q * piv < presupuesto)
            {
                cant = Q.ToString();
            }
            else
            {
                if (esBono)
                    cant = Math.Floor(presupuesto * 100 / piv).ToString();
                else
                    cant = Math.Floor(presupuesto / piv).ToString();
            }

            // Ventana de mercado
            int hhmm = int.Parse(DateTime.Now.ToString("HHmm"));
            if (hhmm >= HorarioApertura && hhmm <= HorarioCierre)
            {
                Operar(simbolo, cant, PIV, P24C);
            }

            Application.DoEvents();
        }

        private async void ToLog(string s)
        {
            lbLog.Items.Add(DateTime.Now.ToLongTimeString() + ": " + s);
            lbLog.SelectedIndex = lbLog.Items.Count - 1;
        }

        private async void tmr_Tick(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(txtUsuarioIOL.Text) || string.IsNullOrWhiteSpace(txtClaveIOL.Text))
                return;

            // Sólo refrescar si está por vencer o vencido
            if (expires == DateTime.MinValue || DateTime.Now >= expires)
                LoginIOL();
        }

        private async void Operar(string simbolo, string q, string PC, string PV)
        {
            int cifrasRedondeo = 0;
            LoginIOL();
            ToLog("Iniciando " + simbolo);

            double preventivoCompra = double.Parse(PC);
            if (preventivoCompra < 100) cifrasRedondeo = 1;
            preventivoCompra = Math.Round(preventivoCompra + ((preventivoCompra / 100) * 0.1), cifrasRedondeo);
            PC = preventivoCompra.ToString().Replace(",", ".");

            double preventivoVenta = double.Parse(PV);
            preventivoVenta = Math.Round(preventivoVenta - ((preventivoVenta / 100) * 0.1), cifrasRedondeo);
            PV = preventivoVenta.ToString().Replace(",", ".");

            string operacionCompra = await Comprar(simbolo, q, PC);
            if (operacionCompra != "Error")
            {
                string estadooperacion = "";
                int intentos = 24;

                for (int i = 1; i <= intentos; i++)
                {
                    ToLog("Intento de compra " + i.ToString() + " de " + simbolo);
                    estadooperacion = GetEstadoOperacion(operacionCompra);
                    if (estadooperacion == "terminada")
                        break;
                    Application.DoEvents();
                }
                if (estadooperacion == "terminada")
                {
                    ToLog("Compra OK " + simbolo);
                    Vender(simbolo, q, PV);
                }
                else
                {
                    ToLog("Venció la compra de " + simbolo);
                    WebRequest request = WebRequest.Create(sURL + "/api/v2/operaciones/" + operacionCompra);
                    request.Method = "DELETE";
                    request.ContentType = "application/json";
                    request.Headers.Add("Authorization", bearer);

                    try
                    {
                        WebResponse response = request.GetResponse();
                    }
                    catch (Exception e)
                    {
                        ToLog(e.Message);
                    }
                }
            }
            else
            {
                ToLog("Error en compra de " + simbolo);
            }
        }

        private void grbLogin_Enter(object sender, EventArgs e) { }
    }

    record Ticker(string IOL, string PrimaryCI, string Primary24);
}
