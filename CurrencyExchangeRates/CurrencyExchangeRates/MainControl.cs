// PROJECT : CurrencyExchangeRates
// This project was developed by Tanguy Touzard
// CODEPLEX: http://xrmtoolbox.codeplex.com
// BLOG: http://mscrmtools.blogspot.com

using System;
using System.ComponentModel;
using System.Drawing;
using System.Windows.Forms;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using XrmToolBox;
using Microsoft.Xrm.Sdk.Query;
using System.Text;
using System.Collections;
using System.Data;

using MsCrmTools.CurrencyExchangeRates.currencyconvertor;
using System.IO;

namespace CurrencyExchangeRates
{
    public partial class MainControl : UserControl, IMsCrmToolsPluginUserControl
    {
        
        

        #region Variables

        /// <summary>
        /// Microsoft Dynamics CRM 2011 Organization Service
        /// </summary>
        private IOrganizationService service;

        /// <summary>
        /// Panel used to display progress information
        /// </summary>
        private Panel infoPanel;

        #endregion Variables

        #region Constructor

        /// <summary>
        /// Initializes a new instance of class <see cref="MainControl"/>
        /// </summary>
        public MainControl()
        {
            InitializeComponent();
            
            
            
        }

        #endregion Constructor

        #region Properties

        /// <summary>
        /// Gets the organization service used by the tool
        /// </summary>
        public IOrganizationService Service
        {
            get { return service; }
        }

        /// <summary>
        /// Gets the logo to display in the tools list
        /// </summary>
        public Image PluginLogo
        {
            get { return toolImageList.Images[0]; }
        }

        #endregion

        #region EventHandlers

        /// <summary>
        /// EventHandler to request a connection to an organization
        /// </summary>
        public event EventHandler OnRequestConnection;

        /// <summary>
        /// EventHandler to close the current tool
        /// </summary>
        public event EventHandler OnCloseTool;

        #endregion EventHandlers

        #region Methods

        /// <summary>
        /// Updates the organization service used by the tool
        /// </summary>
        /// <param name="newService">Organization service</param>
        /// <param name="actionName">Action that requested a service update</param>
        /// <param name="parameter">Parameter passed when requesting a service update</param>
        public void UpdateConnection(IOrganizationService newService, string actionName = "", object parameter = null)
        {
            service = newService;

            if (actionName == "GetCurrentCurrencies")
            {
                GetCurrentCurrencies();
            }
        }


        private void UpdateAllRates()
        {
            infoPanel = InformationPanel.GetInformationPanel(this, "Updating CRM Currencies Rates...", 340, 100);

            var worker = new BackgroundWorker();
            worker.DoWork += WorkerDoWork_UodateAll;
            worker.ProgressChanged += WorkerProgressChanged;
            worker.RunWorkerCompleted += WorkerRunWorkerCompleted;
            worker.WorkerReportsProgress = true;
            worker.RunWorkerAsync();
        }


        private void GetNewCurrencies()
        {
            infoPanel = InformationPanel.GetInformationPanel(this, "Getting New Currencies to add (wait a few seconds)...", 340, 100);

            var worker = new BackgroundWorker();
            worker.DoWork += WorkerDoWork_NewCurrencies;
            worker.ProgressChanged += WorkerProgressChanged;
            worker.RunWorkerCompleted += WorkerRunWorkerCompleted_NewCurrencies;
            worker.WorkerReportsProgress = true;
            worker.RunWorkerAsync();
        }


        private void GetCurrentCurrencies()
        {
            infoPanel = InformationPanel.GetInformationPanel(this, "Getting Current CRM Currencies...", 340, 100);

            var worker = new BackgroundWorker();
            worker.DoWork += WorkerDoWork;
            worker.ProgressChanged += WorkerProgressChanged;
            worker.RunWorkerCompleted += WorkerRunWorkerCompleted;
            worker.WorkerReportsProgress = true;
            worker.RunWorkerAsync();
        }

        private void WorkerProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            InformationPanel.ChangeInformationPanelMessage(infoPanel, e.UserState.ToString());
        }


        private void WorkerDoWork_Add(object sender, DoWorkEventArgs e)
        {
            DataTable dt = (DataTable)dgNewCurrencies.DataSource;
            for (int i = 0; i < dgNewCurrencies.RowCount;i++ )
            {
                if (dgNewCurrencies.Rows[i].Selected)
                {
                    Entity ent=new Entity("transactioncurrency");
                    ent.Attributes.Add("currencyname",dt.Rows[i]["Name"].ToString());
                    ent.Attributes.Add("isocurrencycode", dt.Rows[i]["Code"].ToString());
                    ent.Attributes.Add("currencysymbol", dt.Rows[i]["Symbol"].ToString());
                    ent.Attributes.Add("exchangerate", Convert.ToDecimal( dt.Rows[i]["Rate"].ToString()));
                    ent.Attributes.Add("currencyprecision", Convert.ToInt32( dt.Rows[i]["Precision"].ToString()));
                    service.Create(ent);
                }
                
            }
            WorkerDoWork(sender, e);
        }

        private void WorkerDoWork_UodateAll(object sender, DoWorkEventArgs e)
        {
            DataTable dt=(DataTable)dataGridView1.DataSource;
            foreach (DataRow dr in dt.Rows)
            {
                Entity ent = new Entity("transactioncurrency");
                ent.Attributes.Add("transactioncurrencyid", new Guid(dr["ID"].ToString()));
                if (Convert.ToDecimal(dr["New Rate"].ToString()) == 0)
                { continue; }
                ent.Attributes.Add("exchangerate", Convert.ToDecimal( dr["New Rate"].ToString()));
                service.Update(ent);
            }

            WorkerDoWork(sender,e);
        }
        private void WorkerDoWork_NewCurrencies(object sender, DoWorkEventArgs e)
        {
            DataTable dtOtherCurrencies = GetWebServicesCurrencies(current);
            ((BackgroundWorker)sender).ReportProgress(0, "Retrieved Currencies to add ok.");
            e.Result = dtOtherCurrencies;
        }
        private void WorkerRunWorkerCompleted_NewCurrencies(object sender, RunWorkerCompletedEventArgs e)
        {
            infoPanel.Dispose();
            infoPanel.Dispose();
            Controls.Remove(infoPanel);

            DataTable dtOtherCurrencies = (DataTable)e.Result;
            dgNewCurrencies.DataSource = dtOtherCurrencies;
            dgNewCurrencies.AutoGenerateColumns = true;
            dgNewCurrencies.Refresh();

        }
        private void WorkerDoWork(object sender, DoWorkEventArgs e)
        {
            QueryExpression query = new QueryExpression("organization");
            //"BaseCurrencyCode", "BaseCurrencyName", "BaseCurrencyPrecision", "BaseCurrencySymbol"
           // query.ColumnSet = new ColumnSet(true);
            //EntityCollection orgs = service.RetrieveMultiple(query);
            //Entity org = (Entity)orgs[0];
            StringBuilder str=new StringBuilder();

            str.Append(@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='true'>
                          <entity name='transactioncurrency'>
                            <attribute name='transactioncurrencyid' />
                            <attribute name='currencyname' />
                            <attribute name='isocurrencycode' />
                            <attribute name='currencysymbol' />
                            <attribute name='exchangerate' />
                            <attribute name='currencyprecision' />
                            <order attribute='currencyname' descending='false' />
                            <link-entity name='organization' from='basecurrencyid' to='transactioncurrencyid' alias='ad'></link-entity>
                          </entity>
                        </fetch>");
            EntityCollection default_ents = service.RetrieveMultiple(new FetchExpression(str.ToString()));

            Guid entid = (Guid)default_ents.Entities[0].Attributes["transactioncurrencyid"];
            
          

            str=new StringBuilder();
            str.Append(@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                          <entity name='transactioncurrency'>
                            <attribute name='transactioncurrencyid' />
                            <attribute name='currencyname' />
                            <attribute name='isocurrencycode' />
                            <attribute name='currencysymbol' />
                            <attribute name='exchangerate' />
                            <attribute name='currencyprecision' />
                            <order attribute='currencyname' descending='false' />
                            <filter type='and'>
                                <condition attribute='transactioncurrencyid' operator='ne' uiname='euro' uitype='transactioncurrency' value='" + entid.ToString() + @"' />
                            </filter>
                          </entity>
                        </fetch>");

            EntityCollection ents=service.RetrieveMultiple(new FetchExpression(str.ToString()));

            ArrayList arr = new ArrayList();
            arr.Add(default_ents);
            arr.Add(ents);
            ((BackgroundWorker)sender).ReportProgress(0, "Your default org currency seting has been retrieved!");
            e.Result = arr;
        }
        private EntityCollection default_ent;
        private CurrencyConvertor ws = new CurrencyConvertor();
        ArrayList current = new ArrayList();
        private void WorkerRunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            infoPanel.Dispose();
            infoPanel.Dispose();
            Controls.Remove(infoPanel);

            ArrayList arr = (ArrayList)e.Result;

             default_ent = (EntityCollection)arr[0];
            EntityCollection ents = (EntityCollection)arr[1];

            lblDefaultCurrency.Text = "Default Currency: " + (default_ent[0].Attributes["currencyname"].ToString() + 
                    " (" + default_ent[0].Attributes["currencysymbol"].ToString() + "-" + 
                    default_ent[0].Attributes["isocurrencycode"].ToString() + ")");

            
            DataTable dt = new DataTable();
            dt.Columns.Add("ID");
            dt.Columns.Add("Name");
            dt.Columns.Add("Code");
            dt.Columns.Add("Symbol");
            dt.Columns.Add("Rate");
            dt.Columns.Add("New Rate");
            dt.Columns.Add("Precision");

            
            foreach (Entity ent in ents.Entities)
            {
                DataRow dr=dt.NewRow();
                dr["ID"] = ((Guid)ent.Attributes["transactioncurrencyid"]).ToString();
                dr["Name"] = ent.Attributes["currencyname"].ToString();
                dr["Code"] = ent["isocurrencycode"].ToString();
                current.Add(ent["isocurrencycode"].ToString());
                dr["Symbol"] = ent.Attributes["currencysymbol"].ToString();
                dr["Rate"] = ent.Attributes["exchangerate"];
                double dblRate = 0;
                try
                {
                    Currency cur = (Currency)Enum.Parse(typeof(Currency), ent["isocurrencycode"].ToString());
                    Currency curdef = (Currency)Enum.Parse(typeof(Currency), default_ent[0].Attributes["isocurrencycode"].ToString());
               
                    dblRate = ws.ConversionRate(curdef, cur);
                }
                catch (System.Exception)
                {
                     
                }
                dr["New Rate"] = Convert.ToDecimal( dblRate);
                dr["Precision"] = ent.Attributes["currencyprecision"];
                dt.Rows.Add(dr);       
            }
            dataGridView1.DataSource = dt;
            dataGridView1.AutoGenerateColumns = true;
            dataGridView1.Refresh();
            dataGridView1.Columns[0].Visible = false;

            if (dataGridView1.RowCount > 0)
            {
                tbsUpdateAll.Enabled = true;
                tbsNewRates.Enabled = true;
            }


            
            //dgCurrentCurrencies.DataSource = ents.Entities;

            //try { 
            //    listView1.Clear(); 
            //    foreach (Entity ent in ents.Entities) 
            //    { 
            //        var item2 = new ListViewItem ( Text = ent.Attributes["currencyname"].ToString()}; 
                    
            //    listView1.Items.Add(new ListViewItem(
            //        ListViewDelegates.AddItem(lvEntities, item); 
            //    } 
                
            //}
            //catch (Exception error) { 
            //    string errorMessage = CrmExceptionHelper.GetErrorMessage(error, true); 
            //    CommonDelegates.DisplayMessageBox(ParentForm, errorMessage, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error

        
            //}
         
        }


        private DataTable GetWebServicesCurrencies(ArrayList current)
        {
            ArrayList added = new ArrayList();
            DataTable dt = new DataTable();
            dt.Columns.Add("Name");
            dt.Columns.Add("Code");
            dt.Columns.Add("Symbol");
            dt.Columns.Add("Precision");
            dt.Columns.Add("Rate");

            string[] arr = currencies_crm.Split("\t".ToCharArray());
            
            
            for (int i = 0; i < arr.Length;i++ )
            {
                string[] curr = arr[i].Split("\r\n".ToCharArray());

                if (current.Contains(curr[6]) || added.Contains(curr[6]) || curr[6] == default_ent[0].Attributes["isocurrencycode"].ToString())
                {
                    continue;
                }
                else
                {
                    added.Add(curr[6]);
                }
                DataRow dr = dt.NewRow();
                dr["Name"] = curr[4];
                dr["Code"] = curr[6];
                dr["Symbol"] = curr[8];
                dr["Precision"] = curr[10];
                //double dblRate = ws.ConversionRate(Currency.EUR, Currency.ARS);

                double dblRate = 0;
                try
                {
                    Currency cur = (Currency)Enum.Parse(typeof(Currency), curr[6]);
                    Currency curdef = (Currency)Enum.Parse(typeof(Currency), default_ent[0].Attributes["isocurrencycode"].ToString());

                    dblRate = ws.ConversionRate(curdef, cur);
                }
                catch (System.Exception)
                {
                    continue;
                }
                dr["Rate"] = dblRate;
                dt.Rows.Add(dr);
            }

           
            return dt;
        }
        #endregion Methods



        private void tbsUpdateAll_Click(object sender, EventArgs e)
        {
            
            if (service == null)
            {
                if (OnRequestConnection != null)
                {
                    var args = new RequestConnectionEventArgs { ActionName = "WhoAmI", Control = this, Parameter = null };
                    OnRequestConnection(this, args);
                }
            }
            else
            {
                UpdateAllRates();
            }
        }

        private void tbsNewRates_Click(object sender, EventArgs e)
        {
            
            if (service == null)
            {
                if (OnRequestConnection != null)
                {
                    var args = new RequestConnectionEventArgs { ActionName = "WhoAmI", Control = this, Parameter = null };
                    OnRequestConnection(this, args);
                }
            }
            else
            {
                GetCurrentCurrencies();
            }
        }
        private void tbsNewRates_Click_1(object sender, EventArgs e)
        {
            if (service == null)
            {
                if (OnRequestConnection != null)
                {
                    var args = new RequestConnectionEventArgs { ActionName = "WhoAmI", Control = this, Parameter = null };
                    OnRequestConnection(this, args);
                }
            }
            else
            {
                GetNewCurrencies();
            }
        }
        private void tsbCurrentCurrencies_Click(object sender, EventArgs e)
        {
            
            if (service == null)
            {
                if (OnRequestConnection != null)
                {
                    var args = new RequestConnectionEventArgs { ActionName = "WhoAmI", Control = this, Parameter = null };
                    OnRequestConnection(this, args);
                }
            }
            else
            {
                GetCurrentCurrencies();
            }
        }

        private void toolStripButton1_Click(object sender, EventArgs e)
        {
            if (OnCloseTool != null)
            {
                OnCloseTool(this, null);
            }
        }

        private void btnAdd_Click(object sender, EventArgs e)
        {
            infoPanel = InformationPanel.GetInformationPanel(this, "Adding new Currencies to CRM...", 340, 100);

            var worker = new BackgroundWorker();
            worker.DoWork += WorkerDoWork_Add;
            worker.ProgressChanged += WorkerProgressChanged;
            worker.RunWorkerCompleted += WorkerRunWorkerCompleted;
            worker.WorkerReportsProgress = true;
            worker.RunWorkerAsync();
        }


        private string currencies_crm = @"
Afganistán
افغانى
AFN
؋
2
	
Albania
Lek
ALL
Lek
2
	
Alemania
Euro
EUR
€
2
	
Alemania
euro
EUR
€
2
	
Arabia Saudí
ريال سعودي
SAR
ر.س.‏
2
	
Argelia
Dinar
DZD
DZD
2
	
Argentina
Peso
ARS
$
2
	
Armenia
դրամ
AMD
֏
2
	
Australia
Australian Dollar
AUD
$
2
	
Austria
Euro
EUR
€
2
	
Azerbaiyán
manat
AZN
manat
2
	
Bahrein
دينار بحريني
BHD
د.ب.‏
3
	
Bangladesh
টাকা
BDT
৳
2
	
Belarús
беларускі рубель
BYR
р.
2
	
Bélgica
euro
EUR
€
2
	
Belice
Belize Dollar
BZD
BZ$
2
	
Bolivia
Boliviano
BOB
$b
2
	
Bosnia-Herzegovina
konvertibilna marka
BAM
KM
2
	
Botswana
Pula
BWP
P
2
	
Brasil
Real
BRL
R$
2
	
Brunei
Ringgit
BND
$
0
	
Bulgaria
български лев
BGN
лв.
2
	
Camboya
x179Aៀល
KHR
៛
2
	
Canadá
Canadian Dollar
CAD
$
2

Caribe
Eastern Caribbean Dollar
XCD
EC$
2
	
Chile
Peso
CLP
$
2
	
Colombia
Peso
COP
$
2
	
Corea del Sur
원
KRW
₩
0
	
Costa Rica
Colón
CRC
₡
2
	
Croacia
hrvatska kuna
HRK
kn
2
	
Dinamarca
Dansk krone
DKK
kr.
2
	
Ecuador
US Dollar
USD
$
2
	
Egipto
جنيه مصري
EGP
ج.م.‏
2
	
El Salvador
US Dolar
USD
$
2
	
Emiratos Árabes Unidos
درهم اماراتي
AED
د.إ.‏
2
	
Eritrea
ናቕፋ
ERN
ERN
2
	
Eslovaquia
euro
EUR
EUR
2
	
Eslovenia
evro
EUR
€
2
	
España
Euro
EUR
€
2
	
Estados Unidos
US Dollar
USD
$
2
	
Estonia
euro
EUR
€
2
	
Etiopía
ብር
ETB
ETB
2
	
Filipinas
Philippine Peso
PHP
PhP
2
	
Finlandia
euro
EUR
€
2
	
Francia
euro
EUR
€
2
	
Georgia
ლარი
GEL
ლ.
2
	
Grecia
ευρώ
EUR
€
2
	
Groenlandia
kuruuni
DKK
kr.
2
	
Guatemala
Quetzal
GTQ
Q
2
	
Honduras
Lempira
HNL
L.
2
	
Hong Kong, zona administrativa especial de la República Popular China
港幣
HKD
HK$
2
	
Hungría
forint
HUF
Ft
2
	
India
Rupee
INR
₹
2
	
Indonesia
Rupiah
IDR
Rp
0
	
Irak
دیناری عێراقی
IQD
د.ع.‏
2
	
Irán
ریال
IRR
ريال
2
	
Irlanda
Euro
EUR
€
2
	
Islandia
Króna
ISK
kr.
0
	
Islas Feroe
Dansk krone
DKK
kr.
2
	
Israel
שקל חדש
ILS
₪
2
	
Italia
euro
EUR
€
2
	
Jamaica
Jamaican Dollar
JMD
J$
2
	
Japón
円
JPY
¥
0
	
Jordania
دينار اردني
JOD
د.ا.‏
3
	
Kazajistán
теңге
KZT
₸
2
	
Kenia
Shilingi
KES
S
2
	
Kirguizistán
сом
KGS
сом
2
	
Kuwait
دينار كويتي
KWD
د.ك.‏
3
	
Letonia
Lats
LVL
Ls
2
	
Líbano
ليرة لبناني
LBP
ل.ل.‏
2
	
Libia
دينار ليبي
LYD
د.ل.‏
3
	
Liechtenstein
Schweizer Franken
CHF
CHF
2
	
Lituania
Litas
LTL
Lt
2
	
Luxemburgo
Euro
EUR
€
2
	
Macao (Región Administrativa Especial)
澳門幣
MOP
MOP
2
	
Macedonia (FYROM)
денар
MKD
ден.
2
	
Malasia
Ringgit Malaysia
MYR
RM
0
	
Maldivas
ރުފިޔާ
MVR
ރ.
2
	
Malta
ewro
EUR
€
2
	
Marruecos
درهم مغربي
MAD
د.م.‏
2
	
México
Peso
MXN
$
2
	
Mongolia
Төгрөг
MNT
₮
2
	
Montenegro
evro
EUR
€
2
	
Montenegro
евро
EUR
€
2
	
Nepal
रुपैयाँ
NPR
रु
2
	
Nicaragua
Córdoba
NIO
C$
2
	
Nigeria
Naira
NGN
₦
2
	
Noruega
Norsk krone
NOK
kr
2
	
Nueva Zelanda
New Zealand Dollar
NZD
$
2
	
Omán
ريال عماني
OMR
ر.ع.‏
3
	
Países Bajos
euro
EUR
€
2
	
Panamá
Balboa
PAB
B/.
2
	
Paraguay
Guaraní
PYG
₲
0
	
Perú
Nuevo Sol
PEN
S/.
2
	
Polonia
Złoty
PLN
zł
2
	
Portugal
euro
EUR
€
2
	
Principado de Mónaco
euro
EUR
€
2
	
Puerto Rico
US Dollar
USD
$
2
	
Qatar
ريال قطري
QAR
ر.ق.‏
2
	
R.D.P. de Laos
ກີບ
LAK
₭
2
	
Reino Unido
Pound Sterling
GBP
£
2
	
República Bolivariana de Venezuela
Bolívar Fuerte
VEF
Bs.F.
2
	
República Checa
česká koruna
CZK
Kč
2
	
República Dominicana
Peso
DOP
RD$
2
	
República Islámica de Pakistán
روپيه
PKR
Rs
2
	
República Popular China
人民币
CNY
¥
2
	
Ruanda
Ifaranga
RWF
RWF
2
	
Rumania
Leu
RON
RON
2
	
Rusia
рубль
RUB
р.
2
	
Senegal
CFA
XOF
XOF
2
	
Serbia
dinar
RSD
din.
2
	
Serbia y Montenegro (Ex-República)
dinar
RSD
din.
2
	
Serbia y Montenegro (Ex-República)
динар
RSD
дин.
2
	
Singapur
Singapore Dollar
SGD
$
2
	
Siria
ليرة سوري
SYP
ل.س.‏
2
	
Sri Lanka
ரூபாய்
LKR
Rs
2
	
Sudáfrica
Rand
ZAR
R
2
	
Suecia
Svensk krona
SEK
kr
2

Suiza
franc suisse
CHF
fr.
2
	
Tailandia
บาท
THB
฿
2
	
Taiwán
新台幣
TWD
NT$
2
	
Tayikistán
Сомонӣ
TJS
смн
2
	
Trinidad y Tobago
Trinidad Dollar
TTD
TT$
2
	
Túnez
دينار تونسي
TND
د.ت.‏
3
	
Turkmenistán
manat
TMT
m.
2
	
Turquía
Türk Lirası
TRY
₺
2
	
Ucrania
гривня
UAH
₴
2
	
Uruguay
Peso
UYU
$U
2
	
Uzbekistán
so‘m
UZS
so'm
2
	
Vietnam
Đồng
VND
₫
2
	
Yemen
ريال يمني
YER
ر.ي.‏
2
	
Zimbabue
US Dollar
USD
$
2";

        

        

       
   
    }


}
