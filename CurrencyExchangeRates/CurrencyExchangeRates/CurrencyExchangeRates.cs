using CurrencyExchangeRates.currencyconvertor;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections;
using System.ComponentModel;
using System.Data;
using System.Text;
using System.Web.UI.WebControls;
using System.Windows.Forms;
using XrmToolBox.Extensibility;
using XrmToolBox.Extensibility.Interfaces;

namespace CurrencyExchangeRates
{
    public partial class CurrencyExchangeRates : PluginControlBase
    {
        private ToolStripButton toolStripButton1;
        private ToolStripButton toolStripButton2;
        private ToolStripButton tbsUpdateAll;
        private ToolStripButton tbsNewRates;
        private System.Windows.Forms.Label lblDefaultCurrency;
        private GroupBox grpCurrent;
        private DataGridView dataGridView1;
        private GroupBox groupBox1;
        private DataGridView dgNewCurrencies;
        private System.Windows.Forms.Button btnAdd;
        private System.Windows.Forms.Label label1;
        private LinkLabel linkLabel1;
        private ToolStrip toolStrip1;


        #region Constructor

        /// <summary>
        /// Initializes a new instance of class <see cref="CurrencyExchangeRates"/>
        /// </summary>
        public CurrencyExchangeRates()
        {
            InitializeComponent();



        }
        #endregion
        #region Variables

        /// <summary>
        /// Microsoft Dynamics CRM 2011 Organization Service
        /// </summary>
        //private IOrganizationService service;

        /// <summary>
        /// Panel used to display progress information
        /// </summary>
        private System.Windows.Forms.Panel infoPanel;

        #endregion Variables
        

        #region Properties

        /// <summary>
        /// Gets the organization service used by the tool
        /// </summary>
        //public IOrganizationService Service
        //{
        //    get { return service; }
        //}

        /// <summary>
        /// Gets the logo to display in the tools list
        /// </summary>
        //public Image PluginLogo
        //{
        //    get { null; }
        //}

        #endregion
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(CurrencyExchangeRates));
            this.toolStrip1 = new System.Windows.Forms.ToolStrip();
            this.toolStripButton1 = new System.Windows.Forms.ToolStripButton();
            this.toolStripButton2 = new System.Windows.Forms.ToolStripButton();
            this.tbsUpdateAll = new System.Windows.Forms.ToolStripButton();
            this.tbsNewRates = new System.Windows.Forms.ToolStripButton();
            this.lblDefaultCurrency = new System.Windows.Forms.Label();
            this.grpCurrent = new System.Windows.Forms.GroupBox();
            this.dataGridView1 = new System.Windows.Forms.DataGridView();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.dgNewCurrencies = new System.Windows.Forms.DataGridView();
            this.btnAdd = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.linkLabel1 = new System.Windows.Forms.LinkLabel();
            this.toolStrip1.SuspendLayout();
            this.grpCurrent.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            this.groupBox1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgNewCurrencies)).BeginInit();
            this.SuspendLayout();
            // 
            // toolStrip1
            // 
            this.toolStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.toolStripButton1,
            this.toolStripButton2,
            this.tbsUpdateAll,
            this.tbsNewRates});
            this.toolStrip1.Location = new System.Drawing.Point(0, 0);
            this.toolStrip1.Name = "toolStrip1";
            this.toolStrip1.Size = new System.Drawing.Size(786, 25);
            this.toolStrip1.TabIndex = 0;
            this.toolStrip1.Text = "toolStrip1";
            // 
            // toolStripButton1
            // 
            this.toolStripButton1.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image;
            this.toolStripButton1.Image = ((System.Drawing.Image)(resources.GetObject("toolStripButton1.Image")));
            this.toolStripButton1.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.toolStripButton1.Name = "toolStripButton1";
            this.toolStripButton1.Size = new System.Drawing.Size(23, 22);
            this.toolStripButton1.Tag = "";
            this.toolStripButton1.Text = "Close";
            this.toolStripButton1.Click += new System.EventHandler(this.toolStripButton1_Click);
            // 
            // toolStripButton2
            // 
            this.toolStripButton2.Image = ((System.Drawing.Image)(resources.GetObject("toolStripButton2.Image")));
            this.toolStripButton2.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.toolStripButton2.Name = "toolStripButton2";
            this.toolStripButton2.Size = new System.Drawing.Size(205, 22);
            this.toolStripButton2.Text = "Get Current Currencies from CRM";
            this.toolStripButton2.Click += new System.EventHandler(this.toolStripButton2_Click);
            // 
            // tbsUpdateAll
            // 
            this.tbsUpdateAll.Enabled = false;
            this.tbsUpdateAll.Image = ((System.Drawing.Image)(resources.GetObject("tbsUpdateAll.Image")));
            this.tbsUpdateAll.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.tbsUpdateAll.Name = "tbsUpdateAll";
            this.tbsUpdateAll.Size = new System.Drawing.Size(142, 22);
            this.tbsUpdateAll.Text = "Update All CRM Rates";
            this.tbsUpdateAll.Click += new System.EventHandler(this.tbsUpdateAll_Click);
            // 
            // tbsNewRates
            // 
            this.tbsNewRates.Enabled = false;
            this.tbsNewRates.Image = ((System.Drawing.Image)(resources.GetObject("tbsNewRates.Image")));
            this.tbsNewRates.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.tbsNewRates.Name = "tbsNewRates";
            this.tbsNewRates.Size = new System.Drawing.Size(168, 22);
            this.tbsNewRates.Text = "Get New Currencies to add";
            this.tbsNewRates.Click += new System.EventHandler(this.tbsNewRates_Click);
            // 
            // lblDefaultCurrency
            // 
            this.lblDefaultCurrency.AutoSize = true;
            this.lblDefaultCurrency.Font = new System.Drawing.Font("Microsoft Sans Serif", 15F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblDefaultCurrency.Location = new System.Drawing.Point(25, 29);
            this.lblDefaultCurrency.Name = "lblDefaultCurrency";
            this.lblDefaultCurrency.Size = new System.Drawing.Size(164, 25);
            this.lblDefaultCurrency.TabIndex = 1;
            this.lblDefaultCurrency.Text = "Default Currency:";
            // 
            // grpCurrent
            // 
            this.grpCurrent.Controls.Add(this.dataGridView1);
            this.grpCurrent.Location = new System.Drawing.Point(48, 85);
            this.grpCurrent.Name = "grpCurrent";
            this.grpCurrent.Size = new System.Drawing.Size(714, 227);
            this.grpCurrent.TabIndex = 2;
            this.grpCurrent.TabStop = false;
            this.grpCurrent.Text = "Other Currencies in CRM";
            // 
            // dataGridView1
            // 
            this.dataGridView1.AllowUserToAddRows = false;
            this.dataGridView1.AllowUserToDeleteRows = false;
            this.dataGridView1.AllowUserToResizeColumns = false;
            this.dataGridView1.AllowUserToResizeRows = false;
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.Location = new System.Drawing.Point(7, 20);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.dataGridView1.Size = new System.Drawing.Size(701, 197);
            this.dataGridView1.TabIndex = 1;
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.dgNewCurrencies);
            this.groupBox1.Location = new System.Drawing.Point(48, 352);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(714, 201);
            this.groupBox1.TabIndex = 3;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Add New Currencies";
            // 
            // dgNewCurrencies
            // 
            this.dgNewCurrencies.AllowUserToAddRows = false;
            this.dgNewCurrencies.AllowUserToDeleteRows = false;
            this.dgNewCurrencies.AllowUserToResizeColumns = false;
            this.dgNewCurrencies.AllowUserToResizeRows = false;
            this.dgNewCurrencies.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgNewCurrencies.Location = new System.Drawing.Point(7, 20);
            this.dgNewCurrencies.Name = "dgNewCurrencies";
            this.dgNewCurrencies.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.dgNewCurrencies.Size = new System.Drawing.Size(701, 175);
            this.dgNewCurrencies.TabIndex = 0;
            // 
            // btnAdd
            // 
            this.btnAdd.Location = new System.Drawing.Point(361, 323);
            this.btnAdd.Name = "btnAdd";
            this.btnAdd.Size = new System.Drawing.Size(35, 23);
            this.btnAdd.TabIndex = 4;
            this.btnAdd.Text = "^^";
            this.btnAdd.UseVisualStyleBackColor = true;
            this.btnAdd.Click += new System.EventHandler(this.btnAdd_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(551, 323);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(211, 13);
            this.label1.TabIndex = 5;
            this.label1.Text = "All the Currency rates are downloading from";
            // 
            // linkLabel1
            // 
            this.linkLabel1.AutoSize = true;
            this.linkLabel1.Location = new System.Drawing.Point(571, 340);
            this.linkLabel1.Name = "linkLabel1";
            this.linkLabel1.Size = new System.Drawing.Size(147, 13);
            this.linkLabel1.TabIndex = 6;
            this.linkLabel1.TabStop = true;
            this.linkLabel1.Text = "http://www.webservicex.net/";
            // 
            // CurrencyExchangeRates
            // 
            this.Controls.Add(this.linkLabel1);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.btnAdd);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.grpCurrent);
            this.Controls.Add(this.lblDefaultCurrency);
            this.Controls.Add(this.toolStrip1);
            this.Name = "CurrencyExchangeRates";
            this.Size = new System.Drawing.Size(786, 567);
            this.toolStrip1.ResumeLayout(false);
            this.toolStrip1.PerformLayout();
            this.grpCurrent.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            this.groupBox1.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.dgNewCurrencies)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        private void toolStripButton2_Click(object sender, EventArgs e)
        {
            ExecuteMethod(GetCurrentCurrencies);
            
        }

        private void tbsUpdateAll_Click(object sender, EventArgs e)
        {
            ExecuteMethod(UpdateAllRates);

            
        }

        private void tbsNewRates_Click(object sender, EventArgs e)
        {
            ExecuteMethod(GetNewCurrencies);
        }

        private void toolStripButton1_Click(object sender, EventArgs e)
        {
            CloseTool();
           
        }

        private void btnAdd_Click(object sender, EventArgs e)
        {
            ExecuteMethod(AddCurrencies);
        }

        private void AddCurrencies()
        {
            WorkAsync(new WorkAsyncInfo
            {
                Message = "Adding new Currencies to CRM...",
                Work = (w, e) =>
                {
                    DataTable dt = (DataTable)dgNewCurrencies.DataSource;
                    for (int i = 0; i < dgNewCurrencies.RowCount; i++)
                    {
                        if (dgNewCurrencies.Rows[i].Selected)
                        {
                            Entity ent = new Entity("transactioncurrency");
                            ent.Attributes.Add("currencyname", dt.Rows[i]["Name"].ToString());
                            ent.Attributes.Add("isocurrencycode", dt.Rows[i]["Code"].ToString());
                            ent.Attributes.Add("currencysymbol", dt.Rows[i]["Symbol"].ToString());
                            ent.Attributes.Add("exchangerate", Convert.ToDecimal(dt.Rows[i]["Rate"].ToString()));
                            ent.Attributes.Add("currencyprecision", Convert.ToInt32(dt.Rows[i]["Precision"].ToString()));
                            Service.Create(ent);
                        }

                    }
                    QueryExpression query = new QueryExpression("organization");
                    //"BaseCurrencyCode", "BaseCurrencyName", "BaseCurrencyPrecision", "BaseCurrencySymbol"
                    // query.ColumnSet = new ColumnSet(true);
                    //EntityCollection orgs = service.RetrieveMultiple(query);
                    //Entity org = (Entity)orgs[0];
                    StringBuilder str = new StringBuilder();

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
                    EntityCollection default_ents = Service.RetrieveMultiple(new FetchExpression(str.ToString()));

                    Guid entid = (Guid)default_ents.Entities[0].Attributes["transactioncurrencyid"];



                    str = new StringBuilder();
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

                    EntityCollection ents = Service.RetrieveMultiple(new FetchExpression(str.ToString()));

                    ArrayList arr = new ArrayList();
                    arr.Add(default_ents);
                    arr.Add(ents);
                    w.ReportProgress(0, "Your default org currency seting has been retrieved!");
                    e.Result = arr;
                },
                ProgressChanged = e =>
                {
                    // If progress has to be notified to user, use the following method:
                    SetWorkingMessage("Message to display");
                },
                PostWorkCallBack = e =>
                {
                    //MessageBox.Show(string.Format("You are {0}", (Guid)e.Result));
                    ExecuteMethod(GetCurrentCurrencies);

                },
                AsyncArgument = null,
                IsCancelable = true,
                MessageWidth = 340,
                MessageHeight = 150
            });
        }
        private void GetNewCurrencies() {
            WorkAsync(new WorkAsyncInfo
            {
                Message = "Getting New Currencies to add (wait a few seconds) - reading from http://www.webservicex.net/... ",
                Work = (w, e) =>
                {
                    DataTable dtOtherCurrencies = GetWebServicesCurrencies(current);
                    w.ReportProgress(0, "Retrieved Currencies to add ok.");
                    e.Result = dtOtherCurrencies;
                },
                ProgressChanged = e =>
                {
                    // If progress has to be notified to user, use the following method:
                    SetWorkingMessage("Message to display");
                },
                PostWorkCallBack = e =>
                {
                    //MessageBox.Show(string.Format("You are {0}", (Guid)e.Result));
                    DataTable dtOtherCurrencies = (DataTable)e.Result;
                    dgNewCurrencies.DataSource = dtOtherCurrencies;
                    dgNewCurrencies.AutoGenerateColumns = true;
                    dgNewCurrencies.Refresh();
                },
                AsyncArgument = null,
                IsCancelable = true,
                MessageWidth = 340,
                MessageHeight = 150
            });
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


            for (int i = 0; i < arr.Length; i++)
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


        private void UpdateAllRates()
        {
            WorkAsync(new WorkAsyncInfo
            {
                Message = "Updating CRM Currencies Rates...",
                Work = (w, e) =>
                {
                    DataTable dt = (DataTable)dataGridView1.DataSource;
                    foreach (DataRow dr in dt.Rows)
                    {
                        Entity ent = new Entity("transactioncurrency");
                        ent.Attributes.Add("transactioncurrencyid", new Guid(dr["ID"].ToString()));
                        if (Convert.ToDecimal(dr["New Rate"].ToString()) == 0)
                        { continue; }
                        ent.Attributes.Add("exchangerate", Convert.ToDecimal(dr["New Rate"].ToString()));
                        Service.Update(ent);
                    }

                    //WorkerDoWork(sender, e);
                    QueryExpression query = new QueryExpression("organization");
                    //"BaseCurrencyCode", "BaseCurrencyName", "BaseCurrencyPrecision", "BaseCurrencySymbol"
                    // query.ColumnSet = new ColumnSet(true);
                    //EntityCollection orgs = service.RetrieveMultiple(query);
                    //Entity org = (Entity)orgs[0];
                    StringBuilder str = new StringBuilder();

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
                    EntityCollection default_ents = Service.RetrieveMultiple(new FetchExpression(str.ToString()));

                    Guid entid = (Guid)default_ents.Entities[0].Attributes["transactioncurrencyid"];


                    str = new StringBuilder();
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

                    EntityCollection ents = Service.RetrieveMultiple(new FetchExpression(str.ToString()));

                    ArrayList arr = new ArrayList();
                    arr.Add(default_ents);
                    arr.Add(ents);
                    w.ReportProgress(0, "Your default org currency seting has been retrieved!");
                    e.Result = arr;
                },
                ProgressChanged = e =>
                {
                    // If progress has to be notified to user, use the following method:
                    SetWorkingMessage("Message to display");
                },
                PostWorkCallBack = e =>
                {
                    //MessageBox.Show(string.Format("You are {0}", (Guid)e.Result));
                },
                AsyncArgument = null,
                IsCancelable = true,
                MessageWidth = 340,
                MessageHeight = 150
            });
            
        }

        private void GetCurrentCurrencies()
        {
            WorkAsync(new WorkAsyncInfo
            {
                Message = "Getting Current CRM Currencies...",
                Work = (w, e) =>
                {
                    QueryExpression query = new QueryExpression("organization");
                    //"BaseCurrencyCode", "BaseCurrencyName", "BaseCurrencyPrecision", "BaseCurrencySymbol"
                    // query.ColumnSet = new ColumnSet(true);
                    //EntityCollection orgs = service.RetrieveMultiple(query);
                    //Entity org = (Entity)orgs[0];
                    StringBuilder str = new StringBuilder();

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
                    EntityCollection default_ents = Service.RetrieveMultiple(new FetchExpression(str.ToString()));

                    Guid entid = (Guid)default_ents.Entities[0].Attributes["transactioncurrencyid"];



                    str = new StringBuilder();
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

                    EntityCollection ents = Service.RetrieveMultiple(new FetchExpression(str.ToString()));

                    ArrayList arr = new ArrayList();
                    arr.Add(default_ents);
                    arr.Add(ents);
                    
                    e.Result = arr;
                    
                },
                ProgressChanged = e =>
                {
                    // If progress has to be notified to user, use the following method:
                    SetWorkingMessage(e.UserState.ToString());
                },
                PostWorkCallBack = e =>
                {
                    MessageBox.Show(string.Format("Your default org currency seting has been retrieved!"));
                    //infoPanel.Dispose();
                    //infoPanel.Dispose();
                    //Controls.Remove(infoPanel);

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
                        DataRow dr = dt.NewRow();
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
                        dr["New Rate"] = Convert.ToDecimal(dblRate);
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

                },
                AsyncArgument = null,
                IsCancelable = true,
                MessageWidth = 340,
                MessageHeight = 150
            });

           
            //var worker = new BackgroundWorker();
            //worker.DoWork += WorkerDoWork;
            //worker.ProgressChanged += WorkerProgressChanged;
            //worker.RunWorkerCompleted += WorkerRunWorkerCompleted;
            //worker.WorkerReportsProgress = true;
            //worker.RunWorkerAsync();
        }
       
       
        private EntityCollection default_ent;
        private CurrencyConvertor ws = new CurrencyConvertor();
        ArrayList current = new ArrayList();
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
