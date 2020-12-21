using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Exchange.WebServices.Data;
using Mono.Web;
using JR.Utils.GUI.Forms;
using SpamMailML.Model;
using System.Text.RegularExpressions;
using NPoco.FluentMappings;
using Chilkat;

namespace SpamMail
{
    public partial class Form1 : Form
    {
        ExchangeService exchange = null;
        string[] basliklar = new string[500];
        string[] icerikler = new string[500];
        int i = 0;
        string username;
        string domain;
        public Form1()
        {
            InitializeComponent();
            lstMsg.Clear();
            lstMsg.View = View.Details;
            lstMsg.Columns.Add("Tarih/Saat", 170);
            lstMsg.Columns.Add("Gönderen", 250);
            lstMsg.Columns.Add("Konu", 380);
            lstMsg.Columns.Add("İçerik", 600);
            lstMsg.Columns.Add("Spam", 60);
            lstMsg.Columns.Add("Rate", 70);
            lstMsg.FullRowSelect = true;
            lstMsg.ShowItemToolTips = true;
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        public void ConnectToExchangeServer()
        {
            lblMsg.Text = "Exchange Server'a bağlanılıyor....";
            lblMsg.Refresh();
            try
            {
                exchange = new ExchangeService(ExchangeVersion.Exchange2007_SP1);
                exchange.Credentials = new WebCredentials(textBox1.Text, textBox2.Text, domain);
                exchange.AutodiscoverUrl(textBox1.Text);

                lblMsg.Text = "Exchange Server'a bağlandı : " + exchange.Url.Host+"\n Günlük Mailler Gösteriliyor.";
                lblMsg.Refresh();

            }
            catch (Exception ex)
            {
                lblMsg.Text = "Exchange Server'a bağlanırken hata oluştu.Lütfen maili ve şifreyi kontrol edin.\n" + ex.Message;
                lblMsg.Refresh();
            }

        }
        private void lstMsg_SelectedIndexChanged(object sender, EventArgs e)
        {
            ListView.SelectedIndexCollection indices = lstMsg.SelectedIndices;
            if (indices.Count > 0)
            {
                FlexibleMessageBox.Show(icerikler[indices[0]], basliklar[indices[0]]);
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            

            if (textBox1.Text=="" || textBox2.Text=="")
            {
                MessageBox.Show("Lütfen mail adresi ve şifrenizi giriniz.");
            }
            else
            {
                
                if (Regex.IsMatch(textBox1.Text, @"(@)"))
                {
                    this.Size = new Size(1600, 450);
                    this.Location = new Point(50, 50);
                    username = textBox1.Text.Split('@')[0];
                    domain = textBox1.Text.Split('@')[1];

                    lblMsg.Visible = true;
                    i = 0;
                    lstMsg.Items.Clear();
                    ConnectToExchangeServer();
                    TimeSpan ts = new TimeSpan(0, -24, 0, 0);
                    DateTime date = DateTime.Now.Add(ts);
                    SearchFilter.IsGreaterThanOrEqualTo filter = new SearchFilter.IsGreaterThanOrEqualTo(ItemSchema.DateTimeReceived, date);

                    if (exchange != null)
                    {
                        PropertySet itempropertyset = new PropertySet(BasePropertySet.FirstClassProperties);
                        itempropertyset.RequestedBodyType = BodyType.Text;
                        ItemView itemview = new ItemView(1000);
                        itemview.PropertySet = itempropertyset;

                        //FindItemsResults<Item> findResults = service.FindItems(WellKnownFolderName.Inbox, "subject:TODO", itemview);
                        this.Size = new Size(1600, 800);
                        lstMsg.Width = 1560;
                        lstMsg.Height = 650;
                        try
                        {
                            FindItemsResults<Item> findResults = exchange.FindItems(WellKnownFolderName.Inbox, filter, new ItemView(100));
                            foreach (Item item in findResults)
                            {
                                lstMsg.Visible = true;
                                item.Load(itempropertyset);
                                String content = item.Body;
                                icerikler[i] = content;

                                ModelInput sampleData = new ModelInput()
                                {
                                    Col1 = content,
                                };

                                // Make a single prediction on the sample data and print results
                                var predictionResult = ConsumeModel.Predict(sampleData);

                                String durum;
                                if (predictionResult.Prediction == "spam")
                                {
                                    durum = "YES";                                                                     
                                }
                                else
                                {
                                    durum = "NO";                                    
                                }



                                EmailMessage message = EmailMessage.Bind(exchange, item.Id);
                                basliklar[i] = message.Subject;
                                i++;
                                ListViewItem listitem = new ListViewItem(new[]
                                {
                                message.DateTimeReceived.ToString(), message.From.Name.ToString() + "(" + message.From.Address.ToString() + ")", message.Subject,
                                content,durum,predictionResult.Score.Max().ToString("0.##")
                                });
                                
                                lstMsg.Items.Add(listitem);
                                                              
                            }
                            if (findResults.Items.Count <= 0)
                            {
                                lstMsg.Items.Add("Yeni Mail Bulunamadı.!!");

                            }
                            colorListcolor(lstMsg);
                        }
                        catch
                        {
                            MessageBox.Show("Mail adresi veya şifre yanlış.");
                            textBox1.Text = "";
                            textBox2.Text = "";
                            lblMsg.Text = "";
                            lblMsg.Visible = false;
                            this.Size = new Size(800, 450);
                            Screen screen = Screen.FromControl(this);

                            Rectangle workingArea = screen.WorkingArea;
                            this.Location = new Point()
                            {
                                X = Math.Max(workingArea.X, workingArea.X + (workingArea.Width - this.Width) / 2),
                                Y = Math.Max(workingArea.Y, workingArea.Y + (workingArea.Height - this.Height) / 2)
                            };

                        }                    
                    }
                }
                else
                {
                    MessageBox.Show("Lütfen doğru bir mail formatı giriniz.");
                }              
            }           
        }

        public static void colorListcolor(ListView lsvMain)
        {

            foreach (ListViewItem lvw in lsvMain.Items)
            {
                lvw.UseItemStyleForSubItems = false;

                for (int i = 0; i < lsvMain.Columns.Count; i++)
                {
                    if (lvw.SubItems[4].Text.ToString() == "YES")
                    {
                        lvw.SubItems[4].BackColor = Color.Red;
                        lvw.SubItems[4].ForeColor = Color.White;
                    }
                    else
                    {
                        lvw.SubItems[4].BackColor = Color.Green;
                        lvw.SubItems[4].ForeColor = Color.White;
                    }
                }
            }
        }

    }
}
