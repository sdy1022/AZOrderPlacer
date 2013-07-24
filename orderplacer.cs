using System.Reflection;
using Microsoft.CSharp.RuntimeBinder;
using mshtml;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using AmazonBot;
using Binder = Microsoft.CSharp.RuntimeBinder.Binder;

namespace AmazonBotDY
{

    public partial class OrderPlacer : Form
    {
        [CompilerGenerated]
        private static class SiteContainer1
        {
            // Fields
            public static CallSite<Func<CallSite, object, IHTMLRect>> Site2;
        }

        public delegate void CallbackType(object sender, WebBrowserDocumentCompletedEventArgs e);
        private CallbackType afterLoginCallback;
        private CallbackType documentCompleteCallback;
        private int tries;
        private AmazonOrder amazonOrder = new AmazonOrder();
        private List<DataGridViewCell> statusCells;
        private string debugMsg = "";
        private double paymentWaitTime;
        private Random random = new Random(DateTime.Now.Second);
        private CallbackType oldProcessedbyTitle;
        private debugForm debugForm = new debugForm();

        private int paymentTries;
        public OrderPlacer()
        {
            InitializeComponent();

            this.webBrowser1.ScriptErrorsSuppressed = true;
            debugForm.Show();
        }

        public void startOrder(AmazonOrder op, List<DataGridViewCell> statusCells)
        {
            this.amazonOrder = (AmazonOrder)op.Clone();
            this.statusCells = statusCells;
            this.timer_selfDestruct.Interval = TimerConstant.second;
            this.timer_selfDestruct.Stop();
           // this.timer_selfDestruct.Start();
            this.appendStatus("ordering", false);
            this.processOldCart();
        }

        private void OrderPlacer_FormClosed(object sender, FormClosedEventArgs e)
        {
            if (!this.label1.Text.Equals("logged out"))
            {
                this.appendStatus("Aborted at " + DateTime.Now.ToString(), false);
            }
        }
        private void appendStatus(string msg, bool append)
        {
            foreach (DataGridViewCell cell in this.statusCells)
            {
                if (!append)
                {
                    cell.Value = ("");
                }
                cell.Value = (cell.Value + msg);
            }
        }

        public void processOldCart()
        {
            this.afterLoginCallback = new CallbackType(this.saveOldCartItems);
            if (this.amazonOrder.oldCartPolicy.Equals(OldCartPolicy.Save))
            {
                this.webBrowser1.Navigate(AmazonLinks.logout);
                while (this.webBrowser1.ReadyState != WebBrowserReadyState.Complete)
                {
                    Application.DoEvents();
                }
                this.documentCompleteCallback = new CallbackType(this.signIn);
                this.webBrowser1.Navigate(AmazonLinks.cart);
            }
            else
            {
                this.loadCart();
            }
        }
        private bool checkTitle(string t)
        {
            return this.webBrowser1.Document.Title.ToLower().StartsWith(t.ToLower());
        }
        private void loadCart()
        {
            this.webBrowser1.Navigate(this.amazonOrder.purchaseURL);
            this.label1.Text = "Cart loaded";
            this.documentCompleteCallback = new CallbackType(this.mergeCart);
            this.afterLoginCallback = new CallbackType(this.shippingAddress);
        }
        private void mergeCart(object sender, WebBrowserDocumentCompletedEventArgs e)
        {
            this.webBrowser1.Document.GetElementById("add").InvokeMember("click");
            this.documentCompleteCallback = new CallbackType(this.proceedToCheckout);
        }

        private void proceedToCheckout(object sender, WebBrowserDocumentCompletedEventArgs e)
        {
            try
            {
                this.webBrowser1.Document.GetElementById("proceedToCheckout").InvokeMember("click");
                this.documentCompleteCallback = new CallbackType(this.signIn);
            }
            catch
            {
                this.processPageByTitle(sender, e);
            }
        }


        private void invokeDelagateByTitle(CallbackType d, object sender, WebBrowserDocumentCompletedEventArgs e)
        {
            // if ((this.oldProcessedbyTitle != null) && (this.oldProcessedbyTitle != d))
            // changed by dayang on 12/15
            // found on 12/15 the fisrt call , d is always null
            if ((this.oldProcessedbyTitle != d))
            {
                d(sender, e);
                this.oldProcessedbyTitle = d;
            }
        }

        private void shippingAddress(object sender, WebBrowserDocumentCompletedEventArgs e)
        {
            if (this.checkTitle(AmazonPageTitles.shippingAddress))
            {
                //  this.webBrowser1.Document.GetElementFromPoint()
                foreach (HtmlElement element3 in this.webBrowser1.Document.GetElementsByTagName("a"))
                {
                    string sss = element3.InnerText;

                    if (sss.Contains("Ship to this address"))
                    {
                        element3.InvokeMember("Click");
                        break;
                    }
                }
                // this.webBrowser1.Document.GetElementById("submit").InvokeMember("click");
                this.documentCompleteCallback = new CallbackType(this.shippingOption);

            }
            else
            {
                this.processPageByTitle(sender, e);
            }
        }
        private void shippingOption(object sender, WebBrowserDocumentCompletedEventArgs e)
        {
            if (!this.checkTitle(AmazonPageTitles.shippingOption))
            {
                this.processPageByTitle(sender, e);
            }
            else //FREE Super Saver Shipping
            {
                if (!this.amazonOrder.shipping.Equals(ShippingModeCollection.AccountDefault))
                {

                    HtmlElement elem = this.webBrowser1.Document.GetElementById("order_0_ShippingSpeed_sss-us");
                    if (elem != null)
                        elem.SetAttribute("checked", "checked");
                    /*
                    bool flag = false;
                    HtmlElementCollection elementsByTagName = this.webBrowser1.Document.GetElementsByTagName("div");
                    string str = this.amazonOrder.shipping.ToString().ToLower();
                    foreach (HtmlElement element in elementsByTagName)
                    {
                        // add on 07/23/2013
                        if(element.InnerText==null)
                            continue;
                     //   if (element.InnerText.ToLower().StartsWith(str))
                        if (element.InnerText.ToLower().Contains(str))
                        {
                            string str2 = element.OuterHtml;
                            int index = str2.IndexOf("for=");
                            int num2 = str2.IndexOf(">", index);
                            string str3 = str2.Substring(index + 4, (num2 - index) - 4);
                            HtmlElement elementById = this.webBrowser1.Document.GetElementById(str3);
                            if (elementById != null)
                            {
                                elementById.SetAttribute("checked", "checked");
                                flag = true;
                                break;
                            }
                        }
                    }
                    if (!true)
                    {
                        foreach (HtmlElement element3 in this.webBrowser1.Document.GetElementsByTagName("input"))
                        {
                            if (this.amazonOrder.shipping.compareByValue(element3.GetAttribute("value")))
                            {
                                element3.SetAttribute("checked", "checked");
                                break;
                            }
                        }
                    }
                     * */
                }

                foreach (HtmlElement element3 in this.webBrowser1.Document.GetElementsByTagName("input"))
                {
                    string sss = element3.OuterHtml;

                    if (sss.Contains("a-button-text") && sss.Contains("Continue") && sss.Contains("ubmit"))
                    {
                        element3.InvokeMember("Click");
                        break;
                    }
                }

            //    this.webBrowser1.Document.GetElementById("continue").InvokeMember("click");
                this.documentCompleteCallback = new CallbackType(this.paymentInfo);
            }
        }

        private void logout(object sender, WebBrowserDocumentCompletedEventArgs e)
        {
            this.documentCompleteCallback = null;
            this.webBrowser1.Navigate(AmazonLinks.logout);
            this.label1.Text = "logged out";
            base.Visible = false;
            this.appendStatus(" Finished at " + DateTime.Now.ToString(), true);
            base.Dispose();
        }
        private void paymentInfo(object sender, WebBrowserDocumentCompletedEventArgs e)
        {
            if (this.checkTitle(AmazonPageTitles.paymentInfo))
            {
                if (!this.paymentTimer.Enabled)
                {
                    this.paymentWaitTime = this.amazonOrder.paymentWaitTime + (this.paymentTries * 0x7d0);
                    this.paymentTries++;
                    this.paymentTimer.Stop();
                    this.paymentTimer.Start();
                    this.paymentTimer.Interval = TimerConstant.second / 2;
                    this.paymentTimer.Tick += new EventHandler(this.paymentTimer_Tick);
                    Cursor.Current = Cursors.WaitCursor;
                }
                this.documentCompleteCallback = new CallbackType(this.confirmOrder);
            }
            else
            {
                this.processPageByTitle(sender, e);
            }
        }
        private void closed(object sender, FormClosedEventArgs e)
        {
            if (!this.label1.Text.Equals("logged out"))
            {
                this.appendStatus("Aborted at " + DateTime.Now.ToString(), false);
            }
        }
        private void confirmOrder(object sender, WebBrowserDocumentCompletedEventArgs e)
        {
            this.label1.Text = "";
            if (this.checkTitle(AmazonPageTitles.confirmOrder))
            {
                try
                {
                    foreach (AmazonItem item in this.amazonOrder.Items)
                    {
                        if (item.promotion != "")
                        {
                            this.webBrowser1.Document.GetElementById("enterThemOneAtATime").SetAttribute("value", item.promotion);
                            item.promotion = "";
                            this.webBrowser1.Document.GetElementById("applyGiftCertificate").InvokeMember("click");
                            return;
                        }
                    }
                    HtmlElement elementById = this.webBrowser1.Document.GetElementById("placeYourOrder");
                    // if ((elementById != null) && !elementById.GetAttribute("onclick").Equals(string.Empty))
                    // if ((elementById != null) && !elementById.GetAttribute("onclick").Equals(string.Empty))
                    // chang by dayang on 12/15
                    // found on 12/15 no click arrtribute
                    if ((elementById != null))
                    {
                        if (AmazonBotUtility.submitFinalOrder)
                        {
                            elementById.InvokeMember("click");
                            this.documentCompleteCallback = new CallbackType(this.thankYou);
                        }
                        else
                        {
                            MessageBox.Show("Skiping confirmation step");
                            this.documentCompleteCallback = new CallbackType(this.thankYou);
                            this.thankYou(sender, e);
                        }
                    }
                }
                catch
                {
                }
            }
            else
            {
                this.processPageByTitle(sender, e);
            }
        }



        private void paymentTimer_Tick(object sender, EventArgs e)
        {
            HtmlElement elementById = this.webBrowser1.Document.GetElementById("continue-top");
            if (elementById != null)
            {
                this.label1.Text = "Payment information submitted. Please wait for final confirmation.";
                elementById.InvokeMember("click");
                this.paymentTimer.Enabled = false;
                this.paymentTimer.Stop();
                Cursor.Current = Cursors.Default;
            }
            else
            {
                this.label1.Text = "Waiting for credit card info to load up.";
            }
        }

        [DllImport("user32.dll")]
        private static extern void mouse_event(uint dwFlags, uint dx, uint dy, uint dwData, int dwExtraInfo);

        private void thankYou(object sender, WebBrowserDocumentCompletedEventArgs e)
        {
            if (this.checkTitle(AmazonPageTitles.thankYou))
            {
                try
                {
                    HtmlElement elementById = this.webBrowser1.Document.GetElementById("orders-list");
                    string msg = "";
                    msg = elementById.FirstChild.InnerText.Split(new char[] { '\r' })[0];
                    this.appendStatus(msg, false);
                    this.logout(sender, e);
                }
                catch
                {
                }
            }
            else if (this.checkTitle(AmazonPageTitles.confirmOrder))
            {
                try
                {
                    if (AmazonBotParameters.limited)
                    {
                        this.appendStatus("Duplicated order not supported in this version", false);
                        this.logout(sender, e);
                    }
                    else
                    {
                        this.clickForcePlaceOrder();
                    }
                }
                catch
                {
                }
            }
            else
            {
                this.processPageByTitle(sender, e);
            }
        }
        private void clickForcePlaceOrder()
        {
            // changed a lot 
            appendStatus(this.webBrowser1.DocumentText.ToString(), true);
            IHTMLElement2 element2 = (IHTMLElement2)this.webBrowser1.Document.GetElementById("forcePlaceOrder").DomElement;

            if (SiteContainer1.Site2 == null)
            {
                SiteContainer1.Site2 = CallSite<Func<CallSite, object, IHTMLRect>>.Create(Binder.Convert(CSharpBinderFlags.None, typeof(IHTMLRect), typeof(OrderPlacer)));
            }
            object pvarIndex = 0;
            IHTMLRect rect = SiteContainer1.Site2.Target.Invoke(SiteContainer1.Site2, element2.getClientRects().item(ref pvarIndex));

            Random random = new Random(DateTime.Now.Second);
            Rectangle rectangle = this.webBrowser1.RectangleToScreen(new Rectangle(rect.left, rect.top, rect.right - rect.left, rect.bottom - rect.top));
            int num = random.Next(rectangle.Height);
            int num2 = random.Next(rectangle.Width);
            int x = rectangle.Left + num2;
            int y = rectangle.Top + num;
            Cursor.Position = (new Point(x, y));
            IntPtr handle = this.webBrowser1.Handle;
            IntPtr ptr2 = (IntPtr)((y << 0x10) | x);
            base.Activate();
            mouse_event(2, (uint)x, (uint)y, 0, 0);
            mouse_event(4, (uint)x, (uint)y, 0, 0);
        }

        private void processPageByTitle(object sender, WebBrowserDocumentCompletedEventArgs e)
        {
            if (this.checkTitle(AmazonPageTitles.signIn1) || this.checkTitle(AmazonPageTitles.signIn2))
            {
                this.invokeDelagateByTitle(new CallbackType(this.signIn), sender, e);
                string debugMsg = this.debugMsg;
                this.debugMsg = debugMsg + Environment.NewLine + "[" + DateTime.Now.ToString() + "]\t";
                this.debugMsg = this.debugMsg + "Invoking signIn by processPageByTitle";
            }
            else if (this.checkTitle(AmazonPageTitles.shippingAddress))
            {
                this.invokeDelagateByTitle(new CallbackType(this.shippingAddress), sender, e);
                string str2 = this.debugMsg;
                this.debugMsg = str2 + Environment.NewLine + "[" + DateTime.Now.ToString() + "]\t";
                this.debugMsg = this.debugMsg + "Invoking shippingAddress by processPageByTitle";
            }
            else if (this.checkTitle(AmazonPageTitles.shippingOption))
            {
                this.invokeDelagateByTitle(new CallbackType(this.shippingOption), sender, e);
                string str3 = this.debugMsg;
                this.debugMsg = str3 + Environment.NewLine + "[" + DateTime.Now.ToString() + "]\t";
                this.debugMsg = this.debugMsg + "Invoking shippingOption by processPageByTitle";
            }
            else if (this.checkTitle(AmazonPageTitles.paymentInfo))
            {
                this.invokeDelagateByTitle(new CallbackType(this.paymentInfo), sender, e);
                string str4 = this.debugMsg;
                this.debugMsg = str4 + Environment.NewLine + "[" + DateTime.Now.ToString() + "]\t";
                this.debugMsg = this.debugMsg + "Invoking paymentInfo by processPageByTitle";
            }
            else if (this.checkTitle(AmazonPageTitles.confirmOrder))
            {
                this.invokeDelagateByTitle(new CallbackType(this.confirmOrder), sender, e);
                string str5 = this.debugMsg;
                this.debugMsg = str5 + Environment.NewLine + "[" + DateTime.Now.ToString() + "]\t";
                this.debugMsg = this.debugMsg + "Invoking confirmOrder by processPageByTitle";
            }
            else if (this.checkTitle(AmazonPageTitles.thankYou))
            {
                this.invokeDelagateByTitle(new CallbackType(this.thankYou), sender, e);
                string str6 = this.debugMsg;
                this.debugMsg = str6 + Environment.NewLine + "[" + DateTime.Now.ToString() + "]\t";
                this.debugMsg = this.debugMsg + "Invoking thankYou by processPageByTitle";
            }
            else if (this.webBrowser1.Document.GetElementById("placeYourOrder") != null)
            {
                this.invokeDelagateByTitle(new CallbackType(this.confirmOrder), sender, e);
            }
            else
            {
                string str7 = this.debugMsg;
                this.debugMsg = str7 + Environment.NewLine + "[" + DateTime.Now.ToString() + "]\t";
                this.debugMsg = this.debugMsg + "Cannot find handler for the document titled " + this.webBrowser1.Document.Title;
            }
        }

        private void signIn(object sender, WebBrowserDocumentCompletedEventArgs e)
        {
            if (this.webBrowser1.Document.GetElementById("yaSignIn") != null)
            {
                this.submitLoginInfo("ya-homepage-email", "ya-homepage-password", "yaSignIn", "submit");
            }
            else if (this.checkTitle(AmazonPageTitles.signIn2) && (this.webBrowser1.Document.GetElementById("submit") != null))
            {
                this.submitLoginInfo("email", "password", "submit", "click");
            }
            else if (this.checkTitle(AmazonPageTitles.signIn2) && (this.webBrowser1.Document.GetElementById("signin") != null))
            {
                this.submitLoginInfo("email", "password", "signin", "submit");
            }
            else
            {
                this.afterLoginCallback(sender, e);
            }//Unknown:pipelineJSError
        }

        private void submitLoginInfo(string emailTitle, string passwdTitle, string signinTitle, string signinAction)
        {
            if (this.tries >= 3)
            {
                throw new Exception("Error: Failed to log in amazon");
            }
            this.tries++;
            this.webBrowser1.Document.GetElementById(emailTitle).SetAttribute("value", this.amazonOrder.email);
            this.webBrowser1.Document.GetElementById(passwdTitle).SetAttribute("value", this.amazonOrder.passwd);
            this.webBrowser1.Document.GetElementById(signinTitle).InvokeMember(signinAction);
            this.label1.Text = "Log in finished";
            this.documentCompleteCallback = this.afterLoginCallback;
        }


        public void resetTimer()
        {
            this.paymentTimer.Stop();
            this.timer_selfDestruct.Stop();
        }
        private void saveOldCartItems(object sender, WebBrowserDocumentCompletedEventArgs e)
        {
            if (this.checkTitle("Amazon.com Shopping Cart"))
            {
                try
                {
                    HtmlElement elementById = this.webBrowser1.Document.GetElementById("saveForLater.1");
                    if (elementById != null)
                    {
                        elementById.InvokeMember("click");
                    }
                    else
                    {
                        this.loadCart();
                    }
                }
                catch
                {
                    this.loadCart();
                }
            }
        }

        private void webBrowser1_DocumentCompleted(object sender, WebBrowserDocumentCompletedEventArgs e)
        {
            if (webBrowser1.ReadyState != WebBrowserReadyState.Complete)
                return;


            this.oldProcessedbyTitle = null;
            string debugMsg = this.debugMsg;
            this.debugMsg = debugMsg + Environment.NewLine + "[" + DateTime.Now.ToString() + "]\t";
            if (this.documentCompleteCallback != null)
            {
                this.debugMsg = this.debugMsg + this.documentCompleteCallback.Method.ToString();//  .Method.ToString();
            }
            else
            {
                this.debugMsg = this.debugMsg + "null";
            }
            string str2 = this.debugMsg;
            this.debugMsg = str2 + Environment.NewLine + "[" + DateTime.Now.ToString() + "]\t\t" + e.Url.ToString();
            this.debugForm.setText(this.debugMsg);
            // changed by dayang 07/23/2013

            if(documentCompleteCallback.Method.ToString().Contains("payment"))
            {
                AssignWBHDMethod(e, sender);
                
            }
            if (((WebBrowser)sender).Document.Url.Equals(e.Url))
      //      if(true)
            {
                if ((e.Url.Scheme == "https") || (e.Url.Scheme == "http"))
                {
                    if (e.Url.ToString().ToLower().Contains("pkbonline"))
                    {
                        this.textBox1.Text = "";
                    }
                    else
                    {
                        this.textBox1.Text = e.Url.ToString();
                    }
                }
                /*
                try
                {
                    if (this.documentCompleteCallback != null)
                    {
                        int num = 0;
                        if (AmazonBotParameters.limited)
                        {
                            num = this.random.Next(0xbb8, 0x1388);
                        }
                        else
                        {
                            num = this.random.Next(500, 1000);
                        }
                        this.label1.Text = "Next step in " + ((((double)num) / 1000.0)).ToString() + " seconds.";
                        this.randTimer.Interval = num;
                        this.randTimer.Start();
                        while (this.randTimer.Enabled)
                        {
                            Application.DoEvents();
                        }
                        this.documentCompleteCallback(sender, e);
                        string str3 = this.debugMsg;
                        this.debugMsg = str3 + Environment.NewLine + "[" + DateTime.Now.ToString() + "]\t\t\tInvoked";
                    }
                }
                catch (Exception exception)
                {
                    this.label1.Text = exception.Message;
                }*/
                AssignWBHDMethod(e, sender);
            }
        }

        private void AssignWBHDMethod(WebBrowserDocumentCompletedEventArgs e, object sender)
        {
            try
            {
                if (this.documentCompleteCallback != null)
                {
                    int num = 0;
                    if (AmazonBotParameters.limited)
                    {
                        num = this.random.Next(0xbb8, 0x1388);
                    }
                    else
                    {
                        num = this.random.Next(500, 1000);
                    }
                    this.label1.Text = "Next step in " + ((((double)num) / 1000.0)).ToString() + " seconds.";
                    this.randTimer.Interval = num;
                    this.randTimer.Start();
                    while (this.randTimer.Enabled)
                    {
                        Application.DoEvents();
                    }
                    this.documentCompleteCallback(sender, e);
                    string str3 = this.debugMsg;
                    this.debugMsg = str3 + Environment.NewLine + "[" + DateTime.Now.ToString() + "]\t\t\tInvoked";
                }
            }
            catch (Exception exception)
            {
                this.label1.Text = exception.Message;
            }
        }

        private void randTimer_Tick(object sender, EventArgs e)
        {
            this.randTimer.Stop();
        }
        private string exitMessage()
        {
            int num = this.amazonOrder.destructTime / TimerConstant.second;
            return ("Exiting the current order process in " + num.ToString() + " seconds");
        }
        private void timer_selfDestruct_Tick(object sender, EventArgs e)
        {
            if (this.amazonOrder.destructTime > 0)
            {
                this.amazonOrder.destructTime -= this.timer_selfDestruct.Interval;
                this.label_destruction.Text = this.exitMessage();
            }
            else
            {
                this.logout(null, null);
                this.appendStatus("aborted at " + DateTime.Now.ToString(), false);
                this.resetTimer();
                base.Dispose();
            }

        }

        private void paymentTimer_Tick_1(object sender, EventArgs e)
        {
            HtmlElement elementById = this.webBrowser1.Document.GetElementById("continue-top");
            if (elementById != null)
            {
                this.label1.Text = "Payment information submitted. Please wait for final confirmation.";
                elementById.InvokeMember("click");
                this.paymentTimer.Enabled = false;
                this.paymentTimer.Stop();
                Cursor.Current = Cursors.Default;
            }
            else
            {
                this.label1.Text = "Waiting for credit card info to load up.";
            }
        }

        private void webBrowser1_Navigated(object sender, WebBrowserNavigatedEventArgs e)
        {

            HideScriptErrors(this.webBrowser1, true);
        }


        void HideScriptErrors(WebBrowser wb, bool Hide)
        {
            FieldInfo fiComWebBrowser = typeof(WebBrowser).GetField("_axIWebBrowser2", BindingFlags.Instance | BindingFlags.NonPublic);
            if (fiComWebBrowser == null) return;
            object objComWebBrowser = fiComWebBrowser.GetValue(wb);
            if (objComWebBrowser == null) return;
            objComWebBrowser.GetType().InvokeMember("Silent", BindingFlags.SetProperty, null, objComWebBrowser, new object[] { Hide });
        }


    }
}
