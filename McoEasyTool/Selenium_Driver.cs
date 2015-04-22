using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using NUnit.Framework;
using OpenQA.Selenium;
using System.Diagnostics;
using OpenQA.Selenium.Firefox;
using OpenQA.Selenium.IE;
using OpenQA.Selenium.Support.UI;
using OpenQA.Selenium.Remote;
using MCO_APP = McoEasyTool.Controllers.McoAppController;

namespace McoEasyTool.Controllers
{
    [TestFixture]
    public class Selenium_Driver
    {
        public IWebDriver driver;

        [SetUp]
        public void Setup(string Navigator)
        {
            if (Navigator == "FIREFOX")
            {
                FirefoxProfile profile = new FirefoxProfile
                {
                    AcceptUntrustedCertificates = true,
                    EnableNativeEvents = true,
                    AlwaysLoadNoFocusLibrary = true
                };
                //profile.SetPreference("general.useragent.override", "Mozilla/4.0 (compatible; MSIE 8.0; Win32)");
                driver = new FirefoxDriver();
            }
            else
            {
                string IE_DRIVER_PATH = HomeController.BATCHES_FOLDER;

                InternetExplorerOptions options = new InternetExplorerOptions
                {
                    IntroduceInstabilityByIgnoringProtectedModeSettings = true,
                    UnexpectedAlertBehavior = InternetExplorerUnexpectedAlertBehavior.Dismiss,
                    IgnoreZoomLevel = true,
                    EnableNativeEvents = true,
                    RequireWindowFocus = false,
                    EnsureCleanSession = true,
                    EnablePersistentHover = true,
                    ElementScrollBehavior = InternetExplorerElementScrollBehavior.Top,
                };
                driver = new InternetExplorerDriver(IE_DRIVER_PATH, options);
            }
        }

        [TearDown]
        public void Teardown(string Navigator)
        {
            try
            {
                WebDriverWait page_loaded = new WebDriverWait(driver, new TimeSpan(0, 0, 5));
                page_loaded.IgnoreExceptionTypes(typeof(UnhandledAlertException));
                page_loaded.IgnoreExceptionTypes(typeof(HttpUnhandledException));
                try
                {
                    page_loaded.Until(d => 1 == 1);
                }
                catch (UnhandledAlertException)
                {
                    driver.SwitchTo().Alert().Dismiss();
                }
                catch
                { }
                var handles = driver.WindowHandles.ToList();
                foreach (string handle in handles)
                {
                    driver.SwitchTo().Window(handle);
                    driver.Close();
                }
            }
            catch { }
            driver.Quit();
            try
            {
                string processName = (Navigator == "FIREFOX") ? "firefox" : "iexplore";
                string CurrentUser = Environment.UserName.Replace("*\\", "");
                Process[] processes = Process.GetProcessesByName(processName);
                foreach (Process process in processes)
                {
                    string ProcessUserSID = McoUtilities.GetProcessOwner(process.Id);
                    if (ProcessUserSID == CurrentUser)
                    {
                        process.Kill();
                    }
                }
            }
            catch { }
        }

        public void WaitForPageLoad(int timeout)
        {
            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(timeout));
            bool answer = wait.Until<bool>(
                d =>
                {
                    try
                    {
                        return ((IJavaScriptExecutor)driver)
                                .ExecuteScript("return document.readyState")
                                .Equals("complete");
                    }
                    catch (UnhandledAlertException)
                    {
                        driver.SwitchTo().Alert().Dismiss();
                    }
                    catch { }
                    return false;
                });
        }

        public IWebElement GetNode(MCO_APP.AppHtmlElementInfo element)
        {
            IWebElement node = null;
            List<IWebElement> frames = new List<IWebElement>();
            try
            {
                frames = driver.FindElements(By.TagName("frame")).ToList();
                frames = frames.Concat(driver.FindElements(By.TagName("iframe"))).ToList();
            }
            catch { }
            string selector = element.GetStrongerSelector();
            try
            {
                switch (selector)
                {
                    case "XPATH": node = driver.FindElement(By.XPath(element.AttrXpath)); break;
                    case "ID": node = driver.FindElement(By.Id(element.AttrId)); break;
                    case "NAME": node = driver.FindElement(By.Name(element.AttrName)); break;
                    case "CLASS": node = driver.FindElement(By.ClassName(element.AttrClass)); break;
                    default: node = driver.FindElement(By.TagName(element.TagName)); break;
                }
            }
            catch { }

            if (node == null && frames.Count > 0)
            {
                //SEARCH WITHIN FRAMES
                foreach (IWebElement frame in frames)
                {
                    string frame_selector =
                        (frame.GetAttribute("xpath") != null && frame.GetAttribute("xpath") != "") ? "xpath" :
                        (frame.GetAttribute("id") != null && frame.GetAttribute("id") != "") ? "id" :
                        (frame.GetAttribute("name") != null && frame.GetAttribute("name") != "") ? "name" : "class";
                    try
                    {
                        switch (selector)
                        {
                            case "XPATH": node =
                                driver.SwitchTo().Frame(frame.GetAttribute(frame_selector)).FindElement(By.XPath(element.AttrXpath)); break;
                            case "ID": node =
                                driver.SwitchTo().Frame(frame.GetAttribute(frame_selector)).FindElement(By.Id(element.AttrId)); break;
                            case "NAME": node =
                                driver.SwitchTo().Frame(frame.GetAttribute(frame_selector)).FindElement(By.Name(element.AttrName)); break;
                            case "CLASS": node =
                                driver.SwitchTo().Frame(frame.GetAttribute(frame_selector)).FindElement(By.ClassName(element.AttrClass)); break;
                            default: node =
                                driver.SwitchTo().Frame(frame.GetAttribute(frame_selector)).FindElement(By.TagName(element.TagName)); break;
                        }
                    }
                    catch { }
                    if (node != null)
                    {
                        break;
                    }
                }
            }
            return node;
        }

        public List<IWebElement> GetAllNodesByAttribute(MCO_APP.AppHtmlElementInfo element, string attr_name, string attr_value)
        {
            List<IWebElement> nodes = new List<IWebElement>();
            List<IWebElement> frames = new List<IWebElement>();
            string selector = element.GetStrongerSelector();
            try
            {
                frames = driver.FindElements(By.TagName("frame")).ToList();
                frames = frames.Concat(driver.FindElements(By.TagName("iframe"))).ToList();
            }
            catch { }

            try
            {
                List<IWebElement> some_nodes = driver.FindElements(By.TagName(element.TagName)).ToList();
                foreach (IWebElement a_node in some_nodes)
                {
                    if (a_node.GetAttribute(attr_name) == attr_value)
                    {
                        nodes.Add(a_node);
                    }
                }
            }
            catch { }

            if (nodes.Count == 0 && frames.Count > 0)
            {
                foreach (IWebElement frame in frames)
                {
                    try
                    {
                        string frame_selector =
                            (frame.GetAttribute("xpath") != null && frame.GetAttribute("xpath") != "") ? "xpath" :
                            (frame.GetAttribute("id") != null && frame.GetAttribute("id") != "") ? "id" :
                            (frame.GetAttribute("name") != null && frame.GetAttribute("name") != "") ? "name" : "class";
                        List<IWebElement> some_nodes = driver.SwitchTo().Frame(frame.GetAttribute(frame_selector)).FindElements(By.TagName(element.TagName)).ToList();
                        foreach (IWebElement a_node in some_nodes)
                        {
                            if (a_node.GetAttribute(attr_name) == attr_value)
                            {
                                nodes.Add(a_node);
                            }
                        }
                    }
                    catch { }
                }
            }
            return nodes;
        }

        public IWebElement GetNodeByAttribute(MCO_APP.AppHtmlElementInfo element, string attr_name, string attr_value)
        {
            List<IWebElement> nodes = GetAllNodesByAttribute(element, attr_name, attr_value);
            if (nodes.Count > 0)
            {
                return nodes.FirstOrDefault();
            }
            return null;
        }

        public IWebElement GetNodeByContent(MCO_APP.AppHtmlElementInfo element)
        {
            List<IWebElement> nodes = new List<IWebElement>();
            List<IWebElement> frames = new List<IWebElement>();
            try
            {
                frames = driver.FindElements(By.TagName("frame")).ToList();
                frames = frames.Concat(driver.FindElements(By.TagName("iframe"))).ToList();
            }
            catch { }
            try
            {
                nodes = driver.FindElements(By.TagName(element.TagName)).ToList();
            }
            catch { }
            if (nodes.Count == 0 && frames.Count > 0)
            {
                foreach (IWebElement frame in frames)
                {
                    string frame_selector = (frame.GetAttribute("id") != null && frame.GetAttribute("id") != "") ? "id" :
                                   (frame.GetAttribute("name") != null && frame.GetAttribute("name") != "") ? "name" : "class";
                    try
                    {
                        nodes = driver.SwitchTo().Frame(frame.GetAttribute(frame_selector)).FindElements(By.TagName(element.TagName)).ToList();
                    }
                    catch { }
                }
            }
            foreach (IWebElement node in nodes)
            {
                switch (element.TagName)
                {
                    case "NONE": break;
                    default:
                        if (node.Text.Trim().ToLower() == element.Value.Trim().ToLower())
                        {
                            return node;
                        }
                        break;
                }
            }
            return null;
        }

        //0: first form
        //1: look for form
        //2: manual post
        public bool SubmitForm(int level, MCO_APP.AppHtmlElementInfo element = null)
        {
            IWebElement form = null;
            switch (level)
            {
                case 0:
                //LOOK FOR FORM FOUND
                case 1:
                    form = GetNode(new MCO_APP.AppHtmlElementInfo("FORM", "", "", "", "", ""));
                    List<IWebElement> submits = GetAllNodesByAttribute(new MCO_APP.AppHtmlElementInfo("INPUT", "", "", "", "", ""), "type", "submit");
                    if (submits.Count != 0)
                    {
                        foreach (IWebElement submit in submits)
                        {
                            try
                            {
                                submit.Click();
                            }
                            catch { }
                        }
                        return true;
                    }
                    else
                    {
                        if (form != null)
                        {
                            form.Submit();
                            return true;
                        }
                        form = GetNode(new MCO_APP.AppHtmlElementInfo("BUTTON", "", "", "", "", ""));
                        if (form != null)
                        {
                            form.Click();
                            return true;
                        }
                    }
                    return false;

                //CLICK OBJECT
                case 2:
                    if (element != null && element.Type == "LOGIN")
                    {
                        IWebElement button = GetNode(element);
                        if (button != null)
                        {
                            button.Click();
                            return true;
                        }
                    }
                    return SubmitForm(0);

                //FIRST FORM FOUND
                default:
                    form = GetNode(new MCO_APP.AppHtmlElementInfo("FORM", "", "", "", "", ""));
                    if (form != null)
                    {
                        form.Submit();
                        return true;
                    }
                    else
                    {
                        form = GetNode(new MCO_APP.AppHtmlElementInfo("BUTTON", "", "", "", "", ""));
                        if (form != null)
                        {
                            form.Click();
                            return true;
                        }
                    }
                    return false;
            }
        }

        //SELENIUM STUFFS

        public string GetJsSelector(MCO_APP.AppHtmlElementInfo element)
        {
            string selector = "";
            if (element.AttrId != null && element.AttrId.Trim() != "")
            {
                selector = "document.getElementById('" + element.AttrId + "')";
            }
            else
            {
                if (element.AttrName != null && element.AttrName.Trim() != "")
                {
                    selector = "document.getElementsByName('" + element.AttrName + "')[0]";
                }
                else
                {
                    if (element.AttrClass != null && element.AttrClass.Trim() != "")
                    {
                        selector = "document.getElementsByClassName('" + element.AttrClass + "')[0]";
                    }
                    else
                    {
                        selector = "document.getElementsByTagName('" + element.TagName + "')[0]";
                    }
                }
            }
            return selector;
        }

        public string GetScriptSelector(MCO_APP.AppHtmlElementInfo element)
        {
            string selector = GetJsSelector(element);
            string script = "var frames = document.getElementsByTagName('frame');" +
                "var iframes = document.getElementsByTagName('iframe');" +
                "var node = null;" +
                "node = " + selector + ";" +
                "if(node == null  && frames.length > 0){" +
                "for(index=0;index<frames.length;index++){" +
                "var node = window.frames[index]." + selector + ";" +
                "if(node != null){break;}}}" +
                "else{if(node == null  && iframes.length > 0){" +
                "for(index=0;index<iframes.length;index++){" +
                "var node = window.iframes[index]." + selector + ";" +
                "if(node != null){break;}}}}" +
                "if(node != null){node.innerHTML='" + element.Value + "';}";
            return script;
        }

        public bool SetJsNodeValue(MCO_APP.AppHtmlElementInfo element)
        {
            IWebElement node = GetNode(element);
            IWebElement body = GetNode(new MCO_APP.AppHtmlElementInfo("BODY", "", "", "", "", ""));
            if (node != null && body != null)
            {
                try
                {
                    string script = GetScriptSelector(element);

                    if (script == null || script.Trim() == "")
                    {
                        return false;
                    }
                    IJavaScriptExecutor js = (IJavaScriptExecutor)driver;
                    js.ExecuteScript(script, body);
                    //node.SendKeys(element.Value);
                    return true;
                }
                catch (Exception exception)
                {
                    return false;
                }
            }
            return false;
        }

        public bool SetNodeValue(MCO_APP.AppHtmlElementInfo element)
        {
            IWebElement node = GetNode(element);
            driver.Manage().Window.Maximize();
            if (node != null)
            {
                switch (element.TagName)
                {
                    case "INPUT":
                        /*while (!String.IsNullOrEmpty(node.GetAttribute("value")))
                        {
                            node.SendKeys(OpenQA.Selenium.Keys.Backspace);
                        }*/
                        node.Clear();
                        node.SendKeys(element.Value);
                        return true;
                    case "SELECT":
                        node.SendKeys(element.Value);
                        return true;
                    default:
                        SetJsNodeValue(element);
                        return true;
                }
            }
            return false;
        }

        public string GetNodeValue(MCO_APP.AppHtmlElementInfo element)
        {
            IWebElement node = GetNode(element);
            if (node != null)
            {
                if (HomeController.APP_SRC_ATTR_TAGS_LIST.Contains(element.TagName))
                {
                    return node.GetAttribute("src");
                }
                if (HomeController.APP_TEXT_ATTR_TAGS_LIST.Contains(element.TagName))
                {
                    return node.Text.Trim();
                }
                if (HomeController.APP_VALUE_ATTR_TAGS_LIST.Contains(element.TagName))
                {
                    return node.GetAttribute("value");
                }
                return node.Text.Trim();
            }
            return null;
        }

        public bool isEqualNodeValue(MCO_APP.AppHtmlElementInfo element)
        {
            string Value = GetNodeValue(element);
            if (HomeController.APP_SRC_ATTR_TAGS_LIST.Contains(element.TagName) && Value.IndexOf(element.Value) != -1)
            {
                return true;
            }
            if (HomeController.APP_TEXT_ATTR_TAGS_LIST.Contains(element.TagName) && Value.IndexOf(element.Value) != -1)
            {
                return true;
            }
            if (HomeController.APP_VALUE_ATTR_TAGS_LIST.Contains(element.TagName) && Value == element.Value)
            {
                return true;
            }
            if (Value == element.Value)
            {
                return true;
            }
            return false;
        }

        //END SELENIUM STUFFS
    }
    //END SELENIUM DRIVER
    //END VIRTUALISED METHODS
}