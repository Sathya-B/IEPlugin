using System;
using System.Net;
using SHDocVw;
using mshtml;
using System.Diagnostics;
using Microsoft.Win32;
using System.Runtime.InteropServices;
namespace IEExtension
{
    [
        ComVisible(true),
        InterfaceType(ComInterfaceType.InterfaceIsIUnknown),
        Guid("FC4801A3-2BA9-11CF-A229-00AA003D7352")
    ]
    public interface IObjectWithSite
    {
        [PreserveSig]
        int SetSite([MarshalAs(UnmanagedType.IUnknown)]object site);
        [PreserveSig]
        int GetSite(ref Guid guid, out IntPtr ppvSite);
    }

    [
            ComVisible(true),
            Guid("2159CB25-EF9A-54C1-B43C-E30D1A4A8277"),
            ClassInterface(ClassInterfaceType.None)
    ]
    public class BHO : IObjectWithSite
    {
        private SHDocVw.WebBrowser webBrowser;
        public const string BHO_REGISTRY_KEY_NAME =
               "Software\\Microsoft\\Windows\\" +
               "CurrentVersion\\Explorer\\Browser Helper Objects";
        public int SetSite(object site)
        {
            if (site != null)
            {
                webBrowser = (SHDocVw.WebBrowser)site;
                webBrowser.DocumentComplete +=
                  new DWebBrowserEvents2_DocumentCompleteEventHandler(
                  this.OnDocumentComplete);
            }
            else
            {
                webBrowser.DocumentComplete -=
                  new DWebBrowserEvents2_DocumentCompleteEventHandler(
                  this.OnDocumentComplete);
                webBrowser = null;
            }

            return 0;
        }

        public int GetSite(ref Guid guid, out IntPtr ppvSite)
        {
            IntPtr punk = Marshal.GetIUnknownForObject(webBrowser);
            int hr = Marshal.QueryInterface(punk, ref guid, out ppvSite);
            Marshal.Release(punk);
            return hr;
        }

        public void OnDocumentComplete(object pDisp, ref object URL)
        {
            HTMLDocument document = (HTMLDocument)webBrowser.Document;

            IHTMLElement head = (IHTMLElement)((IHTMLElementCollection)
                                    document.all.tags("head")).item(null, 0);

            IHTMLScriptElement scriptObject =
                (IHTMLScriptElement)document.createElement("script");
            scriptObject.type = @"text/javascript";
            var select = "a[href$= '.pdf']";
            var quote = '"';
            var insertstring = @"var listElem = document.querySelectorAll(" + quote + select + quote + ");";
            scriptObject.text = @"var links = document.links;
                                  " + insertstring + @"
                                  var button = document.createElement('button');
                                  var btntext = document.createTextNode('Open in Revu'); 
                                  button.type = 'button';
                                  button.name = 'button';
                                  var cont = document.getElementById('content');
                                  button.appendChild(btntext);
                                  for(i=0; i <= listElem.length; i++)
                                  {  
                                    button = document.createElement('button');
                                    btntext = document.createTextNode('Open in Revu'); 
                                    button.appendChild(btntext);
                                    button.id = 'button' + i;
                                    button.name = listElem[i];
                                  listElem[i].appendChild(button);
                                  }";

            ((HTMLHeadElement)head).appendChild((IHTMLDOMNode)scriptObject);
            IHTMLElement button = (IHTMLElement)((IHTMLElementCollection)
                                    document.all.tags("button")).item(null, 0);

            DHTMLEventHandler customHandler = new DHTMLEventHandler(document);
            DispHTMLDocument dispDocument = document;
            customHandler.Handler += OnMouseClick;
            dispDocument.onclick = customHandler;
        }

        [ComRegisterFunction]
        public static void RegisterBHO(Type type)
        {
            RegistryKey registryKey =
              Registry.LocalMachine.OpenSubKey(BHO_REGISTRY_KEY_NAME, true);

            if (registryKey == null)
                registryKey = Registry.LocalMachine.CreateSubKey(
                                        BHO_REGISTRY_KEY_NAME);

            string guid = type.GUID.ToString("B");
            RegistryKey ourKey = registryKey.OpenSubKey(guid);

            if (ourKey == null)
            {
                ourKey = registryKey.CreateSubKey(guid);
            }

            ourKey.SetValue("NoExplorer", 1, RegistryValueKind.DWord);

            registryKey.Close();
            ourKey.Close();
        }

        [ComUnregisterFunction]
        public static void UnregisterBHO(Type type)
        {
            RegistryKey registryKey =
              Registry.LocalMachine.OpenSubKey(BHO_REGISTRY_KEY_NAME, true);
            string guid = type.GUID.ToString("B");

            if (registryKey != null)
                registryKey.DeleteSubKey(guid, false);
        }


        // The delegate:
        public delegate void DHTMLEvent(IHTMLEventObj e);
        ///
        /// Generic Event handler for HTML DOM objects.
        /// Handles a basic event object which receives an IHTMLEventObj which
        /// applies to all document events raised.
        ///

        [ComVisible(true)]
        public class DHTMLEventHandler
        {
            public DHTMLEvent Handler;
            private HTMLDocument Document;

            public DHTMLEventHandler(HTMLDocument doc)
            {
                this.Document = doc;
            }

            [DispId(0)]
            public void Call()
            {
                Handler(Document.parentWindow.@event);
            }
        }

        private void OnMouseClick(IHTMLEventObj e)
        {
            if (e.srcElement != null && e.srcElement.tagName.ToLower() == "button")
            {
               var x = e.srcElement.parentElement;

                WebClient webClient = new WebClient();
                webClient.DownloadFile(e.srcElement.getAttribute("name"), (@"D:\Revu\sample.pdf"));
                Process p = new Process();
                Process.Start(@"C:\Program Files\Bluebeam Software\Bluebeam Revu\2017\Revu\Revu.exe", (@"D:\Revu\sample.pdf"));
            }
        }

    }
}
