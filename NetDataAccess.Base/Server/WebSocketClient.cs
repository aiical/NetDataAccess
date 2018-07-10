using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using WebSocketSharp;

namespace NetDataAccess.Base.Server
{
    public class WebSocketClient
    {
        private WebSocket _ClientSocket = null;
        private WebSocket ClientSocket
        {
            get
            {
                return this._ClientSocket;
            }
            set
            {
                this._ClientSocket = value;
            }
        }

        private bool IsOpen
        {
            get
            {
                return this.ClientSocket.ReadyState == WebSocketState.Connecting;
            }
        }

        public void OpenWebSocket(string url)
        {
            WebSocket ws = new WebSocket(url, null);
            ws.OnMessage += ws_OnMessage;
            ws.OnError += ws_OnError;
            ws.OnOpen += ws_OnOpen;
            ws.Connect(); 
            this.ClientSocket = ws;
        }
        void ws_OnOpen(object sender, EventArgs e)
        {
            MessageBox.Show("open");
        }

        void ws_OnError(object sender, ErrorEventArgs e)
        {
            MessageBox.Show(e.Message);
        }

        void ws_OnMessage(object sender, MessageEventArgs e)
        {
            MessageBox.Show(e.Data);
        }
    }
}
