using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Net;
using System.Net.Sockets;
using System.Text;
using System.Threading;

namespace NetDataAccess.TaskService.Server
{
    public abstract class HttpServer
    {
        protected string _IP;
        protected int _Port;
        protected TcpListener _Listener;
        protected bool _IsActive = false;
        private object _ServerLocker = new object();

        public bool IsActive
        {
            get
            {
                return _IsActive;
            }
        }



        public HttpServer(string ip, int port)
        {
            this._IP = ip;
            this._Port = port;
        }

        public void Listen()
        {
            lock (_ServerLocker)
            {
                this._IsActive = true;
                IPAddress ipAddress = IPAddress.Parse(_IP);
                _Listener = new TcpListener(ipAddress, _Port);
                _Listener.Start();
            }
            while (_IsActive)
            {
                lock (_ServerLocker)
                {
                    TcpClient s = _Listener.AcceptTcpClient();
                    HttpProcessor processor = new HttpProcessor(s, this);
                    Thread thread = new Thread(new ThreadStart(processor.Process));
                    thread.Start();
                    Thread.Sleep(1);
                }
            }
        }

        public abstract void HandleGetRequest(HttpProcessor p);
        public abstract void HandlePostRequest(HttpProcessor p, StreamReader inputData);

        public void Start()
        {
            Thread thread = new Thread(new ThreadStart(this.Listen));
            thread.Start();
        }

        public void Stop()
        {
            lock (_ServerLocker)
            {
                _IsActive = false;
                _Listener.Stop();
            }
        }

    }

}
