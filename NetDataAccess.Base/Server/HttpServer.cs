using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Net;
using System.Net.Sockets;
using System.Text;
using System.Threading;

namespace NetDataAccess.Base.Server
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
                TcpClient s = _Listener.AcceptTcpClient();
                HttpProcessor processor = new HttpProcessor(s, this);
                Thread thread = new Thread(new ThreadStart(processor.Process));
                thread.Start();
                Thread.Sleep(1);
            }

            //在此处关闭监听端口，看看是不是能解决端口不释放的问题  added by lixin 20170720
            _Listener.Stop();
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
            _IsActive = false;
            //等待半秒钟，让监听的端口释放掉  added by lixin 20170720
            Thread.Sleep(500);
        }

    }

}
