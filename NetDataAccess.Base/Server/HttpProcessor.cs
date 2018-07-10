using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Net.Sockets;
using System.Text;
using System.Threading;

namespace NetDataAccess.Base.Server
{

    public class HttpProcessor
    {
        public TcpClient _Socket;
        public HttpServer _Server;

        private Stream _InputStream;
        public StreamWriter _OutputStream;

        public String _HttpMethod;
        public String _HttpUrl;
        public String _HttpProtocolVersionstring;
        public Hashtable _HttpHeaders = new Hashtable();

        private static int MAX_POST_SIZE = 10 * 1024 * 1024; // 10MB
        private const int BUF_SIZE = 4096;

        public HttpProcessor(TcpClient s, HttpServer srv)
        {
            this._Socket = s;
            this._Server = srv;
        }


        private string StreamReadLine(Stream inputStream)
        {
            int next_char;
            string data = "";
            while (true)
            {
                next_char = inputStream.ReadByte();
                if (next_char == '\n') { break; }
                if (next_char == '\r') { continue; }
                if (next_char == -1) { Thread.Sleep(1); continue; };
                data += Convert.ToChar(next_char);
            }
            return data;
        }
        public void Process()
        {
            // we can't use a StreamReader for input, because it buffers up extra data on us inside it's
            // "processed" view of the world, and we want the data raw after the headers
            _InputStream = new BufferedStream(_Socket.GetStream());

            // we probably shouldn't be using a streamwriter for all output from handlers either
            _OutputStream = new StreamWriter(new BufferedStream(_Socket.GetStream()));
            try
            {
                ParseRequest();
                ReadHeaders();
                if (_HttpMethod.Equals("GET"))
                {
                    HandleGetRequest();
                }
                else if (_HttpMethod.Equals("POST"))
                {
                    HandlePostRequest();
                }
            }
            catch (Exception e)
            {
                Console.WriteLine("Exception: " + e.ToString());
                WriteFailure();
            }
            _OutputStream.Flush();
            // bs.Flush(); // flush any remaining output
            _InputStream = null; 
            _OutputStream = null; // bs = null;            
            _Socket.Close();
        }

        public void ParseRequest()
        {
            String request = StreamReadLine(_InputStream);
            string[] tokens = request.Split(' ');
            if (tokens.Length != 3)
            {
                throw new Exception("invalid http request line");
            }
            _HttpMethod = tokens[0].ToUpper();
            _HttpUrl = tokens[1];
            _HttpProtocolVersionstring = tokens[2];

            Console.WriteLine("starting: " + request);
        }

        public void ReadHeaders()
        {
            Console.WriteLine("readHeaders()");
            String line;
            while ((line = StreamReadLine(_InputStream)) != null)
            {
                if (line.Equals(""))
                {
                    Console.WriteLine("got headers");
                    return;
                }

                int separator = line.IndexOf(':');
                if (separator == -1)
                {
                    throw new Exception("invalid http header line: " + line);
                }
                String name = line.Substring(0, separator);
                int pos = separator + 1;
                while ((pos < line.Length) && (line[pos] == ' '))
                {
                    pos++; // strip any spaces
                }

                string value = line.Substring(pos, line.Length - pos);
                Console.WriteLine("header: {0}:{1}", name, value);
                _HttpHeaders[name] = value;
            }
        }

        public void HandleGetRequest()
        {
            _Server.HandleGetRequest(this);
        }

        public void HandlePostRequest()
        {
            // this post data processing just reads everything into a memory stream.
            // this is fine for smallish things, but for large stuff we should really
            // hand an input stream to the request processor. However, the input stream 
            // we hand him needs to let him see the "end of the stream" at this content 
            // length, because otherwise he won't know when he's seen it all! 

            Console.WriteLine("get post data start");
            int content_len = 0;
            MemoryStream ms = new MemoryStream();
            if (this._HttpHeaders.ContainsKey("Content-Length"))
            {
                content_len = Convert.ToInt32(this._HttpHeaders["Content-Length"]);
                if (content_len > MAX_POST_SIZE)
                {
                    throw new Exception(
                        String.Format("POST Content-Length({0}) too big for this simple server",
                          content_len));
                }
                byte[] buf = new byte[BUF_SIZE];
                int to_read = content_len;
                while (to_read > 0)
                {
                    Console.WriteLine("starting Read, to_read={0}", to_read);

                    int numread = this._InputStream.Read(buf, 0, Math.Min(BUF_SIZE, to_read));
                    Console.WriteLine("read finished, numread={0}", numread);
                    if (numread == 0)
                    {
                        if (to_read == 0)
                        {
                            break;
                        }
                        else
                        {
                            throw new Exception("client disconnected during post");
                        }
                    }
                    to_read -= numread;
                    ms.Write(buf, 0, numread);
                }
                ms.Seek(0, SeekOrigin.Begin);
            }
            Console.WriteLine("get post data end");
            _Server.HandlePostRequest(this, new StreamReader(ms));

        }

        public void WriteSuccess()
        {
            _OutputStream.WriteLine("HTTP/1.1 200 OK");
            _OutputStream.WriteLine("Content-Type: text/xml");
            _OutputStream.WriteLine("Connection: close");
            _OutputStream.WriteLine("");
        }

        public void WriteFailure()
        {
            _OutputStream.WriteLine("HTTP/1.1 404 File not found");
            _OutputStream.WriteLine("Connection: close");
            _OutputStream.WriteLine("");
        }
    }

}
