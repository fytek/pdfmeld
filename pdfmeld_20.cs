// pdfmeld_20.cs - .NET DLL wrapper for FyTek's PDF Meld

// Call this DLL from the language of your choice as long as it supports
// a COM or .NET DLL.  May be 32 or 64-bit.
// This DLL accepts parameters and then builds the PDF which you may save
// locally, on the box the server is running on (if it's different box) or
// have the PDF returned as a byte array for display on website or for saving
// in a database.
// This DLL calls the PDF Meld executable (32 or 64-bit) or uses sockets
// when PDF Meld is running as a server.  It sends the parameter settings
// made here to build the PDF.  For exapmle, you might want to startup PDF Meld
// with a pool of 5 connections on a Linux box and call it from this DLL on a
// Windows box.  Even if you run PDF Meld on the same box it's recommended
// to start a PDF Meld server in order to keep resource usage in check.
// Note you may start more than one PDF Meld server at a time with different
// port numbers for each one.
// Use startServer to start up a PDF Meld server and stopServer to shut it down.
// You probably want to do that outside of your main routine that will be building
// PDFs as your main routine will link to this DLL to call the already running
// service.


using System.Diagnostics;
using System.Collections.Generic;
using System.IO;
using System;
using System.Net;
using System.Net.Sockets;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Security.Cryptography;
using System.Globalization;

// Compiling this code:
// Microsoft.Net.Compilers.3.4.0\tools\csc /target:library /platform:anycpu /out:pdfmeld_20.dll pdfmeld_20.cs /keyfile:mykey.snk
// C:\Windows\Microsoft.NET\Framework64\v4.0.30319\regasm.exe /codebase c:pdfmeld_20.dll
// cscript.exe (or wscript.exe) rwtest.vbs

namespace FyTek
{
    [ComVisible(true)]
    [Guid("636C1BC1-7CAC-4E43-95AA-D9D716B165B2")]
	[ProgId("FyTek.PDFMeld")]    
    public class PDFMeld : IDisposable
    {
        void IDisposable.Dispose()
        {

        }
        private TcpClient client = new TcpClient();
        private NetworkStream stream;
        private String exe = "pdfmeld64"; // the executable - change with setExe
        private const String srvHost = "localhost";
        private const int srvPort = 7070;
        private const int srvPool = 5;
        private static String srvFile = ""; // the file of servers and ports
        private static int srvNum = 0; // the array index for the next server to use
        private bool useAvailSrv = false; // true when choosing the next available server
        private Dictionary<string,string> opts = new Dictionary<string,string>(); // all of the parameter settings from the method calls
        private Dictionary<string,object> server = new Dictionary<string,object>(); // the server host/port/log file key/values

        private String units = "";
        private Double unitsMult = 1;

        [ComVisible(true)]
        public class Results {
            public byte[] Bytes {get; set;}
            public String Msg {get; set;}
            public String Result {get; set;}
        }        

        private class Server {
            public String Host {get; set;}
            public int Port {get; set;}
            public Server(String host, int port){
                this.Host = host;
                this.Port = port;
            }
        }

        private static List<Server> servers = new List<Server>();

        // Start up PDF Meld as a server
        [ComVisible(true)]
        public String setServerFile(String fileName){
            srvFile = fileName;
            String line = "";
            String[] retCmds = new String[2]; 
            Regex r = new Regex("[\\s\\t]+");
            int port;
            srvNum = 0;
            try {
                System.IO.StreamReader file =   
                new System.IO.StreamReader(fileName);  
                servers = new List<Server>();
                while((line = file.ReadLine()) != null)  
                {  
                    if (!line.Trim().StartsWith("#")
                    && !line.Trim().Equals("")){
                        retCmds = r.Split(line.Trim());
                        if (retCmds[0].Equals("exe")){
                            setExe(retCmds[1]); // passing location of exe instead of a host/port
                        } else {
                            int.TryParse(retCmds[1],out port);
                            if (port > 0)
                                servers.Add(new Server(retCmds[0],port));
                        }
                    }
                }  
                file.Close();                              
            } catch (IOException e) {
                return e.Message;
            }
            return "";
        }

        // Start up PDF Meld as a server
        [ComVisible(true)]
        public String startServer(
            String host = srvHost,
            int port = srvPort,
            int pool = srvPool,
            String log = ""
          )
       {
            byte[] bytes = {};
            String errMsg = "";            
            server["host"] = host;
            server["port"] = port;
            server["pool"] = pool;
            server["log"] = log; // file on server to log the output
            ProcessStartInfo startInfo = new ProcessStartInfo();
            startInfo.CreateNoWindow = false;
            startInfo.UseShellExecute = true;
            startInfo.FileName = exe;
            startInfo.WindowStyle = ProcessWindowStyle.Hidden;
            String cmdsOut = "";
            String s = "";

            if (opts.TryGetValue("licInfo_name", out s)){
                String p = "";
                String d = "";
                opts.TryGetValue("licInfo_name", out p);
                opts.TryGetValue("licInfo_autodl", out d);
                cmdsOut += " -licname " + s + " -licpwd " + p + (d.Equals("1") ? " -licweb" : "");
            } else {
                if (!opts.TryGetValue("keyCode", out s)){
                    setKeyName("demo");
                }
            }
            if (opts.TryGetValue("keyCode", out s))
                cmdsOut += " -kc " + s;
            if (opts.TryGetValue("keyName", out s))
                cmdsOut += " -kn " + s;

            cmdsOut += (!log.Equals("") ? " -log " + '"' + log + '"' : "")
                + " -port " + port + " -pool " + pool + " -host " + host;
            startInfo.Arguments = "-server " + cmdsOut;            
            (bytes, errMsg) = runProcess(startInfo, false);
            opts.Remove("licInfo_name");
            opts.Remove("licInfo_pwd");
            opts.Remove("licInfo_autodl");
            opts.Remove("keyName");
            opts.Remove("keyCode");
            return errMsg;
        }

        [ComVisible(true)]
        public void setServer(
            String host = srvHost,
            int port = srvPort
          )
       {
            server["host"] = host;
            server["port"] = port;
       }

        // Stop the server
        [ComVisible(true)]
        public String stopServer(){
            byte[] bytes;
            String msg;
            String res;
            setOpt("serverCmd","-quit");
            (bytes, msg, res) = callTCP(isStop: true);
            return msg;
        }

        // Get the current stats for the server
        [ComVisible(true)]
        public String serverStatus(bool allServers = false){
            byte[] bytes;
            String msg = "";
            String sMsg = "";
            String res = "";
            if (!allServers || servers.Count == 0){
                if (!isServerRunning()){
                    return "Server " + server["host"] + " is not responding on port " + server["port"] + ".";
                }
                setOpt("serverCmd","-serverstat");            
                (bytes, msg, res) = callTCP(isStatus: true);            
            } else {
                foreach(var item in servers){
                    server["host"] = item.Host;
                    server["port"] = item.Port;
                    setOpt("serverCmd","-serverstat");            
                    (bytes, sMsg, res) = callTCP(isStatus: true);
                    msg += sMsg + "\n";
                }
            }
            return msg;
        }

        // Get the current stats for the server
        [ComVisible(true)]
        public int serverThreads(){
            byte[] bytes;
            String msg;
            String res;
            int threadsAvail = 0;
            if (!isServerRunning()){
                return 0;
            }
            setOpt("serverCmd","-threadsavail");
            (bytes, msg, res) = callTCP(isStatus: true);
            int.TryParse(msg, out threadsAvail);
            return threadsAvail;
        }

        // Stop a process
        [ComVisible(true)]
        public String serverCancelId(int id){
            byte[] bytes;
            String msg;
            String res;
            setOpt("stopId","-stopid " + id);
            (bytes, msg, res) = callTCP();
            return msg;
        }

        // Check server
        [ComVisible(true)]
        public bool isServerRunning(){
            Object host = "";
            int tryCount = 0;
            bool srvRunning = false;
            if (client.Connected){
                return true;
            }            
            if (!server.TryGetValue("host", out host) && servers.Count == 0){
                setServer();
            }
            if (host == null && servers.Count > 0){
                // Loop through list of servers and get the next one that gets connected
                while(tryCount < servers.Count){
                    srvNum++;
                    srvNum %= servers.Count;
                    useAvailSrv = true;
                    srvRunning = true;
                    setServer(servers[srvNum].Host, servers[srvNum].Port);
                    if (client.Connected){
                        // if previous connection, disconnect it
                        try {
                            stream.Close();
                            client.Close();                        
                        } catch (SocketException) { }
                    }
                    try { 
                        client = new TcpClient((String) server["host"], (int) server["port"]);                        
                        stream = client.GetStream();
                        Socket s = client.Client;
                        srvRunning = !((s.Poll(1000, SelectMode.SelectRead) && (s.Available == 0)) || !s.Connected);
                        return true;
                    } catch (SocketException) {
                        srvRunning = false;
                        stream.Close();
                        client.Close();                        
                    }
                    tryCount++;
                }
            }            
            if (!client.Connected && server.TryGetValue("host", out host)){
                try {
                    client = new TcpClient((String) server["host"], (int) server["port"]);
                    stream = client.GetStream();                    
                    Socket s = client.Client;
                    return !((s.Poll(1000, SelectMode.SelectRead) && (s.Available == 0)) || !s.Connected);
                } catch (SocketException) {
                    return false;
                }
            }
            return false;
        }

        // Build the PDF using the server, optionally return the 
        // raw bytes of the PDF for further processing when retBytes = true
        // optional file name as well if saveFile is passed - this allows
        // for saving file on this box if server is running on different box
        private object buildPDFTCP(bool retBytes = false, String saveFile = ""){            
            byte[] bytes = {};
            String errMsg = "";
            String buildResult = "";
            String tempFile = "";
            String s = "";

           if (opts.TryGetValue("datafield", out s)){
             tempFile = $@"{Guid.NewGuid()}.dat"; // come up with unique file name
             byte[] b = System.Text.Encoding.UTF8.GetBytes(s);
             sendFileTCP(tempFile, "", b);
             setInFile(tempFile);
           }

            if (!saveFile.Equals("") || retBytes){
                if (!opts.TryGetValue("outFile", out s)){
                    setOutFile("membuild");
                }
            }

            (bytes, errMsg, buildResult) = callTCP(retBytes: retBytes, saveFile: saveFile);
            Results ret = new Results();
            ret.Bytes = bytes;
            ret.Msg = errMsg.Equals("") ? "OK" : errMsg;
            ret.Result = buildResult.Trim();
            if (useAvailSrv){
                server.Clear();
                useAvailSrv = false;
            }
            return ret;
        }

        // Send all files to server - only necessary if server is on a different box
        [ComVisible(true)]
        public void setAutoSendFiles(){
          setOpt("autoSend","Y");
        }

        // Pass file contents in memory for input type files
        [ComVisible(true)]
        public String setFileDataB(String fileName, byte[] bytes)
        {
            setCmdlineOpts("-data " + fileName);
            sendFileTCP(fileName, "", bytes);
            return fileName;
        }

        // Pass file contents in memory for input type files
        [ComVisible(true)]
        public String setFileDataH(String fileName, String hexString)
        {
            setCmdlineOpts("-data " + fileName);
            sendFileTCP(fileName, "", ConvertHexToByteArray(hexString));
            return fileName;
        }

        // Pass file contents in memory for input type files
        [ComVisible(true)]
        public String setFileDataS(String fileName, String fileContents)
        {
            setCmdlineOpts("-data " + fileName);
            sendFileTCP(fileName, "", System.Text.Encoding.UTF8.GetBytes(fileContents));
            return fileName;
        }

        // Assign the key name
        [ComVisible(true)]
        public String setKeyName(String a)
        {
            if (a.ToLower().Equals("demo")){
                // Get the demo key from website - this only works with the demo pdfmeld executable
                WebClient wClient = new WebClient();
                string res = wClient.DownloadString("http://www.fytek.com/cgi-bin/genkeyw_v2.cgi?prod=meld");
                Regex regex = new Regex("-kc [A-Z0-9]*");
                Match match = regex.Match(res);                
                setOpt("keyName","testkey");
                setOpt("keyCode",match.Value.Substring(4));
            } else {
                setOpt("keyName",a);
            }
            return a;
        }

        // Assign the key code
        [ComVisible(true)]
        public String setKeyCode(String a)
        {
            setOpt("keyCode",a);
            return a;
        }

        // License settings
        [ComVisible(true)]
        public void licInfo(String licName,
            String licPwd,
            int autoDownload)
        {
            setOpt("licInfo_name",licName);
            setOpt("licInfo_pwd",licPwd);
            setOpt("licInfo_autodl",$"{autoDownload}");
        }

        // Assign the user
        [ComVisible(true)]
        public String setUser(String a)
        {
            setOpt("user",a);
            return a;
        }

        // Assign the owner
        [ComVisible(true)]
        public String setOwner(String a)
        {
            setOpt("owner",a);
            return a;
        }

        // Set no annote
        [ComVisible(true)]
        public void setNoAnnote()
        {
            setOpt("noAnnote","Y");
        }

        // Set no copy
        [ComVisible(true)]
        public void setNoCopy()
        {
            setOpt("noCopy","Y");
        }

        // Set no change
        [ComVisible(true)]
        public void setNoChange()
        {
            setOpt("noChange","Y");
        }

        // Set no print
        [ComVisible(true)]
        public void setNoPrint()
        {
            setOpt("noPrint","Y");
        }

        // Set the GUI process window off
        [ComVisible(true)]
        public void setGUIOff()
        {
            setOpt("guiOff","Y");
        }

        // Legacy support
        [ComVisible(true)]
        public String setCmds(String a)
        {
            return setCmdlineOpts(a);
        }

        // Set any other command line type options
        [ComVisible(true)]
        public String setCmdlineOpts(String a)
        {
            String s = "";
            if (opts.TryGetValue("extOpts", out s))
                a = s + " " + a;
            setOpt("extOpts",a);
            return a;
        }

        // Assign the executable
        [ComVisible(true)]
        public String setExe(String a)
        {
            exe = a;
            return a;
        }

        // Assign the input file name
        [ComVisible(true)]
        public String setInFile(String a)
        {
            String s = "";
            a = a.Replace(",","\\,");
            if (opts.TryGetValue("inFile", out s))
                a = s + "," + a;
            setOpt("inFile",a);
            return a;
        }

        // Assign the output file name
        [ComVisible(true)]
        public String setOutFile(String a)
        {
            setOpt("outFile",a);
            return a;
        }

        // Compression 1.5
        [ComVisible(true)]
        public void setOptimize(bool compress = true)
        {
            setOpt("opt",compress ? "Y" : "N");
        }

        // Compression 1.5
        [ComVisible(true)]
        public void setComp15()
        {
            setOpt("comp15","Y");
        }

        // Encrypt 128
        [ComVisible(true)]
        public void setEncrypt128()
        {
            setOpt("enc128","Y");
        }

        // AES 128
        [ComVisible(true)]
        public String setEncryptAES(String a)
        {
            setOpt("aes",a);
            return a;
        }

        // overwrite existing
        [ComVisible(true)]
        public void setForce()
        {
            setOpt("force","Y");
        }

        // open output
        [ComVisible(true)]
        public void setOpen()
        {
            setOpt("open","Y");
        }        

        // print output
        [ComVisible(true)]
        public void setPrint()
        {
            setOpt("print","Y");
        }        

        // buildlog file
        [ComVisible(true)]
        public String setBuildLog(String fileName)
        {
            setOpt("buildLog",fileName);
            return fileName;
        }      

        // debug
        [ComVisible(true)]
        public String setDebug(String fileName)
        {
            setOpt("debug",fileName);
            return fileName;
        }        

        // errFile
        [ComVisible(true)]
        public String setErrFile(String fileName)
        {
            setOpt("errFile",fileName);
            return fileName;
        }        

        // subject
        [ComVisible(true)]
        public String setSubject(String a)
        {
            setOpt("subject",a);
            return a;
        }        

        // author
        [ComVisible(true)]
        public String setAuthor(String a)
        {
            setOpt("author",a);
            return a;
        }        

        // keywords
        [ComVisible(true)]
        public String setKeywords(String a)
        {
            setOpt("keywords",a);
            return a;
        }        

        // creator
        [ComVisible(true)]
        public String setCreator(String a)
        {
            setOpt("creator",a);
            return a;
        }        

        // producer
        [ComVisible(true)]
        public String setProducer(String a)
        {
            setOpt("producer",a);
            return a;
        }      

        [ComVisible(true)]
        public void setOverlay()
        {
            setOpt("overlay","Y");
        }      

        [ComVisible(true)]
        public void setRevOverlay()
        {
            setOpt("revoverlay","Y");
        }      

        [ComVisible(true)]
        public void setOverlay2()
        {
            setOpt("overlay2","Y");
        }      

        [ComVisible(true)]
        public void setRevOverlay2()
        {
            setOpt("revoverlay2","Y");
        }      

        [ComVisible(true)]
        public void setRepeat()
        {
            setOpt("repeat","Y");
        }      

        [ComVisible(true)]
        public void setRepeatLast()
        {
            setOpt("repeatlast","Y");
        }      
        
        [ComVisible(true)]
        public String setPages(String a, bool exclude = false)
        {
            setOpt("pages",a);
            if (exclude){
              setOpt("exclude",a);
            }
            return a;
        }      

        [ComVisible(true)]
        public void setAutoRotate()
        {
            setOpt("autorotate","Y");
        }      

        [ComVisible(true)]
        public String setFieldsFlatten(String a = "")
        {
            setOpt("flatten",a);
            return a;
        }      

        [ComVisible(true)]
        public String setSubset(String a = "")
        {
            setOpt("subset",a);
            return a;
        }      

        [ComVisible(true)]
        public String setCreationDate(String a = "")
        {
            setOpt("creationdate",a);
            return a;
        }      

        [ComVisible(true)]
        public String setModDate(String a = "")
        {
            setOpt("moddate",a);
            return a;
        }      

        [ComVisible(true)]
        public String setUnits(String a)
        {
          units = a.ToLower();
          unitsMult = 1;
          switch (units) {
            case "cm":
              unitsMult = 72 / 2.54;
              break;
            case "mm":
              unitsMult = 72 / 25.4;
              break;
            case "in":
              unitsMult = 72;
              break;
            default:
              unitsMult = 1;
              break;
          }
          setOpt("units",a);
          return a;
        }      

        [ComVisible(true)]
        public String setPageSize(Double width, Double height)
        {
          String a = width * unitsMult + "," + height * unitsMult;
          setOpt("pagesize",a);
          return a;
        }      

        [ComVisible(true)]
        public void setPageNum()
        {
            setOpt("pagenum","Y");
        }      

        [ComVisible(true)]
        public Double setFontSize(Double a)
        {
          setOpt("fontsize",a.ToString());
          return a;
        }      

        [ComVisible(true)]
        public void setPageTop()
        {
            setOpt("pagetop","Y");
        }      

        [ComVisible(true)]
        public String setPageFmt(String a)
        {
          if (!a.Equals("")){
            setOpt("pagefmt",a);
          }
          return a;
        }      

        [ComVisible(true)]
        public String setPageNumAlign(String a)
        {
          if (!a.Equals("")){
            setOpt("pagenumalign",a.ToLower());
          }
          return a;
        }      

        [ComVisible(true)]
        public Double setPageLMargin(Double a)
        {
          setOpt("pagelmargin",(a * unitsMult).ToString());
          return a;
        }      

        [ComVisible(true)]
        public Double setPageBMargin(Double a)
        {
          setOpt("pagebmargin",(a * unitsMult).ToString());
          return a;
        }       

        // optional, pass "nofit" to remove zoom level on bookmark
        [ComVisible(true)]        
        public String setBookmarkKeep(String a = "")
        {          
            setOpt("bmkeep",a);
            return a;
        }      

        [ComVisible(true)]
        public void setNoBookmarks()
        {
            setOpt("nobm","Y");
        }      

        [ComVisible(true)]        
        public String setBookmarkFile(String a)
        {          
            setOpt("bm",a);
            return a;
        }      

        [ComVisible(true)]
        public void setSkipFieldRename()
        {
            setOpt("skipfieldrename","Y");
        }      

        [ComVisible(true)]        
        public String setZoom(String a)
        {          
            setOpt("zoom",a);
            return a;
        }      

        [ComVisible(true)]        
        public String setDataPDF(String fileName, String options)
        {   
            String s = "";
            opts.TryGetValue("datafield", out s);
            fileName = fileName.Replace("\"","\\\"");
            s += $"<PDF SRC=\"{fileName}\" {options}>\n";
            setOpt("datafield",s);
            setOpt("data","Y");
            return fileName;
        }              

        [ComVisible(true)]        
        public String setDataField(String field, String value)
        {   
            String s = "";
            opts.TryGetValue("datafield", out s);
            field = field.Replace("\"","\\\"");
            value = value.Replace("\"","\\\"");
            s += $"<FDFFIELD NAME=\"{field}\" VALUE=\"{value}\">\n";
            setOpt("datafield",s);
            setOpt("data","Y");
            return value;
        }        

        [ComVisible(true)]        
        public String setPageOrder(String a)
        {          
            setOpt("pageord",a);
            return a;
        }      

        [ComVisible(true)]        
        public void setData()
        {   
            setOpt("data","Y");
        }              

        [ComVisible(true)]        
        public void setPageCenter()
        {   
            setOpt("center","Y");
        }     

        [ComVisible(true)]
        public Double setPageScale(Double x, Double y = 0)
        {
          if (y == 0){
            setOpt("scale",x.ToString());
          } else {
            setOpt("scalex",x.ToString());
            setOpt("scaley",y.ToString());
          }
          return x;
        }       

        [ComVisible(true)]
        public Double setPageRight(Double x)
        {
          setOpt("right",(x * unitsMult).ToString());          
          return x;
        }       

        [ComVisible(true)]
        public Double setPageDown(Double x)
        {
          setOpt("down",(x * unitsMult).ToString());          
          return x;
        }  

        [ComVisible(true)]        
        public void setPageNumReset()
        {          
            setOpt("pagenumreset","Y");
        }      

        [ComVisible(true)]        
        public String setStrFileIn(String a)
        {          
            setOpt("strin",a);
            return a;
        }      

        [ComVisible(true)]        
        public String setFileSep(String a)
        {          
            setOpt("filesep",a);
            return a;
        }      

        [ComVisible(true)]        
        public String setPwdList(String a)
        {          
            setOpt("pwdlist",a);
            return a;
        }      

        [ComVisible(true)]        
        public void setBookmarkTitle()
        {          
            setOpt("bmtitle","Y");
        }      

        [ComVisible(true)]        
        public String setLogFile(String a)
        {          
            setOpt("log",a);
            return a;
        }      

        [ComVisible(true)]        
        public String setPageNumFont(String a)
        {          
            setOpt("pagenumfont",a);
            return a;
        }      

        [ComVisible(true)]        
        public String setPageNumColor(String a)
        {          
            setOpt("pagenumcolor",a);
            return a;
        }      

        [ComVisible(true)]        
        public String setPageNumBGColor(String a)
        {          
            setOpt("pagenumbgcolor",a);
            return a;
        }      

        [ComVisible(true)]        
        public String setAutoScale(String a)
        {          
            setOpt("autoscale",a);
            return a;
        }      

        [ComVisible(true)]        
        public void setAutoClip()
        {          
            setOpt("autoclip","Y");
        }      

        [ComVisible(true)]        
        public void setFieldsRemove()
        {          
            setOpt("removeflds","Y");
        }      

        [ComVisible(true)]        
        public void setNoExtract()
        {          
            setOpt("noextract","Y");
        }      

        [ComVisible(true)]        
        public void setNoFillIn()
        {          
            setOpt("nofillin","Y");
        }      

        [ComVisible(true)]        
        public void setNoAssemble()
        {          
            setOpt("noassemble","Y");
        }      

        [ComVisible(true)]        
        public void setNoDigital()
        {          
            setOpt("nodigital","Y");
        }     

        // Calls buildPDF or buildPDFTCP
        [ComVisible(true)]
        public object buildPDF(bool waitForExit = true,
            String saveFile = "")    
        {   
            object host;
            if (server.TryGetValue("host", out host)
                || servers.Count > 0) {
                // if there is a server or servers, build using TCP                 
                // waitForExit means return the byte array of the PDF
                return buildPDFTCP(waitForExit, saveFile);
            } else {
                // otherwise, build using the executable
                if (!saveFile.Equals("")){
                    setOutFile(saveFile); // shorthand for calling setOutFile
                }
                return build(waitForExit);
            }
        }

        // Call the executable (non server mode)
        private object build(bool waitForExit = true)
        {    
            byte[] bytes = {};
            String errMsg = "";
            ProcessStartInfo startInfo = new ProcessStartInfo();
            startInfo.CreateNoWindow = false;
            startInfo.UseShellExecute = !waitForExit;
            if (waitForExit){
                startInfo.RedirectStandardOutput = true;
                startInfo.RedirectStandardInput = true;
            }
            startInfo.FileName = exe;
            startInfo.WindowStyle = ProcessWindowStyle.Hidden;
            startInfo.Arguments = setBaseOpts();  
            (bytes, errMsg) = runProcess(startInfo,waitForExit);
            Results ret = new Results();
            ret.Bytes = bytes;
            ret.Msg = errMsg.Equals("") ? "OK" : errMsg;
            return ret;
        }

        // Reset the options
        [ComVisible(true)]
        public void resetOpts(bool resetServer = false){
            opts.Clear();
            unitsMult = 1;
            if (resetServer){
                server.Clear();
            }
            
        }

        // Passes file to server over socket
        public String sendFileTCP(String fileName,
            String filePath = "",
            byte[] bytes = null){
            String errMsg = "";
            Byte[] data;
            Object host = "";
            // String to store the response ASCII representation.
            String responseData = String.Empty;
            String message = "";
            if (!isServerRunning()){
                return "Server not running";
            }
            try {                         
                // Send a file
                byte[] buffer = new byte[1024];
                int bytesRead = 0;

                message = " -send --binaryname--" + fileName + "--binarybegin--";
                data = System.Text.Encoding.UTF8.GetBytes(message);   
                stream.Write(data, 0, data.Length); 
                if (bytes != null && bytes.Length > 0){
                    stream.Write(bytes, 0, bytes.Length);                    
                }   
                else            
                if (!filePath.Equals("")){
                    BinaryReader br;
                    br = new BinaryReader(new FileStream(filePath, FileMode.Open));

                    while ((bytesRead = br.Read(buffer, 0, buffer.Length)) > 0)
                        stream.Write(buffer, 0, bytesRead);                    

                }
                message = "--binaryend-- ";
                data = System.Text.Encoding.UTF8.GetBytes(message);             
                stream.Write(data, 0, data.Length);                           

            } catch (SocketException e) {
                errMsg = e.Message;
            } catch (IOException e) {
                errMsg = e.Message;
            }
            return (errMsg);
        }

        private void setOpt(String k, String v){
            opts[k] = v;
        }

        private (byte[], String) runProcess(ProcessStartInfo startInfo, bool waitForExit) {
            // Start the process with the info we specified.
            // Call WaitForExit if we are waiting for process to complete. 
            byte[] bytes = {};
            try {              
                using (Process exeProcess = Process.Start(startInfo))
                {
                    MemoryStream memstream = new MemoryStream();
                    byte[] buffer = new byte[1024];
                    int bytesRead = 0;
                    if (waitForExit){
                        BinaryReader br = new BinaryReader(exeProcess.StandardOutput.BaseStream);
                        while ((bytesRead = br.Read(buffer, 0, buffer.Length)) > 0)
                            memstream.Write(buffer, 0, bytesRead);                    
                    }                        
                    if (waitForExit){
                        exeProcess.WaitForExit();
                        bytes = memstream.ToArray();
                    }
                }
            }
            catch (Exception e){
                return (bytes, e.Message);
            }
            return (bytes, "");

        }

        // Passes data to server over socket but does not finalize 
        // (that is, does not send BUILDPDF command)
        private String sendTCP(){
            String errMsg = "";
            Byte[] data;
            Object host = "";
            // String to store the response ASCII representation.
            String responseData = String.Empty;
            String message = "";

            try {         
                
                String s;
                if (opts.TryGetValue("serverCmd", out s)){
                    message = s;
                    opts.Remove("serverCmd");
                } else {
                    message = setBaseOpts();
                }
                // Send commands
                data = System.Text.Encoding.UTF8.GetBytes(message);             
                stream.Write(data, 0, data.Length);      

            } catch (SocketException e) {
                errMsg = e.Message;
            } catch (IOException e) {
                errMsg = e.Message;
            }
            return (errMsg);
        }

        // build the command line string to pass to the executable
        private String setBaseOpts(){
            String s = "";
            String message = "";            
            if (opts.TryGetValue("inFile", out s))
                message += " \"" + s + "\"";
            if (opts.TryGetValue("outFile", out s)) {
                if (!s.Equals(""))
                    message += " \"" + s + "\"";
                }
            if (opts.TryGetValue("data", out s))
                message += " -data";
            if (opts.TryGetValue("keyCode", out s))
                message += " -kc " + s;
            if (opts.TryGetValue("keyName", out s))
                message += " -kn " + s;
            if (opts.TryGetValue("guiOff", out s))
                message += " -guioff";
            if (opts.TryGetValue("comp15", out s))
                message += " -comp15";
            if (opts.TryGetValue("opt", out s))
                message += " -opt" + (s.Equals("Y") ? "15" : "");
            if (opts.TryGetValue("owner", out s))
                message += " -o \"" + s + "\"";
            if (opts.TryGetValue("user", out s))
                message += " -u \"" + s + "\"";
            if (opts.TryGetValue("noAnnote", out s))
                message += " -noannote";
            if (opts.TryGetValue("noCopy", out s))
                message += " -nocopy";
            if (opts.TryGetValue("noPrint", out s))
                message += " -noprint";
            if (opts.TryGetValue("noChange", out s))
                message += " -nochange";
            if (opts.TryGetValue("nofillin", out s))
                message += " -nofillin";
            if (opts.TryGetValue("noassemble", out s))
                message += " -noassemble";
            if (opts.TryGetValue("nodigital", out s))
                message += " -nodigital";
            if (opts.TryGetValue("noextract", out s))
                message += " -noextract";
            if (opts.TryGetValue("enc128", out s))
                message += " -e128";
            if (opts.TryGetValue("aes", out s))
                message += " -aes " + s;
            if (opts.TryGetValue("extOpts", out s))
                message += " " + s;
            if (opts.TryGetValue("force", out s))
                message += " -force";            
            if (opts.TryGetValue("open", out s))
                message += " -open";            
            if (opts.TryGetValue("print", out s))
                message += " -print";            
            if (opts.TryGetValue("autoSend", out s))
                message += " -autosend";                           
            if (opts.TryGetValue("buildLog", out s))            
                message += " -buildlog \"" + s + "\" ";            
            if (opts.TryGetValue("debug", out s))
                message += " -debug \"" + s + "\" ";            
            if (opts.TryGetValue("errFile", out s))
                message += " -e \"" + s + "\" ";            
            if (opts.TryGetValue("licInfo_name", out s)){
                String p = "";
                String d = "";
                opts.TryGetValue("licInfo_name", out p);
                opts.TryGetValue("licInfo_autodl", out d);
                message += " -licname " + s + " -licpwd " + p + (d.Equals("1") ? " -licweb" : "");
            }
            if (opts.TryGetValue("author", out s))
                message += " -author \"" + s + "\" ";            
            if (opts.TryGetValue("keywords", out s))
                message += " -keywords \"" + s + "\" ";            
            if (opts.TryGetValue("creator", out s))
                message += " -creator \"" + s + "\" ";            
            if (opts.TryGetValue("producer", out s))
                message += " -producer \"" + s + "\" ";  
            if (opts.TryGetValue("overlay", out s))
                message += " -overlay";  
            if (opts.TryGetValue("revoverlay", out s))
                message += " -revoverlay";  
            if (opts.TryGetValue("overlay2", out s))
                message += " -overlay2";  
            if (opts.TryGetValue("revoverlay2", out s))
                message += " -revoverlay2";  
            if (opts.TryGetValue("repeat", out s))
                message += " -repeat";  
            if (opts.TryGetValue("repeatlast", out s))
                message += " -repeatlast";  
            if (opts.TryGetValue("pages", out s))
                message += " -pages \"" + s + "\" ";  
            if (opts.TryGetValue("autorotate", out s))
                message += " -autorotate";  
            if (opts.TryGetValue("flatten", out s))
                message += " -flatten" + (!s.Equals("") ? "sig" : "");  
            if (opts.TryGetValue("pagesize", out s))
                message += " -pagesize " + s;  
            if (opts.TryGetValue("pagenum", out s))
                message += " -pagenum";  
            if (opts.TryGetValue("pagenumalign", out s))
                message += " -pagenumalign " + s;  
            if (opts.TryGetValue("pagetop", out s))
                message += " -pagetop";  
            if (opts.TryGetValue("pagelmargin", out s))
                message += " -pagelmargin " + s;  
            if (opts.TryGetValue("pagebmargin", out s))
                message += " -pagebmargin " + s;  
            if (opts.TryGetValue("pagenumreset", out s))
                message += " -pagenumreset";  
            if (opts.TryGetValue("bmkeep", out s))
                message += " -bmkeep" + (!s.Equals("") ? "nofit" : "");  
            if (opts.TryGetValue("nobm", out s))
                message += " -nobm";  
            if (opts.TryGetValue("bm", out s))
                message += " -bm \"" + s + "\" ";  
            if (opts.TryGetValue("skipfieldrename", out s))
                message += " -skipfieldrename";  
            if (opts.TryGetValue("zoom", out s))
                message += " -zoom " + s;  
            if (opts.TryGetValue("fontsize", out s))
                message += " -fontsize " + s;  
            if (opts.TryGetValue("pagefmt", out s))
                message += " -pagefmt \"" + s + "\" ";  
            if (opts.TryGetValue("pageord", out s))
                message += " -pageord \"" + s + "\" ";  
            if (opts.TryGetValue("center", out s))
                message += " -center";  
            if (opts.TryGetValue("scale", out s))
                message += " -scale \"" + s + "\" ";  
            if (opts.TryGetValue("scalex", out s))
                message += " -scalex \"" + s + "\" ";  
            if (opts.TryGetValue("scaley", out s))
                message += " -scaley \"" + s + "\" ";  
            if (opts.TryGetValue("right", out s))
                message += " -right \"" + s + "\" ";  
            if (opts.TryGetValue("down", out s))
                message += " -down \"" + s + "\" ";  
            if (opts.TryGetValue("strin", out s))
                message += " -strin \"" + s + "\" ";  
            if (opts.TryGetValue("filesep", out s))
                message += " -filesep \"" + s + "\" ";  
            if (opts.TryGetValue("exclude", out s))
                message += " -exclude";  
            if (opts.TryGetValue("subset", out s))
                message += " -subset \"" + s + "\" ";  
            if (opts.TryGetValue("pwdlist", out s))
                message += " -pwdlist \"" + s + "\" ";  
            if (opts.TryGetValue("bmtitle", out s))
                message += " -bmtitle";  
            if (opts.TryGetValue("log", out s))
                message += " -log \"" + s + "\" ";  
            if (opts.TryGetValue("pagenumfont", out s))
                message += " -pagenumfont \"" + s + "\" ";  
            if (opts.TryGetValue("pagenumcolor", out s))
                message += " -pagenumcolor \"" + s + "\" ";  
            if (opts.TryGetValue("pagenumbgcolor", out s))
                message += " -pagenumbgcolor \"" + s + "\" ";  
            if (opts.TryGetValue("autoscale", out s))
                message += " -autoscale \"" + s + "\" ";  
            if (opts.TryGetValue("autoclip", out s))
                message += " -autoclip";  
            if (opts.TryGetValue("removeflds", out s))
                message += " -removeflds";  
            if (opts.TryGetValue("creationdate", out s))
                message += " -creationdate \"" + s + "\" ";  
            if (opts.TryGetValue("moddate", out s))
                message += " -moddate \"" + s + "\" ";  
                Console.WriteLine(message);
            return message;                         
        }
    
        // Send the BUILDPDF command to server to run the commands
        private (byte[], String, String) callTCP(bool isStop = false,
            bool isStatus = false,
            bool retBytes = false,
            String saveFile = ""){
            String errMsg = "";
            Byte[] data = {};
            Byte[] bytes = {};
            String buildResult = "";
            Object host = "";
            bool retPDF = retBytes;
            // String to store the response ASCII representation.
            String responseData = String.Empty;

            if (!isServerRunning()){
                return (bytes, "Server not running", "");
            }

            errMsg = sendTCP();
            if (errMsg.Equals("")){
                if (opts.TryGetValue("autoSend", out String s))
                  retPDF = true; // need to keep open and send files
            }
            if (!saveFile.Equals("")){
                retPDF = true; // need to save the PDF
            }

            try {         
                if (retPDF){
                    data = System.Text.Encoding.ASCII.GetBytes(" -return ");             
                    stream.Write(data, 0, data.Length);      
                }
                data = System.Text.Encoding.ASCII.GetBytes("\nBUILDPDF\n");             
                stream.Write(data, 0, data.Length);      
                if (isStatus){ 
                    do {
                        data = new Byte[1024];
                        // Read the first batch of the TcpServer response bytes.
                        Int32 rawData = stream.Read(data, 0, data.Length);
                        responseData += System.Text.Encoding.ASCII.GetString(data, 0, rawData);
                    }
                    while (stream.DataAvailable);
                } else if (!isStop) {
                    MemoryStream memstream = new MemoryStream();
                    Socket s = client.Client;
                    byte[] buffer = new byte[1024];                    
                    int bytesRead = 0;
                    String retStr = "";
                    String[] retCmds = new String[2];                    
                    
                    while(true) {
                        // Read the first batch of the TcpServer response bytes.
                        buffer = new byte[1024];
                        bytesRead = stream.Read(buffer, 0, buffer.Length);
                        retStr = System.Text.Encoding.ASCII.GetString(buffer, 0, buffer.Length);
                        if (retStr.ToLower().StartsWith("content-length:")){
                          // Receiving the PDF back
                          while ((bytesRead = stream.Read(buffer, 0, buffer.Length)) > 0){
                              memstream.Write(buffer, 0, bytesRead);     
                          }                        
                          bytes = memstream.ToArray();
                          memstream.SetLength(0);
                          break;
                        } else if (retStr.ToLower().StartsWith("build-result:")){
                          retCmds = retStr.Split(new string[] { "\n\n", "\n" }, StringSplitOptions.None);
                          retCmds = retCmds[0].Split(new string[] { ":" }, (Int32) 2, StringSplitOptions.None);
                          buildResult = retCmds[1];
                        } else {
                          retCmds = retStr.Split(new string[] { "\n\n", "\n" }, StringSplitOptions.None);
                          retCmds = retCmds[0].Split(new string[] { ":" }, (Int32) 2, StringSplitOptions.None);                           
                          if (retCmds.Length > 0 && 
                              (retCmds[0].ToLower().Equals("send-file")
                              || retCmds[0].ToLower().Equals("send-md5"))) {

                            if (retCmds[0].ToLower().Equals("send-file")){
                                retCmds[1] = retCmds[1].Trim();
                                FileInfo f = new FileInfo(retCmds[1]);
                                data = System.Text.Encoding.ASCII.GetBytes("Content-Length: " + f.Length + "\n\n");
                                stream.Write(data, 0, data.Length);
                                BinaryReader br = new BinaryReader(new FileStream(retCmds[1], FileMode.Open));

                                while ((bytesRead = br.Read(buffer, 0, buffer.Length)) > 0)
                                    stream.Write(buffer, 0, bytesRead);     
                                stream.Flush();
                            } else {
                                    String message = "";
                                    MD5 md5 = MD5.Create();
                                    FileStream fStream = File.OpenRead(retCmds[1]);
                                    byte[] md5Bytes = md5.ComputeHash(fStream);
                                    String fileHash = ByteArrayToString(md5Bytes);
                                    data = System.Text.Encoding.ASCII.GetBytes(message);             
                                    stream.Write(data, 0, data.Length);
                                    stream.Flush();
                                }
                            }
                            if ((s.Poll(1000, SelectMode.SelectRead) && (s.Available == 0)) || !s.Connected)
                            {
                                break;
                            } 
                        }                     
                    }
                }                
                if (!saveFile.Equals("")){
                    File.WriteAllBytes (saveFile, bytes);
                    Array.Clear(bytes, 0, bytes.Length);
                }
                if (!retBytes){
                    Array.Copy(bytes, new byte[0], 0);
                }
            } catch (SocketException e) {
                errMsg = e.Message;
            } catch (IOException e) {
                errMsg = e.Message;                
            }
            stream.Close();
            client.Close();

            return (bytes, (responseData.Equals("") ? errMsg : responseData), buildResult);
        }

        private static string ByteArrayToString(byte[] ba)
        {
            return BitConverter.ToString(ba).Replace("-","");
        }

        private static byte[] ConvertHexToByteArray(string hexString)
        {
            byte[] byteArray = new byte[hexString.Length / 2];
    
            for (int index = 0; index < byteArray.Length; index++)
            {
                string byteValue = hexString.Substring(index * 2, 2);
                byteArray[index] = byte.Parse(byteValue, NumberStyles.HexNumber, CultureInfo.InvariantCulture);
            }
    
            return byteArray;
        }
         
    }
}
