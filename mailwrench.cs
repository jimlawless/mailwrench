/*
 License: MIT / X11
 Copyright (c) 2011-2022 by James K. Lawless
 jimbo@radiks.net http://www.mailsend-online.com

 Permission is hereby granted, free of charge, to any person
 obtaining a copy of this software and associated documentation
 files (the "Software"), to deal in the Software without
 restriction, including without limitation the rights to use,
 copy, modify, merge, publish, distribute, sublicense, and/or sell
 copies of the Software, and to permit persons to whom the
 Software is furnished to do so, subject to the following
 conditions:

 The above copyright notice and this permission notice shall be
 included in all copies or substantial portions of the Software.

 THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND,
 EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES
 OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND
 NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT
 HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY,
 WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING
 FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR
 OTHER DEALINGS IN THE SOFTWARE.
*/

using System ;
using System.Collections;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Net;
using System.Net.Mail;
using System.Net.Mime;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Forms;

namespace MailWrench
{
   public class MailWrench
   {
         ArrayList ops=new ArrayList();
         ArrayList g_cc=new ArrayList();
         ArrayList g_bcc=new ArrayList();
         ArrayList g_to=new ArrayList();
         Dictionary <string,string> g_ch=new Dictionary<string,string>();
         string g_version;
         string g_domain;
         Dictionary <string, string> g_images=new Dictionary<string, string>();
         bool g_ntlm;
         string g_auth_id;
         string g_auth_pass;
         string g_smtp;
         string g_from;
         string g_filename;
         string g_subject;
         string g_attach;
         string g_attachfile;
         string g_filemissingaction;
         string g_bccfile;  
         string g_ccfile;
         string g_bccaddr;
         string g_msg;
         string g_out;
         int g_port;
         string g_body;
         string g_priority;
         bool g_suppress;
         bool g_filter;
         bool g_abort;
         bool g_haveText;
         bool g_haveImages;
         bool g_haveSMTP;
         bool g_haveTo;
         bool g_ssl;
         bool g_html;
         bool g_haveFrom;
         int g_timeout;
         bool g_receipt;
         bool g_disp;
         TextWriter g_tw;



      
        [STAThread]
	   public static void Main(string[] args)
      {    
         new MailWrench(args);
      }     
      
      public MailWrench(string[] args)
      {
         try 
         {
            MainFlow(args);
         }
         catch(Exception e)
         {
            err(e.ToString()+"\r\n");
         }
      }
      public void MainFlow(string[] args)
      {
         int i;
         g_out="";
         g_version="MailWrench v 2.11 (Free, open source version)";
         
         if(args.Length==0) {
            wr(
               "\r\n"  + g_version + "\r\nby Jim Lawless - jimbo@radiks.net\r\n");
            Syntax();        
            return;
         }
         
         LoadCmdLine(args);
         ProcessCmdLine();
         if(!g_suppress) {
            wr(
               "\r\n" + g_version + "\r\nby Jim Lawless - jimbo@radiks.net\r\n");            
         }
         VerifyFilenames();
         
      // send some mail
         wr("Attempting connection to " + g_smtp + "\r\n"); 
         MailMessage mail = new MailMessage();
         SmtpClient smtpClient = new SmtpClient(g_smtp);
 
         mail.From = new MailAddress(g_from);
         for(i=0;i<g_to.Count;i++) 
         {
            mail.To.Add(g_to[i].ToString());
         }
         for(i=0;i<g_cc.Count;i++) 
         {
            mail.CC.Add(g_cc[i].ToString());
         }
         for(i=0;i<g_bcc.Count;i++) 
         {
            mail.Bcc.Add(g_bcc[i].ToString());
         }
         
         mail.Subject = g_subject;

         if(g_filename!="") 
         {
            string s;
            g_body="";
            if(!File.Exists(g_filename))
            {
               wr("-File " + g_filename + " specified by -file option does not exist.");
               Environment.Exit(1);
            }
            TextReader tr=new StreamReader(g_filename);
            while(true) 
            {            
               s=tr.ReadLine();
               if(s==null)
                  break;
               g_body=g_body+s+"\r\n";
            }
            tr.Close();
            
            if(g_haveImages) {
               Random rGen = new Random();
               foreach(var pair in g_images) 
               {
                  // pair.Key,pair.Value 
                  Attachment att=new Attachment(pair.Value);
                  att.ContentId=rGen.Next(100000, 9999999).ToString();
                  g_body=g_body.Replace(pair.Key,"cid:" + att.ContentId);
                  //wr("Replacing " + pair.Key + " with " + "cid:" + att.ContentId);
                  mail.Attachments.Add(att);
               }                              
            }                        
            mail.Body=g_body;
         }
         else
         if(g_msg!="") 
         {
            mail.Body = g_msg;
         }
         else
         if(g_filter)
         {
            string s;
            g_body="";
            while(true) 
            {            
               s=Console.ReadLine();
               if(s==null)
                  break;
               g_body=g_body+s+"\r\n";
            }
            mail.Body=g_body;            
         }
         if(g_html) 
         {
            mail.IsBodyHtml=true;
         }
         smtpClient.Port=g_port;

         if(g_domain!="") 
         {
            smtpClient.Credentials = new System.Net.NetworkCredential(g_auth_id,
               g_auth_pass, g_domain);
         }
         else
         if(g_auth_id!="") 
         {
            smtpClient.Credentials = new System.Net.NetworkCredential(g_auth_id,g_auth_pass);
         }
         if(g_ssl) 
         {
            smtpClient.EnableSsl = true;
         }
         
         if(g_ntlm) 
         {
            smtpClient.Credentials = CredentialCache.DefaultNetworkCredentials;
         }
         
         if(g_attach!="" ) 
         {
            AddAttachments(mail,g_attach);
         }
         if(g_attachfile!="")
         {
            string s;
            TextReader tr=new StreamReader(g_attachfile);
            while(true) 
            {            
               s=tr.ReadLine();
               if(s==null)
                  break;
               if(s.Length>0) {
                  if(s.Substring(0,1)=="#")
                     continue;
               }
               else {
                  continue;
               }
               AddAttachments(mail,s.Trim());
            }
            tr.Close();
         }
         smtpClient.Timeout=g_timeout * 1000;
         if(g_ccfile!="")
         {
            string s;
            TextReader tr=new StreamReader(g_ccfile);
            while(true) 
            {            
               s=tr.ReadLine();
               if(s==null)
                  break;
               if(s.Length>0) {
                  if(s.Substring(0,1)=="#")
                     continue;
               }
               else {
                  continue;
               }
               mail.CC.Add(s.Trim());
            }
            tr.Close();                                             
         }
         if(g_bccfile!="")
         {
            string s;
            TextReader tr=new StreamReader(g_bccfile);
            while(true) 
            {            
               s=tr.ReadLine();
               if(s==null)
                  break;
               if(s.Length>0) {
                  if(s.Substring(0,1)=="#")
                     continue;
               }
               else {
                  continue;
               }
               mail.Bcc.Add(s.Trim());
            }
            tr.Close();                                             
         }
         
         if(g_priority!="")
         {
            if(g_priority=="low")
               mail.Priority=MailPriority.Low;
            else
            if(g_priority=="normal")
               mail.Priority=MailPriority.Normal;
            else
            if(g_priority=="high")
               mail.Priority=MailPriority.High;
            else {
               err("Unknown mail priority " + g_priority + "\r\n");
               Environment.Exit(1);
            }               
         }

         mail.Headers.Add("X-Mailer", g_version );

         foreach(var pair in g_ch) 
         {
            mail.Headers.Add(pair.Key,pair.Value);
         }
         
         if(g_receipt) 
         {
            mail.Headers.Add("Return-Receipt-To",g_from);
         }
         
         if(g_disp)
         {
            mail.Headers.Add("Disposition-Notification-To",g_from);
         }
         
         smtpClient.Send(mail);
         wr("Send complete!\r\n");
         if(g_out != "") {
            g_tw.Close();
         }
      }

      public void AddAttachments(MailMessage mail,string fname)
      {
         if(!File.Exists(fname))
         {
            if(g_filemissingaction=="warn")
            {
                err("Attachment file: " + fname + " is not present.\r\n");
                return;
            }
            if(g_filemissingaction=="silent")
            {
                return;
            }
         }
         Attachment data = new Attachment(fname, MediaTypeNames.Application.Octet);
         ContentDisposition disposition = data.ContentDisposition;
         disposition.CreationDate = System.IO.File.GetCreationTime(fname);
         disposition.ModificationDate = System.IO.File.GetLastWriteTime(fname);
         disposition.ReadDate = System.IO.File.GetLastAccessTime(fname);
         mail.Attachments.Add(data);
      }
      public void wr(string s) 
      {
         if(g_suppress)
            return;
         if(g_out == "" )
         {
            Console.Write(s);
         }
         else
         {
            g_tw.Write(s);
            g_tw.Flush();
         }
      }
      
      public void err(string s)
      {
         Console.Error.Write(s);
         if(g_out!="") {
             g_tw.Write(s);
             g_tw.Flush();
         }
      }
      public void pause() 
      {
         Console.Write("Press ENTER to continue...");
         Console.ReadLine();
      }
      public void Syntax() 
      {
         wr("Syntax:\r\n\tMailWrench [options]    where options are:\r\n");
         wr(" -h                   Display this screen\r\n");
         wr(" -s subject           Specify subject\r\n");
         wr(" -a filename          Attach filename\r\n");
         wr(" -f filename          Attach files in list\r\n");
         wr(" -filemissingaction   'fatal','warn', or 'silent'\r\n");
         wr("\r\n");
         wr(" -smtp smtp_addr      Specify SMTP server address\r\n");
         wr(" -port number         Specify port ( default is 25 ) \r\n");
         wr("\n");
         wr(" -to to_address       Specify 'To:' e-mail address\r\n");
         wr(" -from from_address   Specify 'From: address\r\n");
         wr(" -file text_file      Name of text file to send\r\n");
         wr(" -msg string          Specify immediate message to send via the command-line\r\n");
         wr("\r\n");
         wr(" -filter              Take text file input from the console stdin device\r\n");
         wr(" -out                 Redirect output to specified filename\r\n");
         wr(" -suppress            Suppress output ( except for error messages )\r\n");
         wr("\r\n");
         pause();

         wr(" -cc to_address       Carbon-copy one recipient; Can be specified multiple times\r\n");
         wr(" -ccf filename        Carbon-copy all recipients listed in filename\r\n");
         wr(" -bcc cc_address      Blind-cc one recipient;  Can be specified multiple times.\r\n");
         wr(" -bccf filename       Blind-carbon-copy to all recipients listed in filename\r\n");

         wr(" -receipt             Request return-receipt\r\n");
         wr(" -disp                Request disposition notification\r\n");
         wr(" -ch header value     Add custom header(s)\r\n");
         wr(" -html                Send file as HTML\r\n");
         wr(" -timeout             Maximum timeout ( in seconds ) \r\n");
         wr(" -img number filename Attach an image as an inline resource in HTML\r\n");
         wr(" -priority code       Set priority. { low, normal, or high }\r\n");
         wr(" -ntlm                Perform NTLM authentication with the mail server\r\n");
         wr("                      Used for SMTP authentication.\r\n");
         wr(" -id loginID          User ID\r\n");
         wr(" -password password   User's password\r\n");
         wr(" -domain domainID     Used if authenticating explicitly against a Windows domain\r\n");
         wr(" -ssl                 Use a Secure Sockets Layer connection (SMTPS)\r\n");
         wr("\r\n");
         wr("Options can be combined in a text file to save command-line space.\r\n");
         wr("See the documentation for detailed information.\r\n");
         wr("\r\n");         
      }
      public void LoadCmdLine(string[] args) 
      {
         int i;
         for(i=0;i<args.Length;i++) {
            if(args[i].Substring(0,1)!="@") {
               ops.Add(args[i]);
            }
            else {
               LoadFile(args[i].Substring(1));
            }
         }
      }

      public void ProcessCmdLine()       
      {
         int i;
         
         // Note that I removed the classic command-line capabilities of MailSend
         // in this software.  Everything must use strict keyword parms.
         g_ntlm=false;
         g_domain="";
         g_auth_id="";
         g_auth_pass="";
         g_smtp="";
         g_from="";
         g_filename="";
         g_filemissingaction="fatal"; // could be "silent" or "warn" as well

         g_subject="";
         g_attach="";
         g_attachfile="";
         g_bccfile="";  // blind carbon-copy file
         g_ccfile="";
         g_bccaddr="";
         g_msg="";
         g_body="";
         g_port=25;
         g_priority="";
         g_suppress=false;
         // No longer supported
         //g_pager=0;
         //g_nomime=0;
           // defined headers
         g_filter=false;
         g_abort=false;
         g_haveText=false;
         g_haveImages=false;
         g_haveSMTP=false;
         g_haveTo=false;
         g_ssl=false;
         g_html=false;
         g_haveFrom=false;
         g_timeout=60;
         g_receipt=false;
         g_disp=false;
         
         for(i=0;i<ops.Count;i++) {
            if( ((ops[i].ToString())== "-s")||( (ops[i].ToString())=="-S") ) {
               g_subject=ops[i+1].ToString().ToString();
               i++;
            }
            else
            if( (ops[i].ToString() == "-a")||(ops[i].ToString()=="-A") ) {
               g_attach=ops[i+1].ToString();
               i++;
            }
            else
            if( (ops[i].ToString() == "-f")||(ops[i].ToString()=="-F") ) {
               g_attachfile=ops[i+1].ToString();
               i++;
            }
            else
            if( ops[i].ToString().ToLower() == "-ch")  {
               g_ch.Add( ops[i+1].ToString(),ops[i+2].ToString());
               i+=2;
            }
            else
            if( ops[i].ToString().ToLower() == "-html" ) {
               g_html=true;
            }
            else
            if( ops[i].ToString().ToLower() == "-ssl" ) {
               g_ssl=true;
            }
            else
            if( ops[i].ToString().ToLower() == "-ccf")  {
               g_ccfile=ops[i+1].ToString();
               //g_haveTo=true;
               i++;
            }
            else
            if( ops[i].ToString().ToLower() == "-domain")  {
               g_domain=ops[i+1].ToString();
               i++;
            }
            else
            if( ops[i].ToString().ToLower() == "-ntlm")  {
               g_ntlm=true;
            }
            else
            if( ops[i].ToString().ToLower() == "-id")  {
               g_auth_id=ops[i+1].ToString();
               i++;
            }
            else
            if( ops[i].ToString().ToLower() == "-password")  {
               g_auth_pass=ops[i+1].ToString();
               i++;
            }
            else
            if( ops[i].ToString().ToLower() == "-filemissingaction")  {
               g_filemissingaction=ops[i+1].ToString().ToLower();
               i++;
               if((g_filemissingaction!="fatal")&&(g_filemissingaction!="warn")&&(g_filemissingaction!="silent"))
               {
                   err("Invalid -filemissingaction: " + g_filemissingaction + ". Must be 'fatal','warn', or 'silent'\r\n");
                   Environment.Exit(1);
               }
            }
            else
            if( ops[i].ToString().ToLower() == "-bccf")  {
               g_bccfile=ops[i+1].ToString();
               i++;
            }
            else
            if( ops[i].ToString().ToLower() == "-cc")  {
               g_cc.Add(ops[i+1].ToString());
               //g_haveTo=true;
               i++;
            }
            else
            if( ops[i].ToString().ToLower() == "-bcc")  {
               g_bcc.Add(ops[i+1].ToString());
               i++;
            }
            else
            if( ops[i].ToString().ToLower() == "-smtp")  {
               g_smtp=ops[i+1].ToString();
               g_haveSMTP=true;
               i++;
            }
            else
            if( ops[i].ToString().ToLower()=="-priority") {
               g_priority=ops[i+1].ToString().ToLower();
               i++;
               if(!(( g_priority == "low") 
               ||( g_priority == "normal")
               ||( g_priority == "high")))               
               {
                  err("Error!  Unknown priority  " + g_priority + " \r\n");
                  Environment.Exit(1);
               }
            }
            else
            if( ops[i].ToString().ToLower() == "-port")  {
               bool parsed;
               parsed=Int32.TryParse(ops[i+1].ToString(),out g_port);
               if(!parsed) {
                  err("Error! Port specified (" + ops[i+1].ToString() + ") is not a number.\r\n");
                  Environment.Exit(1);
               }
               i++;
            }
            else
            if( ops[i].ToString().ToLower() == "-to")  {
               g_to.Add(ops[i+1].ToString());
               g_haveTo=true;
               i++;
            }
            else
            if( ops[i].ToString().ToLower() == "-img")  {
               g_images.Add("${img" + ops[i+1].ToString() + "}",ops[i+2].ToString());
               g_haveImages=true;
               i+=2;
            }
            else
            if( ops[i].ToString().ToLower() == "-from")  {
               g_from=ops[i+1].ToString();
               g_haveFrom=true;
               i++;
            }
            else
            if(ops[i].ToString().ToLower() == "-file")  {
               g_filename=ops[i+1].ToString();
               g_haveText=true;
               i++;
            }
            else
            if( ops[i].ToString().ToLower() == "-msg")  {
               g_msg=ops[i+1].ToString();
               g_haveText=true;
               i++;
            }
            else
            if( ops[i].ToString().ToLower() == "-disp")  {
               g_disp=true;
            }
            else
            if( ops[i].ToString().ToLower() == "-receipt")  {
               g_receipt=true;
            }
            else
            if( ops[i].ToString().ToLower() == "-timeout")  {
               bool parsed;
               parsed=Int32.TryParse(ops[i+1].ToString(),out g_timeout);
               i++;
            }
            else
      // alternate output file
            if( ops[i].ToString().ToLower() == "-out")  {
               g_out=ops[i+1].ToString();
               i++;
               try 
               {
                  g_tw=new StreamWriter(g_out);
               }
               catch(Exception e) 
               {
                  err("Could not open output file " + g_out + "\r\n" + e + "\r\n");
                  Environment.Exit(1);
               }
            }
// now, the cmd-line parms that don't take additional args

//      -filter    Take text file input from stdin
            else
            if( ops[i].ToString().ToLower() == "-filter")  {
               g_filter=true;
               g_haveText=true;
            }
//      -suppress  Suppress output
            else
            if( ops[i].ToString().ToLower() == "-suppress")  {
               g_suppress=true;
            }
            else
            if( ops[i].ToString().ToLower() == "-h")  {
               wr(
                  "\r\n" + g_version + "\r\nby Jim Lawless - jimbo@radiks.net\r\n");
               Syntax();
               Environment.Exit(1);
            }
            else {
               err("\aError! Option " + ops[i] + " is not a valid MailWrench option.\r\n");
               Environment.Exit(1);
            }
         }

         if( !g_haveTo) {
            err("\aError! Missing destination 'To:' e-mail address.\r\n") ;
            g_abort=true;
         }
         if( !g_haveSMTP) {
            err("\aError! Missing SMTP server address.\r\n");
            g_abort=true;
         }
         if( !g_haveFrom) {
            err("\aError! Missing 'From:' address.\r\n");
            g_abort=true;
         }
         if(! g_haveText) {
            err("\aError! Missing -msg, -filter, or input text file.\r\n");
            g_abort=true;
         }
         if( g_filter) {
            if( (g_bccaddr!="")||(g_bccfile!="")) {
               err("\aError! Carbon-copy and -filter options are mutually exclusive.\r\n");
               g_abort=true;
            }
         }
         if(g_abort) {
            Environment.Exit(1);
         }
      }
      public void TestFile(string fn,string desc)
      {
         if(fn == "" )
            return;
         if(File.Exists(fn))
            return;
         if(!g_suppress)
         {         
            err("Cannot find " + desc + ":" + fn + "\r\n");
         }
         Environment.Exit(1);
      }
      public void VerifyFilenames() 
      {
         TestFile(g_filename,"input file");
         TestFile(g_attachfile,"attachment file");
         TestFile(g_attach,"attachment");
         TestFile(g_bccfile,"blind CC file");
         TestFile(g_ccfile,"CC file");         
      }
      public void LoadFile(string fn) 
      {
         int ii;
         string s;
         string tok,mode,c;
         if(!File.Exists(fn))
         {
            wr("File " + fn + " does not exist.");
            Environment.Exit(1);
         }
         TextReader tr=new StreamReader(fn);
         while(true) 
         {            
            s=tr.ReadLine();
            if(s==null)
               break;
            if(s.Length>0) {
               if(s.Substring(0,1)=="#")
                  continue;
            }
            tok="";
            mode="n";
            for(ii=0;ii<s.Length;ii++) 
            {
               c=s.Substring(ii,1);
               // skipwhite
               if( mode == "n") 
               {
                  if((c==" ")||(c=="\t")) 
                  {
                     continue;
                  }
                  else
                  if(c=="\"") 
                  {
                     mode="q";
                     continue;
                  }
                  else 
                  {
                     mode="i";
                     tok=c;
                  }
               }
               else
               if(mode=="q") // wait for another quote
               {
                  if(c=="\"") 
                  {
                     ops.Add(tok);
                     tok="";
                     mode="n";
                  }
                  else
                  {
                     tok = tok + c;
                  }
               }
               else
               if(mode == "i")
               {
                  if((c==" ")||(c=="\t")) 
                  {
                     ops.Add(tok);
                     tok="";
                     mode="n";
                  }
                  else
                  {
                     tok = tok + c;
                  }                  
               }               
            }
            if(tok!="") 
            {
               ops.Add(tok);
            }               
         }
         tr.Close();
      }
   }
}